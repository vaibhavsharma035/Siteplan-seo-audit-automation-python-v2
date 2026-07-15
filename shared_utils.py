"""
Shared helper functions used by both streamlit_app.py (SEO Checker) and
pages/1_Redirect_Checker.py (Redirect Checker). Keeping these in one
place means a fix here (e.g. a retry-logic tweak) applies to both tools
without needing to be made twice.
"""

import io
import re
import time
import random
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import pandas as pd
import requests

RETRY_BACKOFF_BASE = 4

GOOGLE_SHEET_ID_PATTERN = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")


def add_cache_buster(url):
    """
    Appends a unique query param so CDNs/proxies/browser-level caches
    don't serve a stale copy of the page. Preserves any existing query
    string on the URL.
    """
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query))
    query["_cb"] = str(int(time.time() * 1000)) + str(random.randint(1000, 9999))
    new_query = urlencode(query)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def fetch_with_retry(session, url, timeout, max_retries, log_fn):
    """
    Fetches a URL with cache-busting and retry logic.

    - Adds a unique query string on every attempt to bypass CDN /
      reverse-proxy caching.
    - Retries on 429 (rate limit), respecting Retry-After if present.
    - Retries on 403 / 404, since some pages briefly return these during
      redirects or WAF challenges even though the page is actually live.
    - Retries on timeout / SSL / connection errors with backoff.

    Returns (response, None) on success, or (None, exception) on
    exhausted retries.
    """
    last_exception = None

    for attempt in range(1, max_retries + 1):
        busted_url = add_cache_buster(url)

        try:
            response = session.get(
                busted_url, timeout=timeout, verify=False, allow_redirects=True
            )

            if response.status_code == 429:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after and retry_after.isdigit() else RETRY_BACKOFF_BASE * attempt
                log_fn(f"  429 received. Waiting {wait_time}s before retry {attempt}/{max_retries}...")
                time.sleep(wait_time)
                continue

            if response.status_code in (403, 404) and attempt < max_retries:
                wait_time = RETRY_BACKOFF_BASE * attempt
                log_fn(f"  {response.status_code} received. Retrying in {wait_time}s ({attempt}/{max_retries})...")
                time.sleep(wait_time)
                continue

            return response, None

        except (requests.exceptions.Timeout,
                requests.exceptions.SSLError,
                requests.exceptions.ConnectionError) as e:
            last_exception = e
            wait_time = RETRY_BACKOFF_BASE * attempt
            log_fn(f"  Request error ({type(e).__name__}). Retrying in {wait_time}s ({attempt}/{max_retries})...")
            time.sleep(wait_time)

    return None, last_exception


def extract_google_sheet_id(url):
    match = GOOGLE_SHEET_ID_PATTERN.search(url)
    return match.group(1) if match else None


def fetch_google_sheet_as_xlsx(sheet_id, timeout=20):
    """
    Downloads a publicly link-shared Google Sheet as .xlsx bytes, using
    Google's built-in export endpoint. Only works for sheets set to
    "Anyone with the link can view" (or more open) — no login/API key
    involved. Returns (bytes, error_message). error_message is None on
    success.
    """
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    try:
        response = requests.get(export_url, timeout=timeout, allow_redirects=True)
    except requests.exceptions.RequestException as e:
        return None, f"Could not reach Google Sheets: {e}"

    if response.status_code != 200:
        return None, f"Google Sheets returned status {response.status_code}. Check the link and sharing settings."

    # A valid .xlsx is a zip archive — starts with 'PK'. If Google instead
    # returned an HTML login/permission page, this check catches it.
    if not response.content[:2] == b"PK":
        return None, (
            "This doesn't look like a valid spreadsheet export. The sheet is likely "
            "not shared as \"Anyone with the link can view\" — check its sharing settings."
        )

    return response.content, None


def guess_header_row(file_bytes, sheet_name, expected_values, max_scan_rows=10):
    """
    Best-effort auto-detect of which row holds the column headers, by
    looking for a row that contains most of `expected_values`. Falls
    back to row 0 if nothing matches well (or if expected_values is
    empty — no scoring signal to go on, e.g. the redirect checker doesn't
    have fixed column-name guesses).
    """
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    best_row, best_score = 0, -1

    for row_idx in range(len(raw)):
        row_values = raw.iloc[row_idx].astype(str).str.strip().tolist()
        score = sum(1 for v in expected_values if v in row_values)
        if score > best_score:
            best_row, best_score = row_idx, score

    return best_row


def best_match_index(columns, guess):
    """Return the index of `guess` in `columns` if present, else 0."""
    if guess in columns:
        return columns.index(guess)
    return 0


def new_requests_session():
    """A requests.Session with the standard headers used across both tools."""
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache"
    })
    return session


def build_data_source_ui(st, key_prefix):
    """
    Renders the "Upload Excel file / Google Sheets link" chooser shared
    by both tools. Returns (file_bytes, source_name) — both None if
    nothing has been loaded yet. `key_prefix` keeps widget keys unique
    when this is called from more than one page in the same app.
    """
    source_mode = st.radio(
        "Sheet source",
        options=["Upload Excel file (.xlsx)", "Google Sheets link"],
        horizontal=True,
        key=f"{key_prefix}_source_mode",
    )

    file_bytes = None
    source_name = None

    if source_mode == "Upload Excel file (.xlsx)":
        uploaded_file = st.file_uploader(
            "Upload your Excel file (.xlsx)", type=["xlsx"], key=f"{key_prefix}_uploader"
        )
        if uploaded_file is not None:
            file_bytes = uploaded_file.getvalue()
            source_name = Path(uploaded_file.name).stem

    else:
        st.caption(
            "Works only for sheets shared as \"Anyone with the link can view\" — "
            "no Google login is used or required."
        )
        sheet_url = st.text_input("Paste the Google Sheets URL", key=f"{key_prefix}_gsheet_url")
        if sheet_url:
            sheet_id = extract_google_sheet_id(sheet_url)
            if not sheet_id:
                st.error("That doesn't look like a Google Sheets URL. Expected a link containing '/spreadsheets/d/<id>/'.")
            else:
                with st.spinner("Fetching sheet from Google..."):
                    fetched_bytes, fetch_error = fetch_google_sheet_as_xlsx(sheet_id)
                if fetch_error:
                    st.error(fetch_error)
                else:
                    file_bytes = fetched_bytes
                    source_name = f"GoogleSheet_{sheet_id[:8]}"
                    st.success("Sheet loaded successfully.")

    return file_bytes, source_name
