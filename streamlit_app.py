import io
import time
import random
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(page_title="Siteplan SEO Checker", layout="wide")

# ==========================================================
# DEFAULTS
# ==========================================================
# Used only as a best-guess when auto-selecting dropdown defaults below —
# not a hard requirement anymore, since columns are now chosen in the UI.
DEFAULT_COLUMN_GUESSES = {
    "url": "42Works - Staging Site Link",
    "h1": "New H1",
    "title": "Live Title Tag",
    "meta": "Live Meta Description",
}

DEFAULT_TIMEOUT = 15
DEFAULT_DELAY = 1
DEFAULT_MAX_RETRIES = 3
RETRY_BACKOFF_BASE = 4

# ==========================================================
# COLORS
# ==========================================================
FILL_FAIL = PatternFill(fill_type="solid", fgColor="FFC7CE")
FONT_FAIL = Font(color="9C0006", bold=True)

FILL_WARN = PatternFill(fill_type="solid", fgColor="FFEB9C")
FONT_WARN = Font(color="9C5700", bold=False)

FILL_HTTP_ERR = PatternFill(fill_type="solid", fgColor="E2CFFE")
FONT_HTTP_ERR = Font(color="4B0082", bold=True)

FILL_ERROR = PatternFill(fill_type="solid", fgColor="D9D9D9")
FONT_ERROR = Font(color="595959", bold=False)

FAIL_VALUES = {"FAIL"}
WARN_VALUES = {"MULTIPLE H1", "H1 MISSING", "TITLE MISSING", "META MISSING",
               "NO EXPECTED H1", "NO EXPECTED TITLE", "NO EXPECTED META"}
HTTP_ERR_CODES = {"404", "403", "500", "502", "503", "301", "302"}
ERROR_VALUES = {"ERROR", "META ERROR"}


def apply_cell_style(cell, value):
    val = str(value).strip().upper()
    if val in FAIL_VALUES:
        cell.fill = FILL_FAIL
        cell.font = FONT_FAIL
    elif val in WARN_VALUES:
        cell.fill = FILL_WARN
        cell.font = FONT_WARN
    elif val in ERROR_VALUES:
        cell.fill = FILL_ERROR
        cell.font = FONT_ERROR


def apply_status_style(cell, value):
    val = str(value).strip()
    if val == "ERROR":
        cell.fill = FILL_ERROR
        cell.font = FONT_ERROR
    elif val in HTTP_ERR_CODES or (val.isdigit() and int(val) >= 400):
        cell.fill = FILL_HTTP_ERR
        cell.font = FONT_HTTP_ERR


def clean_text(text):
    if text is None or pd.isna(text):
        return ""
    return " ".join(str(text).replace("\n", " ").replace("\r", " ").split()).strip()


def add_cache_buster(url):
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query))
    query["_cb"] = str(int(time.time() * 1000)) + str(random.randint(1000, 9999))
    new_query = urlencode(query)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def fetch_with_retry(session, url, timeout, max_retries, log_fn):
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


def guess_header_row(file_bytes, sheet_name, expected_values, max_scan_rows=10):
    """
    Best-effort auto-detect of which row holds the column headers, by
    looking for a row that contains most of the DEFAULT_COLUMN_GUESSES
    values. Falls back to row 0 if nothing matches well.
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


# ==========================================================
# UI — HEADER
# ==========================================================
st.title("🔍 Siteplan SEO Checker")
st.caption("Upload a sheet, map your columns, run the H1 / Title / Meta check against staging URLs, and download a color-coded results workbook.")

with st.sidebar:
    st.header("Run settings")
    request_timeout = st.number_input("Request timeout (seconds)", min_value=5, max_value=60, value=DEFAULT_TIMEOUT)
    delay_between_requests = st.number_input("Delay between requests (seconds)", min_value=0.0, max_value=10.0, value=float(DEFAULT_DELAY), step=0.5)
    max_retries = st.number_input("Max retries per URL", min_value=1, max_value=5, value=DEFAULT_MAX_RETRIES)
    st.markdown("---")
    st.caption("To stop a run in progress, use the ⏹ stop control Streamlit shows at the top of the page while the script is executing.")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    source_name = Path(uploaded_file.name).stem

    # ── Sheet selection ───────────────────────────────────
    try:
        workbook_preview = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = workbook_preview.sheet_names
    except Exception as e:
        st.error(f"Could not read this file as an Excel workbook: {e}")
        st.stop()

    default_sheet_index = sheet_names.index("Ops Center") if "Ops Center" in sheet_names else 0
    sheet_name = st.selectbox("Sheet", options=sheet_names, index=default_sheet_index)

    # ── Header row selection ──────────────────────────────
    auto_row = guess_header_row(file_bytes, sheet_name, list(DEFAULT_COLUMN_GUESSES.values()))
    header_row = st.number_input(
        "Header row (0 = first row of the sheet)",
        min_value=0, max_value=20, value=auto_row,
        help="The row number where your column titles actually appear. Auto-detected as a starting guess — adjust if it looks wrong."
    )

    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
        df_raw.columns = df_raw.columns.astype(str).str.strip()
    except Exception as e:
        st.error(f"Could not read the sheet with that header row: {e}")
        st.stop()

    available_columns = df_raw.columns.tolist()

    with st.expander("Preview of detected columns and first rows", expanded=False):
        st.write("Columns found:", available_columns)
        st.dataframe(df_raw.head(5))

    # ── Column mapping ─────────────────────────────────────
    st.subheader("Map your columns")
    st.caption("Tell the checker which column in your sheet holds each piece of data.")

    map_col1, map_col2 = st.columns(2)

    with map_col1:
        url_column = st.selectbox(
            "Staging URL column",
            options=available_columns,
            index=best_match_index(available_columns, DEFAULT_COLUMN_GUESSES["url"])
        )
        h1_column = st.selectbox(
            "Expected H1 column",
            options=available_columns,
            index=best_match_index(available_columns, DEFAULT_COLUMN_GUESSES["h1"])
        )

    with map_col2:
        title_column = st.selectbox(
            "Expected Title column",
            options=available_columns,
            index=best_match_index(available_columns, DEFAULT_COLUMN_GUESSES["title"])
        )
        meta_column = st.selectbox(
            "Expected Meta Description column",
            options=available_columns,
            index=best_match_index(available_columns, DEFAULT_COLUMN_GUESSES["meta"])
        )

    selected_columns = [url_column, h1_column, title_column, meta_column]
    mapping_has_duplicates = len(set(selected_columns)) != len(selected_columns)

    if mapping_has_duplicates:
        st.warning("You've mapped the same column to more than one field — double check your selections.")

    run_clicked = st.button("▶ Run SEO Check", type="primary", disabled=mapping_has_duplicates)

else:
    run_clicked = False
    st.info("Upload an Excel file to get started.")

# ==========================================================
# MAIN EXECUTION
# ==========================================================
if uploaded_file is not None and run_clicked:

    df = df_raw[[url_column, h1_column, title_column, meta_column]].copy()
    df.columns = ["URL", "Expected_H1", "Expected_Title", "Expected_Meta"]
    df = df[df["URL"].notna()].reset_index(drop=True)

    total_urls = len(df)

    log_box = st.empty()
    log_lines = [f"Input File : {uploaded_file.name}", f"Sheet      : {sheet_name}", f"Total URLs : {total_urls}"]

    def log(msg):
        log_lines.append(msg)
        log_box.code("\n".join(log_lines[-40:]))

    log_box.code("\n".join(log_lines))

    df["Status_Code"] = ""
    df["Actual_H1"] = ""
    df["H1_Count"] = 0
    df["H1_Result"] = ""
    df["Actual_Title"] = ""
    df["Title_Result"] = ""
    df["Actual_Meta"] = ""
    df["Meta_Result"] = ""
    df["Checked_At"] = ""

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache"
    })

    pass_count = fail_count = warn_count = error_count = 0

    progress_bar = st.progress(0.0)
    status_placeholder = st.empty()

    for index, row in df.iterrows():

        url = row["URL"]
        expected_h1 = clean_text(row["Expected_H1"])
        expected_title = clean_text(row["Expected_Title"])
        expected_meta = clean_text(row["Expected_Meta"])

        status_placeholder.write(f"Processing {index + 1} / {total_urls}: {url}")

        try:
            response, fetch_error = fetch_with_retry(session, url, request_timeout, max_retries, log)

            if fetch_error is not None:
                raise fetch_error
            if response is None:
                raise Exception("Max retries exceeded")

            status_code = response.status_code
            soup = BeautifulSoup(response.text, "html.parser")

            h1_tags = soup.find_all("h1")
            h1_count = len(h1_tags)
            actual_h1 = clean_text(h1_tags[0].get_text()) if h1_count > 0 else ""

            actual_title = clean_text(soup.title.get_text()) if soup.title else ""

            meta_tag = soup.find("meta", attrs={"name": "description"})
            if not meta_tag:
                meta_tag = soup.find("meta", attrs={"property": "og:description"})
            actual_meta = clean_text(meta_tag.get("content")) if meta_tag and meta_tag.get("content") else ""

        except requests.exceptions.Timeout:
            log(f"TIMEOUT : {url}")
            status_code, actual_h1, actual_title, actual_meta, h1_count = "ERROR", "", "", "", 0

        except requests.exceptions.SSLError:
            log(f"SSL ERROR : {url}")
            status_code, actual_h1, actual_title, actual_meta, h1_count = "ERROR", "", "", "", 0

        except Exception as e:
            log(f"ERROR : {url} | {e}")
            status_code, actual_h1, actual_title, actual_meta, h1_count = "ERROR", "", "", "", 0

        df.at[index, "Status_Code"] = str(status_code)
        df.at[index, "Actual_H1"] = actual_h1
        df.at[index, "H1_Count"] = h1_count
        df.at[index, "Actual_Title"] = actual_title
        df.at[index, "Actual_Meta"] = actual_meta
        df.at[index, "Checked_At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if status_code == "ERROR":
            h1_result = "ERROR"
        elif expected_h1 == "":
            h1_result = "NO EXPECTED H1"
        elif actual_h1 == "":
            h1_result = "H1 MISSING"
        elif h1_count > 1:
            h1_result = "MULTIPLE H1"
        elif expected_h1.lower() == actual_h1.lower():
            h1_result = "PASS"
        else:
            h1_result = "FAIL"

        if status_code == "ERROR":
            title_result = "ERROR"
        elif expected_title == "":
            title_result = "NO EXPECTED TITLE"
        elif actual_title == "":
            title_result = "TITLE MISSING"
        elif expected_title.lower() == actual_title.lower():
            title_result = "PASS"
        else:
            title_result = "FAIL"

        if status_code == "ERROR":
            meta_result = "ERROR"
        elif expected_meta == "":
            meta_result = "NO EXPECTED META"
        elif actual_meta == "":
            meta_result = "META MISSING"
        elif expected_meta.lower() == actual_meta.lower():
            meta_result = "PASS"
        else:
            meta_result = "FAIL"

        df.at[index, "H1_Result"] = h1_result
        df.at[index, "Title_Result"] = title_result
        df.at[index, "Meta_Result"] = meta_result

        row_results = [h1_result, title_result, meta_result]
        if "FAIL" in row_results:
            fail_count += 1
        elif status_code == "ERROR":
            error_count += 1
        elif any(r in WARN_VALUES for r in row_results):
            warn_count += 1
        else:
            pass_count += 1

        progress_bar.progress((index + 1) / total_urls)
        time.sleep(delay_between_requests)

    status_placeholder.write("Done processing all URLs.")

    # ==========================================================
    # BUILD FORMATTED OUTPUT WORKBOOK (in memory)
    # ==========================================================
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)

    wb = load_workbook(excel_buffer)
    ws = wb.active

    col_map = {}
    for cell in ws[1]:
        if cell.value:
            col_map[str(cell.value).strip()] = cell.column

    result_columns = ["H1_Result", "Title_Result", "Meta_Result"]
    status_column = "Status_Code"

    HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F3864")
    HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
    HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    HEADER_BORDER = Border(bottom=Side(style="medium", color="FFFFFF"))

    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border = HEADER_BORDER

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 35

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_length + 2, 3), 16)

    for row_num in range(2, ws.max_row + 1):
        for col_name in result_columns:
            if col_name in col_map:
                cell = ws.cell(row=row_num, column=col_map[col_name])
                apply_cell_style(cell, cell.value)
        if status_column in col_map:
            cell = ws.cell(row=row_num, column=col_map[status_column])
            apply_status_style(cell, cell.value)

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    # ==========================================================
    # SUMMARY + DOWNLOAD
    # ==========================================================
    st.success("SEO check completed.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Passed", pass_count)
    c2.metric("Failed", fail_count)
    c3.metric("Warnings", warn_count)
    c4.metric("Errors", error_count)

    output_filename = f"{source_name}_SEO_Result_001.xlsx"

    st.download_button(
        label="⬇ Download Results (.xlsx)",
        data=output_buffer,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    with st.expander("View results table"):
        st.dataframe(df)
