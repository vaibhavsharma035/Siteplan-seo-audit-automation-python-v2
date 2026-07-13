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
# DEFAULT CONFIGURATION (adjustable in the sidebar)
# ==========================================================
REQUIRED_COLUMNS = [
    "42Works - Staging Site Link",
    "New H1",
    "Live Title Tag",
    "Live Meta Description"
]

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


def find_header_row(file_bytes, sheet_name, required_columns, max_scan_rows=10):
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    for row_idx in range(len(raw)):
        row_values = raw.iloc[row_idx].astype(str).str.strip().tolist()
        if all(col in row_values for col in required_columns):
            return row_idx
    return None


# ==========================================================
# UI — SIDEBAR SETTINGS
# ==========================================================
st.title("🔍 Siteplan SEO Checker")
st.caption("Upload an Ops Center sheet, run the H1 / Title / Meta check against staging URLs, and download a color-coded results workbook.")

with st.sidebar:
    st.header("Settings")
    sheet_name = st.text_input("Sheet name", value="Ops Center")
    request_timeout = st.number_input("Request timeout (seconds)", min_value=5, max_value=60, value=DEFAULT_TIMEOUT)
    delay_between_requests = st.number_input("Delay between requests (seconds)", min_value=0.0, max_value=10.0, value=float(DEFAULT_DELAY), step=0.5)
    max_retries = st.number_input("Max retries per URL", min_value=1, max_value=5, value=DEFAULT_MAX_RETRIES)
    st.markdown("---")
    st.caption("Required columns in your sheet:")
    for col in REQUIRED_COLUMNS:
        st.caption(f"• {col}")
    st.markdown("---")
    st.caption("To stop a run in progress, use the ⏹ stop control Streamlit shows at the top of the page while the script is executing.")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

run_clicked = st.button("▶ Run SEO Check", type="primary", disabled=uploaded_file is None)

# ==========================================================
# MAIN EXECUTION
# ==========================================================
if run_clicked and uploaded_file is not None:

    file_bytes = uploaded_file.getvalue()
    source_name = Path(uploaded_file.name).stem

    log_box = st.empty()
    log_lines = []

    def log(msg):
        log_lines.append(msg)
        # Keep the log box readable — show the last ~40 lines
        log_box.code("\n".join(log_lines[-40:]))

    log(f"Input File : {uploaded_file.name}")
    log(f"Sheet      : {sheet_name}")

    header_row = find_header_row(file_bytes, sheet_name, REQUIRED_COLUMNS)

    if header_row is None:
        st.error("Could NOT auto-detect the header row. Here are the first 10 rows of the sheet:")
        preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, nrows=10)
        st.dataframe(preview)
        st.stop()

    log(f"Detected header row at Excel row index: {header_row}")

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    df.columns = df.columns.astype(str).str.strip()

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        st.error(f"Missing required column(s): {missing}")
        st.write("Found columns:", df.columns.tolist())
        st.stop()

    df = df[REQUIRED_COLUMNS]
    df.columns = ["URL", "Expected_H1", "Expected_Title", "Expected_Meta"]
    df = df[df["URL"].notna()].reset_index(drop=True)

    total_urls = len(df)
    log(f"Total URLs to check : {total_urls}")

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
                apply_cell_style(ws.cell(row=row_num, column=col_map[col_name]), ws.cell(row=row_num, column=col_map[col_name]).value)
        if status_column in col_map:
            apply_status_style(ws.cell(row=row_num, column=col_map[status_column]), ws.cell(row=row_num, column=col_map[status_column]).value)

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
