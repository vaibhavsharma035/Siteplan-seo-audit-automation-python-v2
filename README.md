# Siteplan SEO & Redirect Checker

QA automation tools used before client website handovers. Checks staging pages
against expected SEO values (H1, Title Tag, Meta Description) and validates
that staging/live URLs redirect to the correct destination — producing
color-coded Excel reports either way.

Available in two forms:
- **A hosted web app** (Streamlit) — no installation needed, works in any browser
- **A local Python script** — for offline/CLI use

🌐 **Live Application:** https://siteplan-seo-audit-automation-python-v2-g9d45dhukap5rprutwx6qv.streamlit.app/

---

## Repo structure

```
├── streamlit_app.py              ← SEO Checker (web app, home page)
├── shared_utils.py                ← helpers shared by both web app pages
├── pages/
│   └── 1_Redirect_Checker.py     ← Redirect Checker (web app, second page)
├── Siteplan_Test_Final.py        ← SEO Checker (local CLI script)
├── requirements.txt
└── README.md
```

---

## Web app (recommended — no install needed)

Open the live link above in any browser. Two tools are available from the
sidebar navigation: **SEO Checker** and **Redirect Checker**.

### SEO Checker

1. Choose a data source: **upload an `.xlsx` file**, or **paste a Google
   Sheets link** (works for sheets shared as "Anyone with the link can view"
   — no Google login required).
2. Pick the sheet and confirm/adjust the auto-detected header row.
3. Choose which fields to check — **H1**, **Title**, and/or **Meta
   Description** — any combination. Unchecked fields are skipped entirely,
   including from the output report.
4. Map your sheet's columns to each selected field using the dropdowns (no
   fixed column names required — works across different clients' sheets).
5. Click **Run SEO Check**. A live progress bar and per-URL status panel
   show Status Code, H1 Result, Title Result, and Meta Result as each page
   is checked.
6. Download the finished, color-coded `.xlsx` report.

### Redirect Checker

1. Same data source options as above (upload or Google Sheets link).
2. Map the **Live/Staging URL** column and the **Expected redirect-to URL**
   column.
3. Click **Run Redirect Check**. Each URL is fetched and its final
   destination (after following redirects) is compared against the expected
   URL.
4. Static asset URLs (`.css`, `.js`, images, fonts, `.pdf`, etc.) and blank
   URLs are automatically skipped and marked accordingly.
5. Download a row-color-coded `.xlsx` report (green = pass, red = fail,
   orange = error, grey = skipped).

### Shared behavior across both tools

- **Cache-busting** — every request appends a unique query parameter so
  CDNs/proxies/browser caches can't serve a stale copy of the page.
- **Automatic retries with backoff** — up to 3 attempts per URL on
  `429`, `403`, `404`, timeout, SSL, or connection errors before marking a
  URL as an error.
- **Nothing is saved to disk on the server** — uploaded files and results
  exist only in your browser session; everything is generated in memory and
  offered as a direct download.
- **Runs can't be paused/resumed** — closing or refreshing the tab stops a
  run in progress, and nothing processed so far is retained. For very large
  sheets, consider checking them in smaller batches.

---

## Local CLI script (`Siteplan_Test_Final.py`)

For offline use or environments without browser access. Covers the SEO
check only (H1 / Title / Meta) — the same core logic as the web app's SEO
Checker page, run from the command line against a fixed local file.

### What it does

1. Reads a list of staging URLs plus their expected H1 / Title / Meta
   Description from an Excel sheet (`Ops Center` tab by default).
2. Auto-detects the header row — scans the first 10 rows of the sheet to
   find the row containing all required column names, instead of assuming
   a fixed row number. If it can't find a match, it prints the first 10
   rows so you can update the column names accordingly.
3. Requests each URL (with cache-busting and automatic retries — see
   below), parses the live HTML, and extracts the actual H1 / Title / Meta
   Description.
4. Compares actual vs. expected and marks each as `PASS`, `FAIL`, or a
   specific warning (e.g. `MULTIPLE H1`, `TITLE MISSING`).
5. Flags HTTP status issues (4xx/5xx) and request errors (timeout, SSL,
   connection failures) separately.
6. Saves a new Excel file with color-coded cells:
   - 🟥 Red — hard `FAIL`
   - 🟧 Orange — warnings (missing/multiple H1, no expected value, etc.)
   - 🟪 Purple — HTTP error status codes
   - ⬜ Grey — request errors (timeout/SSL/connection)
7. Applies consistent formatting to the output workbook: dark navy header
   row, frozen header row, auto-sized columns, sequential output file
   numbering.
8. Prints a running summary and a final pass/fail/warning/error count.
9. Press ESC at any time to stop safely — it finishes the current URL,
   then saves whatever has been processed so far.

### Key features

- **Auto header-row detection** — no manually setting `header=1`; the
  script finds the correct row on its own by matching column names.
- **Cache-busting requests** — every request appends a unique query
  parameter so CDNs, reverse proxies, or browser-level caches don't serve
  a stale copy of the page.
- **Automatic retry logic with backoff** — requests are retried up to 3
  times before being marked as an error:
  - `429 Too Many Requests` — waits based on the `Retry-After` header if
    provided, otherwise backs off progressively.
  - `404 Not Found` — retried, since some staging pages briefly 404
    during redirects or WAF/bot-protection challenges even though the
    page is actually live.
  - Timeouts, SSL errors, and connection errors — retried with
    increasing backoff between attempts.
- **No-cache request headers** — `Cache-Control: no-cache, no-store,
  must-revalidate` and `Pragma: no-cache` are sent on every request in
  addition to the cache-busting query parameter.

### Usage

1. Place your input Excel file in the same folder as the script (or
   update the `INPUT_FILE` path in the config section).
2. Update the configuration constants at the top of the script if needed:

   ```python
   INPUT_FILE  = "SiteplanSheetName.xlsx"
   SHEET_NAME  = "Ops Center"
   ```

3. Your input sheet must contain these exact column headers (the script
   will automatically find the row they're on, within the first 10 rows):
   - `42Works - Staging Site Link`
   - `New H1`
   - `Live Title Tag`
   - `Live Meta Description`
4. Run the script:

   ```
   python Siteplan_Test_Final.py
   ```

5. Output is saved as `<input-filename>_SEO_Result_001.xlsx`
   (auto-incrementing if run multiple times) in the same folder.

### Notes

- SSL certificate verification is disabled (`verify=False`) to allow
  checking staging sites with self-signed or invalid certificates. Only
  run this against sites you trust.
- The `keyboard` library may require elevated/admin permissions on some
  operating systems to detect global key presses. It's only used by this
  local script — the web app doesn't need it.
- A 1-second delay is added between requests (`DELAY_BETWEEN_REQUESTS`) to
  avoid hammering the target server. This is separate from the retry
  backoff delays, which only apply when a request needs to be retried.
- If the header row can't be auto-detected, the script prints the first
  10 rows of the sheet and stops — update `REQUIRED_COLUMNS` to match the
  exact header text shown.

---

## Requirements (for local use — either script or testing the web app locally)

- Python 3.8+
- Install dependencies:

  ```
  pip install -r requirements.txt
  ```

  `requirements.txt` covers both the local script and the web app —
  comments in the file mark which packages belong to which.

To run the web app locally instead of using the hosted version:

```
streamlit run streamlit_app.py
```

---

## License

Personal/internal use.
