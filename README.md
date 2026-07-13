# Siteplan SEO Checker

A Python script that checks staging site pages against expected SEO values
(H1, Title Tag, Meta Description) from an Excel sheet, and produces a
color-coded results workbook.

## What it does

1. Reads a list of staging URLs plus their **expected** H1 / Title / Meta
   Description from an Excel sheet (`Ops Center` tab by default).
2. **Auto-detects the header row** — scans the first 10 rows of the sheet
   to find the row containing all required column names, instead of
   assuming a fixed row number. If it can't find a match, it prints the
   first 10 rows so you can update the column names accordingly.
3. Requests each URL (with cache-busting and automatic retries — see
   below), parses the live HTML, and extracts the **actual** H1 / Title /
   Meta Description.
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
9. Press **ESC** at any time to stop safely — it finishes the current URL,
   then saves whatever has been processed so far.

## Key features (v2)

- **Auto header-row detection** — no more manually setting `header=1`;
  the script finds the correct row on its own by matching column names.
- **Cache-busting requests** — every request appends a unique query
  parameter so CDNs, reverse proxies, or browser-level caches don't serve
  a stale copy of the page, which previously could cause false `FAIL`
  results on staging sites sitting behind a cache.
- **Automatic retry logic with backoff** — requests are retried up to
  3 times before being marked as an error:
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

## Requirements

- Python 3.8+
- Install dependencies:

  ```
  pip install -r requirements.txt
  ```

## Usage

1. Place your input Excel file in the same folder as the script (or update
   the `INPUT_FILE` path in the config section).
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

5. Output is saved as `<input-filename>_SEO_Result_001.xlsx` (auto-incrementing
   if run multiple times) in the same folder.

## Notes

- SSL certificate verification is disabled (`verify=False`) to allow
  checking staging sites with self-signed or invalid certificates. Only
  run this against sites you trust.
- The `keyboard` library may require elevated/admin permissions on some
  operating systems to detect global key presses.
- A 1-second delay is added between requests (`DELAY_BETWEEN_REQUESTS`) to
  avoid hammering the target server. This is separate from the retry
  backoff delays, which only apply when a request needs to be retried.
- If the header row can't be auto-detected, the script prints the first
  10 rows of the sheet and stops — update `REQUIRED_COLUMNS` to match the
  exact header text shown.

## License

Personal/internal use.
