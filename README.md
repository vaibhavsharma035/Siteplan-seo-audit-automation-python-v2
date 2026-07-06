# Siteplan SEO Checker

A Python script that checks staging site pages against expected SEO values
(H1, Title Tag, Meta Description) from an Excel sheet, and produces a
color-coded results workbook.

## What it does

1. Reads a list of staging URLs plus their **expected** H1 / Title / Meta
   Description from an Excel sheet (`Ops Center` tab by default).
2. Requests each URL, parses the live HTML, and extracts the **actual**
   H1 / Title / Meta Description.
3. Compares actual vs. expected and marks each as `PASS`, `FAIL`, or a
   specific warning (e.g. `MULTIPLE H1`, `TITLE MISSING`).
4. Flags HTTP status issues (4xx/5xx) and request errors (timeout, SSL,
   connection failures) separately.
5. Saves a new Excel file with color-coded cells:
   - 🟥 Red — hard `FAIL`
   - 🟧 Orange — warnings (missing/multiple H1, no expected value, etc.)
   - 🟪 Purple — HTTP error status codes
   - ⬜ Grey — request errors (timeout/SSL/connection)
6. Prints a running summary and a final pass/fail/warning/error count.
7. Press **ESC** at any time to stop safely — it finishes the current URL,
   then saves whatever has been processed so far.

## Requirements

- Python 3.8+
- Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```

## Usage

1. Place your input Excel file in the same folder as the script (or update
   the `INPUT_FILE` path in the config section).
2. Update the configuration constants at the top of the script if needed:
   ```python
   INPUT_FILE  = "Brien1.xlsx"
   SHEET_NAME  = "Ops Center"
   ```
3. Your input sheet must contain these exact column headers (row 2, since
   the script reads with `header=1`):
   - `42Works - Staging Site Link`
   - `New H1`
   - `Live Title Tag`
   - `Live Meta Description`
4. Run the script:
   ```bash
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
  avoid hammering the target server.

## License

Personal/internal use.
