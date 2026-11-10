# Changelog

## 0.1.0

- First release.
- Paste as Markdown Table: cells copied from Excel, Google Sheets, Numbers or LibreOffice become a Markdown table at the cursor, with numbers right-aligned, pipes and backslashes escaped and line breaks in cells as `<br>`. No default key; normal paste is not changed.
- Convert CSV/TSV to Markdown Table, Convert Markdown Table to CSV or TSV, and Format Markdown Table for the table under the cursor. Columns line up with Chinese, Japanese, Korean and emoji text.
- Open Spreadsheet as Markdown Table: preview `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.ods`, `.csv` and `.tsv` files, pick a sheet and a range, then copy or insert the table as Markdown, HTML, AsciiDoc, Jira, CSV or TSV.
- Zovo Pro tools, 3 free tries each: convert whole workbooks or folders to `.md` files, save the Markdown tables of a document as an `.xlsx` file, and keep Excel number formats, dates, links and merged cells.
- A Pro tool with nothing to convert uses no free try: a workbook whose sheets are all empty gets no `.md` file, and Keep Excel formats on an empty sheet turns itself off again.
- Workbooks are read in a worker thread with time and memory limits and a zip bomb check. No network, no telemetry.
- Walkthrough with a sample workbook, a rating ask after a few days of use (never after an error), and links to related Zovo extensions.
