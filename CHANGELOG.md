# Changelog

## 0.1.2 (2026-11-13)

- First Marketplace release, under the name Excel to Markdown: Paste & CSV.
- Free: Paste as Markdown Table, CSV/TSV to Markdown and back, Format Markdown Table and the spreadsheet preview with sheet and range, with no limit.
- Zovo Pro (USD 19 once, every Zovo extension), 3 free tries each: Convert Workbooks to Markdown Files, Save Markdown Tables as an Excel File, and Keep Excel Formats, Links and Merged Cells (now also a command).
- The spreadsheet preview has a small Pro tools row under its toolbar with the free tries left for each tool and Get Zovo Pro. The × button hides it until VS Code restarts, and it is not shown with Zovo Pro or when `excelToMarkdown.showProTools` is off.
- Right-click a workbook, a folder or a Markdown file in the Explorer for the "Zovo Pro:" entries.
- The buy choice for Zovo Pro is one short list: card, GitHub Sponsors, PayPal or a license key.
- Rating ask 2.3: it can ask for a rating at a file that comes 30 minutes or more after the first one, and never after an error.
- Nothing opens or pops up when VS Code starts or when you open a file.

## 0.1.0

- Build that was not published.
- Paste as Markdown Table: cells copied from Excel, Google Sheets, Numbers or LibreOffice become a Markdown table at the cursor, with numbers right-aligned, pipes and backslashes escaped and line breaks in cells as `<br>`. No default key; normal paste is not changed.
- Convert CSV/TSV to Markdown Table, Convert Markdown Table to CSV or TSV, and Format Markdown Table for the table under the cursor. Columns line up with Chinese, Japanese, Korean and emoji text.
- Open Spreadsheet as Markdown Table: preview `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.ods`, `.csv` and `.tsv` files, pick a sheet and a range, then copy or insert the table as Markdown, HTML, AsciiDoc, Jira, CSV or TSV.
- Zovo Pro tools, 3 free tries each: convert whole workbooks or folders to `.md` files, save the Markdown tables of a document as an `.xlsx` file, and keep Excel number formats, dates, links and merged cells.
- A Pro tool with nothing to convert uses no free try: a workbook whose sheets are all empty gets no `.md` file, and Keep Excel formats on an empty sheet turns itself off again.
- Workbooks are read in a worker thread with time and memory limits and a zip bomb check. No network, no telemetry.
- Walkthrough with a sample workbook, a rating ask after a few days of use (never after an error), and links to related Zovo extensions.
