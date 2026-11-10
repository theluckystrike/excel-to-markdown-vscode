# Paste cells as a Markdown table

1. Select cells in Excel, Google Sheets, Numbers or LibreOffice Calc and copy them.
2. In a Markdown file, put the cursor on an empty line.
3. Right-click and pick Paste as Markdown Table, or run it from the Command Palette.

The first row becomes the header. Columns that hold only numbers are right-aligned. Pipes, backslashes and `<` `>` in cells are escaped, and a line break inside a cell becomes `<br>`.

Normal paste (Cmd+V or Ctrl+V) is not changed. If you paste tables often, give the command a key of your own, for example Cmd+Alt+V on macOS or Ctrl+Alt+V on Windows and Linux, with Add a keybinding.

The settings `excelToMarkdown.alignment`, `headerRow`, `maxColumnWidth`, `compact`, `escapeHtml`, `escapeMarkdown` and `lineBreaks` change how tables are written.
