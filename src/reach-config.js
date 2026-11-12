'use strict';
// REACH1 (lane NEWSHIP): this listing's Pro tools for the toolbar row and the Explorer entries (src/reach.js).
module.exports = {
  "prefix": "excelToMarkdown",
  "triesKey": "excelToMarkdown.proTrial",
  "slot": ".reach1-line",
  "tools": [
    {
      "id": "batch",
      "title": "Convert Workbooks to Markdown Files",
      "command": "excelToMarkdown.workbookToMarkdown"
    },
    {
      "id": "xlsx",
      "title": "Save Markdown Tables as an Excel File",
      "command": "excelToMarkdown.tablesToXlsx"
    },
    {
      "id": "formats",
      "title": "Keep Excel Formats, Links and Merged Cells",
      "command": "excelToMarkdown.keepFormats"
    }
  ]
};
