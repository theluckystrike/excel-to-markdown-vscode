'use strict';
// Builds samples/sample.xlsx (the walkthrough workbook) with SheetJS from the sibling table-viewer-vscode tree.
// samples/try-it.md is plain text kept in git.
const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', '..', 'table-viewer-vscode', 'node_modules', 'xlsx'));

const prices = XLSX.utils.aoa_to_sheet([
  ['Product', 'Price', 'In stock', 'Margin', 'Launched', 'Page'],
  ['Green tea, 100 g', 6.5, 1240, 0.42, 45292, 'tea'],
  ['Oolong | first flush', 14.25, 310, 0.38, 45366, 'oolong'],
  ['Matcha tin', 28, 85, 0.51, 45444, 'matcha'],
  ['Gift box', 1249.99, 12, 0.3, 45627, 'gift box'],
]);
for (let r = 2; r <= 5; r++) { prices['B' + r].z = '#,##0.00 "USD"'; prices['C' + r].z = '#,##0'; prices['D' + r].z = '0%'; prices['E' + r].z = 'd mmm yyyy'; }
['tea', 'oolong', 'matcha', 'gift-box'].forEach((s, i) => { prices['F' + (i + 2)].l = { Target: `https://example.com/shop/${s}` }; });
prices['!cols'] = [{ wch: 22 }, { wch: 14 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 10 }];

const quarterly = XLSX.utils.aoa_to_sheet([
  ['Region', 'Quarter', 'Orders', 'Revenue'],
  ['North', 'Q1', 410, 18250],
  [null, 'Q2', 455, 20110],
  ['South', 'Q1', 380, 16900],
  [null, 'Q2', 402, 17760],
]);
quarterly['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } }];
for (let r = 2; r <= 5; r++) quarterly['D' + r].z = '#,##0';

const team = XLSX.utils.aoa_to_sheet([
  ['Name', 'Office', 'Role', 'Note'],
  ['Ana Souza', 'São Paulo', 'Lead', 'Owns the price list'],
  ['山田 太郎', '東京', 'Buyer', 'Writes <notes> in cells'],
  ['Lee Min-ji', '서울', 'Design 🎨', 'Line one\nline two'],
]);

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, prices, 'Prices');
XLSX.utils.book_append_sheet(wb, quarterly, 'Quarterly');
XLSX.utils.book_append_sheet(wb, team, 'Team');
const out = path.join(__dirname, '..', 'samples', 'sample.xlsx');
fs.writeFileSync(out, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true }));
console.log('wrote', out, fs.statSync(out).size, 'bytes');
