import ExcelJS from 'exceljs';
import XLSX from 'xlsx';

const tpl =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/seobju/pocos/public/invoice-template.xlsx';
const out =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/seobju/pocos/scripts/test-border-out.xlsx';

const DATA_START = 5;
const LAST = 184;

function hasBorder(cell) {
  const b = cell.border;
  return !!(b?.top?.style || b?.bottom?.style || b?.left?.style || b?.right?.style);
}

function trimRowsBelow(sheet, lastRow) {
  while (sheet.rowCount > lastRow) {
    sheet.spliceRows(sheet.rowCount, 1);
  }
  if (sheet._rows && sheet._rows.length > lastRow) {
    sheet._rows.length = lastRow;
  }
}

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(tpl);
const sheet = workbook.getWorksheet('청구내역서');

// simulate prepareInvoiceDataArea fix
while (sheet.rowCount >= DATA_START) {
  sheet.spliceRows(DATA_START, 1);
}
console.log('after prepare rowCount', sheet.rowCount);

// add fake content to 184
for (let i = 0; i < 180; i++) {
  sheet.getRow(DATA_START + i).getCell(7).value = 'test';
}
trimRowsBelow(sheet, LAST);
await workbook.xlsx.writeFile(out);

const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.readFile(out);
const s2 = wb2.getWorksheet('청구내역서');
console.log('output rowCount', s2.rowCount, 'dims', s2.dimensions);

for (let r = LAST + 1; r <= LAST + 15; r++) {
  let count = 0;
  for (let c = 4; c <= 35; c++) {
    if (hasBorder(s2.getRow(r).getCell(c))) count++;
  }
  if (count) console.log('R' + r, 'borders', count);
}
console.log('done, no borders above means success');
