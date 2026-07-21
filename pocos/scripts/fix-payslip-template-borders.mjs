import ExcelJS from 'exceljs';

const TABLE = { topRow: 3, bottomRow: 14, leftCol: 2, rightCol: 6 };
const THIN = { style: 'thin' };
const MEDIUM = { style: 'medium' };

function tableBorder(row, col) {
  return {
    top: row === TABLE.topRow ? MEDIUM : THIN,
    bottom: row === TABLE.bottomRow ? MEDIUM : THIN,
    left: col === TABLE.leftCol ? MEDIUM : THIN,
    right: col === TABLE.rightCol ? MEDIUM : THIN,
  };
}

function isTable(row, col) {
  return (
    row >= TABLE.topRow &&
    row <= TABLE.bottomRow &&
    col >= TABLE.leftCol &&
    col <= TABLE.rightCol
  );
}

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('public/payslip-template.xlsx');
const sheet = wb.getWorksheet('D(프리랜서_)');

try {
  sheet.unMergeCells('C13:D13');
} catch {}
try {
  sheet.unMergeCells('C14:D14');
} catch {}
sheet.mergeCells('C14:D14');

// remove borders outside table only — do NOT set empty border objects
sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (!isTable(rowNumber, colNumber)) {
      delete cell.border;
    }
  });
});

for (let row = 1; row < TABLE.topRow; row++) {
  for (let col = 1; col <= 10; col++) {
    delete sheet.getCell(row, col).border;
  }
}
for (let row = TABLE.bottomRow + 1; row <= 25; row++) {
  for (let col = 1; col <= 10; col++) {
    delete sheet.getCell(row, col).border;
  }
}

for (let row = TABLE.topRow; row <= TABLE.bottomRow; row++) {
  for (let col = TABLE.leftCol; col <= TABLE.rightCol; col++) {
    sheet.getCell(row, col).border = tableBorder(row, col);
  }
}

sheet.views = [{ showGridLines: false, zoomScale: 100 }];

while (sheet.rowCount > 16) {
  sheet.spliceRows(17, 1);
}

await wb.xlsx.writeFile('public/payslip-template.xlsx');
console.log('template updated, rowCount', sheet.rowCount);
