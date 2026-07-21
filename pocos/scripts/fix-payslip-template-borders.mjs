import ExcelJS from 'exceljs';

const TABLE = { topRow: 3, bottomRow: 14, leftCol: 2, rightCol: 6 };
const THIN = { style: 'thin' };
const MEDIUM = { style: 'medium' };

function isTable(row, col) {
  return (
    row >= TABLE.topRow &&
    row <= TABLE.bottomRow &&
    col >= TABLE.leftCol &&
    col <= TABLE.rightCol
  );
}

function tableBorder(row, col) {
  return {
    top: row === TABLE.topRow ? MEDIUM : THIN,
    bottom: row === TABLE.bottomRow ? MEDIUM : THIN,
    left: col === TABLE.leftCol ? MEDIUM : THIN,
    right: col === TABLE.rightCol ? MEDIUM : THIN,
  };
}

function removeCellBorder(cell) {
  if (!cell.border) return;
  const nextStyle = { ...cell.style };
  delete nextStyle.border;
  cell.style = nextStyle;
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

while (sheet.rowCount > 15) {
  sheet.spliceRows(sheet.rowCount, 1);
}
if (Array.isArray(sheet._rows) && sheet._rows.length > 15) {
  sheet._rows.length = 15;
}

const maxRow = sheet.rowCount;
const maxCol = TABLE.rightCol + 2;
for (let row = 1; row <= maxRow; row++) {
  for (let col = 1; col <= maxCol; col++) {
    const cell = sheet.getCell(row, col);
    if (!isTable(row, col)) {
      removeCellBorder(cell);
    }
  }
}

for (let row = TABLE.topRow; row <= TABLE.bottomRow; row++) {
  for (let col = TABLE.leftCol; col <= TABLE.rightCol; col++) {
    sheet.getCell(row, col).border = tableBorder(row, col);
  }
}

sheet.views = [{ showGridLines: false, zoomScale: 100 }];
sheet.pageSetup = {
  ...sheet.pageSetup,
  showGridLines: false,
  showRowColHeaders: false,
  printArea: 'B1:F15',
  fitToPage: true,
  fitToWidth: 1,
  fitToHeight: 1,
  orientation: 'portrait',
};

await wb.xlsx.writeFile('public/payslip-template.xlsx');

function hasBorder(cell) {
  const b = cell.border || {};
  return !!(b.top?.style || b.bottom?.style || b.left?.style || b.right?.style);
}

const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.readFile('public/payslip-template.xlsx');
const s = wb2.getWorksheet('D(프리랜서_)');
let inside = 0;
let outside = 0;
for (let row = 1; row <= s.rowCount; row++) {
  for (let col = 1; col <= TABLE.rightCol + 2; col++) {
    if (hasBorder(s.getCell(row, col))) {
      if (isTable(row, col)) inside++;
      else outside++;
    }
  }
}
console.log('inside', inside, 'outside', outside, 'rowCount', s.rowCount);
