import ExcelJS from 'exceljs';

function hasBorder(cell) {
  const b = cell.border || {};
  return !!(b.top?.style || b.bottom?.style || b.left?.style || b.right?.style);
}

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('public/payslip-template.xlsx');
const s = wb.getWorksheet('D(프리랜서_)');
console.log('rowCount', s.rowCount, 'lastRow', s.lastRow?.number);

let outside = 0;
for (let row = 1; row <= s.rowCount; row++) {
  const r = s.getRow(row);
  if (!r) continue;
  r.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const inTable = row >= 3 && row <= 14 && colNumber >= 2 && colNumber <= 6;
    if (!inTable && hasBorder(cell)) {
      outside++;
      console.log('outside', row, colNumber, cell.border);
    }
  });
}
console.log('outside count', outside);
