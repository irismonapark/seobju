import XLSX from 'xlsx';

const invoicePath =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/seobju/pocos/scripts/test-invoice-out.xlsx';

function parsePayAmount(value) {
  const text = String(value ?? '')
    .replace(/,/g, '')
    .replace(/["']/g, '')
    .trim();
  if (!text || text === '-') return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function parseNumber(value) {
  const num = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(num) ? num : 0;
}

function findInvoiceHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    if (!row) continue;
    const hasSeq = row.some((cell) => String(cell ?? '').trim() === '순번');
    const hasName = row.some((cell) => String(cell ?? '').trim() === '성명');
    if (hasSeq && hasName) return i;
  }
  return 2;
}

function isValidEmployeeName(name) {
  const trimmed = name.trim();
  if (!trimmed || trimmed === '-') return false;
  const numeric = trimmed.replace(/,/g, '').replace(/\s/g, '');
  if (/^\d+(\.\d+)?$/.test(numeric)) return false;
  return true;
}

function parseEmployees(rows) {
  const headerRowIndex = findInvoiceHeaderRow(rows);
  const headerRow = rows[headerRowIndex] ?? [];
  const seqCol = headerRow.findIndex((cell) => String(cell ?? '').trim() === '순번');
  const nameCol = headerRow.findIndex((cell) => String(cell ?? '').trim() === '성명');
  const eventCol = headerRow.findIndex((cell) => String(cell ?? '').includes('경조'));
  const cols = {
    seqCol,
    nameCol,
    baseA: nameCol + 3,
    otA: nameCol + 5,
    nightA: nameCol + 7,
    holA: nameCol + 9,
    hotA: nameCol + 11,
    eventCol: eventCol >= 0 ? eventCol : nameCol + 22,
  };

  const employees = [];
  for (let r = headerRowIndex + 2; r < rows.length; r++) {
    const row = rows[r];
    if (!row) break;
    const combined = [cols.seqCol, cols.nameCol, cols.baseA]
      .map((c) => String(row[c] ?? '').trim())
      .join('')
      .replace(/\s/g, '');
    if (combined.includes('합계') || combined.includes('공급가액') || combined.includes('청구총액')) break;

    const seq = parseNumber(row[cols.seqCol]);
    const name = String(row[cols.nameCol] ?? '').trim();
    if (!isValidEmployeeName(name) || seq <= 0) continue;

    employees.push({
      name,
      기본급: parsePayAmount(row[cols.baseA]),
      연장: parsePayAmount(row[cols.otA]),
    });
  }
  return employees;
}

const wb = XLSX.readFile(invoicePath);
const rows = XLSX.utils.sheet_to_json(wb.Sheets['청구내역서'], {
  header: 1,
  defval: '',
  raw: false,
});
const employees = parseEmployees(rows);
console.log('count', employees.length);
console.log('first5', employees.slice(0, 5).map((e) => e.name));
console.log('sample filenames', employees.slice(0, 3).map((e) => `${e.name}_4월_급여명세서.pdf`));
