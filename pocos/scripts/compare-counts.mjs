import XLSX from 'xlsx';
import ExcelJS from 'exceljs';

const src =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/26.0.31_섭주네1차_업체파일청구엑셀로_급여명세서로변환/재훈전달파일_원본/포코스 귀속4월 중간근태(260502신아름).xlsx';
const tpl =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/seobju/pocos/public/invoice-template.xlsx';
const out =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/seobju/pocos/scripts/test-invoice-out.xlsx';

const INVOICE_COL = { SEQ: 6, NAME: 7, BASE_A: 10, OT_A: 12, NIGHT_A: 14, HOL_A: 16, HOT_A: 18 };
const DATA_START_ROW = 5;
const INVOICE_XLSX_COL = { SEQ: 5, NAME: 6, BASE_A: 9, OT_A: 11, NIGHT_A: 13, HOL_A: 15, HOT_A: 17, EVENT: 28 };

function parsePayAmount(value) {
  const text = String(value ?? '')
    .replace(/,/g, '')
    .replace(/["']/g, '')
    .trim();
  if (!text || text === '-') return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function sheetToRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
}

function stripSheetFormulas(sheet) {
  sheet.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value;
      if (val === null || val === undefined || typeof val !== 'object') return;
      if ('sharedFormula' in val || 'formula' in val) {
        cell.value = val.result ?? null;
      }
    });
  });
}

function unmergeRowsFrom(sheet, fromRow) {
  for (const range of [...(sheet.model.merges ?? [])]) {
    const match = range.match(/^[A-Z]+(\d+):/);
    if (match && Number(match[1]) >= fromRow) sheet.unMergeCells(range);
  }
}

function buildEmployees(wb) {
  const sheets = ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)'];
  const all = [];
  for (const sn of sheets) {
    const rows = sheetToRows(wb.Sheets[sn]);
    const hdr = rows.findIndex(
      (r) => r.some((c) => String(c).includes('성명')) && r.some((c) => String(c).includes('구분')),
    );
    const header = rows[hdr];
    const nameCol = header.findIndex((c) => String(c).includes('성명'));
    const catCol = header.findIndex((c) => String(c).includes('구분'));
    const payCol = header.findIndex((c) => String(c).includes('급여'));
    const noCol = header.findIndex((c) => String(c).includes('NO'));
    const byNo = new Map();
    let currentNo = null;
    let blockName = '';
    let valid = false;
    for (let r = hdr + 1; r < rows.length; r++) {
      const row = rows[r];
      const name = String(row[nameCol] ?? '').trim();
      const cat = String(row[catCol] ?? '').trim();
      const noText = String(row[noCol] ?? '').trim();
      if ((noText + name).replace(/\s/g, '').includes('합계')) break;
      const no = parsePayAmount(row[noCol]);
      if (no > 0) {
        currentNo = no;
        blockName = name;
        valid = name.length > 0;
      }
      if (!valid || !currentNo || !['정상', '연장', '야간', '휴일', '휴연'].includes(cat)) continue;
      if (!byNo.has(currentNo)) byNo.set(currentNo, { name: blockName, 정상: 0, 연장: 0, 야간: 0, 휴일: 0, 휴연: 0 });
      const rec = byNo.get(currentNo);
      if (blockName) rec.name = blockName;
      const pay = parsePayAmount(row[payCol]);
      const key =
        cat === '정상' ? '정상' : cat === '연장' ? '연장' : cat === '야간' ? '야간' : cat === '휴일' ? '휴일' : '휴연';
      rec[key] += pay;
    }
    all.push(...byNo.values());
  }
  return all;
}

function parseInvoice(rows) {
  const dataStart = 4;
  const included = [];
  for (let r = dataStart; r < rows.length; r++) {
    const row = rows[r];
    let name = String(row[INVOICE_XLSX_COL.NAME] ?? '').trim();
    if (!name) {
      const seqCell = String(row[INVOICE_XLSX_COL.SEQ] ?? '').trim();
      if (seqCell && Number.isNaN(Number(seqCell.replace(/,/g, '')))) name = seqCell;
    }
    if (!name) continue;
    included.push(name);
  }
  return included;
}

const wb = XLSX.readFile(src);
const employees = buildEmployees(wb);
console.log('employees', employees.length);

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(tpl);
const sheet = workbook.getWorksheet('청구내역서');
stripSheetFormulas(sheet);
unmergeRowsFrom(sheet, DATA_START_ROW);
sheet.spliceRows(DATA_START_ROW, sheet.rowCount - DATA_START_ROW + 1);

employees.forEach((emp, i) => {
  const row = sheet.getRow(DATA_START_ROW + i);
  row.getCell(INVOICE_COL.SEQ).value = i + 1;
  row.getCell(INVOICE_COL.NAME).value = emp.name;
  row.getCell(INVOICE_COL.BASE_A).value = emp.정상;
  row.getCell(INVOICE_COL.OT_A).value = emp.연장;
  row.getCell(INVOICE_COL.NIGHT_A).value = emp.야간;
  row.getCell(INVOICE_COL.HOL_A).value = emp.휴일;
  row.getCell(INVOICE_COL.HOT_A).value = emp.휴연;
});

await workbook.xlsx.writeFile(out);

const rows = sheetToRows(XLSX.readFile(out).Sheets['청구내역서']);
const included = parseInvoice(rows);
console.log('parsed', included.length);
const missing = employees.filter((e) => !included.includes(e.name));
console.log('missing', missing.map((m) => m.name));
