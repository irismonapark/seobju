import XLSX from 'xlsx';

const src =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/26.0.31_섭주네1차_업체파일청구엑셀로_급여명세서로변환/재훈전달파일_원본/포코스 귀속4월 중간근태(260502신아름).xlsx';

function sheetToRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
}

const wb = XLSX.readFile(src);
const rows = sheetToRows(wb.Sheets['근무현황(JWL1)']);
const hdr = rows.findIndex(
  (r) => r.some((c) => String(c).includes('성명')) && r.some((c) => String(c).includes('구분')),
);
console.log('header row', hdr + 1);
rows[hdr].forEach((c, i) => {
  const t = String(c).trim();
  if (t) console.log(i, JSON.stringify(t));
});

// find all columns containing 급여
const payCols = rows[hdr]
  .map((c, i) => ({ i, t: String(c).trim() }))
  .filter(({ t }) => t.includes('급여'));
console.log('pay columns', payCols);
