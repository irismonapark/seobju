import ExcelJS from 'exceljs';
import path from 'path';

const src = path.join(
  'c:/Users/paulc/Downloads',
  '26년 6월 중간근태 - 포코스 (7.3-1).xlsx',
);
const dst = path.join(
  'c:/Users/paulc/Downloads',
  '26년 6월 중간근태 - 포코스 (7.3-1)_수정본.xlsx',
);

const RENAMES = [
  ['근무현황(에스씨1)', '근무현황(JWL1)'],
  ['근무현황(에스씨2)', '근무현황(JWL2)'],
  ['근무현황(에스씨3)', '근무현황(JWL3)'],
];

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(src);

console.log('변경 전 탭:');
workbook.worksheets.forEach((ws, i) => console.log(`  ${i + 1}. ${ws.name}`));

const changed = [];
for (const [from, to] of RENAMES) {
  const sheet = workbook.getWorksheet(from);
  if (!sheet) {
    console.log(`없음: ${from}`);
    continue;
  }
  sheet.name = to;
  changed.push(`${from} → ${to}`);
}

await workbook.xlsx.writeFile(dst);

console.log('\n변경한 항목:');
changed.forEach((c) => console.log('  -', c));

console.log('\n변경 후 탭:');
workbook.worksheets.forEach((ws, i) => console.log(`  ${i + 1}. ${ws.name}`));

// verify with xlsx that data is readable
import XLSX from 'xlsx';
const check = XLSX.readFile(dst);
console.log('\n검증 SheetNames:', check.SheetNames);
console.log('검증 Sheets keys:', Object.keys(check.Sheets));
for (const name of ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)']) {
  const rows = XLSX.utils.sheet_to_json(check.Sheets[name], {
    header: 1,
    defval: '',
    raw: false,
  });
  console.log(`${name}: 행수=${rows.length}, 제목=${String(rows[0]?.[2] ?? '').slice(0, 40)}`);
}

console.log('\n저장 위치:', dst);
