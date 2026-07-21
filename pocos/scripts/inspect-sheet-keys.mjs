import XLSX from 'xlsx';

const f =
  'c:/Users/paulc/Downloads/26년 6월 중간근태 - 포코스 (7.3-1)_탭명변경파일.xlsx';
const wb = XLSX.readFile(f);

console.log('SheetNames count', wb.SheetNames.length);
for (const name of wb.SheetNames) {
  const codes = [...name].map((ch) => `${ch}(U+${ch.codePointAt(0).toString(16)})`).join(' ');
  console.log('NAME:', JSON.stringify(name));
  console.log('  codes:', codes);
  console.log('  exists in Sheets?', !!wb.Sheets[name]);
  console.log('  === 근무현황(JWL1)?', name === '근무현황(JWL1)');
}

const expected = ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)'];
for (const e of expected) {
  console.log('\nlookup', JSON.stringify(e), '=>', !!wb.Sheets[e]);
  // try find by includes
  const found = wb.SheetNames.find((n) => n.includes('근무현황') && n.includes('JWL1'));
  if (e.endsWith('1)')) console.log('find includes JWL1:', JSON.stringify(found), 'eq?', found === e);
}
