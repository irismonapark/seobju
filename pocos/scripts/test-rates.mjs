import XLSX from 'xlsx';
import { readFileSync } from 'fs';

const src =
  'c:/Users/paulc/OneDrive/바탕 화면/Iris/3_아웃소싱시스템개발_프리랜서/26.0.31_섭주네1차_업체파일청구엑셀로_급여명세서로변환/재훈전달파일_원본/포코스 귀속4월 중간근태(260502신아름).xlsx';

function parsePayAmount(value) {
  if (value === null || value === undefined || value === '') return 0;
  const text = String(value).replace(/,/g, '').replace(/["']/g, '').trim();
  if (!text || text === '-') return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

const wb = XLSX.readFile(src);
const rates = {};

for (const sn of ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)']) {
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: false });
  const hdr = rows[3];
  const totalCol = hdr.findIndex((h) => String(h).replace(/[\s"']/g, '').includes('합계'));

  let currentNo = null;
  let isValid = false;
  let hourlyRate = 0;

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    const noText = String(row[0] ?? '').trim();
    const name = String(row[2] ?? '').trim();
    const cat = String(row[3] ?? '').trim();
    if (noText.replace(/\s/g, '').includes('합계')) break;

    const no = Number(String(row[0]).replace(/,/g, '')) || 0;
    if (no > 0) {
      currentNo = no;
      isValid = name.length > 0;
      hourlyRate = 0;
      if (cat === '정상' && isValid) {
        hourlyRate = parsePayAmount(row[totalCol]);
      }
    }

    if (!isValid || !currentNo || cat !== '정상' || !name) continue;
    rates[hourlyRate] = (rates[hourlyRate] || 0) + 1;
    if ([10900, 12500].includes(hourlyRate)) {
      console.log(sn, name.replace(/\n/g, ' '), hourlyRate);
    }
  }
}

console.log('rates', rates);
