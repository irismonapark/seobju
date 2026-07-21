import XLSX from 'xlsx';

const f =
  'c:/Users/paulc/Downloads/26년 6월 중간근태 - 포코스 (7.3-1)_탭명변경파일.xlsx';
const f4 =
  'c:/Users/paulc/Downloads/포코스 귀속4월 중간근태(260502신아름).xlsx';

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

function inspect(path, label) {
  const wb = XLSX.readFile(path);
  console.log('\n====', label, '====');
  console.log('tabs:', wb.SheetNames);
  const need = ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)'];
  for (const n of need) {
    console.log(n, wb.SheetNames.includes(n) ? 'OK' : 'MISSING', '| exact match candidates:',
      wb.SheetNames.filter((s) => s.includes('근무현황')).map((s) => JSON.stringify(s)));
  }

  for (const sn of need) {
    const sheet = wb.Sheets[sn];
    if (!sheet) continue;
    const rows = sheetToRows(sheet);
    let hdr = -1;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const cells = (rows[i] || []).map((c) => String(c).trim());
      if (cells.some((c) => c.includes('성명')) && cells.some((c) => c.includes('구분'))) {
        hdr = i;
        break;
      }
    }
    console.log('\n--', sn, 'hdr', hdr, 'rowCount', rows.length);
    if (hdr < 0) continue;
    const header = rows[hdr];
    const idxs = {};
    for (const key of ['NO', '소속', '성명', '구분', '근무시간', '급여', '합계']) {
      idxs[key] = header.findIndex((c) => String(c ?? '').includes(key));
    }
    console.log('col indexes', idxs);
    console.log('header full length', header.length);
    console.log(
      'header cells with content:',
      header
        .map((c, i) => [i, String(c)])
        .filter(([, c]) => c.trim())
        .map(([i, c]) => `${i}:${c}`),
    );

    // mimic parse briefly
    const nameCol = idxs['성명'];
    const catCol = idxs['구분'];
    const noCol = idxs['NO'];
    const payCol = idxs['급여'];
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
      if (!byNo.has(currentNo)) byNo.set(currentNo, { name: blockName, cats: [] });
      byNo.get(currentNo).cats.push(cat);
      if (r < hdr + 8) {
        console.log(' sample', {
          r,
          no: row[noCol],
          name,
          cat,
          pay: row[payCol],
          근무시간: row[idxs['근무시간']],
        });
      }
    }
    console.log('employees', byNo.size);
  }
}

inspect(f, '6월탭변경');
inspect(f4, '4월원본');
