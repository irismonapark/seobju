import XLSX from 'xlsx';

const f =
  'c:/Users/paulc/Downloads/26년 6월 중간근태 - 포코스 (7.3-1)_근무현황탭명변경.xlsx';
const fOrig =
  'c:/Users/paulc/Downloads/26년 6월 중간근태 - 포코스 (7.3-1).xlsx';
const fOk =
  'c:/Users/paulc/Downloads/포코스 귀속4월 중간근태(260502신아름).xlsx';

function analyze(path, label) {
  console.log('\n====================');
  console.log(label);
  console.log('파일:', path.split('/').pop());
  try {
    const wb = XLSX.readFile(path);
    console.log('1) 탭 이름 목록:', wb.SheetNames.join(' | '));
    console.log('2) 실제 데이터 상자 키:', Object.keys(wb.Sheets).join(' | ') || '(비어있음!)');
    console.log('3) 탭 개수:', wb.SheetNames.length, '/ 데이터 상자 개수:', Object.keys(wb.Sheets).length);

    const need = ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)'];
    for (const n of need) {
      const inNames = wb.SheetNames.includes(n);
      const hasData = !!wb.Sheets[n];
      console.log(`   - ${n}: 이름표=${inNames ? 'O' : 'X'}, 데이터=${hasData ? 'O' : 'X'}`);
    }

    // if Sheets empty but names exist - corrupted rename
    if (wb.SheetNames.length > 0 && Object.keys(wb.Sheets).length === 0) {
      console.log('결과: 이름표만 있고 안쪽 데이터가 사라진 깨진 파일');
      return;
    }

    // try parse employees like converter
    let total = 0;
    for (const sn of need) {
      const sheet = wb.Sheets[sn];
      if (!sheet) continue;
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
      let hdr = -1;
      for (let i = 0; i < Math.min(15, rows.length); i++) {
        const cells = (rows[i] || []).map((c) => String(c).trim());
        if (cells.some((c) => c.includes('성명')) && cells.some((c) => c.includes('구분'))) {
          hdr = i;
          break;
        }
      }
      if (hdr < 0) {
        console.log(`   ${sn}: 헤더(성명/구분) 없음`);
        continue;
      }
      const header = rows[hdr];
      const nameCol = header.findIndex((c) => String(c).includes('성명'));
      const catCol = header.findIndex((c) => String(c).includes('구분'));
      const noCol = header.findIndex((c) => String(c).includes('NO'));
      const names = new Set();
      for (let r = hdr + 1; r < rows.length; r++) {
        const row = rows[r];
        const no = Number(String(row[noCol] ?? '').replace(/,/g, ''));
        const name = String(row[nameCol] ?? '').trim();
        const cat = String(row[catCol] ?? '').trim();
        if (name && no > 0 && cat === '정상') names.add(name);
      }
      console.log(`   ${sn}: 헤더행=${hdr}, 정상직원약=${names.size}명, 예=${[...names].slice(0, 3).join(',')}`);
      total += names.size;
    }
    console.log('4) 합계 직원(대략):', total);
  } catch (e) {
    console.log('읽기 오류:', e.message);
  }
}

analyze(f, '지금 올린 파일(근무현황탭명변경)');
analyze(fOrig, '원본 6월(에스씨)');
analyze(fOk, '잘되는 4월 파일');
