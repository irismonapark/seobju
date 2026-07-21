import XLSX from 'xlsx';

const f =
  'c:/Users/paulc/Downloads/26년 6월 중간근태 - 포코스 (7.3-1)_탭명변경파일.xlsx';
const fOrig =
  'c:/Users/paulc/Downloads/26년 6월 중간근태 - 포코스 (7.3-1).xlsx';

function dump(path, label) {
  const wb = XLSX.readFile(path);
  console.log('\n====', label, '====');
  console.log('SheetNames:', wb.SheetNames);
  console.log('Sheets keys:', Object.keys(wb.Sheets));
  console.log('Workbook sheet names from Workbook:', wb.Workbook?.Sheets?.map((s) => s.name));
  for (const key of Object.keys(wb.Sheets)) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[key], { header: 1, defval: '' });
    console.log(' key', JSON.stringify(key), 'rows', rows.length, 'first title', String(rows[0]?.[2] ?? rows[0]?.[0] ?? '').slice(0, 40));
  }
}

dump(f, '탭명변경본');
dump(fOrig, '원본6월');
