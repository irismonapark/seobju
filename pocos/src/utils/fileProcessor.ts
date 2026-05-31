import * as XLSX from 'xlsx';

export const WORK_SHEETS = ['근무현황(JWL1)', '근무현황(JWL2)', '근무현황(JWL3)'] as const;
export const INVOICE_SHEET = '청구내역서';

export class FileValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'FileValidationError';
  }
}

export function validateXlsxFile(file: File): void {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    throw new FileValidationError(
      '파일 형식이 올바르지 않습니다. Excel 파일(.xlsx)을 선택해주세요.',
    );
  }
}

export async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  validateXlsxFile(file);

  try {
    const buffer = await file.arrayBuffer();
    return XLSX.read(buffer, { type: 'array', cellDates: true });
  } catch {
    throw new FileValidationError(
      '데이터가 너무 많거나 손상되었습니다. 파일을 다시 확인해주세요.',
    );
  }
}

export function validateWorkRecordSheets(workbook: XLSX.WorkBook): void {
  for (const sheetName of WORK_SHEETS) {
    if (!workbook.SheetNames.includes(sheetName)) {
      throw new FileValidationError(`[${sheetName}] 시트를 찾을 수 없습니다.`);
    }
  }
}

export function validateInvoiceSheet(workbook: XLSX.WorkBook): void {
  if (!workbook.SheetNames.includes(INVOICE_SHEET)) {
    throw new FileValidationError(`[${INVOICE_SHEET}] 시트를 찾을 수 없습니다.`);
  }
}

export function sheetToRows(sheet: XLSX.WorkSheet): unknown[][] {
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    raw: false,
  }) as unknown[][];
}

export function findHeaderRow(rows: unknown[][], requiredHeaders: string[]): number {
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const row = rows[i];
    if (!row) continue;

    const normalized = row.map((cell) => String(cell ?? '').trim());
    const hasAll = requiredHeaders.every((header) =>
      normalized.some((cell) => cell.includes(header)),
    );
    if (hasAll) return i;
  }

  throw new FileValidationError(
    `필수 컬럼(${requiredHeaders.join(', ')})을 찾을 수 없습니다.`,
  );
}

export function getColumnIndex(row: unknown[], header: string): number {
  const index = row.findIndex((cell) => String(cell ?? '').trim().includes(header));
  if (index === -1) {
    throw new FileValidationError(`[${header}] 컬럼을 찾을 수 없습니다.`);
  }
  return index;
}

export function findTotalColumn(row: unknown[]): number {
  const index = row.findIndex((cell) => {
    const text = String(cell ?? '').replace(/[\s"']/g, '');
    return text.includes('합계') || text === '합계';
  });
  if (index === -1) {
    throw new FileValidationError('[합계] 컬럼을 찾을 수 없습니다.');
  }
  return index;
}

function normalizeNumericText(value: unknown): string {
  return String(value ?? '')
    .replace(/,/g, '')
    .replace(/["']/g, '')
    .trim();
}

export function parsePayAmount(value: unknown): number {
  if (value === null || value === undefined || value === '') return 0;
  const text = normalizeNumericText(value);
  if (!text || text === '-') return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

export function parseNumber(value: unknown): number {
  if (value === null || value === undefined || value === '') return 0;
  const num = Number(normalizeNumericText(value));
  return Number.isFinite(num) ? num : 0;
}

export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/** 1번 파일명에서 월을 추출해 청구서 다운로드 파일명 생성 (예: ★4월 - 포코스 청구서.xlsx) */
export function getInvoiceDownloadFileName(uploadFileName: string): string {
  const baseName = uploadFileName.replace(/\.xlsx$/i, '');
  const monthMatch = baseName.match(/(\d{1,2})월/);
  if (!monthMatch) {
    return '★포코스 청구서.xlsx';
  }
  return `★${monthMatch[1]}월 - 포코스 청구서.xlsx`;
}

export async function downloadZip(files: { fileName: string; blob: Blob }[], zipName: string): Promise<void> {
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();

  for (const file of files) {
    zip.file(file.fileName, file.blob);
  }

  const blob = await zip.generateAsync({ type: 'blob' });
  downloadBlob(blob, zipName);
}
