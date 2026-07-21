import type { PayslipData } from '../types';
import { calculatePayslip } from './calculations';
import { parseNumber, parsePayAmount, sheetToRows } from './fileProcessor';
import type * as XLSX from 'xlsx';

export const INVOICE_SHEET_NAME = '청구내역서';

export interface InvoiceColumnMap {
  seqCol: number;
  nameCol: number;
  baseH: number;
  baseA: number;
  otH: number;
  otA: number;
  nightH: number;
  nightA: number;
  holH: number;
  holA: number;
  hotH: number;
  hotA: number;
  fullAttendCol: number;
  eventCol: number;
}

export interface InvoiceEmployeeAmounts {
  seq: number;
  name: string;
  기본급: number;
  연장: number;
  심야수당: number;
  주특: number;
  특잔: number;
  기본급시간: number;
  연장시간: number;
  심야수당시간: number;
  주특시간: number;
  특잔시간: number;
  만근수당: number;
  경조사비: number;
}

function findInvoiceHeaderRow(rows: unknown[][]): number {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    if (!row) continue;
    const hasSeq = row.some((cell) => String(cell ?? '').trim() === '순번');
    const hasName = row.some((cell) => String(cell ?? '').trim() === '성명');
    if (hasSeq && hasName) return i;
  }
  return 2;
}

function findColumnByKeyword(headerRow: unknown[], keyword: string): number {
  return headerRow.findIndex((cell) => String(cell ?? '').includes(keyword));
}

export function resolveInvoiceColumns(rows: unknown[][]): InvoiceColumnMap {
  const headerRowIndex = findInvoiceHeaderRow(rows);
  const headerRow = rows[headerRowIndex] ?? [];

  const seqCol = headerRow.findIndex((cell) => String(cell ?? '').trim() === '순번');
  const nameCol = headerRow.findIndex((cell) => String(cell ?? '').trim() === '성명');

  if (seqCol < 0 || nameCol < 0) {
    throw new Error('청구내역서에서 [순번]·[성명] 컬럼을 찾을 수 없습니다.');
  }

  const fullAttendCol = findColumnByKeyword(headerRow, '만근');
  const eventCol =
    findColumnByKeyword(headerRow, '경상') >= 0
      ? findColumnByKeyword(headerRow, '경상')
      : findColumnByKeyword(headerRow, '경조');

  return {
    seqCol,
    nameCol,
    baseH: nameCol + 2,
    baseA: nameCol + 3,
    otH: nameCol + 4,
    otA: nameCol + 5,
    nightH: nameCol + 6,
    nightA: nameCol + 7,
    holH: nameCol + 8,
    holA: nameCol + 9,
    hotH: nameCol + 10,
    hotA: nameCol + 11,
    fullAttendCol: fullAttendCol >= 0 ? fullAttendCol : nameCol + 12,
    eventCol: eventCol >= 0 ? eventCol : nameCol + 22,
  };
}

export function extractYearMonthFromInvoiceRows(rows: unknown[][]): { year: number; month: string } {
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    for (const cell of rows[i] ?? []) {
      const text = String(cell ?? '');
      const match = text.match(/(\d{2,4})년\s*(\d{1,2})월/);
      if (match) {
        const yearRaw = Number(match[1]);
        const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
        return { year, month: match[2] };
      }
    }
  }
  const now = new Date();
  return { year: now.getFullYear(), month: String(now.getMonth() + 1) };
}

function isSummaryRow(row: unknown[], cols: InvoiceColumnMap): boolean {
  const parts = [cols.seqCol, cols.nameCol, cols.baseA].map((col) =>
    String(row[col] ?? '').trim(),
  );
  const combined = parts.join('').replace(/\s/g, '');
  return combined.includes('합계') || combined.includes('공급가액') || combined.includes('청구총액');
}

/** 순수 숫자(시급 등)는 성명으로 취급하지 않음 */
function isValidEmployeeName(name: string): boolean {
  const trimmed = name.trim();
  if (!trimmed || trimmed === '-') return false;
  const numeric = trimmed.replace(/,/g, '').replace(/\s/g, '');
  if (/^\d+(\.\d+)?$/.test(numeric)) return false;
  return true;
}

function readEmployeeName(row: unknown[], cols: InvoiceColumnMap): string {
  const raw = String(row[cols.nameCol] ?? '')
    .replace(/\r\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (isValidEmployeeName(raw)) return raw;
  return '';
}

export function parseInvoiceEmployeeRows(rows: unknown[][]): InvoiceEmployeeAmounts[] {
  const cols = resolveInvoiceColumns(rows);
  const headerRow = findInvoiceHeaderRow(rows);
  const dataStartRow = headerRow + 2;
  const employees: InvoiceEmployeeAmounts[] = [];

  for (let r = dataStartRow; r < rows.length; r++) {
    const row = rows[r];
    if (!row || isSummaryRow(row, cols)) break;

    const seq = parseNumber(row[cols.seqCol]);
    const name = readEmployeeName(row, cols);
    if (!name || seq <= 0) continue;

    employees.push({
      seq,
      name,
      기본급: parsePayAmount(row[cols.baseA]),
      연장: parsePayAmount(row[cols.otA]),
      심야수당: parsePayAmount(row[cols.nightA]),
      주특: parsePayAmount(row[cols.holA]),
      특잔: parsePayAmount(row[cols.hotA]),
      기본급시간: parseNumber(row[cols.baseH]),
      연장시간: parseNumber(row[cols.otH]),
      심야수당시간: parseNumber(row[cols.nightH]),
      주특시간: parseNumber(row[cols.holH]),
      특잔시간: parseNumber(row[cols.hotH]),
      만근수당: parsePayAmount(row[cols.fullAttendCol]),
      경조사비: parsePayAmount(row[cols.eventCol]),
    });
  }

  return employees;
}

export function parseInvoiceSheetToPayslips(
  sheet: XLSX.WorkSheet,
  year: number,
  month: string,
): PayslipData[] {
  const rows = sheetToRows(sheet);
  return parseInvoiceEmployeeRows(rows).map((emp) =>
    calculatePayslip(
      emp.name,
      emp.기본급,
      emp.연장,
      emp.심야수당,
      emp.주특,
      emp.특잔,
      emp.경조사비,
      emp.만근수당,
      {
        기본급시간: emp.기본급시간,
        연장시간: emp.연장시간,
        심야수당시간: emp.심야수당시간,
        주특시간: emp.주특시간,
        특잔시간: emp.특잔시간,
      },
      year,
      month,
    ),
  );
}

export function parseInvoiceSheetToNames(sheet: XLSX.WorkSheet): string[] {
  const rows = sheetToRows(sheet);
  return parseInvoiceEmployeeRows(rows).map((emp) => emp.name);
}

export function parseInvoiceMeta(sheet: XLSX.WorkSheet): { year: number; month: string } {
  const rows = sheetToRows(sheet);
  return extractYearMonthFromInvoiceRows(rows);
}
