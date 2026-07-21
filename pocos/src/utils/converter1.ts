import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import type { EmployeePayRecord, InvoiceRow, WorkCategory } from '../types';
import {
  calculateInvoiceRow,
  emptyCategoryPay,
  extractMonthFromFileName,
  extractYearFromWorkbook,
  roundOnesDigit,
} from './calculations';
import {
  findHeaderRow,
  findTotalColumn,
  getColumnIndex,
  parseNumber,
  parsePayAmount,
  readWorkbook,
  sheetToRows,
  validateWorkRecordSheets,
  WORK_SHEETS,
} from './fileProcessor';

const CATEGORY_MAP: Record<WorkCategory, keyof Pick<EmployeePayRecord, '정상' | '연장' | '야간' | '휴일' | '휴연'>> = {
  정상: '정상',
  연장: '연장',
  야간: '야간',
  휴일: '휴일',
  휴연: '휴연',
};

const VALID_CATEGORIES = new Set<string>(Object.keys(CATEGORY_MAP));

const TEMPLATE_PATH = '/invoice-template.xlsx';
const DATA_START_ROW = 5;
const TEMPLATE_TOTAL_ROW = 176;
const TEMPLATE_SUMMARY_ROWS = [178, 179, 180] as const;
const SUMMARY_LABEL_COL = 28;
const SUMMARY_VALUE_COL = 29;

const INVOICE_COL = {
  DEPT: 4,
  SEQ: 6,
  NAME: 7,
  HOURLY: 8,
  BASE_H: 9,
  BASE_A: 10,
  OT_H: 11,
  OT_A: 12,
  NIGHT_H: 13,
  NIGHT_A: 14,
  HOL_H: 15,
  HOL_A: 16,
  HOT_H: 17,
  HOT_A: 18,
  FULL_ATTEND: 19,
  DIRECT: 21,
  PENSION: 22,
  HEALTH: 23,
  LONGTERM: 24,
  EMPLOY: 25,
  INDUST: 26,
  MGMT: 27,
  INDIRECT: 28,
  EVENT: 29,
  TOTAL: 30,
} as const;

const SUM_COLUMNS = [
  INVOICE_COL.BASE_H,
  INVOICE_COL.BASE_A,
  INVOICE_COL.OT_H,
  INVOICE_COL.OT_A,
  INVOICE_COL.NIGHT_H,
  INVOICE_COL.NIGHT_A,
  INVOICE_COL.HOL_H,
  INVOICE_COL.HOL_A,
  INVOICE_COL.HOT_H,
  INVOICE_COL.HOT_A,
  INVOICE_COL.FULL_ATTEND,
  INVOICE_COL.DIRECT,
  INVOICE_COL.PENSION,
  INVOICE_COL.HEALTH,
  INVOICE_COL.LONGTERM,
  INVOICE_COL.EMPLOY,
  INVOICE_COL.INDUST,
  INVOICE_COL.MGMT,
  INVOICE_COL.INDIRECT,
  INVOICE_COL.EVENT,
  INVOICE_COL.TOTAL,
] as const;

const DATA_STYLE_COLUMNS = [
  INVOICE_COL.DEPT,
  INVOICE_COL.SEQ,
  INVOICE_COL.NAME,
  INVOICE_COL.HOURLY,
  INVOICE_COL.BASE_H,
  INVOICE_COL.BASE_A,
  INVOICE_COL.OT_H,
  INVOICE_COL.OT_A,
  INVOICE_COL.NIGHT_H,
  INVOICE_COL.NIGHT_A,
  INVOICE_COL.HOL_H,
  INVOICE_COL.HOL_A,
  INVOICE_COL.HOT_H,
  INVOICE_COL.HOT_A,
  INVOICE_COL.FULL_ATTEND,
  20,
  INVOICE_COL.DIRECT,
  INVOICE_COL.PENSION,
  INVOICE_COL.HEALTH,
  INVOICE_COL.LONGTERM,
  INVOICE_COL.EMPLOY,
  INVOICE_COL.INDUST,
  INVOICE_COL.MGMT,
  INVOICE_COL.INDIRECT,
  INVOICE_COL.EVENT,
  INVOICE_COL.TOTAL,
] as const;

/** 컬럼 너비 — #### 방지 및 긴 성명 표시 */
const COLUMN_WIDTHS: Partial<Record<number, number>> = {
  [INVOICE_COL.DEPT]: 11,
  [INVOICE_COL.SEQ]: 6,
  [INVOICE_COL.NAME]: 16,
  [INVOICE_COL.HOURLY]: 12,
  [INVOICE_COL.BASE_H]: 10,
  [INVOICE_COL.BASE_A]: 14,
  [INVOICE_COL.OT_H]: 10,
  [INVOICE_COL.OT_A]: 14,
  [INVOICE_COL.NIGHT_H]: 10,
  [INVOICE_COL.NIGHT_A]: 14,
  [INVOICE_COL.HOL_H]: 10,
  [INVOICE_COL.HOL_A]: 14,
  [INVOICE_COL.HOT_H]: 10,
  [INVOICE_COL.HOT_A]: 14,
  [INVOICE_COL.FULL_ATTEND]: 12,
  20: 10,
  [INVOICE_COL.DIRECT]: 14,
  [INVOICE_COL.PENSION]: 13,
  [INVOICE_COL.HEALTH]: 13,
  [INVOICE_COL.LONGTERM]: 12,
  [INVOICE_COL.EMPLOY]: 12,
  [INVOICE_COL.INDUST]: 12,
  [INVOICE_COL.MGMT]: 13,
  [INVOICE_COL.INDIRECT]: 14,
  [INVOICE_COL.EVENT]: 10,
  [INVOICE_COL.TOTAL]: 18,
};

const NO_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFFFF' },
};

const THIN_BORDER: Partial<ExcelJS.Border> = { style: 'thin' };
const GRID_BORDER: Partial<ExcelJS.Borders> = {
  top: THIN_BORDER,
  left: THIN_BORDER,
  bottom: THIN_BORDER,
  right: THIN_BORDER,
};
const HOURLY_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF99' },
};
const AMOUNT_NUM_FMT = '#,##0';
const DATA_FONT: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 11 };
const TOTAL_AMOUNT_FONT: Partial<ExcelJS.Font> = { name: '맑은 고딕', size: 9, bold: true };
const LAST_DATA_COL = INVOICE_COL.TOTAL;

function excelColumnLetter(col: number): string {
  let n = col;
  let result = '';
  while (n > 0) {
    n -= 1;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function buildRowTotalFormula(rowNum: number): string {
  const direct = excelColumnLetter(INVOICE_COL.DIRECT);
  const indirect = excelColumnLetter(INVOICE_COL.INDIRECT);
  const fullAttend = excelColumnLetter(INVOICE_COL.FULL_ATTEND);
  const event = excelColumnLetter(INVOICE_COL.EVENT);
  return `ROUNDDOWN(${direct}${rowNum}+${indirect}${rowNum}+${fullAttend}${rowNum}+${event}${rowNum},-1)`;
}

function isSummaryRow(noText: string, name: string): boolean {
  const normalized = (noText + name).replace(/\s/g, '');
  return normalized.includes('합계');
}

function isDayColumn(header: string): boolean {
  const cell = header.replace(/^["']|["']$/g, '').trim();
  return /^\d{1,2}일$/.test(cell) || /^\d{1,2}$/.test(cell);
}

function createEmptyRecord(no: number, dept: string, name: string, hourlyRate: number): EmployeePayRecord {
  return {
    no,
    dept,
    name,
    hourlyRate,
    정상: emptyCategoryPay(),
    연장: emptyCategoryPay(),
    야간: emptyCategoryPay(),
    휴일: emptyCategoryPay(),
    휴연: emptyCategoryPay(),
  };
}

function parseHourlyRate(
  row: unknown[],
  totalCol: number,
  payCol: number,
  workHoursCol: number,
): number {
  const fromTotal = parsePayAmount(row[totalCol]);
  if (fromTotal > 0 && fromTotal < 100_000) {
    return fromTotal;
  }

  const hours = parseNumber(row[workHoursCol]);
  const pay = parsePayAmount(row[payCol]);
  if (hours > 0 && pay > 0) {
    const derived = pay / hours;
    if (derived > 0 && derived < 100_000) {
      return Math.round(derived);
    }
  }

  return 0;
}

function parseWorkSheet(sheet: XLSX.WorkSheet): EmployeePayRecord[] {
  const rows = sheetToRows(sheet);
  const headerRowIndex = findHeaderRow(rows, ['성명', '구분']);
  const headerRow = rows[headerRowIndex];

  const noCol = getColumnIndex(headerRow, 'NO');
  const deptCol = getColumnIndex(headerRow, '소속');
  const nameCol = getColumnIndex(headerRow, '성명');
  const categoryCol = getColumnIndex(headerRow, '구분');
  const workHoursCol = getColumnIndex(headerRow, '근무시간');
  const payCol = getColumnIndex(headerRow, '급여');
  const totalCol = findTotalColumn(headerRow);

  const dayCols: number[] = [];
  for (let i = 0; i < headerRow.length; i++) {
    const cell = String(headerRow[i] ?? '').trim();
    if (isDayColumn(cell)) dayCols.push(i);
  }

  const byNo = new Map<number, EmployeePayRecord>();
  let currentNo: number | null = null;
  let blockName = '';
  let blockDept = '';
  let isValidBlock = false;
  let hourlyRate = 0;

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;

    const noText = String(row[noCol] ?? '').trim();
    const name = String(row[nameCol] ?? '').trim();
    const category = String(row[categoryCol] ?? '').trim();

    if (isSummaryRow(noText, name)) break;

    const noValue = parseNumber(row[noCol]);

    if (noValue > 0) {
      currentNo = noValue;
      blockName = name;
      blockDept = String(row[deptCol] ?? '').trim();
      isValidBlock = name.length > 0;
      hourlyRate = 0;
      if (category === '정상' && isValidBlock) {
        hourlyRate = parseHourlyRate(row, totalCol, payCol, workHoursCol);
      }
    } else if (name && isValidBlock) {
      blockName = name;
    }

    if (!isValidBlock || !currentNo || !VALID_CATEGORIES.has(category)) continue;

    let record = byNo.get(currentNo);
    if (!record) {
      record = createEmptyRecord(currentNo, blockDept, blockName, hourlyRate);
      byNo.set(currentNo, record);
    }

    if (blockName) record.name = blockName;
    if (blockDept) record.dept = blockDept;
    if (category === '정상' && hourlyRate > 0) {
      record.hourlyRate = hourlyRate;
    }

    const hours =
      workHoursCol >= 0
        ? parseNumber(row[workHoursCol])
        : dayCols.reduce((sum, col) => sum + parseNumber(row[col]), 0);
    const amount = payCol >= 0 ? parsePayAmount(row[payCol]) : hours * record.hourlyRate;

    const key = CATEGORY_MAP[category as WorkCategory];
    record[key].hours += hours;
    record[key].amount += amount;
  }

  return [...byNo.values()].sort((a, b) => a.no - b.no);
}

function buildEmployeeRecords(workbook: XLSX.WorkBook): EmployeePayRecord[] {
  const all: EmployeePayRecord[] = [];
  for (const sheetName of WORK_SHEETS) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    all.push(...parseWorkSheet(sheet));
  }
  return all;
}

function buildInvoiceRows(records: EmployeePayRecord[]): InvoiceRow[] {
  return records.map((record, index) => calculateInvoiceRow(record, index + 1));
}

function setCellValue(row: ExcelJS.Row, col: number, value: number | string | null): void {
  if (value === null || value === '') return;
  row.getCell(col).value = value;
}

function fillInvoiceRow(row: ExcelJS.Row, invoice: InvoiceRow, dept: string): void {
  setCellValue(row, INVOICE_COL.DEPT, dept);
  setCellValue(row, INVOICE_COL.SEQ, invoice.순번);
  row.getCell(INVOICE_COL.NAME).value = invoice.성명;
  row.getCell(INVOICE_COL.HOURLY).value = invoice.시급;

  setCellValue(row, INVOICE_COL.BASE_H, invoice.기본급.hours);
  setCellValue(row, INVOICE_COL.BASE_A, invoice.기본급.amount);
  setCellValue(row, INVOICE_COL.OT_H, invoice.연장.hours);
  setCellValue(row, INVOICE_COL.OT_A, invoice.연장.amount);
  setCellValue(row, INVOICE_COL.NIGHT_H, invoice.심야수당.hours);
  setCellValue(row, INVOICE_COL.NIGHT_A, invoice.심야수당.amount);
  setCellValue(row, INVOICE_COL.HOL_H, invoice.주특.hours);
  setCellValue(row, INVOICE_COL.HOL_A, invoice.주특.amount);
  setCellValue(row, INVOICE_COL.HOT_H, invoice.특잔.hours);
  setCellValue(row, INVOICE_COL.HOT_A, invoice.특잔.amount);

  row.getCell(INVOICE_COL.FULL_ATTEND).value = null;
  setCellValue(row, INVOICE_COL.DIRECT, invoice.직접비소계);
  setCellValue(row, INVOICE_COL.PENSION, invoice.국민연금);
  setCellValue(row, INVOICE_COL.HEALTH, invoice.건강보험);
  setCellValue(row, INVOICE_COL.LONGTERM, invoice.장기요양);
  setCellValue(row, INVOICE_COL.EMPLOY, invoice.고용보험);
  setCellValue(row, INVOICE_COL.INDUST, invoice.산재보험);
  setCellValue(row, INVOICE_COL.MGMT, invoice.관리비);
  setCellValue(row, INVOICE_COL.INDIRECT, invoice.간접비소계);
  row.getCell(INVOICE_COL.EVENT).value = null;
  row.getCell(INVOICE_COL.TOTAL).value = {
    formula: buildRowTotalFormula(row.number),
  };
}

async function loadTemplateWorkbook(): Promise<ExcelJS.Workbook> {
  const response = await fetch(TEMPLATE_PATH);
  if (!response.ok) {
    throw new Error('청구서 양식 파일을 불러올 수 없습니다.');
  }
  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

/** 템플릿의 공유 수식(shared formula)을 값으로 변환 — ExcelJS 저장 오류 방지 */
function stripSheetFormulas(sheet: ExcelJS.Worksheet): void {
  sheet.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value;
      if (val === null || val === undefined || typeof val !== 'object') return;

      if ('sharedFormula' in val) {
        const formulaVal = val as { result?: ExcelJS.CellValue };
        cell.value = formulaVal.result ?? null;
      }

      if ('richText' in val && Array.isArray((val as ExcelJS.CellRichTextValue).richText)) {
        cell.value = (val as ExcelJS.CellRichTextValue).richText.map((t) => t.text).join('');
      }
    });
  });
}

type RowStyleSnapshot = Map<number, Partial<ExcelJS.Style>>;

function updateInvoiceHeaders(sheet: ExcelJS.Worksheet): void {
  sheet.getCell(3, INVOICE_COL.FULL_ATTEND).value = '만근수당';
  sheet.getCell(4, INVOICE_COL.FULL_ATTEND).value = '만근수당';
}

function updateSheetTitle(sheet: ExcelJS.Worksheet, year: number, month: string): void {
  sheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (typeof cell.value === 'string' && cell.value.includes('청구내역서')) {
        cell.value = `${year}년 ${month}월 청구내역서`;
      }
    });
  });
}

function captureRowStyles(sheet: ExcelJS.Worksheet, rowNum: number): RowStyleSnapshot {
  const styles: RowStyleSnapshot = new Map();
  const row = sheet.getRow(rowNum);
  row.eachCell({ includeEmpty: false }, (cell, col) => {
    if (cell.style) {
      styles.set(col, JSON.parse(JSON.stringify(cell.style)) as Partial<ExcelJS.Style>);
    }
  });
  return styles;
}

function applySelectedRowStyles(
  row: ExcelJS.Row,
  styles: RowStyleSnapshot,
  cols: readonly number[],
): void {
  for (const col of cols) {
    const style = styles.get(col);
    if (style) {
      row.getCell(col).style = JSON.parse(JSON.stringify(style)) as Partial<ExcelJS.Style>;
    }
  }
}

function applySheetColumnWidths(sheet: ExcelJS.Worksheet): void {
  for (const [colKey, width] of Object.entries(COLUMN_WIDTHS)) {
    const col = Number(colKey);
    sheet.getColumn(col).width = width;
  }
}

function applyDataRowStyles(row: ExcelJS.Row, styles: RowStyleSnapshot): void {
  for (const col of DATA_STYLE_COLUMNS) {
    const cell = row.getCell(col);
    const templateStyle = styles.get(col);
    if (templateStyle) {
      cell.style = JSON.parse(JSON.stringify(templateStyle)) as Partial<ExcelJS.Style>;
    }
    cell.border = { ...GRID_BORDER };
    cell.font = { ...DATA_FONT };
    if (col === INVOICE_COL.HOURLY) {
      cell.fill = { ...HOURLY_FILL };
    } else {
      cell.fill = { ...NO_FILL };
    }
    if (col >= INVOICE_COL.BASE_A && col <= INVOICE_COL.TOTAL) {
      cell.numFmt = AMOUNT_NUM_FMT;
    }
  }
}

function cloneFill(styles: RowStyleSnapshot, col: number): ExcelJS.Fill | undefined {
  const fill = styles.get(col)?.fill;
  if (!fill || typeof fill !== 'object') return undefined;
  return JSON.parse(JSON.stringify(fill)) as ExcelJS.Fill;
}

function clearCellFormatting(cell: ExcelJS.Cell): void {
  cell.value = null;
  cell.border = {};
  cell.fill = { ...NO_FILL };
}

function clearRowArea(row: ExcelJS.Row, fromCol: number, toCol: number): void {
  for (let col = fromCol; col <= toCol; col++) {
    clearCellFormatting(row.getCell(col));
  }
}

/** 총액(30열) 우측 및 청구총액 아래 불필요한 border·행 제거 */
function finalizeSheetLayout(sheet: ExcelJS.Worksheet, lastContentRow: number): void {
  trimRowsBelow(sheet, lastContentRow);

  for (let r = 1; r <= lastContentRow; r++) {
    const row = sheet.getRow(r);
    for (let c = LAST_DATA_COL + 1; c <= 60; c++) {
      clearCellFormatting(row.getCell(c));
    }
  }
}

function unmergeRowsFrom(sheet: ExcelJS.Worksheet, fromRow: number): void {
  const merges = [...(sheet.model.merges ?? [])];
  for (const range of merges) {
    const match = range.match(/^[A-Z]+(\d+):/);
    if (!match) continue;
    const startRow = Number(match[1]);
    if (startRow >= fromRow) {
      sheet.unMergeCells(range);
    }
  }
}

function trimRowsBelow(sheet: ExcelJS.Worksheet, lastRow: number): void {
  while (sheet.rowCount > lastRow) {
    sheet.spliceRows(sheet.rowCount, 1);
  }

  const internal = sheet as unknown as { _rows?: unknown[] };
  if (Array.isArray(internal._rows) && internal._rows.length > lastRow) {
    internal._rows.length = lastRow;
  }
}

function prepareInvoiceDataArea(
  sheet: ExcelJS.Worksheet,
): RowStyleSnapshot {
  const dataRowStyles = captureRowStyles(sheet, DATA_START_ROW);
  unmergeRowsFrom(sheet, DATA_START_ROW);

  while (sheet.rowCount >= DATA_START_ROW) {
    sheet.spliceRows(DATA_START_ROW, 1);
  }

  return dataRowStyles;
}

function getInvoiceColumnValue(invoice: InvoiceRow, col: number): number {
  switch (col) {
    case INVOICE_COL.BASE_H:
      return invoice.기본급.hours;
    case INVOICE_COL.BASE_A:
      return invoice.기본급.amount;
    case INVOICE_COL.OT_H:
      return invoice.연장.hours;
    case INVOICE_COL.OT_A:
      return invoice.연장.amount;
    case INVOICE_COL.NIGHT_H:
      return invoice.심야수당.hours;
    case INVOICE_COL.NIGHT_A:
      return invoice.심야수당.amount;
    case INVOICE_COL.HOL_H:
      return invoice.주특.hours;
    case INVOICE_COL.HOL_A:
      return invoice.주특.amount;
    case INVOICE_COL.HOT_H:
      return invoice.특잔.hours;
    case INVOICE_COL.HOT_A:
      return invoice.특잔.amount;
    case INVOICE_COL.FULL_ATTEND:
      return invoice.만근수당;
    case INVOICE_COL.DIRECT:
      return invoice.직접비소계;
    case INVOICE_COL.PENSION:
      return invoice.국민연금;
    case INVOICE_COL.HEALTH:
      return invoice.건강보험;
    case INVOICE_COL.LONGTERM:
      return invoice.장기요양;
    case INVOICE_COL.EMPLOY:
      return invoice.고용보험;
    case INVOICE_COL.INDUST:
      return invoice.산재보험;
    case INVOICE_COL.MGMT:
      return invoice.관리비;
    case INVOICE_COL.INDIRECT:
      return invoice.간접비소계;
    case INVOICE_COL.EVENT:
      return invoice.경조사비;
    case INVOICE_COL.TOTAL:
      return roundOnesDigit(invoice.급여총액);
    default:
      return 0;
  }
}

function fillTotalRow(
  row: ExcelJS.Row,
  invoices: InvoiceRow[],
  styles: RowStyleSnapshot,
): void {
  const totalFill = cloneFill(styles, INVOICE_COL.BASE_H) ?? cloneFill(styles, INVOICE_COL.BASE_A);

  for (const col of DATA_STYLE_COLUMNS) {
    const cell = row.getCell(col);
    const templateStyle = styles.get(col);
    if (templateStyle) {
      cell.style = JSON.parse(JSON.stringify(templateStyle)) as Partial<ExcelJS.Style>;
    }
    cell.border = { ...GRID_BORDER };
    if (totalFill) {
      cell.fill = totalFill;
    }
    cell.font = col >= INVOICE_COL.BASE_H ? { ...TOTAL_AMOUNT_FONT } : { ...DATA_FONT, bold: true };
  }

  for (const col of SUM_COLUMNS) {
    const cell = row.getCell(col);
    if (col === INVOICE_COL.TOTAL) {
      const firstRow = DATA_START_ROW;
      const lastRow = DATA_START_ROW + invoices.length - 1;
      const totalCol = excelColumnLetter(INVOICE_COL.TOTAL);
      cell.value = { formula: `SUM(${totalCol}${firstRow}:${totalCol}${lastRow})` };
    } else if (col === INVOICE_COL.FULL_ATTEND || col === INVOICE_COL.EVENT) {
      const colLetter = excelColumnLetter(col);
      const firstRow = DATA_START_ROW;
      const lastRow = DATA_START_ROW + invoices.length - 1;
      cell.value = { formula: `SUM(${colLetter}${firstRow}:${colLetter}${lastRow})` };
    } else {
      cell.value = invoices.reduce((acc, invoice) => acc + getInvoiceColumnValue(invoice, col), 0);
    }
    cell.numFmt = AMOUNT_NUM_FMT;
  }
  row.commit();
}

function clearGapRow(sheet: ExcelJS.Worksheet, rowNum: number): void {
  clearRowArea(sheet.getRow(rowNum), 4, 60);
  sheet.getRow(rowNum).commit();
}

function fillBottomSummary(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  supplyAmount: number,
  summaryStyles: RowStyleSnapshot[],
): void {
  const vatAmount = Math.round(supplyAmount * 0.1);
  const grandTotal = supplyAmount + vatAmount;

  const rows = [
    { label: '공급가액', value: supplyAmount, styleIndex: 0 },
    { label: '부가세', value: vatAmount, styleIndex: 1 },
    { label: '청구총액', value: grandTotal, styleIndex: 2 },
  ];

  rows.forEach(({ label, value, styleIndex }, index) => {
    const rowNum = startRow + index;
    try {
      sheet.unMergeCells(rowNum, SUMMARY_VALUE_COL, rowNum, 30);
    } catch {
      // ignore if not merged
    }

    const row = sheet.getRow(rowNum);
    clearRowArea(row, 4, 27);

    const styles = summaryStyles[styleIndex] ?? new Map();
    applySelectedRowStyles(row, styles, [SUMMARY_LABEL_COL, SUMMARY_VALUE_COL, 30]);

    const labelCell = row.getCell(SUMMARY_LABEL_COL);
    labelCell.value = label;
    const labelStyle = styles.get(SUMMARY_LABEL_COL);
    labelCell.border = labelStyle?.border
      ? (JSON.parse(JSON.stringify(labelStyle.border)) as Partial<ExcelJS.Borders>)
      : { ...GRID_BORDER };

    const valueCell = row.getCell(SUMMARY_VALUE_COL);
    valueCell.value = value;
    valueCell.numFmt = AMOUNT_NUM_FMT;
    const valueStyle = styles.get(SUMMARY_VALUE_COL);
    valueCell.border = valueStyle?.border
      ? (JSON.parse(JSON.stringify(valueStyle.border)) as Partial<ExcelJS.Borders>)
      : { ...GRID_BORDER };

    const valueCell2 = row.getCell(30);
    const valueStyle2 = styles.get(30);
    valueCell2.border = valueStyle2?.border
      ? (JSON.parse(JSON.stringify(valueStyle2.border)) as Partial<ExcelJS.Borders>)
      : { ...GRID_BORDER };
    if (styleIndex === 0) {
      valueCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
      };
      valueCell2.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
      };
    }

    try {
      sheet.mergeCells(rowNum, SUMMARY_VALUE_COL, rowNum, 30);
    } catch {
      // ignore
    }

    row.commit();
  });
}

function trimAndAppendSummary(
  sheet: ExcelJS.Worksheet,
  invoices: InvoiceRow[],
  totalRowStyles: RowStyleSnapshot,
  summaryRowStyles: RowStyleSnapshot[],
): number {
  const totalRowNum = DATA_START_ROW + invoices.length;

  trimRowsBelow(sheet, totalRowNum - 1);

  const supplyAmount = invoices.reduce(
    (sum, invoice) => sum + roundOnesDigit(invoice.급여총액),
    0,
  );

  fillTotalRow(sheet.getRow(totalRowNum), invoices, totalRowStyles);
  clearGapRow(sheet, totalRowNum + 1);
  fillBottomSummary(sheet, totalRowNum + 2, supplyAmount, summaryRowStyles);

  return totalRowNum + 4;
}

function removeExtraSheets(workbook: ExcelJS.Workbook): void {
  const keepSheet = '청구내역서';
  const removeIds = workbook.worksheets
    .filter((ws) => ws.name.trim() !== keepSheet)
    .map((ws) => ws.id);

  for (const id of removeIds) {
    workbook.removeWorksheet(id);
  }
}

async function writeInvoiceWorkbook(
  records: EmployeePayRecord[],
  invoices: InvoiceRow[],
  year: number,
  month: string,
): Promise<Blob> {
  const workbook = await loadTemplateWorkbook();
  const sheet = workbook.getWorksheet('청구내역서');
  if (!sheet) {
    throw new Error('[청구내역서] 시트를 찾을 수 없습니다.');
  }

  stripSheetFormulas(sheet);
  updateSheetTitle(sheet, year, month);
  updateInvoiceHeaders(sheet);
  applySheetColumnWidths(sheet);

  const totalRowStyles = captureRowStyles(sheet, TEMPLATE_TOTAL_ROW);
  const summaryRowStyles = TEMPLATE_SUMMARY_ROWS.map((rowNum) => captureRowStyles(sheet, rowNum));
  const dataRowStyles = prepareInvoiceDataArea(sheet);

  invoices.forEach((invoice, index) => {
    const rowNum = DATA_START_ROW + index;
    const row = sheet.getRow(rowNum);
    applyDataRowStyles(row, dataRowStyles);
    fillInvoiceRow(row, invoice, records[index]?.dept ?? '');
    row.commit();
  });

  trimAndAppendSummary(sheet, invoices, totalRowStyles, summaryRowStyles);

  const lastContentRow = DATA_START_ROW + invoices.length + 4;
  finalizeSheetLayout(sheet, lastContentRow);

  removeExtraSheets(workbook);

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

function getSheetTitle(workbook: XLSX.WorkBook): string | undefined {
  const sheet = workbook.Sheets[WORK_SHEETS[0]];
  if (!sheet?.["!ref"]) return undefined;
  const rows = sheetToRows(sheet);
  for (let i = 0; i < Math.min(3, rows.length); i++) {
    for (const cell of rows[i] ?? []) {
      const text = String(cell ?? '');
      if (text.includes('근무') && text.includes('현황')) return text;
    }
  }
  return undefined;
}

export async function convertWorkRecordToInvoice(file: File): Promise<Blob> {
  const workbook = await readWorkbook(file);
  validateWorkRecordSheets(workbook);

  const records = buildEmployeeRecords(workbook);
  if (records.length === 0) {
    throw new Error('변환할 직원 데이터가 없습니다. 파일 내용을 확인해주세요.');
  }

  const invoices = buildInvoiceRows(records);
  const month = extractMonthFromFileName(file.name) ?? String(new Date().getMonth() + 1);
  const year = extractYearFromWorkbook(getSheetTitle(workbook));

  return writeInvoiceWorkbook(records, invoices, year, month);
}

export async function previewWorkRecordEmployees(file: File): Promise<string[]> {
  const workbook = await readWorkbook(file);
  validateWorkRecordSheets(workbook);
  return buildEmployeeRecords(workbook).map((r) => r.name);
}

export { buildInvoiceRows, buildEmployeeRecords };
