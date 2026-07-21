import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import type { GeneratedFile, PayslipData } from '../types';
import { extractMonthFromFileName, formatNumber } from './calculations';
import {
  INVOICE_SHEET,
  readWorkbook,
  validateInvoiceSheet,
} from './fileProcessor';
import {
  parseInvoiceMeta,
  parseInvoiceSheetToNames,
  parseInvoiceSheetToPayslips,
} from './invoiceParser';

const PAYSLIP_TEMPLATE_PATH = '/payslip-template.xlsx';
const PAYSLIP_SHEET = 'D(프리랜서_)';

/** D파일(3번) 양식 셀 위치 — ExcelJS 1-based (B=항목, C=시간, D=금액) */
const PAYSLIP_CELL = {
  TITLE: { row: 1, col: 2 },
  NAME: { row: 2, col: 6 },
  BASE: { row: 4, label: 2, hours: 3, amount: 4 },
  OT: { row: 5, label: 2, hours: 3, amount: 4 },
  NIGHT: { row: 6, label: 2, hours: 3, amount: 4 },
  HOL: { row: 7, label: 2, hours: 3, amount: 4 },
  HOT: { row: 8, label: 2, hours: 3, amount: 4 },
  EVENT: { row: 9, label: 2, hours: 3, amount: 4 },
  FULL_ATTEND: { row: 10, label: 2, hours: 3, amount: 4 },
  BLANK_DEDUCT_ROWS: [4, 5, 6, 7] as const,
  INCOME_TAX: { row: 8, col: 6 },
  LOCAL_TAX: { row: 9, col: 6 },
  DEDUCT_TOTAL: { row: 13, col: 6 },
  GROSS: { row: 14, label: 2, hours: 3, amount: 4 },
  NET: { row: 14, col: 6 },
} as const;

/** 급여명세서 본문 테이블 (B3:F14) — 외곽 medium, 내부 thin */
const PAYSLIP_TABLE = {
  topRow: 3,
  bottomRow: 14,
  leftCol: 2,
  rightCol: 6,
} as const;

const PAYSLIP_BORDER_THIN: Partial<ExcelJS.Border> = { style: 'thin' };
const PAYSLIP_BORDER_MEDIUM: Partial<ExcelJS.Border> = { style: 'medium' };

function payslipCellBorder(row: number, col: number): Partial<ExcelJS.Borders> {
  const { topRow, bottomRow, leftCol, rightCol } = PAYSLIP_TABLE;
  return {
    top: row === topRow ? PAYSLIP_BORDER_MEDIUM : PAYSLIP_BORDER_THIN,
    bottom: row === bottomRow ? PAYSLIP_BORDER_MEDIUM : PAYSLIP_BORDER_THIN,
    left: col === leftCol ? PAYSLIP_BORDER_MEDIUM : PAYSLIP_BORDER_THIN,
    right: col === rightCol ? PAYSLIP_BORDER_MEDIUM : PAYSLIP_BORDER_THIN,
  };
}

function isPayslipTableCell(row: number, col: number): boolean {
  const { topRow, bottomRow, leftCol, rightCol } = PAYSLIP_TABLE;
  return row >= topRow && row <= bottomRow && col >= leftCol && col <= rightCol;
}

function removeCellBorder(cell: ExcelJS.Cell): void {
  if (!cell.border) return;
  const nextStyle = { ...cell.style } as Partial<ExcelJS.Style> & { border?: Partial<ExcelJS.Borders> };
  delete nextStyle.border;
  cell.style = nextStyle;
}

/** 테이블 밖 셀 테두리 제거 — rowCount/col 범위만 처리 (getCell로 빈 행 생성 방지) */
function clearPayslipExtraBorders(sheet: ExcelJS.Worksheet): void {
  const maxRow = sheet.rowCount;
  const maxCol = PAYSLIP_TABLE.rightCol + 2;

  for (let row = 1; row <= maxRow; row++) {
    for (let col = 1; col <= maxCol; col++) {
      if (!isPayslipTableCell(row, col)) {
        removeCellBorder(sheet.getCell(row, col));
      }
    }
  }
}

function configurePayslipPrint(sheet: ExcelJS.Worksheet): void {
  sheet.views = [{ showGridLines: false, zoomScale: 100 }];
  sheet.pageSetup = {
    ...sheet.pageSetup,
    showGridLines: false,
    showRowColHeaders: false,
    printArea: 'B1:F15',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    orientation: 'portrait',
  };
}

function applyPayslipTableBorders(sheet: ExcelJS.Worksheet): void {
  clearPayslipExtraBorders(sheet);
  for (let row = PAYSLIP_TABLE.topRow; row <= PAYSLIP_TABLE.bottomRow; row++) {
    for (let col = PAYSLIP_TABLE.leftCol; col <= PAYSLIP_TABLE.rightCol; col++) {
      sheet.getCell(row, col).border = payslipCellBorder(row, col);
    }
  }
  configurePayslipPrint(sheet);
}

function trimPayslipSheetRows(sheet: ExcelJS.Worksheet): void {
  const lastRow = 15;
  while (sheet.rowCount > lastRow) {
    sheet.spliceRows(sheet.rowCount, 1);
  }

  const internal = sheet as unknown as { _rows?: unknown[] };
  if (Array.isArray(internal._rows) && internal._rows.length > lastRow) {
    internal._rows.length = lastRow;
  }
}

function sanitizeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
}

function buildPayslipLabel(payslip: PayslipData): string {
  const safeName = sanitizeFileName(payslip.성명);
  return `${safeName}_${payslip.month}월_급여명세서`;
}

function buildPayslipFileName(payslip: PayslipData, ext: string): string {
  return `${buildPayslipLabel(payslip)}.${ext}`;
}

/** Excel 시트 탭명 (최대 31자, 특수문자 제한) */
function buildPayslipSheetName(payslip: PayslipData): string {
  return buildPayslipLabel(payslip).replace(/[\[\]]/g, '_').slice(0, 31);
}

function setAmountCell(sheet: ExcelJS.Worksheet, row: number, col: number, value: number): void {
  sheet.getCell(row, col).value = value > 0 ? value : null;
}

function setPayItemCell(
  sheet: ExcelJS.Worksheet,
  row: number,
  labelCol: number,
  hoursCol: number,
  amountCol: number,
  label: string,
  hours: number,
  amount: number,
): void {
  sheet.getCell(row, labelCol).value = label;
  sheet.getCell(row, hoursCol).value = hours;
  setAmountCell(sheet, row, amountCol, amount);
}

function fillPayRow(
  sheet: ExcelJS.Worksheet,
  cell: { row: number; label: number; hours: number; amount: number },
  label: string,
  hours: number,
  amount: number,
): void {
  setPayItemCell(sheet, cell.row, cell.label, cell.hours, cell.amount, label, hours, amount);
}

function fillPayslipSheet(sheet: ExcelJS.Worksheet, payslip: PayslipData): void {
  sheet.getCell(PAYSLIP_CELL.TITLE.row, PAYSLIP_CELL.TITLE.col).value =
    `${payslip.year}년 ${payslip.month}월 급여명세서`;
  sheet.getCell(PAYSLIP_CELL.NAME.row, PAYSLIP_CELL.NAME.col).value = `${payslip.성명} 님`;

  fillPayRow(sheet, PAYSLIP_CELL.BASE, '기본급', payslip.기본급시간, payslip.기본급);
  fillPayRow(sheet, PAYSLIP_CELL.OT, '연장', payslip.연장시간, payslip.연장);
  fillPayRow(sheet, PAYSLIP_CELL.NIGHT, '심야수당', payslip.심야수당시간, payslip.심야수당);
  fillPayRow(sheet, PAYSLIP_CELL.HOL, '주특', payslip.주특시간, payslip.주특);
  fillPayRow(sheet, PAYSLIP_CELL.HOT, '특잔', payslip.특잔시간, payslip.특잔);
  fillPayRow(sheet, PAYSLIP_CELL.EVENT, '경조사비', 0, payslip.경조사비);
  fillPayRow(sheet, PAYSLIP_CELL.FULL_ATTEND, '만근수당', 0, payslip.만근수당);

  for (const row of PAYSLIP_CELL.BLANK_DEDUCT_ROWS) {
    sheet.getCell(row, 6).value = null;
  }

  setAmountCell(
    sheet,
    PAYSLIP_CELL.INCOME_TAX.row,
    PAYSLIP_CELL.INCOME_TAX.col,
    payslip.근로소득세,
  );
  setAmountCell(
    sheet,
    PAYSLIP_CELL.LOCAL_TAX.row,
    PAYSLIP_CELL.LOCAL_TAX.col,
    payslip.지방소득세,
  );
  setAmountCell(
    sheet,
    PAYSLIP_CELL.DEDUCT_TOTAL.row,
    PAYSLIP_CELL.DEDUCT_TOTAL.col,
    payslip.공제총액,
  );
  sheet.getCell(PAYSLIP_CELL.GROSS.row, PAYSLIP_CELL.GROSS.label).value = '급여총액';
  sheet.getCell(PAYSLIP_CELL.GROSS.row, PAYSLIP_CELL.GROSS.hours).value = null;
  setAmountCell(sheet, PAYSLIP_CELL.GROSS.row, PAYSLIP_CELL.GROSS.amount, payslip.급여총액);
  sheet.getCell(PAYSLIP_CELL.DEDUCT_TOTAL.row, 5).value = '공제총액';
  sheet.getCell(PAYSLIP_CELL.GROSS.row, 5).value = '실수령액';
  sheet.getCell(PAYSLIP_CELL.INCOME_TAX.row, 5).value = '근로소득세(3%)';
  sheet.getCell(PAYSLIP_CELL.LOCAL_TAX.row, 5).value = '지방소득세(0.3%)';
  setAmountCell(sheet, PAYSLIP_CELL.NET.row, PAYSLIP_CELL.NET.col, payslip.실수령액);

  trimPayslipSheetRows(sheet);
  applyPayslipTableBorders(sheet);
}

async function loadPayslipTemplate(): Promise<ExcelJS.Workbook> {
  const response = await fetch(PAYSLIP_TEMPLATE_PATH);
  if (!response.ok) {
    throw new Error('급여명세서 양식 파일을 불러올 수 없습니다.');
  }
  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

async function createPayslipExcel(payslip: PayslipData): Promise<Blob> {
  const workbook = await loadPayslipTemplate();
  const sheet = workbook.getWorksheet(PAYSLIP_SHEET);
  if (!sheet) {
    throw new Error(`[${PAYSLIP_SHEET}] 시트를 찾을 수 없습니다.`);
  }

  const removeIds = workbook.worksheets
    .filter((ws) => ws.name !== PAYSLIP_SHEET)
    .map((ws) => ws.id);
  for (const id of removeIds) {
    workbook.removeWorksheet(id);
  }

  fillPayslipSheet(sheet, payslip);
  sheet.name = buildPayslipSheetName(payslip);

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

function buildPayslipHtml(payslip: PayslipData): string {
  const fmt = (n: number) => formatNumber(n);
  const fmtHours = (n: number) => n.toFixed(1);
  const fmtAmount = (n: number) => (n > 0 ? fmt(n) : '');
  const title = `${payslip.year}년 ${payslip.month}월 급여명세서`;

  const payRow = (
    payLabel: string,
    hours: number,
    payAmount: number,
    deductLabel: string,
    deductAmount?: number,
  ) => `<tr>
    <td class="label pay">${payLabel}</td>
    <td class="hours">${fmtHours(hours)}</td>
    <td class="amount pay">${fmtAmount(payAmount)}</td>
    <td class="label deduct">${deductLabel}</td>
    <td class="amount deduct">${deductAmount === undefined ? '' : fmt(deductAmount)}</td>
  </tr>`;

  return `
    <div class="payslip">
      <div class="title">${title}</div>
      <div class="name">${payslip.성명} 님</div>
      <table>
        <colgroup>
          <col class="col-pay-label" />
          <col class="col-hours" />
          <col class="col-pay-amount" />
          <col class="col-deduct-label" />
          <col class="col-deduct-amount" />
        </colgroup>
        <thead>
          <tr>
            <th colspan="3">급&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;여</th>
            <th colspan="2">세&nbsp;&nbsp;&nbsp;액&nbsp;&nbsp;&nbsp;&nbsp;및&nbsp;&nbsp;&nbsp;&nbsp;공&nbsp;&nbsp;&nbsp;제</th>
          </tr>
        </thead>
        <tbody>
          ${payRow('기본급', payslip.기본급시간, payslip.기본급, '국민연금')}
          ${payRow('연장', payslip.연장시간, payslip.연장, '건강보험')}
          ${payRow('심야수당', payslip.심야수당시간, payslip.심야수당, '장기요양보험')}
          ${payRow('주특', payslip.주특시간, payslip.주특, '고용보험')}
          ${payRow('특잔', payslip.특잔시간, payslip.특잔, '근로소득세(3%)', payslip.근로소득세)}
          ${payRow('경조사비', 0, payslip.경조사비, '지방소득세(0.3%)', payslip.지방소득세)}
          ${payRow('만근수당', 0, payslip.만근수당, '', undefined)}
          <tr class="summary">
            <td class="label pay"></td>
            <td class="hours"></td>
            <td class="amount pay"></td>
            <td class="label deduct">공제총액</td>
            <td class="amount deduct">${fmt(payslip.공제총액)}</td>
          </tr>
          <tr class="summary">
            <td class="label pay">급여총액</td>
            <td class="hours"></td>
            <td class="amount pay">${fmt(payslip.급여총액)}</td>
            <td class="label deduct">실수령액</td>
            <td class="amount deduct">${fmt(payslip.실수령액)}</td>
          </tr>
        </tbody>
      </table>
      <div class="footer">한 달 동안 고생 많으셨습니다.</div>
    </div>
  `;
}

const PAYSLIP_CSS = `
  .payslip {
    width: 680px;
    max-width: 100%;
    padding: 24px 28px;
    background: #fff;
    color: #000;
    font-size: 13px;
    line-height: 1.4;
    box-sizing: border-box;
  }
  .title {
    text-align: left;
    font-size: 21px;
    font-weight: bold;
    margin-bottom: 8px;
    text-decoration: underline;
    letter-spacing: 0.04em;
  }
  .name {
    text-align: right;
    font-size: 14px;
    margin-bottom: 12px;
    text-decoration: underline;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    border: 1px solid #000;
  }
  col.col-pay-label { width: 17%; }
  col.col-hours { width: 48px; }
  col.col-pay-amount { width: 21%; }
  col.col-deduct-label { width: 30%; }
  col.col-deduct-amount { width: 18%; }
  th {
    text-align: center;
    font-weight: normal;
    padding: 7px 4px;
    border: 1px solid #000;
    background: #d9d9d9;
    font-size: 13px;
  }
  td {
    padding: 6px 4px;
    vertical-align: middle;
    border: 1px solid #000;
    overflow: hidden;
  }
  .label.pay {
    text-align: center;
    font-size: 13px;
    white-space: nowrap;
  }
  .label.deduct {
    text-align: center;
    font-size: 11.5px;
    letter-spacing: -0.03em;
    white-space: nowrap;
    padding: 6px 3px;
  }
  .hours {
    text-align: center;
    font-size: 12px;
    padding: 6px 2px;
    white-space: nowrap;
  }
  .amount {
    text-align: right;
    padding-right: 8px;
    font-size: 13px;
    white-space: nowrap;
  }
  tr.summary td {
    background: #d9d9d9;
    font-weight: bold;
    font-size: 12.5px;
  }
  .footer {
    text-align: center;
    margin-top: 16px;
    padding: 10px 8px;
    background: #b4c7e7;
    border: 1px solid #000;
    font-size: 13px;
  }
`;

/** A4 출력용 — scale 2 PNG 대비 용량 대폭 절감 */
const PDF_CANVAS_SCALE = 1.5;
const PDF_JPEG_QUALITY = 0.85;

async function createPayslipPdf(payslip: PayslipData): Promise<Blob> {
  const wrapper = document.createElement('div');
  wrapper.style.cssText =
    'position:fixed;left:-9999px;top:0;z-index:-1;font-family:"Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif;';

  const style = document.createElement('style');
  style.textContent = PAYSLIP_CSS;
  wrapper.appendChild(style);

  const content = document.createElement('div');
  content.innerHTML = buildPayslipHtml(payslip);
  wrapper.appendChild(content);
  document.body.appendChild(wrapper);

  try {
    const canvas = await html2canvas(content.firstElementChild as HTMLElement, {
      scale: PDF_CANVAS_SCALE,
      backgroundColor: '#ffffff',
      useCORS: true,
    });

    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4', compress: true });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 15;
    const imgWidth = pageWidth - margin * 2;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    const imgData = canvas.toDataURL('image/jpeg', PDF_JPEG_QUALITY);
    doc.addImage(imgData, 'JPEG', margin, margin, imgWidth, imgHeight);
    return doc.output('blob');
  } finally {
    document.body.removeChild(wrapper);
  }
}

function resolvePayslipPeriod(
  fileName: string,
  sheetMeta: { year: number; month: string },
): { year: number; month: string } {
  const fromFile = extractMonthFromFileName(fileName);
  return {
    year: sheetMeta.year,
    month: fromFile ?? sheetMeta.month,
  };
}

export async function convertInvoiceToPayslip(
  file: File,
  format: 'pdf' | 'excel',
): Promise<GeneratedFile[]> {
  const workbook = await readWorkbook(file);
  validateInvoiceSheet(workbook);

  const sheet = workbook.Sheets[INVOICE_SHEET];
  if (!sheet) {
    throw new Error(`[${INVOICE_SHEET}] 시트를 찾을 수 없습니다.`);
  }

  const sheetMeta = parseInvoiceMeta(sheet);
  const { year, month } = resolvePayslipPeriod(file.name, sheetMeta);
  const payslips = parseInvoiceSheetToPayslips(sheet, year, month);

  if (payslips.length === 0) {
    throw new Error('변환할 직원 데이터가 없습니다. 청구내역서 시트를 확인해주세요.');
  }

  const ext = format === 'pdf' ? 'pdf' : 'xlsx';
  const usedNames = new Map<string, number>();

  const files: GeneratedFile[] = [];
  for (const payslip of payslips) {
    let fileName = buildPayslipFileName(payslip, ext);
    const count = usedNames.get(fileName) ?? 0;
    if (count > 0) {
      const safeName = sanitizeFileName(payslip.성명);
      fileName = `${safeName}_${payslip.month}월_급여명세서_${count + 1}.${ext}`;
    }
    usedNames.set(buildPayslipFileName(payslip, ext), count + 1);

    const blob =
      format === 'pdf' ? await createPayslipPdf(payslip) : await createPayslipExcel(payslip);
    files.push({ fileName, blob });
  }

  return files;
}

export async function previewInvoiceEmployees(file: File): Promise<string[]> {
  const workbook = await readWorkbook(file);
  validateInvoiceSheet(workbook);
  const sheet = workbook.Sheets[INVOICE_SHEET];
  if (!sheet) {
    throw new Error(`[${INVOICE_SHEET}] 시트를 찾을 수 없습니다.`);
  }
  return parseInvoiceSheetToNames(sheet);
}
