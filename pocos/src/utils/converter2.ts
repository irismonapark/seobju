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

/** D파일(3번) 양식 셀 위치 — ExcelJS 1-based */
const PAYSLIP_CELL = {
  TITLE: { row: 1, col: 2 },
  NAME: { row: 2, col: 6 },
  BASE: { row: 4, hours: 2, amount: 3 },
  OT: { row: 5, hours: 2, amount: 3 },
  NIGHT: { row: 6, hours: 2, amount: 3 },
  HOL: { row: 7, hours: 2, amount: 3 },
  HOT: { row: 8, hours: 2, amount: 3 },
  EVENT: { row: 9, hours: 2, amount: 3 },
  FULL_ATTEND: { row: 10, hours: 2, amount: 3 },
  BLANK_DEDUCT_ROWS: [4, 5, 6, 7] as const,
  INCOME_TAX: { row: 8, col: 6 },
  LOCAL_TAX: { row: 9, col: 6 },
  DEDUCT_TOTAL: { row: 13, col: 6 },
  GROSS: { row: 14, hours: 2, amount: 3 },
  NET: { row: 14, col: 6 },
} as const;

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
  hoursCol: number,
  amountCol: number,
  hours: number,
  amount: number,
): void {
  sheet.getCell(row, hoursCol).value = hours;
  setAmountCell(sheet, row, amountCol, amount);
}

function fillPayslipSheet(sheet: ExcelJS.Worksheet, payslip: PayslipData): void {
  sheet.getCell(PAYSLIP_CELL.TITLE.row, PAYSLIP_CELL.TITLE.col).value =
    `${payslip.year}년 ${payslip.month}월 급여명세서`;
  sheet.getCell(PAYSLIP_CELL.NAME.row, PAYSLIP_CELL.NAME.col).value = `${payslip.성명} 님`;

  setPayItemCell(
    sheet,
    PAYSLIP_CELL.BASE.row,
    PAYSLIP_CELL.BASE.hours,
    PAYSLIP_CELL.BASE.amount,
    payslip.기본급시간,
    payslip.기본급,
  );
  setPayItemCell(
    sheet,
    PAYSLIP_CELL.OT.row,
    PAYSLIP_CELL.OT.hours,
    PAYSLIP_CELL.OT.amount,
    payslip.연장시간,
    payslip.연장,
  );
  setPayItemCell(
    sheet,
    PAYSLIP_CELL.NIGHT.row,
    PAYSLIP_CELL.NIGHT.hours,
    PAYSLIP_CELL.NIGHT.amount,
    payslip.심야수당시간,
    payslip.심야수당,
  );
  setPayItemCell(
    sheet,
    PAYSLIP_CELL.HOL.row,
    PAYSLIP_CELL.HOL.hours,
    PAYSLIP_CELL.HOL.amount,
    payslip.주특시간,
    payslip.주특,
  );
  setPayItemCell(
    sheet,
    PAYSLIP_CELL.HOT.row,
    PAYSLIP_CELL.HOT.hours,
    PAYSLIP_CELL.HOT.amount,
    payslip.특잔시간,
    payslip.특잔,
  );
  setPayItemCell(
    sheet,
    PAYSLIP_CELL.EVENT.row,
    PAYSLIP_CELL.EVENT.hours,
    PAYSLIP_CELL.EVENT.amount,
    0,
    payslip.경조사비,
  );
  setPayItemCell(
    sheet,
    PAYSLIP_CELL.FULL_ATTEND.row,
    PAYSLIP_CELL.FULL_ATTEND.hours,
    PAYSLIP_CELL.FULL_ATTEND.amount,
    0,
    payslip.만근수당,
  );

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
  setAmountCell(sheet, PAYSLIP_CELL.GROSS.row, PAYSLIP_CELL.GROSS.amount, payslip.급여총액);
  sheet.getCell(PAYSLIP_CELL.GROSS.row, PAYSLIP_CELL.GROSS.hours).value = null;
  setAmountCell(sheet, PAYSLIP_CELL.NET.row, PAYSLIP_CELL.NET.col, payslip.실수령액);
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
    <td class="label">${payLabel}</td>
    <td class="hours">${fmtHours(hours)}</td>
    <td class="amount">${fmtAmount(payAmount)}</td>
    <td class="label deduct">${deductLabel}</td>
    <td class="amount">${deductAmount === undefined ? '' : fmt(deductAmount)}</td>
  </tr>`;

  return `
    <div class="payslip">
      <div class="title">${title}</div>
      <div class="name">${payslip.성명} 님</div>
      <table>
        <thead>
          <tr>
            <th colspan="3">급&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;여</th>
            <th colspan="2">세&nbsp;&nbsp;&nbsp;액&nbsp;&nbsp;&nbsp;&nbsp;및&nbsp;&nbsp;&nbsp;&nbsp;공&nbsp;&nbsp;&nbsp;제</th>
          </tr>
        </thead>
        <tbody>
          ${payRow('기&nbsp;&nbsp;본&nbsp;&nbsp;급', payslip.기본급시간, payslip.기본급, '국&nbsp;&nbsp;민&nbsp;&nbsp;연&nbsp;&nbsp;금')}
          ${payRow('연장', payslip.연장시간, payslip.연장, '건&nbsp;&nbsp;강&nbsp;&nbsp;보&nbsp;&nbsp;험')}
          ${payRow('심야수당', payslip.심야수당시간, payslip.심야수당, '장 기 요 양 보 험')}
          ${payRow('주특', payslip.주특시간, payslip.주특, '고용보험')}
          ${payRow('특잔', payslip.특잔시간, payslip.특잔, '근로소득세 (3%)', payslip.근로소득세)}
          ${payRow('경조사비', 0, payslip.경조사비, '지방소득세 (0.3%)', payslip.지방소득세)}
          ${payRow('만근수당', 0, payslip.만근수당, '', undefined)}
          <tr class="summary">
            <td class="label"></td>
            <td class="hours"></td>
            <td class="amount"></td>
            <td class="label deduct">공&nbsp;&nbsp;제&nbsp;&nbsp;총&nbsp;&nbsp;액</td>
            <td class="amount">${fmt(payslip.공제총액)}</td>
          </tr>
          <tr class="summary">
            <td class="label">급&nbsp;&nbsp;여&nbsp;&nbsp;총&nbsp;&nbsp;액</td>
            <td class="hours"></td>
            <td class="amount">${fmt(payslip.급여총액)}</td>
            <td class="label deduct">실&nbsp;&nbsp;수&nbsp;&nbsp;령&nbsp;&nbsp;액</td>
            <td class="amount">${fmt(payslip.실수령액)}</td>
          </tr>
        </tbody>
      </table>
      <div class="footer">한 달 동안 고생 많으셨습니다.</div>
    </div>
  `;
}

const PAYSLIP_CSS = `
  .payslip { width: 560px; padding: 28px 36px; background: #fff; color: #000; font-size: 14px; line-height: 1.45; }
  .title { text-align: left; font-size: 22px; font-weight: bold; margin-bottom: 8px; text-decoration: underline; letter-spacing: 0.05em; }
  .name { text-align: right; font-size: 15px; margin-bottom: 14px; text-decoration: underline; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; border: 1px solid #000; }
  th { text-align: center; font-weight: normal; padding: 8px 4px; border: 1px solid #000; background: #d9d9d9; }
  td { padding: 7px 6px; vertical-align: middle; border: 1px solid #000; }
  .label { width: 18%; text-align: center; white-space: nowrap; }
  .label.deduct { width: 22%; }
  .hours { width: 12%; text-align: center; }
  .amount { width: 18%; text-align: right; padding-right: 10px; }
  tr.summary td { background: #d9d9d9; font-weight: bold; }
  .footer { text-align: center; margin-top: 18px; padding: 10px 8px; background: #b4c7e7; border: 1px solid #000; font-size: 14px; }
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
