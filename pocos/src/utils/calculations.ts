import type { CategoryPay, EmployeePayRecord, InvoiceRow, PayslipData } from '../types';

const round = (value: number, digits = 0): number => {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
};

/** 10원 단위 내림 (1원 단위 버림) */
export function roundOnesDigit(value: number): number {
  return Math.floor(Math.trunc(value) / 10) * 10;
}

export function emptyCategoryPay(): CategoryPay {
  return { hours: 0, amount: 0 };
}

export function calculateInvoiceRow(record: EmployeePayRecord, 순번: number): InvoiceRow {
  const 기본급 = record.정상;
  const 연장 = record.연장;
  const 심야수당 = record.야간;
  const 주특 = record.휴일;
  const 특잔 = record.휴연;
  const 만근수당 = 0;
  const 경조사비 = 0;

  const 직접비소계 =
    기본급.amount + 연장.amount + 심야수당.amount + 주특.amount + 특잔.amount;
  const 국민연금 = round(직접비소계 * 0.0475);
  const 건강보험 = round(직접비소계 * 0.03595);
  const 장기요양 = round(건강보험 * 0.1314);
  const 고용보험 = round(직접비소계 * 0.009);
  const 산재보험 = round(직접비소계 * 0.007);
  const 관리비 = round(직접비소계 * 0.06);
  const 간접비소계 = 국민연금 + 건강보험 + 장기요양 + 고용보험 + 산재보험 + 관리비;
  const 급여총액 = 직접비소계 + 간접비소계;

  return {
    순번,
    성명: record.name,
    시급: record.hourlyRate,
    기본급,
    연장,
    심야수당,
    주특,
    특잔,
    만근수당,
    직접비소계,
    국민연금,
    건강보험,
    장기요양,
    고용보험,
    산재보험,
    관리비,
    간접비소계,
    경조사비,
    급여총액,
  };
}

export function calculatePayslip(
  성명: string,
  기본급: number,
  연장: number,
  심야수당: number,
  주특: number,
  특잔: number,
  경조사비 = 0,
  만근수당 = 0,
  hours: {
    기본급시간?: number;
    연장시간?: number;
    심야수당시간?: number;
    주특시간?: number;
    특잔시간?: number;
  } = {},
  year = new Date().getFullYear(),
  month = String(new Date().getMonth() + 1),
): PayslipData {
  const 급여총액 = 기본급 + 연장 + 심야수당 + 주특 + 특잔 + 경조사비 + 만근수당;
  const 근로소득세 = roundOnesDigit(급여총액 * 0.03);
  const 지방소득세 = roundOnesDigit(근로소득세 / 10);
  const 공제총액 = roundOnesDigit(근로소득세 + 지방소득세);
  const 실수령액 = roundOnesDigit(급여총액 - 공제총액);

  return {
    성명,
    year,
    month,
    기본급,
    연장,
    심야수당,
    주특,
    특잔,
    기본급시간: hours.기본급시간 ?? 0,
    연장시간: hours.연장시간 ?? 0,
    심야수당시간: hours.심야수당시간 ?? 0,
    주특시간: hours.주특시간 ?? 0,
    특잔시간: hours.특잔시간 ?? 0,
    경조사비,
    만근수당,
    급여총액,
    근로소득세,
    지방소득세,
    공제총액,
    실수령액,
  };
}

export function formatNumber(value: number): string {
  return value.toLocaleString('ko-KR');
}

export function extractMonthFromFileName(fileName: string): string | null {
  const baseName = fileName.replace(/\.xlsx$/i, '');
  const monthMatch = baseName.match(/(\d{1,2})월/);
  return monthMatch ? monthMatch[1] : null;
}

export function extractYearFromWorkbook(sheetTitle?: string): number {
  if (sheetTitle) {
    const match = sheetTitle.match(/(\d{2,4})년/);
    if (match) {
      const year = Number(match[1]);
      return year < 100 ? 2000 + year : year;
    }
  }
  return new Date().getFullYear();
}
