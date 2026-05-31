export type WorkCategory = '정상' | '연장' | '야간' | '휴일' | '휴연';

export interface CategoryPay {
  hours: number;
  amount: number;
}

export interface EmployeePayRecord {
  no: number;
  dept: string;
  name: string;
  hourlyRate: number;
  정상: CategoryPay;
  연장: CategoryPay;
  야간: CategoryPay;
  휴일: CategoryPay;
  휴연: CategoryPay;
}

export interface InvoiceRow {
  순번: number;
  성명: string;
  시급: number;
  기본급: CategoryPay;
  연장: CategoryPay;
  심야수당: CategoryPay;
  주특: CategoryPay;
  특잔: CategoryPay;
  직접비소계: number;
  국민연금: number;
  건강보험: number;
  장기요양: number;
  고용보험: number;
  산재보험: number;
  관리비: number;
  간접비소계: number;
  경조사비: number;
  급여총액: number;
}

export interface PayslipData {
  성명: string;
  year: number;
  month: string;
  기본급: number;
  연장: number;
  심야수당: number;
  주특: number;
  특잔: number;
  경조사비: number;
  급여총액: number;
  근로소득세: number;
  지방소득세: number;
  공제총액: number;
  실수령액: number;
}

export interface GeneratedFile {
  fileName: string;
  blob: Blob;
}
