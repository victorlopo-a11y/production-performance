
export interface ProductionPlanRow {
  line: string;
  shift: number;
  code: string;
  product: string;
  date: string;
  meta: number;
}

export interface ActualProductionRecord {
  line: string;
  material: string;
  quantity: number;
  date: string;
  time: string;
  shift: number;
  productionDay: string;
}

export interface FailureReportRecord {
  line: string;
  material: string;
  origin: string;
  quantity: number;
  date: string;
  time: string;
  shift: number;
  productionDay: string;
}

export interface ComparisonResult {
  date: string;
  shift: number;
  line: string;
  material: string;
  product: string;
  meta: number;
  produced: number;
  failures: number;
  failuresByOrigin: Record<string, number>;
  yield: number;
  difference: number;
  efficiency: number;
}

export enum ShiftType {
  Shift1 = 1,
  Shift2 = 2
}
