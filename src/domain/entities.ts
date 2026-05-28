import type { MonthKey } from "./month";

export const UNMAPPED_SEGMENT = "Unmapped";

export interface SourceTransaction {
  rowNumber: number;
  date: Date | null;
  month: MonthKey;
  customer: string;
  product: string;
  segment: string;
  dsr: string;
  unitPrice: number;
  quantity: number;
  amount: number;
  revenue: number;
}

export interface SourceParseResult {
  fileName: string;
  transactions: SourceTransaction[];
  availableMonths: MonthKey[];
  warnings: string[];
  skippedRows: number;
  segmentColumnFound: boolean;
}

export type ValueKind = "money" | "quantity";

export interface MetricRow {
  label: string;
  monthValues: Record<MonthKey, number>;
  previousYearShare?: number;
  currentYearShare?: number;
  previousYearTotal: number;
  currentYearTotal: number;
  p3m: number;
  p6m: number;
  p9m: number;
  trend: number;
  ifytd: number;
  icytd: number;
  iya: number;
}

export interface CustomerSection {
  customer: string;
  rows: MetricRow[];
  total: MetricRow;
}

export interface SkuCustomerSection {
  product: string;
  rows: MetricRow[];
  total: MetricRow;
}

export interface QecReport {
  reportMonth: MonthKey;
  periodMonths: MonthKey[];
  previousYear: number;
  currentYear: number;
  summary: ReportSummary;
  qecRows: MetricRow[];
  dsrRows: MetricRow[];
  customerBaseRows: MetricRow[];
  skuRevenueRows: MetricRow[];
  skuQuantityRows: MetricRow[];
  customerRevenueSections: CustomerSection[];
  customerQuantitySections: CustomerSection[];
  skuCustomerRevenueSections: SkuCustomerSection[];
  skuCustomerQuantitySections: SkuCustomerSection[];
}

export interface ReportSummary {
  transactionCount: number;
  customerCount: number;
  productCount: number;
  unmappedCustomerCount: number;
  unmappedCustomers: string[];
  availableMonths: MonthKey[];
}
