import {
  compareMonths,
  monthNumber,
  monthYear,
  sameMonthPreviousYear,
  trailingMonths,
  ytdMonths,
  type MonthKey
} from "./month";
import { UNMAPPED_SEGMENT, type CustomerSection, type MetricRow, type SourceTransaction } from "./entities";

const SEGMENT_ORDER = ["B&M", "WS", "Ecom", "ETC", "IDP", "KEY", "OTC", "MT", "AK", UNMAPPED_SEGMENT];

type ValueSelector = (transaction: SourceTransaction) => number;

export function safeRatio(numerator: number, denominator: number): number {
  return denominator === 0 ? 0 : numerator / denominator;
}

export function sumMonths(values: Record<MonthKey, number>, months: MonthKey[]): number {
  return months.reduce((sum, month) => sum + (values[month] ?? 0), 0);
}

export function averageMonths(values: Record<MonthKey, number>, months: MonthKey[]): number {
  return months.length === 0 ? 0 : sumMonths(values, months) / months.length;
}

export function aggregateMetricRows(
  transactions: SourceTransaction[],
  periodMonths: MonthKey[],
  reportMonth: MonthKey,
  groupSelector: (transaction: SourceTransaction) => string,
  valueSelector: ValueSelector,
  sortMode: "segment" | "total" | "alpha" = "total"
): MetricRow[] {
  const groups = new Map<string, Record<MonthKey, number>>();

  for (const transaction of transactions) {
    const group = groupSelector(transaction).trim() || UNMAPPED_SEGMENT;
    const monthValues = groups.get(group) ?? {};
    monthValues[transaction.month] = (monthValues[transaction.month] ?? 0) + valueSelector(transaction);
    groups.set(group, monthValues);
  }

  const rows = Array.from(groups.entries()).map(([label, values]) => {
    return buildMetricRow(label, values, periodMonths, reportMonth);
  });

  rows.sort((a, b) => {
    if (sortMode === "segment") {
      const left = SEGMENT_ORDER.indexOf(a.label);
      const right = SEGMENT_ORDER.indexOf(b.label);
      if (left !== -1 || right !== -1) {
        return (left === -1 ? 999 : left) - (right === -1 ? 999 : right);
      }
    }

    if (sortMode === "alpha") {
      return a.label.localeCompare(b.label, "vi");
    }

    const totalDiff = b.currentYearTotal + b.previousYearTotal - (a.currentYearTotal + a.previousYearTotal);
    return totalDiff === 0 ? a.label.localeCompare(b.label, "vi") : totalDiff;
  });

  return rows;
}

export function buildMetricRow(
  label: string,
  sourceValues: Record<MonthKey, number>,
  periodMonths: MonthKey[],
  reportMonth: MonthKey
): MetricRow {
  const currentYear = monthYear(reportMonth);
  const previousYear = currentYear - 1;
  const throughMonth = monthNumber(reportMonth);
  const previousYearMonths = periodMonths.filter((month) => monthYear(month) === previousYear);
  const currentYearMonths = periodMonths.filter((month) => monthYear(month) === currentYear);
  const p3Months = trailingMonths(reportMonth, 3);
  const p6Months = trailingMonths(reportMonth, 6);
  const p9Months = trailingMonths(reportMonth, 9);
  const previousYtdMonths = ytdMonths(previousYear, throughMonth);
  const currentYtdMonths = ytdMonths(currentYear, throughMonth);

  const monthValues = Object.fromEntries(periodMonths.map((month) => [month, sourceValues[month] ?? 0])) as Record<
    MonthKey,
    number
  >;

  const previousYearTotal = sumMonths(sourceValues, previousYearMonths);
  const currentYearTotal = sumMonths(sourceValues, currentYearMonths);
  const p3m = averageMonths(sourceValues, p3Months);
  const p6m = averageMonths(sourceValues, p6Months);
  const p9m = averageMonths(sourceValues, p9Months);
  const currentYtd = sumMonths(sourceValues, currentYtdMonths);
  const previousYtd = sumMonths(sourceValues, previousYtdMonths);
  const currentMonth = sourceValues[reportMonth] ?? 0;
  const previousYearSameMonth = sourceValues[sameMonthPreviousYear(reportMonth)] ?? 0;

  return {
    label,
    monthValues,
    previousYearTotal,
    currentYearTotal,
    p3m,
    p6m,
    p9m,
    trend: safeRatio(p3m * 2, p6m + p9m),
    ifytd: safeRatio(currentYtd, previousYtd),
    icytd: safeRatio(currentYtd, previousYtd),
    iya: safeRatio(currentMonth, previousYearSameMonth)
  };
}

export function addShareColumns(rows: MetricRow[]): MetricRow[] {
  const previousTotal = rows.reduce((sum, row) => sum + row.previousYearTotal, 0);
  const currentTotal = rows.reduce((sum, row) => sum + row.currentYearTotal, 0);

  return rows.map((row) => ({
    ...row,
    previousYearShare: safeRatio(row.previousYearTotal, previousTotal),
    currentYearShare: safeRatio(row.currentYearTotal, currentTotal)
  }));
}

export function totalMetricRow(label: string, rows: MetricRow[], periodMonths: MonthKey[], reportMonth: MonthKey): MetricRow {
  const monthValues = periodMonths.reduce<Record<MonthKey, number>>((values, month) => {
    values[month] = rows.reduce((sum, row) => sum + (row.monthValues[month] ?? 0), 0);
    return values;
  }, {} as Record<MonthKey, number>);

  return buildMetricRow(label, monthValues, periodMonths, reportMonth);
}

export function buildCustomerSections(
  transactions: SourceTransaction[],
  periodMonths: MonthKey[],
  reportMonth: MonthKey,
  valueSelector: ValueSelector
): CustomerSection[] {
  const transactionsByCustomer = new Map<string, SourceTransaction[]>();

  for (const transaction of transactions) {
    const customer = transaction.customer.trim() || "Unknown customer";
    const customerRows = transactionsByCustomer.get(customer) ?? [];
    customerRows.push(transaction);
    transactionsByCustomer.set(customer, customerRows);
  }

  return Array.from(transactionsByCustomer.entries())
    .sort(([a], [b]) => a.localeCompare(b, "vi"))
    .map(([customer, customerTransactions]) => {
      const rows = aggregateMetricRows(
        customerTransactions,
        periodMonths,
        reportMonth,
        (transaction) => transaction.product,
        valueSelector,
        "total"
      );

      return {
        customer,
        rows,
        total: totalMetricRow("Total", rows, periodMonths, reportMonth)
      };
    });
}

export function filterTransactionsThroughMonth(
  transactions: SourceTransaction[],
  reportMonth: MonthKey,
  periodMonths: MonthKey[]
): SourceTransaction[] {
  const periodSet = new Set(periodMonths);
  return transactions.filter((transaction) => {
    return compareMonths(transaction.month, reportMonth) <= 0 && periodSet.has(transaction.month);
  });
}
