export type MonthKey = `${number}-${string}`;

export function parseMonthKey(input: unknown): MonthKey | null {
  if (input instanceof Date && !Number.isNaN(input.valueOf())) {
    return `${input.getFullYear()}-${String(input.getMonth() + 1).padStart(2, "0")}`;
  }

  if (typeof input === "number" && Number.isFinite(input)) {
    return null;
  }

  const text = String(input ?? "").trim();
  const match = text.match(/^(\d{4})\s*[-/]\s*(\d{1,2})$/);
  if (!match) {
    return null;
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  if (month < 1 || month > 12) {
    return null;
  }

  return `${year}-${String(month).padStart(2, "0")}`;
}

export function displayMonth(month: MonthKey): string {
  const [year, monthNumber] = month.split("-");
  return `${year} -${monthNumber}`;
}

export function monthYear(month: MonthKey): number {
  return Number(month.slice(0, 4));
}

export function monthNumber(month: MonthKey): number {
  return Number(month.slice(5, 7));
}

export function compareMonths(a: MonthKey, b: MonthKey): number {
  return a.localeCompare(b);
}

export function addMonths(month: MonthKey, delta: number): MonthKey {
  const year = monthYear(month);
  const zeroBasedMonth = monthNumber(month) - 1;
  const date = new Date(Date.UTC(year, zeroBasedMonth + delta, 1));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

export function monthRange(start: MonthKey, end: MonthKey): MonthKey[] {
  const months: MonthKey[] = [];
  for (let cursor = start; compareMonths(cursor, end) <= 0; cursor = addMonths(cursor, 1)) {
    months.push(cursor);
  }
  return months;
}

export function reportPeriodMonths(reportMonth: MonthKey): MonthKey[] {
  const previousYear = monthYear(reportMonth) - 1;
  return monthRange(`${previousYear}-01`, reportMonth);
}

export function trailingMonths(reportMonth: MonthKey, count: number): MonthKey[] {
  const start = addMonths(reportMonth, -(count - 1));
  return monthRange(start, reportMonth);
}

export function sameMonthPreviousYear(reportMonth: MonthKey): MonthKey {
  return `${monthYear(reportMonth) - 1}-${String(monthNumber(reportMonth)).padStart(2, "0")}`;
}

export function ytdMonths(year: number, throughMonthNumber: number): MonthKey[] {
  return Array.from({ length: throughMonthNumber }, (_, index) => {
    return `${year}-${String(index + 1).padStart(2, "0")}` as MonthKey;
  });
}
