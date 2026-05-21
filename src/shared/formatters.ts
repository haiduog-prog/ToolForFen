import { displayMonth, type MonthKey } from "../domain/month";

export function formatNumber(value: number, maximumFractionDigits = 0): string {
  return new Intl.NumberFormat("vi-VN", {
    maximumFractionDigits
  }).format(value);
}

export function formatRatio(value: number): string {
  return new Intl.NumberFormat("vi-VN", {
    style: "percent",
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  }).format(value);
}

export function formatMonth(month: MonthKey): string {
  return displayMonth(month);
}

export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
