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

export function formatDate(date: Date | null | undefined): string {
  if (!date || Number.isNaN(date.getTime())) {
    return "-";
  }
  const pad = (n: number) => String(n).padStart(2, "0");
  const day = pad(date.getDate());
  const month = pad(date.getMonth() + 1);
  const year = date.getFullYear();
  const hours = pad(date.getHours());
  const minutes = pad(date.getMinutes());
  const seconds = pad(date.getSeconds());
  
  if (hours === "00" && minutes === "00" && seconds === "00") {
    return `${day}/${month}/${year}`;
  }
  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
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
