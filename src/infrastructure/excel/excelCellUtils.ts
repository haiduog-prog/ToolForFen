import type ExcelJS from "exceljs";
import { parseMonthKey, type MonthKey } from "../../domain/month";

export function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value == null) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === "object") {
    if ("text" in value && typeof value.text === "string") {
      return value.text;
    }

    if ("result" in value) {
      return String(value.result ?? "");
    }

    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join("");
    }

    if ("hyperlink" in value && "text" in value) {
      return String(value.text ?? "");
    }
  }

  return String(value).trim();
}

export function cellNumber(cell: ExcelJS.Cell): number {
  const value = cell.value;
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "object" && value != null && "result" in value) {
    const result = value.result;
    if (typeof result === "number" && Number.isFinite(result)) {
      return result;
    }
  }

  const text = cellText(cell)
    .replace(/\s/g, "")
    .replace(/,/g, "");
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function cellDate(cell: ExcelJS.Cell): Date | null {
  const value = cell.value;
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return value;
  }

  return null;
}

export function normalizeHeader(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

export function parseMonthFromCell(cell: ExcelJS.Cell): MonthKey | null {
  const value = cell.value;
  if (value instanceof Date) {
    return parseMonthKey(value);
  }

  return parseMonthKey(cellText(cell));
}
