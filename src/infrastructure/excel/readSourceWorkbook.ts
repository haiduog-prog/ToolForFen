import * as xlsx from "xlsx";
import { UNMAPPED_SEGMENT, type SourceParseResult, type SourceTransaction } from "../../domain/entities";
import { compareMonths, type MonthKey, parseMonthKey, displayMonth } from "../../domain/month";
import { lookupDsr, lookupSegment, normalizeCustomerName } from "../../domain/customerMapping";

interface HeaderMap {
  headerRow: number; // 0-based row index in sheet rows array
  monthColumn: number | null;
  customerColumn: number;
  productColumn: number;
  unitPriceColumn: number | null;
  quantityColumn: number;
  amountColumn: number | null;
  revenueColumn: number;
  segmentColumn: number | null;
  dateColumn: number | null;
  dsrColumn: number | null;
}

export function normalizeHeaderName(value: unknown): string {
  if (value == null) return "";
  return String(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ")
    .trim();
}

export async function readSourceWorkbook(file: File): Promise<SourceParseResult> {
  const arrayBuffer = await file.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook = xlsx.read(data, { type: "array", cellDates: true });

  // 1. Detect sheet: Try "Data nguồn" first, fallback to first worksheet
  const sheetName = workbook.SheetNames.includes("Data nguồn")
    ? "Data nguồn"
    : workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("Không tìm thấy sheet nào trong file Excel.");
  }

  const worksheet = workbook.Sheets[sheetName]!;
  // Read all cells as a 2D array of values
  const rows = xlsx.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "" });

  const headerMap = detectHeaderMap(rows);
  const transactions: SourceTransaction[] = [];
  let skippedRows = 0;

  for (let i = headerMap.headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    // Skip empty lines
    const hasValue = row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== "");
    if (!hasValue) continue;

    const customerVal = row[headerMap.customerColumn];
    const productVal = row[headerMap.productColumn];

    const customer = customerVal != null ? String(customerVal).trim() : "";
    const product = productVal != null ? String(productVal).trim() : "";
    const month = getRowMonth(row, headerMap);

    if (!month || !customer || !product) {
      skippedRows += 1;
      continue;
    }

    const rawSegment =
      headerMap.segmentColumn === null ? "" : String(row[headerMap.segmentColumn] ?? "").trim();

    const rawDsr =
      headerMap.dsrColumn === null ? "" : String(row[headerMap.dsrColumn] ?? "").trim();

    const normalizedCustomer = normalizeCustomerName(customer);
    const segment = rawSegment || lookupSegment(normalizedCustomer);
    const dsr = lookupDsr(rawDsr);

    const dateVal = headerMap.dateColumn === null ? null : row[headerMap.dateColumn];
    const date = dateVal != null ? parseDateValue(dateVal) : null;

    const unitPrice = headerMap.unitPriceColumn === null ? 0 : parseNumber(row[headerMap.unitPriceColumn]);
    const quantity = parseNumber(row[headerMap.quantityColumn]);
    const amount = headerMap.amountColumn === null ? 0 : parseNumber(row[headerMap.amountColumn]);
    const revenue = parseNumber(row[headerMap.revenueColumn]);

    transactions.push({
      rowNumber: i + 1, // 1-based row number in spreadsheet
      date,
      month,
      customer: normalizedCustomer,
      product,
      segment,
      dsr,
      unitPrice,
      quantity,
      amount: headerMap.amountColumn === null ? revenue : amount,
      revenue
    });
  }

  const availableMonths = Array.from(new Set(transactions.map((t) => t.month))).sort(compareMonths);
  const warnings: string[] = [];

  if (availableMonths.length === 0) {
    warnings.push("Không nhận diện được tháng nào trong Data nguồn.");
  }

  if (headerMap.segmentColumn === null) {
    warnings.push("Data nguồn không có cột Segment/Phân khúc; toàn bộ khách hàng sẽ được gom vào Unmapped.");
  }

  if (skippedRows > 0) {
    warnings.push(`${skippedRows} dòng bị bỏ qua vì thiếu tháng, nhà thuốc hoặc tên sản phẩm.`);
  }

  return {
    fileName: file.name,
    transactions,
    availableMonths,
    warnings,
    skippedRows,
    segmentColumnFound: headerMap.segmentColumn !== null
  };
}

function detectHeaderMap(rows: any[][]): HeaderMap {
  const maxRows = Math.min(20, rows.length);

  for (let r = 0; r < maxRows; r++) {
    const row = rows[r];
    if (!row) continue;

    const headers = row.map((cell) => normalizeHeaderName(cell));

    const customerColumn = findBestColumn(headers, ["nha thuoc", "khach hang", "customer"]);
    const productColumn = findBestColumn(headers, ["ten san pham", "ten hang", "product", "product name"]);
    const quantityColumn = findBestColumn(headers, ["sl", "so luong", "quantity"]);
    const revenueColumn = findBestColumn(headers, ["doanh thu chi tiet", "doanh thu", "revenue", "doanh thu thuan"]);

    if (customerColumn !== null && productColumn !== null && quantityColumn !== null && revenueColumn !== null) {
      const monthColumn = findBestColumn(headers, ["thang", "month"]);
      const dateColumn = findBestColumn(headers, ["thoi gian", "ngay", "date"]);

      return {
        headerRow: r,
        monthColumn,
        customerColumn,
        productColumn,
        unitPriceColumn: findBestColumn(headers, ["don gia", "unit price"]),
        quantityColumn,
        amountColumn: findBestColumn(headers, ["gia tri niem yet chi tiet", "thanh tien", "amount"]),
        revenueColumn,
        segmentColumn: findBestColumn(headers, ["segment", "phan khuc", "kenh"]),
        dateColumn,
        dsrColumn: findBestColumn(headers, ["nguoi thuc hien", "dsr", "nhan vien"])
      };
    }
  }

  throw new Error("Không nhận diện được header của Data nguồn. Cần có Nhà Thuốc/Khách hàng, Tên sản phẩm/Tên hàng, Số lượng/SL và DOANH THU.");
}

function findBestColumn(headers: string[], priorityList: string[]): number | null {
  // 1. Try to match exact normalized headers first
  for (const key of priorityList) {
    const idx = headers.indexOf(key);
    if (idx !== -1) return idx;
  }
  // 2. Fallback to partial matching
  for (const key of priorityList) {
    const idx = headers.findIndex((h) => h.includes(key));
    if (idx !== -1) return idx;
  }
  return null;
}

function parseDateValue(value: any): Date | null {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  // 1. Handle number or string containing a pure Excel serial number
  const numVal = Number(value);
  if (value !== "" && value !== null && value !== undefined && !isNaN(numVal) && isFinite(numVal) && numVal > 20000 && numVal < 100000) {
    const epoch = new Date(1899, 11, 30);
    const date = new Date(epoch.getTime() + numVal * 86400 * 1000);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }

  if (value != null && String(value).trim()) {
    const text = String(value).trim();
    
    // 2. Try direct parsing
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }

    // 3. Try match DD/MM/YYYY or DD-MM-YYYY (with optional 2-digit or 4-digit year, and flexible time hours:mins:secs)
    const matchDMY = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (matchDMY) {
      const d = Number(matchDMY[1]);
      const m = Number(matchDMY[2]) - 1; // 0-based month
      let y = Number(matchDMY[3]);
      if (y < 100) {
        y += y < 50 ? 2000 : 1900;
      }
      const hr = matchDMY[4] ? Number(matchDMY[4]) : 0;
      const min = matchDMY[5] ? Number(matchDMY[5]) : 0;
      const sec = matchDMY[6] ? Number(matchDMY[6]) : 0;
      const date = new Date(y, m, d, hr, min, sec);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    // 4. Try match YYYY/MM/DD or YYYY-MM-DD (with flexible time hours:mins:secs)
    const matchYMD = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (matchYMD) {
      const y = Number(matchYMD[1]);
      const m = Number(matchYMD[2]) - 1;
      const d = Number(matchYMD[3]);
      const hr = matchYMD[4] ? Number(matchYMD[4]) : 0;
      const min = matchYMD[5] ? Number(matchYMD[5]) : 0;
      const sec = matchYMD[6] ? Number(matchYMD[6]) : 0;
      const date = new Date(y, m, d, hr, min, sec);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
  }
  return null;
}

function monthKeyFromDate(date: Date): MonthKey {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}` as MonthKey;
}

function getRowMonth(row: any[], headerMap: HeaderMap): MonthKey | null {
  // 1. Try dedicated monthColumn first
  if (headerMap.monthColumn !== null) {
    const val = row[headerMap.monthColumn];
    if (val != null && String(val).trim()) {
      const parsed = parseMonthKey(val);
      if (parsed) return parsed;
      
      const d = parseDateValue(val);
      if (d) return monthKeyFromDate(d);
    }
  }
  // 2. Fallback to dateColumn
  if (headerMap.dateColumn !== null) {
    const val = row[headerMap.dateColumn];
    const d = parseDateValue(val);
    if (d) return monthKeyFromDate(d);
  }
  return null;
}

function parseNumber(value: any): number {
  if (typeof value === "number" && !isNaN(value)) {
    return value;
  }
  if (value == null) {
    return 0;
  }
  const text = String(value)
    .replace(/\s/g, "")
    .replace(/,/g, "");
  const parsed = Number(text);
  return isFinite(parsed) ? parsed : 0;
}

