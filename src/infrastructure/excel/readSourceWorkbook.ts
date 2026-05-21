import type ExcelJS from "exceljs";
import { UNMAPPED_SEGMENT, type SourceParseResult, type SourceTransaction } from "../../domain/entities";
import { compareMonths, type MonthKey } from "../../domain/month";
import { lookupDsr, lookupSegment, normalizeCustomerName } from "../../domain/customerMapping";
import { cellDate, cellNumber, cellText, normalizeHeader, parseMonthFromCell } from "./excelCellUtils";

interface HeaderMap {
  headerRow: number;
  monthColumn: number;
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

export async function readSourceWorkbook(file: File): Promise<SourceParseResult> {
  const ExcelJSModule = await import("exceljs");
  const workbook = new ExcelJSModule.default.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  const worksheet = workbook.getWorksheet("Data nguồn");
  if (!worksheet) {
    throw new Error("Không tìm thấy sheet Data nguồn trong file Excel.");
  }

  const headerMap = detectHeaderMap(worksheet);
  const transactions: SourceTransaction[] = [];
  let skippedRows = 0;

  for (let rowNumber = headerMap.headerRow + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const month = parseMonthFromCell(row.getCell(headerMap.monthColumn));
    const customer = cellText(row.getCell(headerMap.customerColumn)).trim();
    const product = cellText(row.getCell(headerMap.productColumn)).trim();

    if (!month || !customer || !product) {
      skippedRows += hasAnyValue(row) ? 1 : 0;
      continue;
    }

    const rawSegment =
      headerMap.segmentColumn == null ? "" : cellText(row.getCell(headerMap.segmentColumn)).trim();

    const rawDsr =
      headerMap.dsrColumn == null ? "" : cellText(row.getCell(headerMap.dsrColumn)).trim();

    const normalizedCustomer = normalizeCustomerName(customer);
    const segment = rawSegment || lookupSegment(normalizedCustomer);
    const dsr = lookupDsr(rawDsr);

    transactions.push({
      rowNumber,
      date: headerMap.dateColumn == null ? null : cellDate(row.getCell(headerMap.dateColumn)),
      month,
      customer: normalizedCustomer,
      product,
      segment,
      dsr,
      unitPrice: headerMap.unitPriceColumn == null ? 0 : cellNumber(row.getCell(headerMap.unitPriceColumn)),
      quantity: cellNumber(row.getCell(headerMap.quantityColumn)),
      amount: headerMap.amountColumn == null ? 0 : cellNumber(row.getCell(headerMap.amountColumn)),
      revenue: cellNumber(row.getCell(headerMap.revenueColumn))
    });
  }

  const availableMonths = Array.from(new Set(transactions.map((transaction) => transaction.month))).sort(compareMonths);
  const warnings: string[] = [];

  if (availableMonths.length === 0) {
    warnings.push("Không nhận diện được tháng nào trong Data nguồn.");
  }

  if (headerMap.segmentColumn == null) {
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
    segmentColumnFound: headerMap.segmentColumn != null
  };
}

function detectHeaderMap(worksheet: ExcelJS.Worksheet): HeaderMap {
  const candidates = Array.from({ length: Math.min(20, worksheet.rowCount) }, (_, index) => index + 1);

  for (const headerRow of candidates) {
    const headers = readHeaders(worksheet.getRow(headerRow));
    const customerColumn = findColumn(headers, (header) => header.includes("nha thuoc") || header.includes("customer"));
    const productColumn = findColumn(headers, (header) => header.includes("ten san pham") || header.includes("product"));
    const revenueColumn = findColumn(headers, (header) => header.includes("doanh thu") || header.includes("revenue"));
    const quantityColumn = findColumn(headers, (header) => header.includes("so luong") || header.includes("quantity"));

    if (customerColumn && productColumn && revenueColumn && quantityColumn) {
      const monthColumn = detectMonthColumn(worksheet, headerRow, headers);
      if (!monthColumn) {
        continue;
      }

      const dateCandidates = findColumns(headers, (header) => header === "ngay" || header === "date");
      const dateColumn = dateCandidates.find((column) => column !== monthColumn) ?? null;

      return {
        headerRow,
        monthColumn,
        customerColumn,
        productColumn,
        revenueColumn,
        quantityColumn,
        dateColumn,
        unitPriceColumn: findColumn(headers, (header) => header.includes("don gia") || header.includes("unit price")),
        amountColumn: findColumn(headers, (header) => header.includes("thanh tien") || header.includes("amount")),
        segmentColumn: findColumn(
          headers,
          (header) => header.includes("segment") || header.includes("phan khuc") || header.includes("kenh")
        ),
        dsrColumn: findColumn(
          headers,
          (header) => header.includes("nguoi thuc hien") || header.includes("dsr") || header.includes("nhan vien")
        )
      };
    }
  }

  throw new Error("Không nhận diện được header của Data nguồn. Cần có Nhà Thuốc, Tên sản phẩm, Số lượng và DOANH THU.");
}

function readHeaders(row: ExcelJS.Row): Map<number, string> {
  const headers = new Map<number, string>();
  row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
    const normalized = normalizeHeader(cellText(cell));
    if (normalized) {
      headers.set(columnNumber, normalized);
    }
  });
  return headers;
}

function detectMonthColumn(
  worksheet: ExcelJS.Worksheet,
  headerRow: number,
  headers: Map<number, string>
): number | null {
  const likelyColumns = findColumns(headers, (header) => header === "ngay" || header.includes("month") || header.includes("thang"));
  const candidates = likelyColumns.length > 0 ? likelyColumns : Array.from(headers.keys());

  let bestColumn: number | null = null;
  let bestScore = 0;

  for (const column of candidates) {
    let score = 0;
    for (let rowNumber = headerRow + 1; rowNumber <= Math.min(headerRow + 50, worksheet.rowCount); rowNumber += 1) {
      const month = parseMonthFromCell(worksheet.getRow(rowNumber).getCell(column));
      if (month) {
        score += 1;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestColumn = column;
    }
  }

  return bestScore > 0 ? bestColumn : null;
}

function findColumn(headers: Map<number, string>, predicate: (header: string) => boolean): number | null {
  return findColumns(headers, predicate)[0] ?? null;
}

function findColumns(headers: Map<number, string>, predicate: (header: string) => boolean): number[] {
  return Array.from(headers.entries())
    .filter(([, header]) => predicate(header))
    .map(([column]) => column);
}

function hasAnyValue(row: ExcelJS.Row): boolean {
  let hasValue = false;
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (cellText(cell).trim() !== "") {
      hasValue = true;
    }
  });
  return hasValue;
}
