import type ExcelJS from "exceljs";
import { type MetricRow, type QecReport, type SourceTransaction, type ValueKind } from "../../domain/entities";
import { displayMonth, type MonthKey } from "../../domain/month";

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF183B56" } } as const;
const SECTION_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F1F2" } } as const;
const TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } } as const;
const MONEY_FORMAT = '#,##0;[Red]-#,##0;"-"';
const QUANTITY_FORMAT = '#,##0.00;[Red]-#,##0.00;"-"';
const RATIO_FORMAT = '0.00%;[Red]-0.00%;"-"';

export async function writeReportWorkbook(report: QecReport, sourceRows: SourceTransaction[]): Promise<Blob> {
  const ExcelJSModule = await import("exceljs");
  const workbook = new ExcelJSModule.default.Workbook();
  workbook.creator = "QEC Export Builder";
  workbook.created = new Date();

  addQecWorksheet(workbook, report);
  addSkuWorksheet(workbook, report);
  addCustomerWorksheet(workbook, "SKU - Customer review", report, "money");
  addCustomerWorksheet(workbook, "SKU customer review", report, "quantity");
  addSourceWorksheet(workbook, sourceRows);

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function addQecWorksheet(workbook: ExcelJS.Workbook, report: QecReport): void {
  const sheet = workbook.addWorksheet("QEC review", {
    views: [{ state: "frozen", ySplit: 1, xSplit: 1 }]
  });

  const segmentHeader = [
    "Segment",
    ...report.periodMonths.map(displayMonth),
    "",
    `Share ${report.previousYear}`,
    `Share ${report.currentYear}`,
    `Total ${report.previousYear}`,
    `Total ${report.currentYear}`,
    "P3M",
    "P6M",
    "P9M",
    "TREND",
    "IFYTD",
    "ICYTD",
    "IYA"
  ];
  sheet.addRow(segmentHeader);
  styleHeader(sheet.getRow(1));

  for (const row of report.qecRows) {
    const excelRow = sheet.addRow([
      row.label,
      ...report.periodMonths.map((month) => row.monthValues[month] ?? 0),
      "",
      row.previousYearShare ?? "",
      row.currentYearShare ?? "",
      row.previousYearTotal,
      row.currentYearTotal,
      row.p3m,
      row.p6m,
      row.p9m,
      row.trend,
      row.ifytd,
      row.icytd,
      row.iya
    ]);

    if (row.label === "Total") {
      styleTotal(excelRow);
    }
  }

  // --- DSR table (stacked below Segment, separated by 1 blank row) ---
  sheet.addRow([]); // blank separator

  const dsrHeader = [
    "DSR",
    ...report.periodMonths.map(displayMonth),
    "",
    `Total ${report.previousYear}`,
    `Total ${report.currentYear}`,
    "P3M",
    "P6M",
    "P9M",
    "TREND",
    "IFYTD",
    "ICYTD",
    "IYA"
  ];
  const dsrHeaderRow = sheet.addRow(dsrHeader);
  styleHeader(dsrHeaderRow);

  for (const row of report.dsrRows) {
    const excelRow = sheet.addRow([
      row.label,
      ...report.periodMonths.map((month) => row.monthValues[month] ?? 0),
      "",
      row.previousYearTotal,
      row.currentYearTotal,
      row.p3m,
      row.p6m,
      row.p9m,
      row.trend,
      row.ifytd,
      row.icytd,
      row.iya
    ]);

    if (row.label === "Total") {
      styleTotal(excelRow);
    }
  }

  applyMetricFormats(sheet, report.periodMonths.length, "money", true);
  autoFit(sheet);
}

function addSkuWorksheet(workbook: ExcelJS.Workbook, report: QecReport): void {
  const sheet = workbook.addWorksheet("SKU review", {
    views: [{ state: "frozen", ySplit: 2, xSplit: 2 }]
  });

  // Row 1: column index numbers (matching chuan.xlsx structure)
  const indexRow: (number | string)[] = [
    "",
    1,
    ...report.periodMonths.map((_, i) => i + 3),
    "",
    report.periodMonths.length + 4,
    report.periodMonths.length + 5
  ];
  sheet.addRow(indexRow);

  // Row 2: actual header labels
  sheet.addRow([
    "Code",
    "BRAND_OF_PRODUCT (PCS)",
    ...report.periodMonths.map(displayMonth),
    "",
    report.previousYear,
    report.currentYear,
    "P3M",
    "P6M",
    "P9M",
    "TREND",
    "IFYTD",
    "ICYTD",
    "IYA"
  ]);
  styleHeader(sheet.getRow(2));

  report.skuRevenueRows.forEach((row, index) => {
    sheet.addRow([
      index + 1,
      row.label,
      ...report.periodMonths.map((month) => row.monthValues[month] ?? 0),
      "",
      row.previousYearTotal,
      row.currentYearTotal,
      row.p3m,
      row.p6m,
      row.p9m,
      row.trend,
      row.ifytd,
      row.icytd,
      row.iya
    ]);
  });

  applyMetricFormats(sheet, report.periodMonths.length, "money", false, 3);
  autoFit(sheet);
}

function addCustomerWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  report: QecReport,
  valueKind: ValueKind
): void {
  const sections = valueKind === "money" ? report.customerRevenueSections : report.customerQuantitySections;
  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 1, xSplit: 2 }]
  });
  const header = [
    "Customer",
    "Name SKU",
    ...report.periodMonths.map(displayMonth),
    "",
    report.previousYear,
    report.currentYear,
    "P3M",
    "P6M",
    "P9M",
    "TREND",
    "IFYTD",
    "ICYTD",
    "IYA"
  ];
  sheet.addRow(header);
  styleHeader(sheet.getRow(1));

  for (const section of sections) {
    const sectionRow = sheet.addRow([section.customer]);
    sheet.mergeCells(sectionRow.number, 1, sectionRow.number, header.length);
    styleSection(sectionRow);

    for (const row of section.rows) {
      sheet.addRow([
        section.customer,
        row.label,
        ...report.periodMonths.map((month) => row.monthValues[month] ?? 0),
        "",
        row.previousYearTotal,
        row.currentYearTotal,
        row.p3m,
        row.p6m,
        row.p9m,
        row.trend,
        row.ifytd,
        row.icytd,
        row.iya
      ]);
    }

    const totalRow = sheet.addRow([
      section.customer,
      section.total.label,
      ...report.periodMonths.map((month) => section.total.monthValues[month] ?? 0),
      "",
      section.total.previousYearTotal,
      section.total.currentYearTotal,
      section.total.p3m,
      section.total.p6m,
      section.total.p9m,
      section.total.trend,
      section.total.ifytd,
      section.total.icytd,
      section.total.iya
    ]);
    styleTotal(totalRow);
  }

  applyMetricFormats(sheet, report.periodMonths.length, valueKind, false, 3);
  autoFit(sheet);
}

function addSourceWorksheet(workbook: ExcelJS.Workbook, sourceRows: SourceTransaction[]): void {
  const sheet = workbook.addWorksheet("Data nguồn", {
    views: [{ state: "frozen", ySplit: 1 }]
  });

  sheet.addRow([
    "Row",
    "Ngày",
    "Tháng",
    "Nhà Thuốc",
    "Segment",
    "Tên sản phẩm",
    "Đơn giá",
    "Số lượng (lọ)",
    "Thành tiền",
    "DOANH THU"
  ]);
  styleHeader(sheet.getRow(1));

  for (const transaction of sourceRows) {
    sheet.addRow([
      transaction.rowNumber,
      transaction.date ?? "",
      displayMonth(transaction.month),
      transaction.customer,
      transaction.segment,
      transaction.product,
      transaction.unitPrice,
      transaction.quantity,
      transaction.amount,
      transaction.revenue
    ]);
  }

  sheet.getColumn(2).numFmt = "dd/mm/yyyy";
  sheet.getColumn(7).numFmt = MONEY_FORMAT;
  sheet.getColumn(8).numFmt = QUANTITY_FORMAT;
  sheet.getColumn(9).numFmt = MONEY_FORMAT;
  sheet.getColumn(10).numFmt = MONEY_FORMAT;
  autoFit(sheet);
}

function applyMetricFormats(
  sheet: ExcelJS.Worksheet,
  monthCount: number,
  valueKind: ValueKind,
  includesShareColumns: boolean,
  monthStartColumn = 2
): void {
  const valueFormat = valueKind === "money" ? MONEY_FORMAT : QUANTITY_FORMAT;
  const monthEndColumn = monthStartColumn + monthCount - 1;
  const afterSpacerColumn = monthEndColumn + 2;
  const valueColumnsBeforeRatios = includesShareColumns ? 7 : 5;
  const firstRatioColumn = afterSpacerColumn + valueColumnsBeforeRatios;

  for (let column = monthStartColumn; column <= monthEndColumn; column += 1) {
    sheet.getColumn(column).numFmt = valueFormat;
  }

  if (includesShareColumns) {
    sheet.getColumn(afterSpacerColumn).numFmt = RATIO_FORMAT;
    sheet.getColumn(afterSpacerColumn + 1).numFmt = RATIO_FORMAT;
    sheet.getColumn(afterSpacerColumn + 2).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 3).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 4).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 5).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 6).numFmt = valueFormat;
  } else {
    sheet.getColumn(afterSpacerColumn).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 1).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 2).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 3).numFmt = valueFormat;
    sheet.getColumn(afterSpacerColumn + 4).numFmt = valueFormat;
  }

  for (let column = firstRatioColumn; column <= firstRatioColumn + 3; column += 1) {
    sheet.getColumn(column).numFmt = RATIO_FORMAT;
  }
}

function styleHeader(row: ExcelJS.Row): void {
  row.height = 24;
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = thinBorder();
  });
}

function styleSection(row: ExcelJS.Row): void {
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FF183B56" } };
    cell.fill = SECTION_FILL;
    cell.alignment = { vertical: "middle" };
    cell.border = thinBorder();
  });
}

function styleTotal(row: ExcelJS.Row): void {
  row.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = TOTAL_FILL;
    cell.border = thinBorder();
  });
}

function thinBorder(): Partial<ExcelJS.Borders> {
  return {
    top: { style: "thin", color: { argb: "FFD8DEE4" } },
    left: { style: "thin", color: { argb: "FFD8DEE4" } },
    bottom: { style: "thin", color: { argb: "FFD8DEE4" } },
    right: { style: "thin", color: { argb: "FFD8DEE4" } }
  };
}

function autoFit(sheet: ExcelJS.Worksheet): void {
  const columnCount = sheet.columnCount;
  for (let i = 1; i <= columnCount; i += 1) {
    const column = sheet.getColumn(i);
    let maxLength = 10;
    column.eachCell?.({ includeEmpty: true }, (cell) => {
      const rawValue = cell.value;
      if (rawValue == null) {
        return;
      }
      const text =
        rawValue instanceof Date
          ? "dd/mm/yyyy"
          : typeof rawValue === "object" && "richText" in rawValue
            ? (rawValue.richText as Array<{ text: string }>).map((part) => part.text).join("")
            : String(rawValue);
      maxLength = Math.max(maxLength, Math.min(text.length + 2, 48));
    });
    column.width = maxLength;
  }
}
