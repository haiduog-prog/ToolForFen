import type ExcelJS from "exceljs";
import { type MetricRow, type QecReport, type SourceTransaction, type ValueKind } from "../../domain/entities";
import { displayMonth, type MonthKey } from "../../domain/month";
import {
  STATIC_CUSTOMERS,
  STATIC_PCS_PRODUCTS,
  STATIC_VND_PRODUCTS,
  lookupSegment
} from "../../domain/customerMapping";
import { totalMetricRow } from "../../domain/reportCalculations";

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF183B56" } } as const;
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
    views: [{ state: "frozen", ySplit: 1, xSplit: 6 }]
  });

  // Set standard column widths to match chuan.xlsx as closely as possible
  const colWidths = [
    12.57, 73.43, 13.43, 24.57, 26.57, 21.57, // C1-C6
    14.43, 14.43, 14.43, 14.43, 14.43, 14.43, 14.43, 14.43, 15, 15, 15, 14.43, 15, 15, 15, 15, // C7-C22 (16 months)
    21, 21, // C23-C24 (Share)
    15.43, 15.43, // C25-C26 (Total)
    14.43, 14.43, 14.43, 10.14, 9.43, 9.57, 6.57 // C27-C33 (metrics)
  ];
  colWidths.forEach((w, i) => {
    sheet.getColumn(i + 1).width = w;
  });

  // --- BẢNG SEGMENT (Dòng 1-11) ---
  const segHeaderRow = sheet.getRow(1);
  segHeaderRow.getCell(6).value = "Segment";
  report.periodMonths.forEach((month, idx) => {
    segHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  segHeaderRow.getCell(23).value = report.previousYear;
  segHeaderRow.getCell(24).value = report.currentYear;
  segHeaderRow.getCell(25).value = report.previousYear;
  segHeaderRow.getCell(26).value = report.currentYear;
  segHeaderRow.getCell(27).value = "P3M";
  segHeaderRow.getCell(28).value = "P6M";
  segHeaderRow.getCell(29).value = "P9M";
  segHeaderRow.getCell(30).value = "TREND";
  segHeaderRow.getCell(31).value = "IFYTD";
  segHeaderRow.getCell(32).value = "ICYTD";
  segHeaderRow.getCell(33).value = "IYA";
  styleHeaderCells(segHeaderRow, 6, 33);

  for (let i = 0; i < report.qecRows.length; i++) {
    const rowData = report.qecRows[i];
    const excelRow = sheet.getRow(2 + i);
    excelRow.getCell(6).value = rowData.label;
    report.periodMonths.forEach((month, idx) => {
      excelRow.getCell(7 + idx).value = rowData.monthValues[month] ?? 0;
    });
    excelRow.getCell(23).value = rowData.previousYearShare ?? 0;
    excelRow.getCell(24).value = rowData.currentYearShare ?? 0;
    excelRow.getCell(25).value = rowData.previousYearTotal;
    excelRow.getCell(26).value = rowData.currentYearTotal;
    excelRow.getCell(27).value = rowData.p3m;
    excelRow.getCell(28).value = rowData.p6m;
    excelRow.getCell(29).value = rowData.p9m;
    excelRow.getCell(30).value = rowData.trend;
    excelRow.getCell(31).value = rowData.ifytd;
    excelRow.getCell(32).value = rowData.icytd;
    excelRow.getCell(33).value = rowData.iya;

    styleDataCells(excelRow, 6, 33);
    formatQecRow(excelRow);

    if (rowData.label === "Total") {
      styleTotalCells(excelRow, 6, 33);
    }
  }

  // --- DÒNG 12 (Index Tĩnh) ---
  const r12 = sheet.getRow(12);
  r12.getCell(6).value = 1;
  r12.getCell(7).value = 56;
  r12.getCell(8).value = 56;
  r12.getCell(23).value = 57;
  r12.getCell(25).value = 62;
  r12.getCell(27).value = 56;
  r12.getCell(28).value = 55.25;
  r12.getCell(29).value = 53.857142857142854;
  r12.height = 14.25;

  // --- BẢNG REGION (Dòng 13-20) ---
  const regHeaderRow = sheet.getRow(13);
  regHeaderRow.getCell(6).value = "REGION";
  report.periodMonths.forEach((month, idx) => {
    regHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  regHeaderRow.getCell(25).value = report.previousYear;
  regHeaderRow.getCell(26).value = report.currentYear;
  regHeaderRow.getCell(27).value = "P3M";
  regHeaderRow.getCell(28).value = "P6M";
  regHeaderRow.getCell(29).value = "P9M";
  regHeaderRow.getCell(30).value = "TREND";
  regHeaderRow.getCell(31).value = "IFYTD";
  regHeaderRow.getCell(32).value = "ICYTD";
  regHeaderRow.getCell(33).value = "IYA";
  styleHeaderCells(regHeaderRow, 6, 33);

  const regionNames = [
    "1. HA NOI",
    "2. N-PROVINCE",
    "3. CENTRAL",
    "4. S-EAST",
    "5. HCM",
    "6. MKD"
  ];
  for (let i = 0; i < regionNames.length; i++) {
    const excelRow = sheet.getRow(14 + i);
    excelRow.getCell(6).value = regionNames[i];
    for (let c = 7; c <= 33; c++) {
      if (c === 23 || c === 24) continue;
      excelRow.getCell(c).value = 0;
    }
    styleDataCells(excelRow, 6, 33);
    formatQecRow(excelRow);
  }

  // Dòng Total Region (dòng 20)
  const regTotalRow = sheet.getRow(20);
  regTotalRow.getCell(6).value = "Total";
  for (let c = 7; c <= 33; c++) {
    if (c === 23 || c === 24) continue;
    regTotalRow.getCell(c).value = 0;
  }
  styleTotalCells(regTotalRow, 6, 33);
  formatQecRow(regTotalRow);

  // --- DÒNG 21 (Ẩn) ---
  const r21 = sheet.getRow(21);
  r21.getCell(7).value = 34;
  r21.getCell(8).value = 38;
  r21.getCell(25).value = 194;
  r21.getCell(27).value = 38;
  r21.getCell(28).value = 24;
  r21.getCell(29).value = 26.285714285714285;
  r21.height = 14.25;
  r21.hidden = true;

  // --- BẢNG DSR (Dòng 22-26) ---
  const dsrHeaderRow = sheet.getRow(22);
  dsrHeaderRow.getCell(4).value = "Mã DSR";
  dsrHeaderRow.getCell(5).value = "TOTAL STAFF";
  dsrHeaderRow.getCell(6).value = "DSR";
  report.periodMonths.forEach((month, idx) => {
    dsrHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  dsrHeaderRow.getCell(25).value = report.previousYear;
  dsrHeaderRow.getCell(26).value = report.currentYear;
  dsrHeaderRow.getCell(27).value = "P3M";
  dsrHeaderRow.getCell(28).value = "P6M";
  dsrHeaderRow.getCell(29).value = "P9M";
  dsrHeaderRow.getCell(30).value = "TREND";
  dsrHeaderRow.getCell(31).value = "IFYTD";
  dsrHeaderRow.getCell(32).value = "ICYTD";
  dsrHeaderRow.getCell(33).value = "IYA";
  styleHeaderCells(dsrHeaderRow, 4, 33);

  const dsrMeta = [
    { code: "NV1", staff: "SM" },
    { code: "NV2", staff: "AE-HCM" },
    { code: "NV3", staff: "MT-HCM" }
  ];

  for (let i = 0; i < report.dsrRows.length - 1; i++) {
    const rowData = report.dsrRows[i];
    const excelRow = sheet.getRow(23 + i);
    excelRow.getCell(4).value = dsrMeta[i]?.code ?? "";
    excelRow.getCell(5).value = dsrMeta[i]?.staff ?? "";
    excelRow.getCell(6).value = rowData.label;

    report.periodMonths.forEach((month, idx) => {
      excelRow.getCell(7 + idx).value = rowData.monthValues[month] ?? 0;
    });
    excelRow.getCell(25).value = rowData.previousYearTotal;
    excelRow.getCell(26).value = rowData.currentYearTotal;
    excelRow.getCell(27).value = rowData.p3m;
    excelRow.getCell(28).value = rowData.p6m;
    excelRow.getCell(29).value = rowData.p9m;
    excelRow.getCell(30).value = rowData.trend;
    excelRow.getCell(31).value = rowData.ifytd;
    excelRow.getCell(32).value = rowData.icytd;
    excelRow.getCell(33).value = rowData.iya;

    styleDataCells(excelRow, 4, 33);
    formatQecRow(excelRow);
  }

  // Dòng Total DSR (dòng 26)
  const dsrTotalRow = sheet.getRow(26);
  const totalDsrData = report.dsrRows[report.dsrRows.length - 1]!;
  dsrTotalRow.getCell(6).value = "Total";
  report.periodMonths.forEach((month, idx) => {
    dsrTotalRow.getCell(7 + idx).value = totalDsrData.monthValues[month] ?? 0;
  });
  dsrTotalRow.getCell(25).value = totalDsrData.previousYearTotal;
  dsrTotalRow.getCell(26).value = totalDsrData.currentYearTotal;
  dsrTotalRow.getCell(27).value = totalDsrData.p3m;
  dsrTotalRow.getCell(28).value = totalDsrData.p6m;
  dsrTotalRow.getCell(29).value = totalDsrData.p9m;
  dsrTotalRow.getCell(30).value = totalDsrData.trend;
  dsrTotalRow.getCell(31).value = totalDsrData.ifytd;
  dsrTotalRow.getCell(32).value = totalDsrData.icytd;
  dsrTotalRow.getCell(33).value = totalDsrData.iya;
  styleTotalCells(dsrTotalRow, 4, 33);
  formatQecRow(dsrTotalRow);

  // --- DÒNG 27 (Tĩnh) ---
  const r27 = sheet.getRow(27);
  r27.getCell(7).value = 39;
  r27.getCell(8).value = 43;
  r27.getCell(26).value = 0;
  r27.getCell(32).value = 0.6267029972752044;
  r27.getCell(33).value = 0;
  r27.height = 14.25;

  // --- BẢNG CUSTOMER DETAIL (Dòng 28-324) ---
  const custHeaderRow = sheet.getRow(28);
  custHeaderRow.getCell(1).value = "Customer code";
  custHeaderRow.getCell(2).value = "CUSTOMER_CHANGE";
  custHeaderRow.getCell(3).value = "Segment";
  custHeaderRow.getCell(4).value = "Tên DSR";
  custHeaderRow.getCell(5).value = "CITY/PROVINCE";
  custHeaderRow.getCell(6).value = "REGION";
  report.periodMonths.forEach((month, idx) => {
    custHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  custHeaderRow.getCell(25).value = report.previousYear;
  custHeaderRow.getCell(26).value = report.currentYear;
  custHeaderRow.getCell(27).value = "P3M";
  custHeaderRow.getCell(28).value = "P6M";
  custHeaderRow.getCell(29).value = "P9M";
  custHeaderRow.getCell(30).value = "TREND";
  custHeaderRow.getCell(31).value = "IFYTD";
  custHeaderRow.getCell(32).value = "ICYTD";
  custHeaderRow.getCell(33).value = "IYA";
  styleHeaderCells(custHeaderRow, 1, 33);

  for (let i = 0; i < report.customerBaseRows.length; i++) {
    const rowData = report.customerBaseRows[i]!;
    const excelRow = sheet.getRow(29 + i);
    excelRow.getCell(2).value = rowData.label;
    excelRow.getCell(3).value = lookupSegment(rowData.label);

    report.periodMonths.forEach((month, idx) => {
      excelRow.getCell(7 + idx).value = rowData.monthValues[month] ?? 0;
    });
    excelRow.getCell(25).value = rowData.previousYearTotal;
    excelRow.getCell(26).value = rowData.currentYearTotal;
    excelRow.getCell(27).value = rowData.p3m;
    excelRow.getCell(28).value = rowData.p6m;
    excelRow.getCell(29).value = rowData.p9m;
    excelRow.getCell(30).value = rowData.trend;
    excelRow.getCell(31).value = rowData.ifytd;
    excelRow.getCell(32).value = rowData.icytd;
    excelRow.getCell(33).value = rowData.iya;

    styleDataCells(excelRow, 1, 33);
    formatQecRow(excelRow);
  }

  // Dòng TOTAL Customer Detail (dòng 324)
  const custTotalRow = sheet.getRow(324);
  const totalCustData = totalMetricRow("TOTAL", report.customerBaseRows, report.periodMonths, report.reportMonth);
  custTotalRow.getCell(2).value = "TOTAL";
  report.periodMonths.forEach((month, idx) => {
    custTotalRow.getCell(7 + idx).value = totalCustData.monthValues[month] ?? 0;
  });
  custTotalRow.getCell(25).value = totalCustData.previousYearTotal;
  custTotalRow.getCell(26).value = totalCustData.currentYearTotal;
  custTotalRow.getCell(27).value = totalCustData.p3m;
  custTotalRow.getCell(28).value = totalCustData.p6m;
  custTotalRow.getCell(29).value = totalCustData.p9m;
  custTotalRow.getCell(30).value = totalCustData.trend;
  custTotalRow.getCell(31).value = totalCustData.ifytd;
  custTotalRow.getCell(32).value = totalCustData.icytd;
  custTotalRow.getCell(33).value = totalCustData.iya;

  styleTotalCells(custTotalRow, 1, 33);
  formatQecRow(custTotalRow);
}

function addSkuWorksheet(workbook: ExcelJS.Workbook, report: QecReport): void {
  const sheet = workbook.addWorksheet("SKU review", {
    views: [{ state: "frozen", ySplit: 2, xSplit: 2 }]
  });

  // Set widths
  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 50;
  for (let c = 3; c <= 18; c++) sheet.getColumn(c).width = 14;
  sheet.getColumn(19).width = 5;
  sheet.getColumn(20).width = 14;
  sheet.getColumn(21).width = 16;
  sheet.getColumn(22).width = 16;
  for (let c = 23; c <= 30; c++) sheet.getColumn(c).width = 12;

  // --- DÒNG 1 (Index Tĩnh cho PCS) ---
  writeSkuIndexRow(sheet, 1, true);

  // --- HEADER BẢNG PCS (Dòng 2) ---
  const headerPcs = sheet.getRow(2);
  headerPcs.getCell(1).value = "Code";
  headerPcs.getCell(2).value = "BRAND_OF_PRODUCT (PCS)";
  report.periodMonths.forEach((month, idx) => {
    headerPcs.getCell(3 + idx).value = displayMonth(month);
  });
  headerPcs.getCell(21).value = report.previousYear;
  headerPcs.getCell(22).value = report.currentYear;
  headerPcs.getCell(23).value = "P3M";
  headerPcs.getCell(24).value = "P6M";
  headerPcs.getCell(25).value = "P9M";
  headerPcs.getCell(26).value = "TREND";
  headerPcs.getCell(27).value = "IFYTD";
  headerPcs.getCell(28).value = "ICYTD";
  headerPcs.getCell(29).value = "IYA";
  styleHeaderCells(headerPcs, 1, 29);

  // --- DỮ LIỆU BẢNG PCS (Dòng 3-34) ---
  report.skuQuantityRows.forEach((row, idx) => {
    const excelRow = sheet.getRow(3 + idx);
    excelRow.getCell(1).value = idx;
    excelRow.getCell(2).value = row.label;
    report.periodMonths.forEach((month, mIdx) => {
      excelRow.getCell(3 + mIdx).value = row.monthValues[month] ?? 0;
    });
    excelRow.getCell(21).value = row.previousYearTotal;
    excelRow.getCell(22).value = row.currentYearTotal;
    excelRow.getCell(23).value = row.p3m;
    excelRow.getCell(24).value = row.p6m;
    excelRow.getCell(25).value = row.p9m;
    excelRow.getCell(26).value = row.trend;
    excelRow.getCell(27).value = row.ifytd;
    excelRow.getCell(28).value = row.icytd;
    excelRow.getCell(29).value = row.iya;

    styleDataCells(excelRow, 1, 29);
    formatSkuRow(excelRow, QUANTITY_FORMAT);
  });

  // --- GRAND TOTAL BẢNG PCS (Dòng 35) ---
  const pcsTotalExcelRow = sheet.getRow(35);
  const pcsTotalData = totalMetricRow("Grand Total", report.skuQuantityRows, report.periodMonths, report.reportMonth);
  pcsTotalExcelRow.getCell(2).value = "Grand Total";
  report.periodMonths.forEach((month, idx) => {
    pcsTotalExcelRow.getCell(3 + idx).value = pcsTotalData.monthValues[month] ?? 0;
  });
  pcsTotalExcelRow.getCell(21).value = pcsTotalData.previousYearTotal;
  pcsTotalExcelRow.getCell(22).value = pcsTotalData.currentYearTotal;
  pcsTotalExcelRow.getCell(23).value = pcsTotalData.p3m;
  pcsTotalExcelRow.getCell(24).value = pcsTotalData.p6m;
  pcsTotalExcelRow.getCell(25).value = pcsTotalData.p9m;
  pcsTotalExcelRow.getCell(26).value = pcsTotalData.trend;
  pcsTotalExcelRow.getCell(27).value = pcsTotalData.ifytd;
  pcsTotalExcelRow.getCell(28).value = pcsTotalData.icytd;
  pcsTotalExcelRow.getCell(29).value = pcsTotalData.iya;

  styleTotalCells(pcsTotalExcelRow, 1, 29);
  formatSkuRow(pcsTotalExcelRow, QUANTITY_FORMAT);

  // --- DÒNG 36 (Index Tĩnh cho VND) ---
  writeSkuIndexRow(sheet, 36, false);

  // --- HEADER BẢNG VND (Dòng 37) ---
  const headerVnd = sheet.getRow(37);
  headerVnd.getCell(1).value = "Code";
  headerVnd.getCell(2).value = "BRAND_OF_PRODUCT (VND)";
  report.periodMonths.forEach((month, idx) => {
    headerVnd.getCell(3 + idx).value = displayMonth(month);
  });
  headerVnd.getCell(20).value = "2025 -01";
  headerVnd.getCell(21).value = "CY 2024";
  headerVnd.getCell(22).value = "CY 2025";
  headerVnd.getCell(23).value = "P3M";
  headerVnd.getCell(24).value = "P6M";
  headerVnd.getCell(25).value = "P9M";
  headerVnd.getCell(26).value = "TREND";
  headerVnd.getCell(27).value = "IFYTD";
  headerVnd.getCell(28).value = "ICYTD";
  headerVnd.getCell(29).value = "IYA";
  styleHeaderCells(headerVnd, 1, 29);

  // --- DỮ LIỆU BẢNG VND (Dòng 38-89) ---
  report.skuRevenueRows.forEach((row, idx) => {
    const excelRow = sheet.getRow(38 + idx);
    excelRow.getCell(1).value = idx;
    excelRow.getCell(2).value = row.label;
    report.periodMonths.forEach((month, mIdx) => {
      excelRow.getCell(3 + mIdx).value = row.monthValues[month] ?? 0;
    });
    excelRow.getCell(21).value = row.previousYearTotal;
    excelRow.getCell(22).value = row.currentYearTotal;
    excelRow.getCell(23).value = row.p3m;
    excelRow.getCell(24).value = row.p6m;
    excelRow.getCell(25).value = row.p9m;
    excelRow.getCell(26).value = row.trend;
    excelRow.getCell(27).value = row.ifytd;
    excelRow.getCell(28).value = row.icytd;
    excelRow.getCell(29).value = row.iya;

    styleDataCells(excelRow, 1, 29);
    formatSkuRow(excelRow, MONEY_FORMAT);
  });

  // --- TOTAL BẢNG VND (Dòng 90) ---
  const vndTotalExcelRow = sheet.getRow(90);
  const vndTotalData = totalMetricRow("TOTAL", report.skuRevenueRows, report.periodMonths, report.reportMonth);
  vndTotalExcelRow.getCell(2).value = "TOTAL";
  report.periodMonths.forEach((month, idx) => {
    vndTotalExcelRow.getCell(3 + idx).value = vndTotalData.monthValues[month] ?? 0;
  });
  vndTotalExcelRow.getCell(21).value = vndTotalData.previousYearTotal;
  vndTotalExcelRow.getCell(22).value = vndTotalData.currentYearTotal;
  vndTotalExcelRow.getCell(23).value = vndTotalData.p3m;
  vndTotalExcelRow.getCell(24).value = vndTotalData.p6m;
  vndTotalExcelRow.getCell(25).value = vndTotalData.p9m;
  vndTotalExcelRow.getCell(26).value = vndTotalData.trend;
  vndTotalExcelRow.getCell(27).value = vndTotalData.ifytd;
  vndTotalExcelRow.getCell(28).value = vndTotalData.icytd;
  vndTotalExcelRow.getCell(29).value = vndTotalData.iya;

  styleTotalCells(vndTotalExcelRow, 1, 29);
  formatSkuRow(vndTotalExcelRow, MONEY_FORMAT);
}

function addCustomerWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  report: QecReport,
  valueKind: ValueKind
): void {
  const rawSections = valueKind === "money" ? report.customerRevenueSections : report.customerQuantitySections;
  const staticProducts = valueKind === "money" ? STATIC_VND_PRODUCTS : STATIC_PCS_PRODUCTS;
  const valueFormat = valueKind === "money" ? MONEY_FORMAT : QUANTITY_FORMAT;

  // Sync customer blocks order with STATIC_CUSTOMERS static order
  const customerOrderMap = new Map<string, number>();
  STATIC_CUSTOMERS.forEach((name, index) => {
    customerOrderMap.set(name, index);
  });

  const sections = [...rawSections].sort((a, b) => {
    const idxA = customerOrderMap.get(a.customer);
    const idxB = customerOrderMap.get(b.customer);
    return (idxA ?? 999) - (idxB ?? 999);
  });

  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 1, xSplit: 4 }]
  });

  // Set widths
  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 50;
  sheet.getColumn(3).width = 5;
  sheet.getColumn(4).width = 50;
  for (let c = 5; c <= 20; c++) sheet.getColumn(c).width = 14;
  sheet.getColumn(21).width = 5;
  sheet.getColumn(22).width = 16;
  sheet.getColumn(23).width = 16;
  for (let c = 24; c <= 30; c++) sheet.getColumn(c).width = 12;

  // We write the blocks stacked.
  // Each block starts with a customer header, then 52 product rows, then a spacer.
  let currentRowNum = 1;

  for (const section of sections) {
    // --- 1. DÒNG HEADER CHO KHÁCH HÀNG ---
    const headerRow = sheet.getRow(currentRowNum);
    headerRow.getCell(1).value = "Customer";
    headerRow.getCell(2).value = section.customer;
    headerRow.getCell(4).value = "Name SKU";
    report.periodMonths.forEach((month, idx) => {
      headerRow.getCell(5 + idx).value = displayMonth(month);
    });
    headerRow.getCell(22).value = report.previousYear;
    headerRow.getCell(23).value = report.currentYear;
    headerRow.getCell(24).value = "P3M";
    headerRow.getCell(25).value = "P6M";
    headerRow.getCell(26).value = "P9M";
    headerRow.getCell(27).value = "TREND";
    headerRow.getCell(28).value = "IFYTD";
    headerRow.getCell(29).value = "ICYTD";
    headerRow.getCell(30).value = "IYA";

    styleCustomerHeader(headerRow);
    currentRowNum++;

    // Create a quick lookup map of actual rows for speed
    const prodMap = new Map<string, MetricRow>();
    section.rows.forEach(r => {
      prodMap.set(r.label, r);
    });

    // --- 2. 52 DÒNG SKU TĨNH ---
    staticProducts.forEach((prodName) => {
      const excelRow = sheet.getRow(currentRowNum);
      excelRow.getCell(4).value = prodName;

      const rowData = prodName ? prodMap.get(prodName) : null;
      if (rowData) {
        report.periodMonths.forEach((month, mIdx) => {
          excelRow.getCell(5 + mIdx).value = rowData.monthValues[month] ?? 0;
        });
        excelRow.getCell(22).value = rowData.previousYearTotal;
        excelRow.getCell(23).value = rowData.currentYearTotal;
        excelRow.getCell(24).value = rowData.p3m;
        excelRow.getCell(25).value = rowData.p6m;
        excelRow.getCell(26).value = rowData.p9m;
        excelRow.getCell(27).value = rowData.trend;
        excelRow.getCell(28).value = rowData.ifytd;
        excelRow.getCell(29).value = rowData.icytd;
        excelRow.getCell(30).value = rowData.iya;
      } else {
        // Fallback to zeros
        report.periodMonths.forEach((month, mIdx) => {
          excelRow.getCell(5 + mIdx).value = 0;
        });
        excelRow.getCell(22).value = 0;
        excelRow.getCell(23).value = 0;
        excelRow.getCell(24).value = 0;
        excelRow.getCell(25).value = 0;
        excelRow.getCell(26).value = 0;
        excelRow.getCell(27).value = 0;
        excelRow.getCell(28).value = 0;
        excelRow.getCell(29).value = 0;
        excelRow.getCell(30).value = 0;
      }

      styleCustomerDataCells(excelRow);
      formatCustomerProductRow(excelRow, valueFormat);
      currentRowNum++;
    });

    // --- 3. DÒNG TRỐNG SPACER ---
    sheet.getRow(currentRowNum).height = 15;
    currentRowNum++;
  }
}

function addSourceWorksheet(workbook: ExcelJS.Workbook, sourceRows: SourceTransaction[]): void {
  const sheet = workbook.addWorksheet("Data nguồn", {
    views: [{ state: "frozen", ySplit: 1 }]
  });

  const header = sheet.getRow(1);
  header.getCell(1).value = "Row";
  header.getCell(2).value = "Ngày";
  header.getCell(3).value = "Tháng";
  header.getCell(4).value = "Nhà Thuốc";
  header.getCell(5).value = "Segment";
  header.getCell(6).value = "Tên sản phẩm";
  header.getCell(7).value = "Đơn giá";
  header.getCell(8).value = "Số lượng (lọ)";
  header.getCell(9).value = "Thành tiền";
  header.getCell(10).value = "DOANH THU";
  styleHeaderCells(header, 1, 10);

  for (let i = 0; i < sourceRows.length; i++) {
    const transaction = sourceRows[i]!;
    const excelRow = sheet.getRow(2 + i);
    excelRow.getCell(1).value = transaction.rowNumber;
    excelRow.getCell(2).value = transaction.date ?? "";
    excelRow.getCell(3).value = displayMonth(transaction.month);
    excelRow.getCell(4).value = transaction.customer;
    excelRow.getCell(5).value = transaction.segment;
    excelRow.getCell(6).value = transaction.product;
    excelRow.getCell(7).value = transaction.unitPrice;
    excelRow.getCell(8).value = transaction.quantity;
    excelRow.getCell(9).value = transaction.amount;
    excelRow.getCell(10).value = transaction.revenue;

    // Apply basic thin borders to transaction rows
    for (let c = 1; c <= 10; c++) {
      excelRow.getCell(c).border = thinBorder();
    }
  }

  sheet.getColumn(2).numFmt = "dd/mm/yyyy";
  sheet.getColumn(7).numFmt = MONEY_FORMAT;
  sheet.getColumn(8).numFmt = QUANTITY_FORMAT;
  sheet.getColumn(9).numFmt = MONEY_FORMAT;
  sheet.getColumn(10).numFmt = MONEY_FORMAT;

  autoFit(sheet);
}

// === HELPERS ===

function writeSkuIndexRow(sheet: ExcelJS.Worksheet, rowNum: number, isFirst: boolean): void {
  const row = sheet.getRow(rowNum);
  row.getCell(2).value = 1;
  for (let c = 3; c <= 18; c++) {
    row.getCell(c).value = c;
  }
  if (isFirst) {
    row.getCell(20).value = 43;
    row.getCell(21).value = 47;
    for (let c = 23; c <= 29; c++) {
      row.getCell(c).value = c + 25; // 48..54
    }
    row.getCell(30).value = "Forecast 1H CY2023";
  } else {
    row.getCell(20).value = 44;
    row.getCell(21).value = { error: "#REF!" } as any;
    for (let c = 23; c <= 25; c++) {
      row.getCell(c).value = c + 26; // 49..51
    }
    row.getCell(30).value = 56;
  }
  row.height = 14.25;
}

function styleRowCells(row: ExcelJS.Row, startCol: number, endCol: number, styleFn: (cell: ExcelJS.Cell) => void): void {
  for (let c = startCol; c <= endCol; c++) {
    styleFn(row.getCell(c));
  }
}

function styleHeaderCells(row: ExcelJS.Row, startCol: number, endCol: number): void {
  row.height = 24;
  styleRowCells(row, startCol, endCol, (cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = thinBorder();
  });
}

function styleTotalCells(row: ExcelJS.Row, startCol: number, endCol: number): void {
  styleRowCells(row, startCol, endCol, (cell) => {
    cell.font = { bold: true };
    cell.fill = TOTAL_FILL;
    cell.border = thinBorder();
  });
}

function styleDataCells(row: ExcelJS.Row, startCol: number, endCol: number): void {
  styleRowCells(row, startCol, endCol, (cell) => {
    cell.border = thinBorder();
  });
}

function styleCustomerDataCells(row: ExcelJS.Row): void {
  for (const c of [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 23, 24, 25, 26, 27, 28, 29, 30]) {
    row.getCell(c).border = thinBorder();
  }
}

function styleCustomerHeader(row: ExcelJS.Row): void {
  row.height = 24;
  for (const c of [1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 23, 24, 25, 26, 27, 28, 29, 30]) {
    const cell = row.getCell(c);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = thinBorder();
  }
}

function formatQecRow(row: ExcelJS.Row): void {
  // Cột 7-22: tháng
  for (let c = 7; c <= 22; c++) {
    row.getCell(c).numFmt = MONEY_FORMAT;
  }
  // Cột 23-24: share
  row.getCell(23).numFmt = RATIO_FORMAT;
  row.getCell(24).numFmt = RATIO_FORMAT;
  // Cột 25-29: Total & P3M/P6M/P9M
  for (let c = 25; c <= 29; c++) {
    row.getCell(c).numFmt = MONEY_FORMAT;
  }
  // Cột 30-33: ratio metrics
  for (let c = 30; c <= 33; c++) {
    row.getCell(c).numFmt = RATIO_FORMAT;
  }
}

function formatSkuRow(row: ExcelJS.Row, valueFormat: string): void {
  // Cột 3-18: tháng
  for (let c = 3; c <= 18; c++) {
    row.getCell(c).numFmt = valueFormat;
  }
  // Cột 21-25: Total & P3M/P6M/P9M
  for (let c = 21; c <= 25; c++) {
    row.getCell(c).numFmt = valueFormat;
  }
  // Cột 26-29: ratio metrics
  for (let c = 26; c <= 29; c++) {
    row.getCell(c).numFmt = RATIO_FORMAT;
  }
}

function formatCustomerProductRow(row: ExcelJS.Row, valueFormat: string): void {
  // Cột 5-20: tháng
  for (let c = 5; c <= 20; c++) {
    row.getCell(c).numFmt = valueFormat;
  }
  // Cột 22-26: Total & P3M/P6M/P9M
  for (let c = 22; c <= 26; c++) {
    row.getCell(c).numFmt = valueFormat;
  }
  // Cột 27-30: ratio metrics
  for (let c = 27; c <= 30; c++) {
    row.getCell(c).numFmt = RATIO_FORMAT;
  }
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
