import type ExcelJS from "exceljs";
import { type MetricRow, type QecReport, type SourceTransaction, type ValueKind } from "../../domain/entities";
import { displayMonth, monthNumber, type MonthKey } from "../../domain/month";
import { lookupSegment } from "../../domain/customerMapping";
import { totalMetricRow } from "../../domain/reportCalculations";

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF183B56" } } as const;
const TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3CD" } } as const;
const MONEY_FORMAT = '#,##0;[Red]-#,##0;"-"';
const QUANTITY_FORMAT = '#,##0.00;[Red]-#,##0.00;"-"';
const RATIO_FORMAT = '0.00%;[Red]-0.00%;"-"';

export async function writeReportWorkbook(report: QecReport, sourceRows: SourceTransaction[]): Promise<Blob> {
  const ExcelJSModule = await import("exceljs/dist/exceljs.min.js");
  const workbook = new ExcelJSModule.default.Workbook();
  workbook.creator = "QEC Export Builder";
  workbook.created = new Date();

  // Tạo danh sách khách hàng duy nhất (đã lọc và sort)
  const rawCustomers = report.customerRevenueSections.map(s => s.customer);
  const customerNames = [...rawCustomers].sort((a, b) => a.localeCompare(b, "vi"));

  // Tạo config sheet ẩn để làm data source cho dropdown
  const configSheet = workbook.addWorksheet("Config");
  configSheet.state = "hidden";
  configSheet.getRow(1).getCell(1).value = "Customer List";
  customerNames.forEach((name, idx) => {
    configSheet.getRow(2 + idx).getCell(1).value = name;
  });

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

  const numMonths = report.periodMonths.length;
  const endCol = 7 + numMonths + 10;

  // Set standard column widths dynamically
  const colWidths = [
    12.57, 73.43, 13.43, 24.57, 26.57, 21.57 // C1-C6
  ];
  for (let i = 0; i < numMonths; i++) {
    colWidths.push(14.43); // Cột tháng
  }
  colWidths.push(21, 21); // Share
  colWidths.push(15.43, 15.43); // Total
  colWidths.push(14.43, 14.43, 14.43, 10.14, 9.43, 9.57, 6.57); // các metrics
  
  colWidths.forEach((w, i) => {
    sheet.getColumn(i + 1).width = w;
  });

  let currentRowNum = 1;

  // --- BẢNG SEGMENT (Bắt đầu từ dòng 1) ---
  const segHeaderRow = sheet.getRow(currentRowNum);
  segHeaderRow.getCell(6).value = "Segment";
  report.periodMonths.forEach((month, idx) => {
    segHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  
  segHeaderRow.getCell(7 + numMonths).value = report.previousYear;
  segHeaderRow.getCell(7 + numMonths + 1).value = report.currentYear;
  segHeaderRow.getCell(7 + numMonths + 2).value = report.previousYear;
  segHeaderRow.getCell(7 + numMonths + 3).value = report.currentYear;
  segHeaderRow.getCell(7 + numMonths + 4).value = "P3M";
  segHeaderRow.getCell(7 + numMonths + 5).value = "P6M";
  segHeaderRow.getCell(7 + numMonths + 6).value = "P9M";
  segHeaderRow.getCell(7 + numMonths + 7).value = "TREND";
  segHeaderRow.getCell(7 + numMonths + 8).value = "IFYTD";
  segHeaderRow.getCell(7 + numMonths + 9).value = "ICYTD";
  segHeaderRow.getCell(7 + numMonths + 10).value = "IYA";
  
  styleHeaderCells(segHeaderRow, 6, endCol);
  currentRowNum++;

  for (let i = 0; i < report.qecRows.length; i++) {
    const rowData = report.qecRows[i]!;
    const excelRow = sheet.getRow(currentRowNum);
    excelRow.getCell(6).value = rowData.label;
    
    const isTotalRow = rowData.label === "Total";

    report.periodMonths.forEach((month, idx) => {
      if (isTotalRow) {
        const colLetter = getColumnLetter(7 + idx);
        excelRow.getCell(7 + idx).value = {
          formula: `SUM(${colLetter}2:${colLetter}${currentRowNum - 1})`
        };
      } else {
        excelRow.getCell(7 + idx).value = rowData.monthValues[month] ?? 0;
      }
    });

    excelRow.getCell(7 + numMonths).value = rowData.previousYearShare ?? 0;
    excelRow.getCell(7 + numMonths + 1).value = rowData.currentYearShare ?? 0;

    if (isTotalRow) {
      excelRow.getCell(7 + numMonths + 2).value = {
        formula: `SUM(${getColumnLetter(7 + numMonths + 2)}2:${getColumnLetter(7 + numMonths + 2)}${currentRowNum - 1})`
      };
      excelRow.getCell(7 + numMonths + 3).value = {
        formula: `SUM(${getColumnLetter(7 + numMonths + 3)}2:${getColumnLetter(7 + numMonths + 3)}${currentRowNum - 1})`
      };
      excelRow.getCell(7 + numMonths + 4).value = {
        formula: `SUM(${getColumnLetter(7 + numMonths + 4)}2:${getColumnLetter(7 + numMonths + 4)}${currentRowNum - 1})`
      };
      excelRow.getCell(7 + numMonths + 5).value = {
        formula: `SUM(${getColumnLetter(7 + numMonths + 5)}2:${getColumnLetter(7 + numMonths + 5)}${currentRowNum - 1})`
      };
      excelRow.getCell(7 + numMonths + 6).value = {
        formula: `SUM(${getColumnLetter(7 + numMonths + 6)}2:${getColumnLetter(7 + numMonths + 6)}${currentRowNum - 1})`
      };
    } else {
      excelRow.getCell(7 + numMonths + 2).value = {
        formula: `SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + 11)}${currentRowNum})`
      };
      excelRow.getCell(7 + numMonths + 3).value = {
        formula: `SUM(${getColumnLetter(7 + 12)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
      };
      excelRow.getCell(7 + numMonths + 4).value = {
        formula: `AVERAGE(${getColumnLetter(7 + numMonths - 3)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
      };
      excelRow.getCell(7 + numMonths + 5).value = {
        formula: `AVERAGE(${getColumnLetter(7 + numMonths - 6)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
      };
      excelRow.getCell(7 + numMonths + 6).value = {
        formula: `AVERAGE(${getColumnLetter(7 + numMonths - 9)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
      };
    }

    excelRow.getCell(7 + numMonths + 7).value = {
      formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 8).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 9).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 10).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
    };

    styleDataCells(excelRow, 6, endCol);
    formatQecRow(excelRow, numMonths);

    if (isTotalRow) {
      styleTotalCells(excelRow, 6, endCol);
    }
    currentRowNum++;
  }

  // --- DÒNG INDEX TĨNH (Phía dưới bảng Segment) ---
  const r12 = sheet.getRow(currentRowNum);
  r12.getCell(6).value = 1;
  r12.getCell(7).value = 56;
  r12.getCell(8).value = 56;
  r12.getCell(7 + numMonths).value = 57;
  r12.getCell(7 + numMonths + 2).value = 62;
  r12.getCell(7 + numMonths + 4).value = 56;
  r12.getCell(7 + numMonths + 5).value = 55.25;
  r12.getCell(7 + numMonths + 6).value = 53.857142857142854;
  r12.height = 14.25;
  currentRowNum++;

  // --- BẢNG REGION ---
  const regHeaderRow = sheet.getRow(currentRowNum);
  regHeaderRow.getCell(6).value = "REGION";
  report.periodMonths.forEach((month, idx) => {
    regHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  regHeaderRow.getCell(7 + numMonths + 2).value = report.previousYear;
  regHeaderRow.getCell(7 + numMonths + 3).value = report.currentYear;
  regHeaderRow.getCell(7 + numMonths + 4).value = "P3M";
  regHeaderRow.getCell(7 + numMonths + 5).value = "P6M";
  regHeaderRow.getCell(7 + numMonths + 6).value = "P9M";
  regHeaderRow.getCell(7 + numMonths + 7).value = "TREND";
  regHeaderRow.getCell(7 + numMonths + 8).value = "IFYTD";
  regHeaderRow.getCell(7 + numMonths + 9).value = "ICYTD";
  regHeaderRow.getCell(7 + numMonths + 10).value = "IYA";
  styleHeaderCells(regHeaderRow, 6, endCol);
  currentRowNum++;

  const regionNames = [
    "1. HA NOI",
    "2. N-PROVINCE",
    "3. CENTRAL",
    "4. S-EAST",
    "5. HCM",
    "6. MKD"
  ];
  const regionStartRow = currentRowNum;
  for (let i = 0; i < regionNames.length; i++) {
    const excelRow = sheet.getRow(currentRowNum);
    excelRow.getCell(6).value = regionNames[i];
    report.periodMonths.forEach((month, idx) => {
      excelRow.getCell(7 + idx).value = 0;
    });

    excelRow.getCell(7 + numMonths + 2).value = {
      formula: `SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + 11)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 3).value = {
      formula: `SUM(${getColumnLetter(7 + 12)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 4).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 3)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 5).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 6)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 6).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 9)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 7).value = {
      formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 8).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 9).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 10).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
    };

    styleDataCells(excelRow, 6, endCol);
    formatQecRow(excelRow, numMonths);
    currentRowNum++;
  }

  // Dòng Total Region
  const regTotalRow = sheet.getRow(currentRowNum);
  regTotalRow.getCell(6).value = "Total";
  report.periodMonths.forEach((month, idx) => {
    regTotalRow.getCell(7 + idx).value = {
      formula: `SUM(${getColumnLetter(7 + idx)}${regionStartRow}:${getColumnLetter(7 + idx)}${currentRowNum - 1})`
    };
  });
  regTotalRow.getCell(7 + numMonths + 2).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 2)}${regionStartRow}:${getColumnLetter(7 + numMonths + 2)}${currentRowNum - 1})`
  };
  regTotalRow.getCell(7 + numMonths + 3).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 3)}${regionStartRow}:${getColumnLetter(7 + numMonths + 3)}${currentRowNum - 1})`
  };
  regTotalRow.getCell(7 + numMonths + 4).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 4)}${regionStartRow}:${getColumnLetter(7 + numMonths + 4)}${currentRowNum - 1})`
  };
  regTotalRow.getCell(7 + numMonths + 5).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 5)}${regionStartRow}:${getColumnLetter(7 + numMonths + 5)}${currentRowNum - 1})`
  };
  regTotalRow.getCell(7 + numMonths + 6).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 6)}${regionStartRow}:${getColumnLetter(7 + numMonths + 6)}${currentRowNum - 1})`
  };

  regTotalRow.getCell(7 + numMonths + 7).value = {
    formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
  };
  regTotalRow.getCell(7 + numMonths + 8).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };
  regTotalRow.getCell(7 + numMonths + 9).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };
  regTotalRow.getCell(7 + numMonths + 10).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
  };

  styleTotalCells(regTotalRow, 6, endCol);
  formatQecRow(regTotalRow, numMonths);
  currentRowNum++;

  // --- DÒNG ẨN ---
  const r21 = sheet.getRow(currentRowNum);
  r21.getCell(7).value = 34;
  r21.getCell(8).value = 38;
  r21.getCell(7 + numMonths + 2).value = 194;
  r21.getCell(7 + numMonths + 4).value = 38;
  r21.getCell(7 + numMonths + 5).value = 24;
  r21.getCell(7 + numMonths + 6).value = 26.285714285714285;
  r21.height = 14.25;
  r21.hidden = true;
  currentRowNum++;

  // --- BẢNG DSR ---
  const dsrHeaderRow = sheet.getRow(currentRowNum);
  dsrHeaderRow.getCell(4).value = "Mã DSR";
  dsrHeaderRow.getCell(5).value = "TOTAL STAFF";
  dsrHeaderRow.getCell(6).value = "DSR";
  report.periodMonths.forEach((month, idx) => {
    dsrHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  dsrHeaderRow.getCell(7 + numMonths + 2).value = report.previousYear;
  dsrHeaderRow.getCell(7 + numMonths + 3).value = report.currentYear;
  dsrHeaderRow.getCell(7 + numMonths + 4).value = "P3M";
  dsrHeaderRow.getCell(7 + numMonths + 5).value = "P6M";
  dsrHeaderRow.getCell(7 + numMonths + 6).value = "P9M";
  dsrHeaderRow.getCell(7 + numMonths + 7).value = "TREND";
  dsrHeaderRow.getCell(7 + numMonths + 8).value = "IFYTD";
  dsrHeaderRow.getCell(7 + numMonths + 9).value = "ICYTD";
  dsrHeaderRow.getCell(7 + numMonths + 10).value = "IYA";
  styleHeaderCells(dsrHeaderRow, 4, endCol);
  currentRowNum++;

  const dsrMeta = [
    { code: "NV1", staff: "SM" },
    { code: "NV2", staff: "AE-HCM" },
    { code: "NV3", staff: "MT-HCM" }
  ];

  const dsrStartRow = currentRowNum;
  for (let i = 0; i < report.dsrRows.length - 1; i++) {
    const rowData = report.dsrRows[i]!;
    const excelRow = sheet.getRow(currentRowNum);
    excelRow.getCell(4).value = dsrMeta[i]?.code ?? "";
    excelRow.getCell(5).value = dsrMeta[i]?.staff ?? "";
    excelRow.getCell(6).value = rowData.label;

    report.periodMonths.forEach((month, idx) => {
      excelRow.getCell(7 + idx).value = rowData.monthValues[month] ?? 0;
    });
    
    excelRow.getCell(7 + numMonths + 2).value = {
      formula: `SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + 11)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 3).value = {
      formula: `SUM(${getColumnLetter(7 + 12)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 4).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 3)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 5).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 6)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 6).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 9)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 7).value = {
      formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 8).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 9).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 10).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
    };

    styleDataCells(excelRow, 4, endCol);
    formatQecRow(excelRow, numMonths);
    currentRowNum++;
  }

  // Dòng Total DSR
  const dsrTotalRow = sheet.getRow(currentRowNum);
  dsrTotalRow.getCell(6).value = "Total";
  report.periodMonths.forEach((month, idx) => {
    dsrTotalRow.getCell(7 + idx).value = {
      formula: `SUM(${getColumnLetter(7 + idx)}${dsrStartRow}:${getColumnLetter(7 + idx)}${currentRowNum - 1})`
    };
  });
  dsrTotalRow.getCell(7 + numMonths + 2).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 2)}${dsrStartRow}:${getColumnLetter(7 + numMonths + 2)}${currentRowNum - 1})`
  };
  dsrTotalRow.getCell(7 + numMonths + 3).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 3)}${dsrStartRow}:${getColumnLetter(7 + numMonths + 3)}${currentRowNum - 1})`
  };
  dsrTotalRow.getCell(7 + numMonths + 4).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 4)}${dsrStartRow}:${getColumnLetter(7 + numMonths + 4)}${currentRowNum - 1})`
  };
  dsrTotalRow.getCell(7 + numMonths + 5).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 5)}${dsrStartRow}:${getColumnLetter(7 + numMonths + 5)}${currentRowNum - 1})`
  };
  dsrTotalRow.getCell(7 + numMonths + 6).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 6)}${dsrStartRow}:${getColumnLetter(7 + numMonths + 6)}${currentRowNum - 1})`
  };

  dsrTotalRow.getCell(7 + numMonths + 7).value = {
    formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
  };
  dsrTotalRow.getCell(7 + numMonths + 8).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };
  dsrTotalRow.getCell(7 + numMonths + 9).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };
  dsrTotalRow.getCell(7 + numMonths + 10).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
  };

  styleTotalCells(dsrTotalRow, 4, endCol);
  formatQecRow(dsrTotalRow, numMonths);
  currentRowNum++;

  // --- DÒNG TĨNH ---
  const r27 = sheet.getRow(currentRowNum);
  r27.getCell(7).value = 39;
  r27.getCell(8).value = 43;
  r27.getCell(7 + numMonths + 3).value = 0;
  r27.getCell(7 + numMonths + 9).value = 0.6267029972752044;
  r27.getCell(7 + numMonths + 10).value = 0;
  r27.height = 14.25;
  currentRowNum++;

  // --- BẢNG CUSTOMER DETAIL ---
  const custHeaderRow = sheet.getRow(currentRowNum);
  custHeaderRow.getCell(1).value = "Customer code";
  custHeaderRow.getCell(2).value = "CUSTOMER_CHANGE";
  custHeaderRow.getCell(3).value = "Segment";
  custHeaderRow.getCell(4).value = "Tên DSR";
  custHeaderRow.getCell(5).value = "CITY/PROVINCE";
  custHeaderRow.getCell(6).value = "REGION";
  report.periodMonths.forEach((month, idx) => {
    custHeaderRow.getCell(7 + idx).value = displayMonth(month);
  });
  custHeaderRow.getCell(7 + numMonths + 2).value = report.previousYear;
  custHeaderRow.getCell(7 + numMonths + 3).value = report.currentYear;
  custHeaderRow.getCell(7 + numMonths + 4).value = "P3M";
  custHeaderRow.getCell(7 + numMonths + 5).value = "P6M";
  custHeaderRow.getCell(7 + numMonths + 6).value = "P9M";
  custHeaderRow.getCell(7 + numMonths + 7).value = "TREND";
  custHeaderRow.getCell(7 + numMonths + 8).value = "IFYTD";
  custHeaderRow.getCell(7 + numMonths + 9).value = "ICYTD";
  custHeaderRow.getCell(7 + numMonths + 10).value = "IYA";
  styleHeaderCells(custHeaderRow, 1, endCol);
  currentRowNum++;

  const custStartRow = currentRowNum;
  for (let i = 0; i < report.customerBaseRows.length; i++) {
    const rowData = report.customerBaseRows[i]!;
    const excelRow = sheet.getRow(currentRowNum);
    excelRow.getCell(1).value = i + 1;
    excelRow.getCell(2).value = rowData.label;
    excelRow.getCell(3).value = lookupSegment(rowData.label);

    report.periodMonths.forEach((month, idx) => {
      excelRow.getCell(7 + idx).value = rowData.monthValues[month] ?? 0;
    });
    
    excelRow.getCell(7 + numMonths + 2).value = {
      formula: `SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + 11)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 3).value = {
      formula: `SUM(${getColumnLetter(7 + 12)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 4).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 3)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 5).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 6)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 6).value = {
      formula: `AVERAGE(${getColumnLetter(7 + numMonths - 9)}${currentRowNum}:${getColumnLetter(7 + numMonths - 1)}${currentRowNum})`
    };
    excelRow.getCell(7 + numMonths + 7).value = {
      formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 8).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 9).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };
    excelRow.getCell(7 + numMonths + 10).value = {
      formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
    };

    styleDataCells(excelRow, 1, endCol);
    formatQecRow(excelRow, numMonths);
    currentRowNum++;
  }

  // Dòng TOTAL Customer Detail
  const custTotalRow = sheet.getRow(currentRowNum);
  custTotalRow.getCell(2).value = "TOTAL";
  report.periodMonths.forEach((month, idx) => {
    custTotalRow.getCell(7 + idx).value = {
      formula: `SUM(${getColumnLetter(7 + idx)}${custStartRow}:${getColumnLetter(7 + idx)}${currentRowNum - 1})`
    };
  });
  custTotalRow.getCell(7 + numMonths + 2).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 2)}${custStartRow}:${getColumnLetter(7 + numMonths + 2)}${currentRowNum - 1})`
  };
  custTotalRow.getCell(7 + numMonths + 3).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 3)}${custStartRow}:${getColumnLetter(7 + numMonths + 3)}${currentRowNum - 1})`
  };
  custTotalRow.getCell(7 + numMonths + 4).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 4)}${custStartRow}:${getColumnLetter(7 + numMonths + 4)}${currentRowNum - 1})`
  };
  custTotalRow.getCell(7 + numMonths + 5).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 5)}${custStartRow}:${getColumnLetter(7 + numMonths + 5)}${currentRowNum - 1})`
  };
  custTotalRow.getCell(7 + numMonths + 6).value = {
    formula: `SUM(${getColumnLetter(7 + numMonths + 6)}${custStartRow}:${getColumnLetter(7 + numMonths + 6)}${currentRowNum - 1})`
  };

  custTotalRow.getCell(7 + numMonths + 7).value = {
    formula: `IFERROR((${getColumnLetter(7 + numMonths + 4)}${currentRowNum}*2)/(${getColumnLetter(7 + numMonths + 5)}${currentRowNum}+${getColumnLetter(7 + numMonths + 6)}${currentRowNum}), 0)`
  };
  custTotalRow.getCell(7 + numMonths + 8).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };
  custTotalRow.getCell(7 + numMonths + 9).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths + 3)}${currentRowNum}/SUM(${getColumnLetter(7)}${currentRowNum}:${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };
  custTotalRow.getCell(7 + numMonths + 10).value = {
    formula: `IFERROR(${getColumnLetter(7 + numMonths - 1)}${currentRowNum}/${getColumnLetter(7 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
  };

  styleTotalCells(custTotalRow, 1, endCol);
  formatQecRow(custTotalRow, numMonths);
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

  let currentRowNum = 1;

  // --- DÒNG 1 (Index Tĩnh cho PCS) ---
  writeSkuIndexRow(sheet, currentRowNum, true);
  currentRowNum++;

  // --- HEADER BẢNG PCS (Dòng 2) ---
  const headerPcs = sheet.getRow(currentRowNum);
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
  currentRowNum++;

  // --- DỮ LIỆU BẢNG PCS ---
  report.skuQuantityRows.forEach((row, idx) => {
    const excelRow = sheet.getRow(currentRowNum);
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
    currentRowNum++;
  });

  // --- GRAND TOTAL BẢNG PCS ---
  const pcsTotalExcelRow = sheet.getRow(currentRowNum);
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
  currentRowNum++;

  // --- Index Tĩnh cho VND ---
  writeSkuIndexRow(sheet, currentRowNum, false);
  currentRowNum++;

  // --- HEADER BẢNG VND ---
  const headerVnd = sheet.getRow(currentRowNum);
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
  currentRowNum++;

  // --- DỮ LIỆU BẢNG VND ---
  report.skuRevenueRows.forEach((row, idx) => {
    const excelRow = sheet.getRow(currentRowNum);
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
    currentRowNum++;
  });

  // --- TOTAL BẢNG VND ---
  const vndTotalExcelRow = sheet.getRow(currentRowNum);
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

function getColumnLetter(colNumber: number): string {
  let temp = colNumber;
  let letter = "";
  while (temp > 0) {
    const modulo = (temp - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    temp = Math.floor((temp - modulo) / 26);
  }
  return letter;
}

function addCustomerWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  report: QecReport,
  valueKind: ValueKind
): void {
  const rawCustomers = report.customerRevenueSections.map((s) => s.customer);
  const customerNames = [...rawCustomers].sort((a, b) => a.localeCompare(b, "vi"));
  const products = valueKind === "money" ? report.skuRevenueRows : report.skuQuantityRows;
  const productNames = products.map((p) => p.label);
  const valueFormat = valueKind === "money" ? MONEY_FORMAT : QUANTITY_FORMAT;
  const numMonths = report.periodMonths.length;

  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 3, xSplit: 4 }]
  });

  // Set widths dynamically
  sheet.getColumn(1).width = 15;
  sheet.getColumn(2).width = 50;
  sheet.getColumn(3).width = 5;
  sheet.getColumn(4).width = 50;
  const monthEnd = 5 + numMonths - 1;
  for (let c = 5; c <= monthEnd; c++) sheet.getColumn(c).width = 14;
  
  const spacerCol = 5 + numMonths;
  sheet.getColumn(spacerCol).width = 5;
  
  const indexStart = 5 + numMonths + 1;
  sheet.getColumn(indexStart).width = 16;     // Previous Year Total
  sheet.getColumn(indexStart + 1).width = 16; // Current Year Total
  for (let c = indexStart + 2; c <= indexStart + 8; c++) {
    sheet.getColumn(c).width = 12; // P3M..IYA
  }

  // --- DÒNG 1 (Customer Dropdown) ---
  const r1 = sheet.getRow(1);
  r1.getCell(1).value = "Customer";
  r1.getCell(2).value = customerNames[0] ?? "";

  r1.getCell(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  r1.getCell(1).fill = HEADER_FILL;
  r1.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
  r1.getCell(1).border = thinBorder();

  r1.getCell(2).font = { bold: true, color: { argb: "FF183B56" } };
  r1.getCell(2).alignment = { vertical: "middle", horizontal: "left" };
  r1.getCell(2).border = thinBorder();

  if (customerNames.length > 0) {
    r1.getCell(2).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`Config!$A$2:$A$${customerNames.length + 1}`]
    };
  }
  r1.height = 24;

  // --- DÒNG 2 (Dòng trống) ---
  sheet.getRow(2).height = 15;

  // --- DÒNG 3 (Header Bảng) ---
  const headerRow = sheet.getRow(3);
  headerRow.getCell(4).value = "Name SKU";
  report.periodMonths.forEach((month, idx) => {
    headerRow.getCell(5 + idx).value = displayMonth(month);
  });
  
  headerRow.getCell(indexStart).value = report.previousYear;
  headerRow.getCell(indexStart + 1).value = report.currentYear;
  headerRow.getCell(indexStart + 2).value = "P3M";
  headerRow.getCell(indexStart + 3).value = "P6M";
  headerRow.getCell(indexStart + 4).value = "P9M";
  headerRow.getCell(indexStart + 5).value = "TREND";
  headerRow.getCell(indexStart + 6).value = "IFYTD";
  headerRow.getCell(indexStart + 7).value = "ICYTD";
  headerRow.getCell(indexStart + 8).value = "IYA";

  const colsToStyle = [4];
  for (let c = 5; c <= monthEnd; c++) colsToStyle.push(c);
  for (let c = indexStart; c <= indexStart + 8; c++) colsToStyle.push(c);

  for (const c of colsToStyle) {
    const cell = headerRow.getCell(c);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = thinBorder();
  }
  headerRow.height = 24;

  let currentRowNum = 4;

  // --- CÁC DÒNG SKU DÙNG CÔNG THỨC SUMIFS ---
  productNames.forEach((prodName) => {
    const excelRow = sheet.getRow(currentRowNum);
    excelRow.getCell(4).value = prodName;

    report.periodMonths.forEach((month, mIdx) => {
      const colLetter = getColumnLetter(5 + mIdx);
      const dataCol = valueKind === "money" ? "J" : "H";
      excelRow.getCell(5 + mIdx).value = {
        formula: `SUMIFS('Data nguồn'!${dataCol}:${dataCol}, 'Data nguồn'!D:D, $B$1, 'Data nguồn'!F:F, $D${currentRowNum}, 'Data nguồn'!C:C, ${colLetter}$3)`
      };
    });

    // Cột CY Y-1
    excelRow.getCell(indexStart).value = {
      formula: `SUM(${getColumnLetter(5)}${currentRowNum}:${getColumnLetter(5 + 11)}${currentRowNum})`
    };

    // Cột CY Y
    excelRow.getCell(indexStart + 1).value = {
      formula: `SUM(${getColumnLetter(5 + 12)}${currentRowNum}:${getColumnLetter(monthEnd)}${currentRowNum})`
    };

    // Cột P3M
    excelRow.getCell(indexStart + 2).value = {
      formula: `AVERAGE(${getColumnLetter(monthEnd - 2)}${currentRowNum}:${getColumnLetter(monthEnd)}${currentRowNum})`
    };

    // Cột P6M
    excelRow.getCell(indexStart + 3).value = {
      formula: `AVERAGE(${getColumnLetter(monthEnd - 5)}${currentRowNum}:${getColumnLetter(monthEnd)}${currentRowNum})`
    };

    // Cột P9M
    excelRow.getCell(indexStart + 4).value = {
      formula: `AVERAGE(${getColumnLetter(monthEnd - 8)}${currentRowNum}:${getColumnLetter(monthEnd)}${currentRowNum})`
    };

    // Cột TREND
    const p3mCol = getColumnLetter(indexStart + 2);
    const p6mCol = getColumnLetter(indexStart + 3);
    const p9mCol = getColumnLetter(indexStart + 4);
    excelRow.getCell(indexStart + 5).value = {
      formula: `IFERROR((${p3mCol}${currentRowNum}*2)/(${p6mCol}${currentRowNum}+${p9mCol}${currentRowNum}), 0)`
    };

    // Cột IFYTD
    const cyYCol = getColumnLetter(indexStart + 1);
    excelRow.getCell(indexStart + 6).value = {
      formula: `IFERROR(${cyYCol}${currentRowNum}/SUM(${getColumnLetter(5)}${currentRowNum}:${getColumnLetter(5 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };

    // Cột ICYTD
    excelRow.getCell(indexStart + 7).value = {
      formula: `IFERROR(${cyYCol}${currentRowNum}/SUM(${getColumnLetter(5)}${currentRowNum}:${getColumnLetter(5 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
    };

    // Cột IYA
    excelRow.getCell(indexStart + 8).value = {
      formula: `IFERROR(${getColumnLetter(monthEnd)}${currentRowNum}/${getColumnLetter(5 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
    };

    styleCustomerDataCells(excelRow, numMonths);
    formatCustomerProductRow(excelRow, valueFormat, numMonths);
    excelRow.height = 20;
    currentRowNum++;
  });

  // --- DÒNG GRAND TOTAL ---
  const totalRow = sheet.getRow(currentRowNum);
  totalRow.getCell(4).value = "TOTAL";

  report.periodMonths.forEach((month, mIdx) => {
    const colLetter = getColumnLetter(5 + mIdx);
    totalRow.getCell(5 + mIdx).value = {
      formula: `SUM(${colLetter}4:${colLetter}${currentRowNum - 1})`
    };
  });

  // Cột CY Y-1
  totalRow.getCell(indexStart).value = {
    formula: `SUM(${getColumnLetter(indexStart)}4:${getColumnLetter(indexStart)}${currentRowNum - 1})`
  };

  // Cột CY Y
  totalRow.getCell(indexStart + 1).value = {
    formula: `SUM(${getColumnLetter(indexStart + 1)}4:${getColumnLetter(indexStart + 1)}${currentRowNum - 1})`
  };

  // Cột P3M
  totalRow.getCell(indexStart + 2).value = {
    formula: `SUM(${getColumnLetter(indexStart + 2)}4:${getColumnLetter(indexStart + 2)}${currentRowNum - 1})`
  };

  // Cột P6M
  totalRow.getCell(indexStart + 3).value = {
    formula: `SUM(${getColumnLetter(indexStart + 3)}4:${getColumnLetter(indexStart + 3)}${currentRowNum - 1})`
  };

  // Cột P9M
  totalRow.getCell(indexStart + 4).value = {
    formula: `SUM(${getColumnLetter(indexStart + 4)}4:${getColumnLetter(indexStart + 4)}${currentRowNum - 1})`
  };

  // Cột TREND
  const p3mColTotal = getColumnLetter(indexStart + 2);
  const p6mColTotal = getColumnLetter(indexStart + 3);
  const p9mColTotal = getColumnLetter(indexStart + 4);
  totalRow.getCell(indexStart + 5).value = {
    formula: `IFERROR((${p3mColTotal}${currentRowNum}*2)/(${p6mColTotal}${currentRowNum}+${p9mColTotal}${currentRowNum}), 0)`
  };

  // Cột IFYTD
  const cyYColTotal = getColumnLetter(indexStart + 1);
  totalRow.getCell(indexStart + 6).value = {
    formula: `IFERROR(${cyYColTotal}${currentRowNum}/SUM(${getColumnLetter(5)}${currentRowNum}:${getColumnLetter(5 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };

  // Cột ICYTD
  totalRow.getCell(indexStart + 7).value = {
    formula: `IFERROR(${cyYColTotal}${currentRowNum}/SUM(${getColumnLetter(5)}${currentRowNum}:${getColumnLetter(5 + monthNumber(report.reportMonth) - 1)}${currentRowNum}), 0)`
  };

  // Cột IYA
  totalRow.getCell(indexStart + 8).value = {
    formula: `IFERROR(${getColumnLetter(monthEnd)}${currentRowNum}/${getColumnLetter(5 + monthNumber(report.reportMonth) - 1)}${currentRowNum}, 0)`
  };

  for (const c of colsToStyle) {
    const cell = totalRow.getCell(c);
    cell.font = { bold: true };
    cell.fill = TOTAL_FILL;
    cell.border = thinBorder();
  }
  formatCustomerProductRow(totalRow, valueFormat, numMonths);
  totalRow.height = 22;
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

function styleCustomerDataCells(row: ExcelJS.Row, numMonths: number): void {
  const monthEnd = 5 + numMonths - 1;
  const indexStart = 5 + numMonths + 1;
  const indexEnd = 5 + numMonths + 9;

  row.getCell(4).border = thinBorder();
  for (let c = 5; c <= monthEnd; c++) {
    row.getCell(c).border = thinBorder();
  }
  for (let c = indexStart; c <= indexEnd; c++) {
    row.getCell(c).border = thinBorder();
  }
}

function styleCustomerHeader(row: ExcelJS.Row, numMonths: number): void {
  row.height = 24;
  const monthEnd = 5 + numMonths - 1;
  const indexStart = 5 + numMonths + 1;
  const indexEnd = 5 + numMonths + 9;

  const cols = [1, 2, 4];
  for (let c = 5; c <= monthEnd; c++) cols.push(c);
  for (let c = indexStart; c <= indexEnd; c++) cols.push(c);

  for (const c of cols) {
    const cell = row.getCell(c);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = thinBorder();
  }
}

function formatQecRow(row: ExcelJS.Row, numMonths: number): void {
  const lastMonthCol = 7 + numMonths - 1;
  
  // Cột tháng
  for (let c = 7; c <= lastMonthCol; c++) {
    row.getCell(c).numFmt = MONEY_FORMAT;
  }
  // Cột Share
  row.getCell(lastMonthCol + 1).numFmt = RATIO_FORMAT;
  row.getCell(lastMonthCol + 2).numFmt = RATIO_FORMAT;
  // Cột Total & P3M/P6M/P9M
  for (let c = lastMonthCol + 3; c <= lastMonthCol + 7; c++) {
    row.getCell(c).numFmt = MONEY_FORMAT;
  }
  // Cột metrics (TREND, IFYTD, ICYTD, IYA)
  for (let c = lastMonthCol + 8; c <= lastMonthCol + 11; c++) {
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

function formatCustomerProductRow(row: ExcelJS.Row, valueFormat: string, numMonths: number): void {
  const monthEnd = 5 + numMonths - 1;
  const indexStart = 5 + numMonths + 1;

  // Cột tháng
  for (let c = 5; c <= monthEnd; c++) {
    row.getCell(c).numFmt = valueFormat;
  }
  // Cột Total & P3M/P6M/P9M
  for (let c = indexStart; c <= indexStart + 4; c++) {
    row.getCell(c).numFmt = valueFormat;
  }
  // Cột metrics (TREND, IFYTD, ICYTD, IYA)
  for (let c = indexStart + 5; c <= indexStart + 8; c++) {
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
