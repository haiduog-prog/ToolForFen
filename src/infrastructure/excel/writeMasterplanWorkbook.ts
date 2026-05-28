import type ExcelJS from "exceljs";
import { type MasterplanData, type TimelineBlock, type SalesPlanCustomer, CHANNEL_OPTIONS } from "../../domain/masterplan";

const TITLE_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } } as const;
const SUBHEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE7E6E6" } } as const;
const YELLOW_TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } } as const;
const GREEN_TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAD3" } } as const;
const PURPLE_QUARTER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF6A329F" } } as const;
const TEAL_YEAR_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F6B5F" } } as const;
const BLUE_TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } } as const;

const BORDER_COLOR = "FFD9D9D9";
const MONEY_FORMAT = '#,##0';

export async function writeMasterplanWorkbook(data: MasterplanData): Promise<Blob> {
  const { customers } = data;

  // Extract unique channels dynamically from customer list
  const uniqueChannelsMap = new Map<string, { code: number; channel: string; sub: string }>();
  customers.forEach((cust) => {
    const key = `${cust.codeChannel}_${cust.channel}_${cust.subChannel}`;
    if (!uniqueChannelsMap.has(key)) {
      uniqueChannelsMap.set(key, {
        code: cust.codeChannel,
        channel: cust.channel,
        sub: cust.subChannel
      });
    }
  });
  
  // Sort channels by code ascending for neat layout
  const channels = Array.from(uniqueChannelsMap.values()).sort((a, b) => a.code - b.code);
  const numChannels = channels.length;

  const channelsStartRow = 5;
  const nationwideRow = channelsStartRow + numChannels;
  const timelineHeadersStartRow = nationwideRow + 3; // 2 spacer rows
  const dataStartRow = timelineHeadersStartRow + 3; // Timeline headers take 3 rows

  const ExcelJSModule = await import("exceljs/dist/exceljs.min.js");
  const workbook = new ExcelJSModule.default.Workbook();
  workbook.creator = "Sales Plan Creator";
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  // Create sheets in order: DetailPlan_Template, Products, Mapping
  const dataEndRow = addDetailPlanSheet(workbook, data, channels, channelsStartRow, nationwideRow, timelineHeadersStartRow, dataStartRow);
  addProductsReferenceSheet(workbook, data);
  addMappingSheet(workbook, data, channelsStartRow, nationwideRow, timelineHeadersStartRow, dataStartRow, dataEndRow);

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function addDetailPlanSheet(
  workbook: ExcelJS.Workbook,
  data: MasterplanData,
  channels: Array<{ code: number; channel: string; sub: string }>,
  channelsStartRow: number,
  nationwideRow: number,
  timelineHeadersStartRow: number,
  dataStartRow: number
): number {
  const { products, customers, timelineBlocks, startMonthIndex, startYear } = data;
  const numCustomers = customers.length;
  const productCount = products.length;
  let dataEndRow = dataStartRow;

  // Set worksheet pane freezes and grid options
  const sheet = workbook.addWorksheet("DetailPlan_Template", {
    views: [
      {
        state: "frozen",
        xSplit: 9,
        ySplit: dataStartRow - 1,
        topLeftCell: `J${dataStartRow}`
      }
    ]
  });

  const lastCol = timelineBlocks[timelineBlocks.length - 1]!.endCol;
  const lastColName = getColLetter(lastCol);

  // Standard column widths
  const colWidths = [12, 20, 24, 30, 10, 18, 14, 18, 12];
  colWidths.forEach((w, i) => {
    sheet.getColumn(i + 1).width = w;
  });
  for (let c = 10; c <= lastCol; c++) {
    sheet.getColumn(c).width = 18;
  }

  const thinBorder = {
    top: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    left: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    bottom: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    right: { style: "thin" as const, color: { argb: BORDER_COLOR } }
  };

  // Row 1: Title block
  sheet.mergeCells(`A1:${lastColName}1`);
  const titleCell = sheet.getCell("A1");
  const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  titleCell.value = `SALES PLAN EXPORT TEMPLATE - START ${monthNames[startMonthIndex]} ${startYear}`;
  titleCell.font = { name: "Calibri", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
  titleCell.fill = TITLE_FILL;
  titleCell.alignment = { vertical: "middle", horizontal: "center" };
  sheet.getRow(1).height = 18;

  // Row 2: Subtitle / Rules description
  sheet.mergeCells(`A2:${lastColName}2`);
  const subtitleCell = sheet.getCell("A2");
  subtitleCell.value = "Rule: first month = current month + 1. Quarter totals are created even when the first quarter is incomplete.";
  subtitleCell.font = { name: "Calibri", size: 11, bold: true };
  subtitleCell.alignment = { vertical: "middle", horizontal: "left" };
  sheet.getRow(2).height = 18;

  // Row 3: Customer headers (Header level 1)
  const fixedHeaders = ["Code channel", "Channel", "Sub-channel", "Customer", "Stores", "Province", "Regional", "Sub-Channel 1", "STAFF"];
  const row3 = sheet.getRow(3);
  fixedHeaders.forEach((h, idx) => {
    const cell = row3.getCell(idx + 1);
    cell.value = h;
    cell.font = { name: "Calibri", size: 11, bold: true };
    cell.fill = SUBHEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = thinBorder;
  });
  sheet.getRow(3).height = 18;
  sheet.getRow(4).height = 18; // Blank spacer row 4

  // Initialize summary rows
  channels.forEach((ch, idx) => {
    const r = channelsStartRow + idx;
    const row = sheet.getRow(r);
    row.height = 18;
    row.getCell(1).value = ch.code;
    row.getCell(1).numFmt = "0";
    row.getCell(2).value = ch.channel;
    row.getCell(3).value = ch.sub;

    for (let c = 1; c <= lastCol; c++) {
      const cell = row.getCell(c);
      cell.font = { name: "Calibri", size: 11, bold: c <= 3 };
      cell.fill = SUBHEADER_FILL;
      cell.border = thinBorder;
      if (c <= 3) {
        cell.alignment = { vertical: "middle", horizontal: c === 1 ? "center" : "left" };
      }
    }
  });

  // Nationwide Row
  const row12 = sheet.getRow(nationwideRow);
  row12.height = 18;
  row12.getCell(1).value = 1000;
  row12.getCell(1).numFmt = "0";
  row12.getCell(2).value = "Nationwide";
  for (let c = 1; c <= lastCol; c++) {
    const cell = row12.getCell(c);
    cell.font = { name: "Calibri", size: 11, bold: true };
    cell.fill = SUBHEADER_FILL;
    cell.border = thinBorder;
    if (c <= 2) {
      cell.alignment = { vertical: "middle", horizontal: c === 1 ? "center" : "left" };
    }
  }

  // Row 15, 16, 17: Timeline Block Headers and Product Information
  const row15 = sheet.getRow(timelineHeadersStartRow);
  const row16 = sheet.getRow(timelineHeadersStartRow + 1);
  const row17 = sheet.getRow(timelineHeadersStartRow + 2);

  row15.height = 24;
  row16.height = 32;
  row17.height = 18;

  // Fixed labels in row 16
  fixedHeaders.forEach((h, idx) => {
    const cell = row16.getCell(idx + 1);
    cell.value = h;
    cell.font = { name: "Calibri", size: 11, bold: true };
    cell.fill = SUBHEADER_FILL;
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = thinBorder;

    const cell17 = row17.getCell(idx + 1);
    cell17.border = thinBorder;
  });

  // Populate Timeline Blocks in Row 15, 16, 17
  timelineBlocks.forEach((block) => {
    // Row 15: Block label merged across start and end columns
    sheet.mergeCells(timelineHeadersStartRow, block.startCol, timelineHeadersStartRow, block.endCol);
    const labelCell = row15.getCell(block.startCol);
    labelCell.value = block.label;
    labelCell.alignment = { vertical: "middle", horizontal: "center" };
    labelCell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
    
    // Custom label fill color based on block type
    if (block.type === "month") {
      labelCell.fill = TITLE_FILL;
    } else if (block.type === "quarter") {
      labelCell.fill = PURPLE_QUARTER_FILL;
    } else {
      labelCell.fill = TEAL_YEAR_FILL;
    }

    // Row 16 & 17: Products columns in the block
    for (let i = 0; i < productCount; i++) {
      const col = block.productCols[i]!;
      row16.getCell(col).value = products[i]!.name;
      row16.getCell(col).font = { name: "Calibri", size: 11, bold: true };
      row16.getCell(col).fill = SUBHEADER_FILL;
      row16.getCell(col).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      row16.getCell(col).border = thinBorder;

      row17.getCell(col).value = products[i]!.price;
      row17.getCell(col).numFmt = MONEY_FORMAT;
      row17.getCell(col).font = { name: "Calibri", size: 11 };
      row17.getCell(col).fill = YELLOW_TOTAL_FILL;
      row17.getCell(col).alignment = { vertical: "middle", horizontal: "center" };
      row17.getCell(col).border = thinBorder;
    }

    // Month Total columns
    if (block.type === "quarter") {
      row16.getCell(block.totalCol).value = "Quarter Total Volume";
      row16.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
      row16.getCell(block.totalCol).fill = GREEN_TOTAL_FILL;
      row16.getCell(block.totalCol).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      row16.getCell(block.totalCol).border = thinBorder;

      row17.getCell(block.totalCol).border = thinBorder;
      row17.getCell(block.totalCol).fill = YELLOW_TOTAL_FILL;
    } else {
      // Month or Year
      row16.getCell(block.totalCol).value = "Total Volume";
      row16.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
      row16.getCell(block.totalCol).fill = GREEN_TOTAL_FILL;
      row16.getCell(block.totalCol).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      row16.getCell(block.totalCol).border = thinBorder;

      row16.getCell(block.grossCol!).value = "Gross Sales\n+VAT";
      row16.getCell(block.grossCol!).font = { name: "Calibri", size: 11, bold: true };
      row16.getCell(block.grossCol!).fill = GREEN_TOTAL_FILL;
      row16.getCell(block.grossCol!).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      row16.getCell(block.grossCol!).border = thinBorder;

      row17.getCell(block.totalCol).border = thinBorder;
      row17.getCell(block.totalCol).fill = YELLOW_TOTAL_FILL;

      row17.getCell(block.grossCol!).border = thinBorder;
      row17.getCell(block.grossCol!).fill = YELLOW_TOTAL_FILL;
    }
  });

  // Sort customers in exactly the same order as channels, so all customers of the same channel are consecutive
  const sortedCustomers = [...customers].sort((a, b) => {
    if (a.codeChannel !== b.codeChannel) return a.codeChannel - b.codeChannel;
    if (a.channel !== b.channel) return a.channel.localeCompare(b.channel);
    return (a.subChannel || "").localeCompare(b.subChannel || "");
  });

  // Group sorted customers by channel key: `${cust.codeChannel}_${cust.channel}_${cust.subChannel}`
  interface CustomerGroup {
    key: string;
    codeChannel: number;
    channelName: string;
    subChannelName: string;
    items: SalesPlanCustomer[];
  }
  const groupsOfCustomers: CustomerGroup[] = [];

  sortedCustomers.forEach((cust) => {
    const key = `${cust.codeChannel}_${cust.channel}_${cust.subChannel}`;
    let group = groupsOfCustomers.find((g) => g.key === key);
    if (!group) {
      group = {
        key,
        codeChannel: cust.codeChannel,
        channelName: cust.channel,
        subChannelName: cust.subChannel,
        items: []
      };
      groupsOfCustomers.push(group);
    }
    group.items.push(cust);
  });

  let currentR = dataStartRow;

  // Populate Customer Rows & Group TOTAL Rows
  groupsOfCustomers.forEach((group) => {
    const groupStartRow = currentR;

    // Write all customer rows in this group
    group.items.forEach((cust) => {
      const r = currentR;
      const row = sheet.getRow(r);
      row.height = 18;

      row.getCell(1).value = cust.codeChannel;
      row.getCell(1).numFmt = "0";
      row.getCell(2).value = cust.channel;
      row.getCell(3).value = cust.subChannel;
      row.getCell(4).value = cust.customer;
      row.getCell(5).value = cust.stores;
      row.getCell(6).value = cust.province;
      row.getCell(7).value = cust.regional;
      row.getCell(8).value = cust.subChannel1;
      row.getCell(9).value = cust.staff;

      // Apply borders and fonts to Customer Info Columns A-I
      for (let c = 1; c <= 9; c++) {
        const cell = row.getCell(c);
        cell.font = { name: "Calibri", size: 11 };
        cell.border = thinBorder;
        cell.alignment = { vertical: "middle", horizontal: c === 1 ? "center" : "left" };
      }

      const monthBlocks = timelineBlocks.filter((b) => b.type === "month");

      // Populate timeline block formulas for customer row `r`
      timelineBlocks.forEach((block) => {
        if (block.type === "month") {
          const monthIdx = monthBlocks.indexOf(block);
          // Month Product Quantities are dynamically filled based on product threshold
          for (let i = 0; i < productCount; i++) {
            const col = block.productCols[i]!;
            const prod = products[i]!;
            
            if (cust.isEcomSpecial) {
              if (monthIdx === 0) {
                row.getCell(col).value = prod.ecomInitialQty ?? 0;
              } else {
                const prevMonthBlock = monthBlocks[monthIdx - 1]!;
                const prevColLetter = getColLetter(prevMonthBlock.productCols[i]!);
                row.getCell(col).value = {
                  formula: `${prevColLetter}${r}*1.2`
                };
              }
            } else {
              const qty = cust.customerTotalAmount > prod.amountThreshold ? prod.quantityA : prod.quantityB;
              row.getCell(col).value = qty;
            }
            row.getCell(col).numFmt = "#,##0";
            row.getCell(col).border = thinBorder;
          }

          // Month Total Volume = SUM(first_product:last_product)
          const firstProd = getColLetter(block.productCols[0]!);
          const lastProd = getColLetter(block.productCols[productCount - 1]!);
          row.getCell(block.totalCol).value = {
            formula: `SUM(${firstProd}${r}:${lastProd}${r})`
          };
          row.getCell(block.totalCol).numFmt = "#,##0";
          row.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
          row.getCell(block.totalCol).border = thinBorder;

          // Month Gross Sales +VAT = SUMPRODUCT(prices, quantities)
          const grossColLetter = getColLetter(block.grossCol!);
          row.getCell(block.grossCol!).value = {
            formula: `SUMPRODUCT(${firstProd}$${timelineHeadersStartRow + 2}:${lastProd}$${timelineHeadersStartRow + 2}, ${firstProd}${r}:${lastProd}${r})`
          };
          row.getCell(block.grossCol!).numFmt = MONEY_FORMAT;
          row.getCell(block.grossCol!).font = { name: "Calibri", size: 11, bold: true };
          row.getCell(block.grossCol!).border = thinBorder;

        } else if (block.type === "quarter") {
          // Quarter Product values are sum of months in this quarter
          for (let i = 0; i < productCount; i++) {
            const col = block.productCols[i]!;
            const monthRefs = block.monthBlocks!.map((mb) => `${getColLetter(mb.productCols[i]!)}${r}`);
            row.getCell(col).value = { formula: monthRefs.join("+") };
            row.getCell(col).numFmt = "#,##0";
            row.getCell(col).border = thinBorder;
          }

          // Quarter Total Volume = SUM(first_qproduct:last_qproduct)
          const firstProd = getColLetter(block.productCols[0]!);
          const lastProd = getColLetter(block.productCols[productCount - 1]!);
          row.getCell(block.totalCol).value = {
            formula: `SUM(${firstProd}${r}:${lastProd}${r})`
          };
          row.getCell(block.totalCol).numFmt = "#,##0";
          row.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
          row.getCell(block.totalCol).border = thinBorder;

        } else {
          // Year Block
          // Year Product values are sum of quarters
          for (let i = 0; i < productCount; i++) {
            const col = block.productCols[i]!;
            const qRefs = block.quarterBlocks!.map((qb) => `${getColLetter(qb.productCols[i]!)}${r}`);
            row.getCell(col).value = { formula: qRefs.join("+") };
            row.getCell(col).numFmt = "#,##0";
            row.getCell(col).border = thinBorder;
          }

          // Year Total Volume = SUM(first_yproduct:last_yproduct)
          const firstProd = getColLetter(block.productCols[0]!);
          const lastProd = getColLetter(block.productCols[productCount - 1]!);
          row.getCell(block.totalCol).value = {
            formula: `SUM(${firstProd}${r}:${lastProd}${r})`
          };
          row.getCell(block.totalCol).numFmt = "#,##0";
          row.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
          row.getCell(block.totalCol).border = thinBorder;

          // Year Gross Sales +VAT = SUMPRODUCT(prices, year quantities)
          const month1Block = timelineBlocks.find((b) => b.type === "month")!;
          const firstPriceCol = getColLetter(month1Block.productCols[0]!);
          const lastPriceCol = getColLetter(month1Block.productCols[productCount - 1]!);
          row.getCell(block.grossCol!).value = {
            formula: `SUMPRODUCT(${firstPriceCol}$${timelineHeadersStartRow + 2}:${lastPriceCol}$${timelineHeadersStartRow + 2}, ${firstProd}${r}:${lastProd}${r})`
          };
          row.getCell(block.grossCol!).numFmt = MONEY_FORMAT;
          row.getCell(block.grossCol!).font = { name: "Calibri", size: 11, bold: true };
          row.getCell(block.grossCol!).border = thinBorder;
        }
      });

      currentR++;
    });

    const groupEndRow = currentR - 1;

    // Write TOTAL Row for this group
    const rTotal = currentR;
    const rowT = sheet.getRow(rTotal);
    rowT.height = 18;

    // A, B, C left blank/style (to avoid double-counting in SUMIF summary rows)
    rowT.getCell(1).value = "";
    rowT.getCell(2).value = "";
    rowT.getCell(3).value = "";

    // D: Total Name mapping
    let totalLabel = `TOTAL ${group.channelName.toUpperCase()}`;
    const normChanName = group.channelName.toLowerCase().trim();
    const normSubChanName = group.subChannelName.toLowerCase().trim();
    if (group.codeChannel === 1003 || normChanName.includes("ecom") || normSubChanName.includes("ecom")) {
      totalLabel = "TOTAL E-COMMERCE";
    } else if (normChanName.includes("baby") || normSubChanName.includes("baby")) {
      totalLabel = "TOTAL BABY & MOM";
    } else if (normChanName.includes("otc") || normSubChanName.includes("otc")) {
      totalLabel = "TOTAL OTC";
    } else if (normChanName.includes("mtc") || normSubChanName.includes("mtc")) {
      totalLabel = "TOTAL MTC";
    } else if (normChanName.includes("etc") || normSubChanName.includes("etc")) {
      totalLabel = "TOTAL ETC";
    }
    rowT.getCell(4).value = totalLabel;

    // E: Store sum
    rowT.getCell(5).value = {
      formula: `SUM(E${groupStartRow}:E${groupEndRow})`
    };
    rowT.getCell(5).numFmt = "#,##0";

    // F, G, H, I: blank but styled
    rowT.getCell(6).value = "";
    rowT.getCell(7).value = "";
    rowT.getCell(8).value = "";
    rowT.getCell(9).value = "";

    // Style A-I for TOTAL row
    for (let c = 1; c <= 9; c++) {
      const cell = rowT.getCell(c);
      cell.font = { name: "Calibri", size: 11, bold: true };
      cell.fill = BLUE_TOTAL_FILL;
      cell.border = thinBorder;
      cell.alignment = { vertical: "middle", horizontal: (c === 1 || c === 5) ? "center" : "left" };
    }

    // Populate dynamic SUM formulas for all timeline blocks in the TOTAL row
    timelineBlocks.forEach((block) => {
      // Products sum
      for (let i = 0; i < productCount; i++) {
        const col = block.productCols[i]!;
        const colLetter = getColLetter(col);
        rowT.getCell(col).value = {
          formula: `SUM(${colLetter}${groupStartRow}:${colLetter}${groupEndRow})`
        };
        rowT.getCell(col).numFmt = "#,##0";
        rowT.getCell(col).font = { name: "Calibri", size: 11, bold: true };
        rowT.getCell(col).fill = BLUE_TOTAL_FILL;
        rowT.getCell(col).border = thinBorder;
        rowT.getCell(col).alignment = { vertical: "middle", horizontal: "right" };
      }

      // Total Volume sum
      const totalColLetter = getColLetter(block.totalCol);
      rowT.getCell(block.totalCol).value = {
        formula: `SUM(${totalColLetter}${groupStartRow}:${totalColLetter}${groupEndRow})`
      };
      rowT.getCell(block.totalCol).numFmt = "#,##0";
      rowT.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
      rowT.getCell(block.totalCol).fill = BLUE_TOTAL_FILL;
      rowT.getCell(block.totalCol).border = thinBorder;
      rowT.getCell(block.totalCol).alignment = { vertical: "middle", horizontal: "right" };

      // Gross Sales sum (if month/year block)
      if (block.grossCol !== null) {
        const grossColLetter = getColLetter(block.grossCol);
        rowT.getCell(block.grossCol).value = {
          formula: `SUM(${grossColLetter}${groupStartRow}:${grossColLetter}${groupEndRow})`
        };
        rowT.getCell(block.grossCol).numFmt = MONEY_FORMAT;
        rowT.getCell(block.grossCol).font = { name: "Calibri", size: 11, bold: true };
        rowT.getCell(block.grossCol).fill = BLUE_TOTAL_FILL;
        rowT.getCell(block.grossCol).border = thinBorder;
        rowT.getCell(block.grossCol).alignment = { vertical: "middle", horizontal: "right" };
      }
    });

    currentR++;
  });

  dataEndRow = currentR - 1;

  // Populate dynamic formulas for Summary rows
  for (let idx = 0; idx < channels.length; idx++) {
    const r = channelsStartRow + idx;
    const row = sheet.getRow(r);

    // 1. Column E (Stores) SUMIF
    row.getCell(5).value = {
      formula: `SUMIF($A$${dataStartRow}:$A$${dataEndRow},$A${r},E$${dataStartRow}:E$${dataEndRow})`
    };
    row.getCell(5).numFmt = "#,##0";
    row.getCell(5).font = { name: "Calibri", size: 11, bold: true };
    row.getCell(5).border = thinBorder;
    row.getCell(5).fill = SUBHEADER_FILL;

    // 2. All timeline blocks columns SUMIF
    timelineBlocks.forEach((block) => {
      // Products SUMIF
      for (let i = 0; i < productCount; i++) {
        const col = block.productCols[i]!;
        const colLetter = getColLetter(col);
        row.getCell(col).value = {
          formula: `SUMIF($A$${dataStartRow}:$A$${dataEndRow},$A${r},${colLetter}$${dataStartRow}:${colLetter}$${dataEndRow})`
        };
        row.getCell(col).numFmt = "#,##0";
        row.getCell(col).font = { name: "Calibri", size: 11 };
        row.getCell(col).border = thinBorder;
      }

      // Total Volume SUMIF
      const totalColLetter = getColLetter(block.totalCol);
      row.getCell(block.totalCol).value = {
        formula: `SUMIF($A$${dataStartRow}:$A$${dataEndRow},$A${r},${totalColLetter}$${dataStartRow}:${totalColLetter}$${dataEndRow})`
      };
      row.getCell(block.totalCol).numFmt = "#,##0";
      row.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
      row.getCell(block.totalCol).border = thinBorder;
      row.getCell(block.totalCol).fill = SUBHEADER_FILL;

      // Gross Sales SUMIF (if month/year block)
      if (block.grossCol !== null) {
        const grossColLetter = getColLetter(block.grossCol);
        row.getCell(block.grossCol).value = {
          formula: `SUMIF($A$${dataStartRow}:$A$${dataEndRow},$A${r},${grossColLetter}$${dataStartRow}:${grossColLetter}$${dataEndRow})`
        };
        row.getCell(block.grossCol).numFmt = MONEY_FORMAT;
        row.getCell(block.grossCol).font = { name: "Calibri", size: 11, bold: true };
        row.getCell(block.grossCol).border = thinBorder;
        row.getCell(block.grossCol).fill = SUBHEADER_FILL;
      }
    });
  }

  // Populate dynamic formulas for Nationwide Row (SUM of summary channels rows)
  const rowN = sheet.getRow(nationwideRow);

  // 1. Column E (Stores) SUM for Nationwide
  rowN.getCell(5).value = {
    formula: `SUM(E${channelsStartRow}:E${nationwideRow - 1})`
  };
  rowN.getCell(5).numFmt = "#,##0";
  rowN.getCell(5).font = { name: "Calibri", size: 11, bold: true };
  rowN.getCell(5).border = thinBorder;
  rowN.getCell(5).fill = SUBHEADER_FILL;
  rowN.getCell(5).alignment = { vertical: "middle", horizontal: "center" };

  // 2. All timeline blocks columns SUM for Nationwide
  timelineBlocks.forEach((block) => {
    // Product columns SUM
    for (let i = 0; i < productCount; i++) {
      const col = block.productCols[i]!;
      const colLetter = getColLetter(col);
      rowN.getCell(col).value = { formula: `SUM(${colLetter}${channelsStartRow}:${colLetter}${nationwideRow - 1})` };
      rowN.getCell(col).numFmt = "#,##0";
      rowN.getCell(col).font = { name: "Calibri", size: 11, bold: true };
      rowN.getCell(col).fill = SUBHEADER_FILL;
      rowN.getCell(col).border = thinBorder;
      rowN.getCell(col).alignment = { vertical: "middle", horizontal: "right" };
    }

    // Total Volume column SUM
    const totalColLetter = getColLetter(block.totalCol);
    rowN.getCell(block.totalCol).value = { formula: `SUM(${totalColLetter}${channelsStartRow}:${totalColLetter}${nationwideRow - 1})` };
    rowN.getCell(block.totalCol).numFmt = "#,##0";
    rowN.getCell(block.totalCol).font = { name: "Calibri", size: 11, bold: true };
    rowN.getCell(block.totalCol).fill = SUBHEADER_FILL;
    rowN.getCell(block.totalCol).border = thinBorder;
    rowN.getCell(block.totalCol).alignment = { vertical: "middle", horizontal: "right" };

    // Gross Sales column SUM (if month/year block)
    if (block.grossCol !== null) {
      const grossColLetter = getColLetter(block.grossCol);
      rowN.getCell(block.grossCol).value = { formula: `SUM(${grossColLetter}${channelsStartRow}:${grossColLetter}${nationwideRow - 1})` };
      rowN.getCell(block.grossCol).numFmt = MONEY_FORMAT;
      rowN.getCell(block.grossCol).font = { name: "Calibri", size: 11, bold: true };
      rowN.getCell(block.grossCol).fill = SUBHEADER_FILL;
      rowN.getCell(block.grossCol).border = thinBorder;
      rowN.getCell(block.grossCol).alignment = { vertical: "middle", horizontal: "right" };
    }
  });

  // Apply autofilter across the columns
  sheet.autoFilter = `A${timelineHeadersStartRow + 1}:${lastColName}${dataEndRow}`;

  return dataEndRow;
}

function addProductsReferenceSheet(workbook: ExcelJS.Workbook, data: MasterplanData): void {
  const sheet = workbook.addWorksheet("Products");
  const { products } = data;

  const thinBorder = {
    top: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    left: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    bottom: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    right: { style: "thin" as const, color: { argb: BORDER_COLOR } }
  };

  sheet.getColumn(1).width = 28;
  sheet.getColumn(2).width = 14;
  sheet.getColumn(3).width = 18;
  sheet.getColumn(4).width = 14;
  sheet.getColumn(5).width = 14;

  // Header Row 1
  const headerRow = sheet.getRow(1);
  
  headerRow.getCell(1).value = "Product Name";
  headerRow.getCell(1).font = { name: "Calibri", size: 11, bold: true };
  headerRow.getCell(1).fill = SUBHEADER_FILL;
  headerRow.getCell(1).border = thinBorder;

  headerRow.getCell(2).value = "Price";
  headerRow.getCell(2).font = { name: "Calibri", size: 11, bold: true };
  headerRow.getCell(2).fill = SUBHEADER_FILL;
  headerRow.getCell(2).border = thinBorder;

  headerRow.getCell(3).value = "Amount Threshold";
  headerRow.getCell(3).font = { name: "Calibri", size: 11, bold: true };
  headerRow.getCell(3).fill = SUBHEADER_FILL;
  headerRow.getCell(3).border = thinBorder;

  headerRow.getCell(4).value = "Quantity A";
  headerRow.getCell(4).font = { name: "Calibri", size: 11, bold: true };
  headerRow.getCell(4).fill = SUBHEADER_FILL;
  headerRow.getCell(4).border = thinBorder;

  headerRow.getCell(5).value = "Quantity B";
  headerRow.getCell(5).font = { name: "Calibri", size: 11, bold: true };
  headerRow.getCell(5).fill = SUBHEADER_FILL;
  headerRow.getCell(5).border = thinBorder;

  // Data rows
  products.forEach((p, idx) => {
    const r = idx + 2;
    const row = sheet.getRow(r);
    
    row.getCell(1).value = p.name;
    row.getCell(1).font = { name: "Calibri", size: 11 };
    row.getCell(1).border = thinBorder;

    row.getCell(2).value = p.price;
    row.getCell(2).numFmt = MONEY_FORMAT;
    row.getCell(2).font = { name: "Calibri", size: 11 };
    row.getCell(2).alignment = { vertical: "middle", horizontal: "right" };
    row.getCell(2).border = thinBorder;

    row.getCell(3).value = p.amountThreshold;
    row.getCell(3).numFmt = MONEY_FORMAT;
    row.getCell(3).font = { name: "Calibri", size: 11 };
    row.getCell(3).alignment = { vertical: "middle", horizontal: "right" };
    row.getCell(3).border = thinBorder;

    row.getCell(4).value = p.quantityA;
    row.getCell(4).numFmt = "#,##0";
    row.getCell(4).font = { name: "Calibri", size: 11 };
    row.getCell(4).alignment = { vertical: "middle", horizontal: "right" };
    row.getCell(4).border = thinBorder;

    row.getCell(5).value = p.quantityB;
    row.getCell(5).numFmt = "#,##0";
    row.getCell(5).font = { name: "Calibri", size: 11 };
    row.getCell(5).alignment = { vertical: "middle", horizontal: "right" };
    row.getCell(5).border = thinBorder;
  });
}

function addMappingSheet(
  workbook: ExcelJS.Workbook,
  data: MasterplanData,
  channelsStartRow: number,
  nationwideRow: number,
  timelineHeadersStartRow: number,
  dataStartRow: number,
  dataEndRow: number
): void {
  const sheet = workbook.addWorksheet("Mapping");
  const { startMonthIndex, startYear, customers } = data;
  const numCustomers = customers.length;

  const thinBorder = {
    top: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    left: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    bottom: { style: "thin" as const, color: { argb: BORDER_COLOR } },
    right: { style: "thin" as const, color: { argb: BORDER_COLOR } }
  };

  sheet.getColumn(1).width = 28;
  sheet.getColumn(2).width = 110;

  const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];

  const mapLines = [
    ["Field", "Value"],
    ["Main sheet", "DetailPlan_Template"],
    ["Start month rule", "Current month + 1"],
    ["Template start month", `${monthNames[startMonthIndex]} ${startYear}`],
    ["Customer data starts at row", String(dataStartRow)],
    ["Customer data prepared rows", `${dataStartRow}:${dataEndRow}`],
    ["Fixed customer columns", "A:I = Code channel, Channel, Sub-channel, Customer, Stores, Province, Regional, Sub-Channel 1, STAFF"],
    ["Product input source", "UI should provide Product Name + Price; this template uses Products sheet as sample values"],
    ["Month block rule", "For each month: product quantity columns, Total Volume, Gross Sales +VAT"],
    ["Quarter block rule", "After each quarter appearing in the timeline: product quantity totals, Quarter Total Volume. Create it even if the quarter is incomplete."],
    ["Year block rule", "At year end: product quantity totals, Total Volume, Gross Sales +VAT"],
    ["Channel summary rows", `Rows ${channelsStartRow}:${nationwideRow - 1} by channel code; row ${nationwideRow} Nationwide total`],
    ["Do not overwrite", `Rows ${timelineHeadersStartRow}:${timelineHeadersStartRow + 2} contain labels, headers, prices; write customer rows from row ${dataStartRow}`]
  ];

  mapLines.forEach((line, idx) => {
    const r = idx + 1;
    const row = sheet.getRow(r);
    const isHeader = r === 1;

    row.getCell(1).value = line[0];
    row.getCell(1).font = { name: "Calibri", size: 11, bold: isHeader };
    row.getCell(1).border = thinBorder;

    row.getCell(2).value = line[1];
    row.getCell(2).font = { name: "Calibri", size: 11, bold: isHeader };
    row.getCell(2).border = thinBorder;

    if (isHeader) {
      row.getCell(1).fill = SUBHEADER_FILL;
      row.getCell(2).fill = SUBHEADER_FILL;
    }
  });
}

function getColLetter(colNumber: number): string {
  let temp = colNumber;
  let letter = "";
  while (temp > 0) {
    const modulo = (temp - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    temp = Math.floor((temp - modulo) / 26);
  }
  return letter;
}

/**
 * Automatically adjusts column widths dynamically based on content length.
 */
function autoFitColumns(sheet: ExcelJS.Worksheet, maxCols: number): void {
  for (let c = 1; c <= maxCols; c++) {
    const column = sheet.getColumn(c);
    let maxLength = 10;

    sheet.eachRow((row) => {
      const cell = row.getCell(c);
      let cellLength = 0;

      if (cell.value) {
        if (typeof cell.value === "object") {
          if ("formula" in cell.value) {
            cellLength = 12;
          } else if (cell.value instanceof Date) {
            cellLength = 12;
          }
        } else {
          cellLength = String(cell.value).length;
        }
      }
      
      if (cellLength > maxLength) {
        maxLength = cellLength;
      }
    });

    column.width = Math.min(Math.max(maxLength + 4, 14), 40);
  }
}
