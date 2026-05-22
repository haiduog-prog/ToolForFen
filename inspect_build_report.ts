import { readSourceWorkbook } from "./src/infrastructure/excel/readSourceWorkbook.js";
import { buildQecReport } from "./src/application/buildQecReport.js";
import * as fs from "fs";

async function run() {
  try {
    const fileBuffer = fs.readFileSync("d:\\work\\ToolForFen\\input.xlsx");
    
    const mockFile = {
      name: "input.xlsx",
      arrayBuffer: async () => {
        const ab = new ArrayBuffer(fileBuffer.length);
        const view = new Uint8Array(ab);
        for (let i = 0; i < fileBuffer.length; ++i) {
          view[i] = fileBuffer[i];
        }
        return ab;
      }
    };

    const source = await readSourceWorkbook(mockFile as any);
    console.log("Parsed successful. Transactions:", source.transactions.length);
    
    // Choose 2026-05 as the report month
    const reportMonth = "2026-05";
    console.log(`Building QEC Report for month ${reportMonth}...`);
    
    const report = buildQecReport(source, reportMonth as any);
    console.log("REPORT BUILT SUCCESSFULLY!");
    console.log("QEC Rows:", report.qecRows.length);
    console.log("SKU Revenue Rows:", report.skuRevenueRows.length);
    console.log("Customer Revenue Sections:", report.customerRevenueSections.length);
    console.log("Customer Quantity Sections:", report.customerQuantitySections.length);
    
    // Let's print summary info
    console.log("Summary:", report.summary);
  } catch (err) {
    console.error("CRASHED WHILE BUILDING REPORT:", err);
  }
}

run();
