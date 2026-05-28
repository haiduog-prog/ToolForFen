import { readSourceWorkbook } from "./src/infrastructure/excel/readSourceWorkbook.js";
import { buildQecReport } from "./src/application/buildQecReport.js";
import * as fs from "fs";

async function run() {
  try {
    const fileBuffer = fs.readFileSync("d:\\work\\ToolForFen\\chuan.xlsx");
    
    const mockFile = {
      name: "chuan.xlsx",
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
    
    // Choose 2025-04 as the report month (based on chuan.xlsx headers)
    const reportMonth = "2025-04";
    console.log(`Building QEC Report for month ${reportMonth}...`);
    
    const report = buildQecReport(source, reportMonth as any);
    console.log("REPORT BUILT SUCCESSFULLY!");
    console.log("QEC Rows Count:", report.qecRows.length);
    
    console.log("\n=== SEGMENT ROWS IFYTD, ICYTD, IYA ===");
    for (const row of report.qecRows) {
      console.log(`Label: ${row.label}`);
      console.log(`  CY Total: ${row.currentYearTotal}`);
      console.log(`  PY Total: ${row.previousYearTotal}`);
      console.log(`  IFYTD: ${row.ifytd}`);
      console.log(`  ICYTD: ${row.icytd}`);
      console.log(`  IYA: ${row.iya}`);
    }
  } catch (err) {
    console.error("CRASHED WHILE BUILDING REPORT:", err);
  }
}

run();
