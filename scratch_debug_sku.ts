import { readSourceWorkbook } from "./src/infrastructure/excel/readSourceWorkbook.js";
import { buildQecReport } from "./src/application/buildQecReport.js";
import { STATIC_VND_PRODUCTS } from "./src/domain/customerMapping.js";
import * as fs from "fs";

async function run() {
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
  console.log("Transactions:", source.transactions.length);
  const report = buildQecReport(source, "2026-05");

  console.log("STATIC_VND_PRODUCTS length:", STATIC_VND_PRODUCTS.length);
  console.log("skuRevenueRows length:", report.skuRevenueRows.length);
  
  console.log("Products in parsed transactions:", Array.from(new Set(source.transactions.map(t => t.product))));

  const babyRow = report.skuRevenueRows.find(r => r.label.includes("BABY"));
  console.log("BABY row found in skuRevenueRows:", babyRow ? babyRow.label : "NOT FOUND");
  if (babyRow) {
    console.log("BABY row values:", babyRow.monthValues);
    console.log("BABY row totals:", {
      prevTotal: babyRow.previousYearTotal,
      currTotal: babyRow.currentYearTotal
    });
  }

  console.log("\nQEC Rows:");
  report.qecRows.forEach(r => {
    console.log(`- ${r.label}: prevTotal=${r.previousYearTotal}, currTotal=${r.currentYearTotal}`);
  });
}

run().catch(console.error);
