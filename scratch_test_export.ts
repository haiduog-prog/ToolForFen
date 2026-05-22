import { readSourceWorkbook } from "./src/infrastructure/excel/readSourceWorkbook.js";
import { buildQecReport } from "./src/application/buildQecReport.js";
import { writeReportWorkbook } from "./src/infrastructure/excel/writeReportWorkbook.js";
import * as fs from "fs";
import ExcelJS from "exceljs";

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

  console.log("Reading source workbook...");
  const source = await readSourceWorkbook(mockFile as any);
  console.log("Building report for 2026-05...");
  const report = buildQecReport(source, "2026-05");

  console.log("Writing report to Blob...");
  const blob = await writeReportWorkbook(report, source.transactions);
  const buffer = Buffer.from(await blob.arrayBuffer());
  fs.writeFileSync("d:\\work\\ToolForFen\\test_input_export.xlsx", buffer);
  console.log("Saved test_input_export.xlsx successfully!");

  // Now, read the file back with ExcelJS to verify that there is non-zero data
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\test_input_export.xlsx");

  console.log("\n=== VERIFYING GENERATED EXCEL FILE ===");

  const sheets = wb.worksheets.map(w => w.name);
  console.log("Sheets found:", sheets);

  // 1. Check QEC review sheet
  const qec = wb.getWorksheet("QEC review")!;
  console.log("\n[QEC review] sheet rows count:", qec.rowCount);
  // Find rows with Unmapped and Total
  for (let r = 1; r <= qec.rowCount; r++) {
    const label = qec.getRow(r).getCell(6).value;
    if (label === "Unmapped" || label === "Total" || label === "TOTAL") {
      const currYearVal = qec.getRow(r).getCell(26).value;
      const prevYearVal = qec.getRow(r).getCell(25).value;
      console.log(`  Row ${r} (${label}): 2025=${prevYearVal}, 2026=${currYearVal}`);
    }
  }

  // 2. Check SKU review sheet
  const sku = wb.getWorksheet("SKU review")!;
  console.log("\n[SKU review] sheet rows count:", sku.rowCount);
  // Print rows where product name contains BABY LACTOFERRIN or has non-zero total
  for (let r = 1; r <= sku.rowCount; r++) {
    const skuName = sku.getRow(r).getCell(2).value;
    if (skuName && String(skuName).includes("BABY")) {
      const cyVal = sku.getRow(r).getCell(22).value;
      const pyVal = sku.getRow(r).getCell(21).value;
      console.log(`  Row ${r} (${skuName}): CY=${cyVal}, PY=${pyVal}`);
    }
  }

  // 3. Check SKU - Customer review sheet
  const skuCust = wb.getWorksheet("SKU - Customer review")!;
  console.log("\n[SKU - Customer review] sheet rows count:", skuCust.rowCount);
  let positiveCells = 0;
  for (let r = 1; r <= Math.min(1000, skuCust.rowCount); r++) {
    const row = skuCust.getRow(r);
    for (let c = 5; c <= 30; c++) {
      const val = Number(row.getCell(c).value || 0);
      if (val > 0) {
        positiveCells++;
      }
    }
  }
  console.log(`  Number of non-zero value cells in first 1000 rows of SKU - Customer review: ${positiveCells}`);

  // 4. Check Data nguồn sheet
  const src = wb.getWorksheet("Data nguồn")!;
  console.log("\n[Data nguồn] sheet rows count:", src.rowCount);
  const firstRow = src.getRow(2);
  console.log("  Row 2 data:", [1,2,3,4,5,6,7,8,9,10].map(c => firstRow.getCell(c).value));
}

run().catch(console.error);
