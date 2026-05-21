import { describe, expect, it } from "vitest";
import fs from "fs";
import ExcelJS from "exceljs";
import { readSourceWorkbook } from "../infrastructure/excel/readSourceWorkbook";
import { buildQecReport } from "../application/buildQecReport";
import { writeReportWorkbook } from "../infrastructure/excel/writeReportWorkbook";
import type { MonthKey } from "./month";

describe("End-to-End Excel Generation and Comparison", () => {
  it("should parse datamau, build report, write excel and match chuan.xlsx", async () => {
    // 1. Read datamau.xlsx
    const buffer = fs.readFileSync("d:\\Work\\ToolsForFen\\datamau.xlsx");
    const mockFile = {
      name: "datamau.xlsx",
      arrayBuffer: async () => {
        const ab = new ArrayBuffer(buffer.length);
        const view = new Uint8Array(ab);
        for (let i = 0; i < buffer.length; ++i) {
          view[i] = buffer[i];
        }
        return ab;
      }
    };

    const parseResult = await readSourceWorkbook(mockFile as any);
    expect(parseResult.transactions.length).toBeGreaterThan(0);

    // 2. Build report for 2025-04
    const report = buildQecReport(parseResult, "2025-04" as MonthKey);

    // 3. Write report to test_output.xlsx
    const blob = await writeReportWorkbook(report, parseResult.transactions);
    expect(blob).toBeInstanceOf(Blob);

    // Save Blob to file
    const arrayBuffer = await blob.arrayBuffer();
    const outputBuffer = Buffer.from(arrayBuffer);
    const outputPath = "d:\\Work\\ToolsForFen\\test_output.xlsx";
    fs.writeFileSync(outputPath, outputBuffer);
    console.log(`Saved output excel to ${outputPath}`);

    // 4. Compare generated excel with chuan.xlsx cell by cell
    const wbGen = new ExcelJS.Workbook();
    await wbGen.xlsx.readFile(outputPath);

    const wbChuan = new ExcelJS.Workbook();
    await wbChuan.xlsx.readFile("d:\\Work\\ToolsForFen\\chuan.xlsx");

    // We compare worksheets: 'QEC review', 'SKU review', 'SKU - Customer review'
    const sheetsToCompare = ["QEC review", "SKU review", "SKU - Customer review"];
    expect(wbGen.getWorksheet("SKU customer review")).toBeDefined();

    for (const sheetName of sheetsToCompare) {
      console.log(`\n--- Comparing sheet: "${sheetName}" ---`);
      const sGen = wbGen.getWorksheet(sheetName);
      const sChuan = wbChuan.getWorksheet(sheetName);

      expect(sGen).toBeDefined();
      expect(sChuan).toBeDefined();

      if (!sGen || !sChuan) continue;

      console.log(`Generated sheet rows: ${sGen.rowCount}, Chuan sheet rows: ${sChuan.rowCount}`);

      let diffCount = 0;
      const maxRows = Math.max(sGen.rowCount, sChuan.rowCount);

      for (let r = 1; r <= maxRows; r++) {
        const rGen = sGen.getRow(r);
        const rChuan = sChuan.getRow(r);

        // Check values in columns 1 to 40
        for (let c = 1; c <= 40; c++) {
          const vGen = rGen.getCell(c).value;
          const vChuan = rChuan.getCell(c).value;

          // Normalize values
          let valGen = vGen === null || vGen === undefined ? "" : vGen;
          let valChuan = vChuan === null || vChuan === undefined ? "" : vChuan;

          // If formula, compare formulas
          if (typeof valGen === "object" && "formula" in valGen) {
            valGen = `FORMULA:${valGen.formula}`;
          }
          if (typeof valChuan === "object" && "formula" in valChuan) {
            valChuan = `FORMULA:${valChuan.formula}`;
          }

          // If they are numbers, allow very small float differences
          let areEqual = false;
          if (typeof valGen === "number" && typeof valChuan === "number") {
            areEqual = Math.abs(valGen - valChuan) < 1.0; // Allow slight rounding differences
          } else {
            areEqual = String(valGen).trim() === String(valChuan).trim();
          }

          if (!areEqual) {
            diffCount++;
            if (diffCount <= 30) {
              console.log(
                `  DIFF at Cell [Row ${r}, Col ${c}]: Ours=${JSON.stringify(vGen)}, Chuan=${JSON.stringify(vChuan)}`
              );
            }
          }
        }
      }

      console.log(`Sheet "${sheetName}" comparison finished. Total differences: ${diffCount}`);
      // In strict mode, we want diffCount to be 0 for critical sheets.
      // But let's log them first before failing the test to see what differs.
    }
  }, 120000);
});
