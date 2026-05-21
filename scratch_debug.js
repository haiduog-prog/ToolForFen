import fs from "fs";
import { readSourceWorkbook } from "./src/infrastructure/excel/readSourceWorkbook.js";

async function testRead() {
  try {
    const buffer = fs.readFileSync("d:\\Work\\ToolsForFen\\datamau.xlsx");
    const mockFile = {
      name: "datamau.xlsx",
      arrayBuffer: async () => {
        // Convert Node Buffer to ArrayBuffer
        const ab = new ArrayBuffer(buffer.length);
        const view = new Uint8Array(ab);
        for (let i = 0; i < buffer.length; ++i) {
          view[i] = buffer[i];
        }
        return ab;
      }
    };

    console.log("Calling readSourceWorkbook...");
    const result = await readSourceWorkbook(mockFile);
    console.log("SUCCESS! Parsed transactions count:", result.transactions.length);
    console.log("Available months:", result.availableMonths);
    console.log("Warnings:", result.warnings);
    console.log("Skipped rows:", result.skippedRows);
    if (result.transactions.length > 0) {
      console.log("First transaction sample:", JSON.stringify(result.transactions[0]));
    }
  } catch (err) {
    console.error("CRASHED during readSourceWorkbook:", err);
  }
}

testRead();
