import { readSourceWorkbook } from "./src/infrastructure/excel/readSourceWorkbook.js";
import * as fs from "fs";

async function run() {
  try {
    const fileBuffer = fs.readFileSync("d:\\work\\ToolForFen\\input.xlsx");
    
    // Mock the browser File object interface
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

    const result = await readSourceWorkbook(mockFile as any);
    console.log("SUCCESSFULLY PARSED!");
    console.log("Parsed Transactions Count:", result.transactions.length);
    console.log("Skipped Rows Count:", result.skippedRows);
    console.log("Warnings:", result.warnings);
    console.log("Available Months:", result.availableMonths);
    if (result.transactions.length > 0) {
      console.log("First Transaction:", result.transactions[0]);
    }
  } catch (err) {
    console.error("Error running parser:", err);
  }
}

run();
