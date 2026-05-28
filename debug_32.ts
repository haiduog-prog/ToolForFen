/**
 * Debug script: find all cells with value = 32 in the generated Excel output
 */
import { buildMasterplan } from "./src/application/buildMasterplan";
import type { SalesPlanProduct, SalesPlanCustomer } from "./src/domain/masterplan";

// Simulate a minimal scenario
const products: SalesPlanProduct[] = [
  { name: "SPR EPO", price: 590000, amountThreshold: 1500000, quantityA: 10, quantityB: 2 },
  { name: "SPR Collagen", price: 590000, amountThreshold: 2000000, quantityA: 8, quantityB: 1 },
];

const customers: SalesPlanCustomer[] = [
  {
    codeChannel: 1004,
    channel: "Baby & Mom",
    subChannel: "Baby & Mom",
    customer: "Test Customer 1",
    stores: "1",
    province: "HCM",
    regional: "South",
    subChannel1: "B&M",
    staff: "NV01",
    customerTotalAmount: 2000000,
  },
  {
    codeChannel: 1004,
    channel: "Baby & Mom",
    subChannel: "Baby & Mom",
    customer: "Test Customer 2",
    stores: "3",
    province: "HN",
    regional: "North",
    subChannel1: "B&M",
    staff: "NV02",
    customerTotalAmount: 500000,
  },
];

const data = buildMasterplan(customers, products, false);

console.log("=== MASTERPLAN DATA INSPECTION ===");
console.log("Products:", data.products.length);
console.log("Customers:", data.customers.length);
console.log("Timeline blocks:", data.timelineBlocks.length);
console.log("Start month:", data.startMonthIndex, "Start year:", data.startYear);

// Now let's check the actual Excel writer output
async function checkExcel() {
  const { writeMasterplanWorkbook } = await import("./src/infrastructure/excel/writeMasterplanWorkbook");
  const ExcelJSModule = await import("exceljs/dist/exceljs.min.js");
  
  const blob = await writeMasterplanWorkbook(data);
  const arrayBuffer = await blob.arrayBuffer();
  
  const workbook = new ExcelJSModule.default.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  const sheet = workbook.getWorksheet("DetailPlan_Template");
  if (!sheet) {
    console.log("No DetailPlan_Template sheet found!");
    return;
  }
  
  console.log("\n=== ALL CELLS IN THE SHEET ===");
  let cellsWith32: string[] = [];
  
  sheet.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    row.eachCell({ includeEmpty: false }, (cell: any, colNumber: number) => {
      const val = cell.value;
      
      // Check for value = 32
      if (val === 32 || val === "32" || (typeof val === "number" && val === 32)) {
        cellsWith32.push(`Row ${rowNumber}, Col ${colNumber}: value=${JSON.stringify(val)}, type=${typeof val}`);
      }
    });
  });
  
  if (cellsWith32.length > 0) {
    console.log("\n=== CELLS WITH VALUE = 32 ===");
    cellsWith32.forEach(s => console.log("  " + s));
  } else {
    console.log("\nNo cells with value = 32 found.");
  }
  
  // Print ALL cells to trace any weird values
  console.log("\n=== DUMP ALL CELL VALUES (non-formula, non-string for data rows) ===");
  sheet.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    row.eachCell({ includeEmpty: false }, (cell: any, colNumber: number) => {
      const val = cell.value;
      // Only print numeric values that might be suspicious
      if (typeof val === "number" && ![18, 24, 1000, 1001, 1002, 1003, 1004, 1005, 1006, 1007].includes(val)) {
        // Skip known prices
        if (val !== 590000 && val !== 750000 && val !== 745000 && val !== 1500000 && val !== 2000000 && val !== 1000000 && val !== 3000000 && val !== 1200000) {
          // Print anything that isn't a known quantity (10, 8, 15, 12, 6, 2, 1, 3, 0)
          if (![0, 1, 2, 3, 5, 6, 8, 10, 12, 15].includes(val)) {
            console.log(`  Row ${rowNumber}, Col ${colNumber}: value=${val}`);
          }
        }
      }
    });
  });
  
  // Write the file to disk for manual inspection
  const fs = await import("fs");
  const buffer = Buffer.from(arrayBuffer);
  fs.writeFileSync("debug_output.xlsx", buffer);
  console.log("\nFile saved to debug_output.xlsx for manual inspection.");
}

checkExcel().catch(console.error);
