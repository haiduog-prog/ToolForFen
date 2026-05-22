import ExcelJS from "exceljs";

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\QEC_2025-04.xlsx");
  const sheet = wb.getWorksheet("SKU - Customer review");
  if (!sheet) {
    console.error("Sheet not found");
    return;
  }
  
  console.log("=== INSPECTING SKU - CUSTOMER REVIEW (FIRST 70 ROWS) ===");
  for (let r = 1; r <= 70; r++) {
    const row = sheet.getRow(r);
    const cells = [];
    for (let c = 1; c <= 30; c++) {
      cells.push(row.getCell(c).value);
    }
    
    // Check if it's a customer header, product row, or total row
    const label = cells[3]; // Column D is Name SKU (4th column, 1-indexed)
    const custVal = cells[1]; // Column B is customer (2nd column, 1-indexed)
    
    // Check background color of Column D cell
    const cellD = row.getCell(4);
    const bgFill = cellD.fill ? JSON.stringify(cellD.fill) : "none";
    
    console.log(`Row ${r}: B=${cells[1]}, D=${cells[3]}, E=${cells[4]}, V=${cells[21]}, bgD=${bgFill}`);
  }
}

run().catch(err => console.error(err));
