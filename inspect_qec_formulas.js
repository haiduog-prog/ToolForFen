import ExcelJS from "exceljs";

async function inspectFormulas() {
  const wb = new ExcelJS.Workbook();
  // Read the actual template file QEC_2025-04.xlsx!
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\QEC_2025-04.xlsx");
  
  const sheet = wb.getWorksheet("QEC review");
  if (!sheet) {
    console.error("QEC review sheet not found in QEC_2025-04.xlsx");
    return;
  }
  
  console.log("=== QEC REVIEW CELLS WITH FORMULAS ===");
  let count = 0;
  for (let r = 1; r <= 30; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= 40; c++) {
      const cell = row.getCell(c);
      if (cell.value && typeof cell.value === "object" && "formula" in cell.value) {
        count++;
        if (count <= 20) {
          console.log(`Cell [Row ${r}, Col ${c}] (${sheet.getRow(2).getCell(c).value}):`, cell.value);
        }
      }
    }
  }
  console.log(`Total formulas printed (first 20): ${count}`);
}

inspectFormulas().catch(err => console.error(err));
