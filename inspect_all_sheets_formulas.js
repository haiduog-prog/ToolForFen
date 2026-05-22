import ExcelJS from "exceljs";

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\QEC_2025-04.xlsx");
  console.log("Workbook sheets:", wb.worksheets.map(s => s.name));
  
  for (const sheet of wb.worksheets) {
    let formulaCount = 0;
    for (let r = 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      for (let c = 1; c <= sheet.columnCount; c++) {
        const cell = row.getCell(c);
        if (cell.value && typeof cell.value === "object" && "formula" in cell.value) {
          formulaCount++;
        }
      }
    }
    console.log(`Sheet "${sheet.name}": found ${formulaCount} formulas`);
  }
}

run().catch(err => console.error(err));
