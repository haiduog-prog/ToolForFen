import ExcelJS from "exceljs";

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\QEC_2025-04.xlsx");
  const sheet = wb.getWorksheet("QEC review");
  if (!sheet) {
    console.error("Sheet not found");
    return;
  }
  
  console.log("=== CUSTOMER DETAIL ROWS IN QEC_2025-04.XLSX ===");
  for (let r = 280; r <= 320; r++) {
    const row = sheet.getRow(r);
    const c1 = row.getCell(1).value;
    const c2 = row.getCell(2).value;
    if (c1 || c2) {
      console.log(`Row ${r}: Col1=${JSON.stringify(c1)} | Col2=${JSON.stringify(c2)}`);
    }
  }
}

run().catch(console.error);
