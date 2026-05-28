import ExcelJS from "exceljs";

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\chuan.xlsx");
  const sheet = wb.getWorksheet("Data nguồn");
  if (!sheet) {
    console.error("Sheet not found");
    return;
  }
  
  console.log("=== DATA NGUON COLUMNS IN CHUAN.XLSX ===");
  const r1 = sheet.getRow(1);
  for (let c = 1; c <= 30; c++) {
    const val = r1.getCell(c).value;
    if (val) console.log(`Col ${c}: ${val}`);
  }
}

run().catch(console.error);
