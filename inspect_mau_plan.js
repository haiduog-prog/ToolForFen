import ExcelJS from "exceljs";

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\Work\\ToolsForFen\\mau_xuat_sales_plan_start_2026-06.xlsx");
  const sheet = wb.getWorksheet("DetailPlan_Template");
  if (!sheet) return;

  console.log("=== CHECK ROWS 18-30 CUSTOMER DATA ===");
  for (let r = 18; r <= 30; r++) {
    const row = sheet.getRow(r);
    const rowVals = [];
    for (let c = 1; c <= 9; c++) {
      rowVals.push(row.getCell(c).value || "");
    }
    console.log(`Row ${r}: ${rowVals.join(" | ")}`);
  }
}

run().catch(console.error);
