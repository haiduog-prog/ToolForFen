import ExcelJS from "exceljs";

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile("d:\\work\\ToolForFen\\chuan.xlsx");
  const sheet = wb.getWorksheet("QEC review");
  if (!sheet) {
    console.error("Sheet not found");
    return;
  }
  
  // Ta biết bảng Customer Detail bắt đầu ở khoảng dòng 280+.
  // Hãy quét từ dòng 280 đến 310, in ra cột 1, 2, 3, 4, 5, 6
  console.log("=== CUSTOMER DETAIL ROWS IN CHUAN.XLSX ===");
  for (let r = 280; r <= 310; r++) {
    const row = sheet.getRow(r);
    const c1 = row.getCell(1).value;
    const c2 = row.getCell(2).value;
    const c3 = row.getCell(3).value;
    const c4 = row.getCell(4).value;
    const c5 = row.getCell(5).value;
    const c6 = row.getCell(6).value;
    if (c1 || c2) {
      console.log(`Row ${r}: Col1=${JSON.stringify(c1)} | Col2=${JSON.stringify(c2)} | Col3=${JSON.stringify(c3)} | Col4=${JSON.stringify(c4)} | Col5=${JSON.stringify(c5)} | Col6=${JSON.stringify(c6)}`);
    }
  }
}

run().catch(console.error);
