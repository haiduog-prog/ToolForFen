import ExcelJS from "exceljs";

async function inspectQecDetails() {
  const wb = new ExcelJS.Workbook();
  // Hãy thử đọc chuan.xlsx
  try {
    await wb.xlsx.readFile("d:\\work\\ToolForFen\\chuan.xlsx");
    console.log("Successfully loaded chuan.xlsx");
  } catch (e) {
    console.log("Failed to load chuan.xlsx, trying QEC_2025-04.xlsx...");
    await wb.xlsx.readFile("d:\\work\\ToolForFen\\QEC_2025-04.xlsx");
    console.log("Successfully loaded QEC_2025-04.xlsx");
  }
  
  const qecReview = wb.getWorksheet("QEC review");
  if (!qecReview) {
    console.error("QEC review sheet not found");
    return;
  }
  
  console.log("Header row 1 (Segment):");
  const r1 = qecReview.getRow(1);
  for (let c = 1; c <= 30; c++) {
    const val = r1.getCell(c).value;
    if (val) console.log(`Col ${c}: ${val}`);
  }
  
  console.log("\nFirst 15 rows of QEC review:");
  for (let r = 1; r <= 15; r++) {
    const row = qecReview.getRow(r);
    const label = row.getCell(6).value;
    // IFYTD, ICYTD, IYA nằm ở cột nào?
    // Giả sử có 16 tháng (từ cột 7 đến 22).
    // Share 2024: cột 23
    // Share 2025: cột 24
    // CY 2024: cột 25
    // CY 2025: cột 26
    // P3M: cột 27, P6M: cột 28, P9M: cột 29
    // TREND: cột 30, IFYTD: cột 31, ICYTD: cột 32, IYA: cột 33
    const cells = [6, 25, 26, 27, 28, 29, 30, 31, 32, 33].map(c => `${c}: ${JSON.stringify(row.getCell(c).value)}`);
    console.log(`Row ${r}: ${cells.join(" | ")}`);
  }
  
  console.log("\nRows from 290 to 305 of QEC review (Customer details total area):");
  for (let r = 290; r <= 310; r++) {
    const row = qecReview.getRow(r);
    const label = row.getCell(2).value || row.getCell(6).value;
    if (label) {
      const cells = [2, 6, 25, 26, 27, 28, 29, 30, 31, 32, 33].map(c => `${c}: ${JSON.stringify(row.getCell(c).value)}`);
      console.log(`Row ${r} (${label}): ${cells.join(" | ")}`);
    }
  }
}

inspectQecDetails().catch(console.error);
