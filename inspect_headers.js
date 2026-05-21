import ExcelJS from 'exceljs';

async function inspectHeaders() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\datamau.xlsx');
  
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) return;
  
  console.log('=== DATA NGUỒN HEADER ROW ===');
  const row = sheet.getRow(1);
  for (let c = 1; c <= 30; c++) {
    const val = row.getCell(c).value;
    if (val !== null && val !== undefined) {
      console.log(`Col ${c}: "${val}"`);
    }
  }
  
  // Also print a sample data row
  console.log('\n=== SAMPLE DATA ROW (Row 2) ===');
  const dataRow = sheet.getRow(2);
  for (let c = 1; c <= 30; c++) {
    const val = dataRow.getCell(c).value;
    if (val !== null && val !== undefined) {
      console.log(`Col ${c}: "${val}"`);
    }
  }
}

inspectHeaders().catch(err => console.error(err));
