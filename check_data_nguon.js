import ExcelJS from 'exceljs';

async function checkDataNguonColumns() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) {
    console.error('Data nguồn sheet not found');
    return;
  }
  
  console.log('=== DATA NGUON ROW 2 CELLS 1 TO 7 ===');
  const row2 = sheet.getRow(2);
  for (let c = 1; c <= 7; c++) {
    const cell = row2.getCell(c);
    console.log(`Col ${c}:`, {
      type: cell.type,
      value: cell.value,
      formula: cell.formula,
      result: cell.result
    });
  }
}

checkDataNguonColumns().catch(err => console.error(err));
