import ExcelJS from 'exceljs';

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\work\\ToolForFen\\QEC_2025-04.xlsx');
  
  const sheet = wb.getWorksheet('QEC review');
  console.log('Sheet name: QEC review');
  console.log('Row count:', sheet.rowCount);
  
  // Find where Customer Detail table starts
  let startRow = -1;
  for (let r = 1; r <= sheet.rowCount; r++) {
    const val = sheet.getRow(r).getCell(2).value;
    if (val && String(val).includes('CUSTOMER_CHANGE')) {
      startRow = r;
      break;
    }
  }
  
  if (startRow === -1) {
    console.log('Could not find CUSTOMER_CHANGE header row.');
    return;
  }
  
  console.log('Customer Detail table starts at row:', startRow);
  
  // Print headers
  const headerRow = sheet.getRow(startRow);
  const headers = [];
  for (let c = 1; c <= 10; c++) {
    headers.push(headerRow.getCell(c).value);
  }
  console.log('Headers (C1-C10):', headers);
  
  // Print first 25 data rows
  for (let r = startRow + 1; r <= startRow + 25; r++) {
    const row = sheet.getRow(r);
    const cells = [];
    for (let c = 1; c <= 8; c++) {
      cells.push(row.getCell(c).value);
    }
    console.log(`Row ${r}:`, cells);
  }
}

run().catch(console.error);
