import ExcelJS from 'exceljs';

async function inspectSourceErrors() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) {
    console.error('Data nguồn sheet not found');
    return;
  }
  
  let errorCount = 0;
  let nonNullCount = 0;
  
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= 6; c++) {
      const val = row.getCell(c).value;
      if (val !== null && val !== undefined) {
        nonNullCount++;
        if (typeof val === 'object' && val.error) {
          errorCount++;
        }
        if (nonNullCount <= 10) {
          console.log(`Row ${r} Col ${c}:`, val);
        }
      }
    }
  }
  
  console.log(`Total non-null cells in columns 1-6 of Data nguồn: ${nonNullCount}`);
  console.log(`Total cells with error in columns 1-6 of Data nguồn: ${errorCount}`);
}

inspectSourceErrors().catch(err => console.error(err));
