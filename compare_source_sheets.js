import ExcelJS from 'exceljs';

async function compareSourceSheets() {
  const wbChuan = new ExcelJS.Workbook();
  await wbChuan.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  const sheetChuan = wbChuan.getWorksheet('Data nguồn');
  
  const wbMau = new ExcelJS.Workbook();
  await wbMau.xlsx.readFile('d:\\Work\\ToolsForFen\\datamau.xlsx');
  const sheetMau = wbMau.getWorksheet('Data nguồn');
  
  if (!sheetChuan || !sheetMau) {
    console.error('One of the source worksheets was not found');
    return;
  }
  
  console.log('=== COMPARING SOURCE SHEETS ===');
  console.log(`Chuan row count: ${sheetChuan.rowCount}, Mau row count: ${sheetMau.rowCount}`);
  
  let diffCount = 0;
  const maxRows = Math.max(sheetChuan.rowCount, sheetMau.rowCount);
  
  // Compare row 1 (headers) and up to 10 rows of data first, then count total differences
  for (let r = 1; r <= maxRows; r++) {
    const rowChuan = sheetChuan.getRow(r);
    const rowMau = sheetMau.getRow(r);
    
    let rowDiff = false;
    for (let c = 1; c <= 30; c++) {
      const valChuan = rowChuan.getCell(c).value;
      const valMau = rowMau.getCell(c).value;
      
      const strChuan = valChuan !== null && valChuan !== undefined ? String(valChuan) : '';
      const strMau = valMau !== null && valMau !== undefined ? String(valMau) : '';
      
      if (strChuan !== strMau) {
        rowDiff = true;
        diffCount++;
        if (diffCount <= 5) {
          console.log(`Difference ${diffCount} at Row ${r} Col ${c}: Chuan='${strChuan}', Mau='${strMau}'`);
        }
      }
    }
  }
  
  console.log(`Total differences in first 30 columns: ${diffCount}`);
}

compareSourceSheets().catch(err => console.error(err));
