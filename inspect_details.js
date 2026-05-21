import ExcelJS from 'exceljs';

async function inspectChuan() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  console.log('Sheet names:', wb.worksheets.map(s => s.name));
  
  // 1. Inspect "Data nguồn"
  const dataNguon = wb.getWorksheet('Data nguồn');
  if (dataNguon) {
    console.log('\n=== DATA NGUON SHEET ===');
    console.log('Row count:', dataNguon.rowCount);
    
    // Print header columns
    const headerRow = dataNguon.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers.push(`Col ${colNumber}: ${cell.value}`);
    });
    console.log('Headers (Row 1):', headers.slice(0, 25).join(' | '));
    
    // Print row 2
    const row2 = dataNguon.getRow(2);
    const row2Vals = [];
    row2.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      row2Vals.push(`Col ${colNumber}: ${cell.value}`);
    });
    console.log('Row 2 Values:', row2Vals.slice(0, 25).join(' | '));
  }
  
  // 2. Inspect "QEC review"
  const qecReview = wb.getWorksheet('QEC review');
  if (qecReview) {
    console.log('\n=== QEC REVIEW SHEET ===');
    console.log('Row count:', qecReview.rowCount);
    for (let r = 1; r <= 50; r++) {
      const row = qecReview.getRow(r);
      const rowVals = [];
      for (let c = 1; c <= 15; c++) {
        const val = row.getCell(c).value;
        rowVals.push(val !== null && val !== undefined ? String(val) : '');
      }
      if (rowVals.some(v => v !== '')) {
        console.log(`Row ${r}:`, rowVals.map((v, i) => `C${i+1}:${v}`).filter(x => !x.endsWith(':')).join(' | '));
      }
    }
  }
}

inspectChuan().catch(err => console.error(err));
