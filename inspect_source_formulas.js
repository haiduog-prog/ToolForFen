import ExcelJS from 'exceljs';

async function inspectSourceFormulas() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) {
    console.error('Data nguồn sheet not found');
    return;
  }
  
  console.log('=== DATA NGUON COLUMN 1-5 FORMULAS ===');
  
  for (let r = 4; r <= 15; r++) {
    const row = sheet.getRow(r);
    const rowVals = [];
    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c);
      rowVals.push(`Col ${c}: val=${JSON.stringify(cell.value)}, formula=${cell.formula}`);
    }
    console.log(`Row ${r}:`, rowVals.join(' | '));
  }
}

inspectSourceFormulas().catch(err => console.error(err));
