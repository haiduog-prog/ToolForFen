import ExcelJS from 'exceljs';

async function findAllFormulas() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('QEC review');
  if (!sheet) {
    console.error('QEC review sheet not found');
    return;
  }
  
  console.log('=== SEARCHING FOR FORMULAS IN QEC REVIEW ===');
  let formulaCount = 0;
  
  for (let r = 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= sheet.columnCount; c++) {
      const cell = row.getCell(c);
      if (cell.type === 4 || (cell.value && typeof cell.value === 'object' && cell.value.formula)) {
        formulaCount++;
        console.log(`Cell ${r},${c} (${sheet.getRow(1).getCell(c).value || 'Col ' + c}): Formula = "${cell.value.formula}", Result = ${cell.value.result}`);
      }
    }
  }
  
  console.log(`Total formulas found in QEC review: ${formulaCount}`);
}

findAllFormulas().catch(err => console.error(err));
