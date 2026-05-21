import ExcelJS from 'exceljs';

async function inspectQecCells() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('QEC review');
  if (!sheet) {
    console.error('QEC review sheet not found');
    return;
  }
  
  console.log('=== CELL DETAILS IN QEC REVIEW ===');
  
  // 1. Segment table - Row 2 (B&M), Column 7 (2024 -01)
  console.log('Row 2 Col 7 (B&M 2024-01):', getCellDetails(sheet, 2, 7));
  
  // 2. Region table - Row 14 (1. HA NOI), Column 7 (2024 -01)
  console.log('Row 14 Col 7 (1. HA NOI 2024-01):', getCellDetails(sheet, 14, 7));
  
  // 3. DSR table - Row 23 (TRUONG MINH CANH), Column 7 (2024 -01)
  console.log('Row 23 Col 7 (TRUONG MINH CANH 2024-01):', getCellDetails(sheet, 23, 7));
  
  // 4. Customer table - Row 29 (Guardian), Column 7 (2024 -01)
  console.log('Row 29 Col 7 (Guardian 2024-01):', getCellDetails(sheet, 29, 7));
  console.log('Row 29 Col 14 (Guardian 2024-08):', getCellDetails(sheet, 29, 14));
}

function getCellDetails(sheet, rowNum, colNum) {
  const cell = sheet.getRow(rowNum).getCell(colNum);
  return {
    type: cell.type, // 0: null, 2: number, 3: string, 4: formula, etc.
    value: cell.value,
    formula: cell.formula,
    result: cell.result
  };
}

inspectQecCells().catch(err => console.error(err));
