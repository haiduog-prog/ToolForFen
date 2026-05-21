import ExcelJS from 'exceljs';

async function inspectFormulas() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('QEC review');
  if (!sheet) {
    console.error('QEC review sheet not found');
    return;
  }
  
  console.log('=== INSPECTING FORMULAS IN QEC REVIEW ===');
  
  // Let's print formulas for:
  // Segment table (rows 2-10)
  // Region table (rows 14-19)
  // DSR table (rows 23-25)
  // for Column 7 (2024 -01) and Column 14 (2024 -08)
  
  const rowsToInspect = [
    { row: 2, name: 'Segment B&M' },
    { row: 9, name: 'Segment MT' },
    { row: 14, name: 'Region 1. HA NOI' },
    { row: 18, name: 'Region 5. HCM' },
    { row: 23, name: 'DSR TRUONG MINH CANH' },
    { row: 24, name: 'DSR TRINH NGOC HA' }
  ];
  
  for (const item of rowsToInspect) {
    console.log(`\n--- ${item.name} (Row ${item.row}) ---`);
    for (let c = 7; c <= 22; c++) {
      const cell = sheet.getRow(item.row).getCell(c);
      const colLetter = sheet.getRow(1).getCell(c).value || `Col ${c}`;
      if (cell.type === 4 || (cell.value && typeof cell.value === 'object' && cell.value.formula)) {
        console.log(`Month [${colLetter}]: Formula = "${cell.value.formula}", Result = ${cell.value.result}`);
      } else {
        console.log(`Month [${colLetter}]: Static Value = ${JSON.stringify(cell.value)}`);
      }
    }
  }
}

inspectFormulas().catch(err => console.error(err));
