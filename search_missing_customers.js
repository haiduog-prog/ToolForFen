import ExcelJS from 'exceljs';

async function searchMissingCustomers() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const searchNames = ['Vũ Thị Hòa', 'Như Trần Ngọc'];
  
  for (const sheet of wb.worksheets) {
    console.log(`\nSearching in sheet: ${sheet.name}`);
    for (let r = 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      for (let c = 1; c <= sheet.columnCount; c++) {
        const val = row.getCell(c).value;
        if (val && typeof val === 'string') {
          for (const name of searchNames) {
            if (val.includes(name)) {
              console.log(`Found [${name}] at Row ${r} Col ${c}: "${val}"`);
              // Print nearby cells in this row
              const rowVals = [];
              for (let i = 1; i <= Math.min(10, sheet.columnCount); i++) {
                rowVals.push(`C${i}:${row.getCell(i).value}`);
              }
              console.log(`  Row ${r} details:`, rowVals.join(' | '));
            }
          }
        }
      }
    }
  }
}

searchMissingCustomers().catch(err => console.error(err));
