import ExcelJS from 'exceljs';

async function listSheets() {
  for (const file of ['datamau.xlsx', 'chuan.xlsx']) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(`d:\\Work\\ToolsForFen\\${file}`);
    console.log(`\n=== ${file} SHEETS ===`);
    wb.eachSheet((sheet, id) => {
      console.log(`  Sheet ${id}: "${sheet.name}" (${sheet.rowCount} rows, ${sheet.columnCount} cols)`);
    });
  }
}

listSheets().catch(err => console.error(err));
