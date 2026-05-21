import ExcelJS from 'exceljs';

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sku = wb.getWorksheet('SKU review');
  
  console.log('=== ROW 2 CELLS (1-30) ===');
  const r2 = sku.getRow(2);
  for (let c = 1; c <= 30; c++) {
    const val = r2.getCell(c).value;
    if (val !== null) {
      console.log(`Col ${c}: ${JSON.stringify(val)}`);
    }
  }

  console.log('\n=== ROW 37 CELLS (1-30) ===');
  const r37 = sku.getRow(37);
  for (let c = 1; c <= 30; c++) {
    const val = r37.getCell(c).value;
    if (val !== null) {
      console.log(`Col ${c}: ${JSON.stringify(val)}`);
    }
  }
}

main().catch(err => console.error(err));
