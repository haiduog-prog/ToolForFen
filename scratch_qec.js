import ExcelJS from 'exceljs';

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('SKU - Customer review');
  console.log('Total rows in SKU - Customer review:', sheet.rowCount);
  
  let validSkuRows = 0;
  const otherCustomers = new Set();
  
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const c2 = row.getCell(2).value;
    const c4 = row.getCell(4).value;
    if (c4 !== null) {
      validSkuRows++;
    }
    if (c2 !== null) {
      otherCustomers.add(c2);
    }
  }
  
  console.log('Valid SKU rows (Col 4 not null):', validSkuRows);
  console.log('Other customer names found in Col 2 (rows 2+):', [...otherCustomers]);
}

main().catch(err => console.error(err));
