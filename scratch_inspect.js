import ExcelJS from 'exceljs';

async function printHeaders() {
  const chuanPath = 'd:\\Work\\ToolsForFen\\chuan.xlsx';
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(chuanPath);
  
  const skuReview = wb.getWorksheet('SKU review');
  if (skuReview) {
    console.log('=== SKU REVIEW TOP 5 ROWS ===');
    for (let r = 1; r <= 5; r++) {
      const row = skuReview.getRow(r);
      const vals = [];
      for (let c = 1; c <= 22; c++) {
        const val = row.getCell(c).value;
        if (val !== null) vals.push(`C${c}: ${val}`);
      }
      console.log(`Row ${r}:`, vals.join(' | '));
    }
  }

  const skuCustReview = wb.getWorksheet('SKU - Customer review');
  if (skuCustReview) {
    console.log('=== SKU - CUSTOMER REVIEW TOP 5 ROWS ===');
    for (let r = 1; r <= 5; r++) {
      const row = skuCustReview.getRow(r);
      const vals = [];
      for (let c = 1; c <= 22; c++) {
        const val = row.getCell(c).value;
        if (val !== null) vals.push(`C${c}: ${val}`);
      }
      console.log(`Row ${r}:`, vals.join(' | '));
    }
  }
}

printHeaders().catch(err => console.error(err));
