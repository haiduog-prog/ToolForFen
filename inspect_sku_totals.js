import ExcelJS from 'exceljs';

async function inspectSkuTotals() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  const skuReview = wb.getWorksheet('SKU review');
  
  if (!qecReview || !skuReview) {
    console.error('Worksheets not found');
    return;
  }
  
  // Read months
  const months = [];
  const headerRow = qecReview.getRow(1);
  for (let c = 7; c <= 22; c++) {
    const val = headerRow.getCell(c).value;
    if (val) months.push(String(val).trim());
  }
  
  // Sum monthly values in SKU review
  const skuMonthlySums = {};
  for (const m of months) skuMonthlySums[m] = 0;
  
  // Row 1 in SKU review is headers. We read from row 2 onwards.
  for (let r = 2; r <= skuReview.rowCount; r++) {
    const row = skuReview.getRow(r);
    const skuName = row.getCell(2).value;
    if (!skuName || skuName === 'Total' || skuName === 'TOTAL') continue;
    
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const val = Number(row.getCell(3 + i).value || 0); // Month values start at Col 3 in SKU review
      skuMonthlySums[m] += val;
    }
  }
  
  console.log('Month     | Seg Total   | Reg Total   | SKU Sum');
  console.log('-------------------------------------------------');
  const segRow = qecReview.getRow(11);
  const regRow = qecReview.getRow(20);
  
  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    const segTot = Number(segRow.getCell(7 + i).value || 0);
    const regTot = Number(regRow.getCell(7 + i).value || 0);
    console.log(`${m} | ${Math.round(segTot).toString().padStart(11)} | ${Math.round(regTot).toString().padStart(11)} | ${Math.round(skuMonthlySums[m]).toString().padStart(11)}`);
  }
}

inspectSkuTotals().catch(err => console.error(err));
