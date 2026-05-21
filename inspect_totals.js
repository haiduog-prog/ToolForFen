import ExcelJS from 'exceljs';

async function inspectTotals() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const dataNguon = wb.getWorksheet('Data nguồn');
  const qecReview = wb.getWorksheet('QEC review');
  
  if (!dataNguon || !qecReview) {
    console.error('Worksheets not found');
    return;
  }
  
  const months = [];
  const headerRow = qecReview.getRow(1);
  for (let c = 7; c <= 22; c++) {
    const val = headerRow.getCell(c).value;
    if (val) months.push(String(val).trim());
  }
  
  // 1. Calculate raw sales by month from Data nguồn
  const rawSales = {};
  for (const m of months) rawSales[m] = 0;
  
  for (let r = 2; r <= dataNguon.rowCount; r++) {
    const row = dataNguon.getRow(r);
    const month = row.getCell(9).value ? String(row.getCell(9).value).trim() : '';
    const revenue = Number(row.getCell(18).value || 0);
    
    if (month && rawSales[month] !== undefined) {
      rawSales[month] += revenue;
    }
  }
  
  // 2. Read QEC Totals
  const segmentTotals = [];
  const regionTotals = [];
  const dsrTotals = [];
  
  const segRow = qecReview.getRow(11);
  const regRow = qecReview.getRow(20);
  const dsrRow = qecReview.getRow(26);
  
  for (let c = 7; c < 7 + months.length; c++) {
    segmentTotals.push(Number(segRow.getCell(c).value || 0));
    regionTotals.push(Number(regRow.getCell(c).value || 0));
    dsrTotals.push(Number(dsrRow.getCell(c).value || 0));
  }
  
  console.log('Month     | Raw Sales   | Seg Total   | Reg Total   | DSR Total');
  console.log('---------------------------------------------------------------');
  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    console.log(`${m} | ${Math.round(rawSales[m]).toString().padStart(11)} | ${Math.round(segmentTotals[i]).toString().padStart(11)} | ${Math.round(regionTotals[i]).toString().padStart(11)} | ${Math.round(dsrTotals[i]).toString().padStart(11)}`);
  }
}

inspectTotals().catch(err => console.error(err));
