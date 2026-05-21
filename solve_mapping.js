import ExcelJS from 'exceljs';

async function solveMapping() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const dataNguon = wb.getWorksheet('Data nguồn');
  const qecReview = wb.getWorksheet('QEC review');
  
  if (!dataNguon || !qecReview) {
    console.error('Worksheets not found');
    return;
  }
  
  // 1. Read monthly columns from QEC review
  // Header row 1 has months
  const months = [];
  const headerRow = qecReview.getRow(1);
  for (let c = 7; c <= 22; c++) {
    const val = headerRow.getCell(c).value;
    if (val) months.push(String(val).trim());
  }
  console.log('QEC months:', months);
  
  // Read DSR rows in QEC
  const qecDsrs = {};
  for (let r = 23; r <= 25; r++) {
    const row = qecReview.getRow(r);
    const name = String(row.getCell(6).value).trim();
    const monthlyVals = [];
    for (let c = 7; c < 7 + months.length; c++) {
      monthlyVals.push(Number(row.getCell(c).value || 0));
    }
    qecDsrs[name] = monthlyVals;
  }
  
  // Read Region rows in QEC
  const qecRegions = {};
  for (let r = 14; r <= 19; r++) {
    const row = qecReview.getRow(r);
    const name = String(row.getCell(6).value).trim();
    const monthlyVals = [];
    for (let c = 7; c < 7 + months.length; c++) {
      monthlyVals.push(Number(row.getCell(c).value || 0));
    }
    qecRegions[name] = monthlyVals;
  }
  
  // 2. Aggregate raw transactions in Data nguồn by DSR (Người thực hiện)
  const rawDsrMonthlySales = {};
  for (let r = 2; r <= dataNguon.rowCount; r++) {
    const row = dataNguon.getRow(r);
    const dsr = row.getCell(21).value ? String(row.getCell(21).value).trim() : 'UNSPECIFIED';
    const month = row.getCell(9).value ? String(row.getCell(9).value).trim() : '';
    const revenue = Number(row.getCell(18).value || 0);
    
    if (!month) continue;
    
    if (!rawDsrMonthlySales[dsr]) {
      rawDsrMonthlySales[dsr] = {};
      for (const m of months) rawDsrMonthlySales[dsr][m] = 0;
    }
    
    if (rawDsrMonthlySales[dsr][month] !== undefined) {
      rawDsrMonthlySales[dsr][month] += revenue;
    }
  }
  
  console.log('\n=== SOLVING DSR MAPPING ===');
  // Check which raw DSR maps to which QEC DSR by comparing monthly values
  const dsrMapping = {};
  for (const rawDsr of Object.keys(rawDsrMonthlySales)) {
    const rawSales = months.map(m => rawDsrMonthlySales[rawDsr][m]);
    
    // Find matching QEC DSR
    let matchedDsr = null;
    // We check if this raw DSR's sales are a subset of or match exactly any QEC DSR
    // More simply, let's see which QEC DSR is active when this raw DSR is active,
    // or let's calculate correlation/matrix.
    // Let's print raw DSR monthly values to analyze
    console.log(`Raw DSR [${rawDsr}] monthly sales:`, rawSales.map(v => Math.round(v)));
  }
  
  console.log('\n=== TARGET QEC DSR SALES ===');
  for (const name of Object.keys(qecDsrs)) {
    console.log(`Target DSR [${name}]:`, qecDsrs[name].map(v => Math.round(v)));
  }
  
  console.log('\n=== TARGET QEC REGION SALES ===');
  for (const name of Object.keys(qecRegions)) {
    console.log(`Target Region [${name}]:`, qecRegions[name].map(v => Math.round(v)));
  }
}

solveMapping().catch(err => console.error(err));
