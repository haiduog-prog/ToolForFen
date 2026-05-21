import ExcelJS from 'exceljs';

async function findCustomerDiffs() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  const dataNguon = wb.getWorksheet('Data nguồn');
  
  if (!qecReview || !dataNguon) {
    console.error('Worksheets not found');
    return;
  }
  
  const months = [];
  const headerRow = qecReview.getRow(1);
  for (let c = 7; c <= 22; c++) {
    const val = headerRow.getCell(c).value;
    if (val) months.push(String(val).trim());
  }
  
  // 1. Read QEC customer monthly sales
  const qecSales = {}; // customer -> month -> val
  for (let r = 29; r <= qecReview.rowCount; r++) {
    const row = qecReview.getRow(r);
    const customer = row.getCell(2).value ? String(row.getCell(2).value).trim() : '';
    if (!customer || customer === 'Total' || customer === 'TOTAL') continue;
    
    if (!qecSales[customer]) {
      qecSales[customer] = {};
      for (const m of months) qecSales[customer][m] = 0;
    }
    
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const val = Number(row.getCell(7 + i).value || 0);
      qecSales[customer][m] += val;
    }
  }
  
  // 2. Read raw sales by customer and month from Data nguồn
  const rawSales = {}; // customer -> month -> val
  for (let r = 2; r <= dataNguon.rowCount; r++) {
    const row = dataNguon.getRow(r);
    const customer = row.getCell(10).value ? String(row.getCell(10).value).trim() : '';
    const month = row.getCell(9).value ? String(row.getCell(9).value).trim() : '';
    const revenue = Number(row.getCell(18).value || 0);
    
    if (!customer || !month) continue;
    
    if (!rawSales[customer]) {
      rawSales[customer] = {};
      for (const m of months) rawSales[customer][m] = 0;
    }
    
    if (rawSales[customer][month] !== undefined) {
      rawSales[customer][month] += revenue;
    }
  }
  
  // 3. Compare and print differences
  console.log('=== CUSTOMER DIFFERENCES (RAW vs QEC) ===');
  let diffCount = 0;
  
  const allCustomers = new Set([...Object.keys(qecSales), ...Object.keys(rawSales)]);
  
  for (const customer of allCustomers) {
    const qSales = qecSales[customer] || {};
    const rSales = rawSales[customer] || {};
    
    const diffs = [];
    for (const m of months) {
      const qVal = qSales[m] || 0;
      const rVal = rSales[m] || 0;
      const diff = rVal - qVal;
      if (Math.abs(diff) > 0.01) {
        diffs.push(`${m}: Raw=${Math.round(rVal)}, Qec=${Math.round(qVal)}, Diff=${Math.round(diff)}`);
      }
    }
    
    if (diffs.length > 0) {
      diffCount++;
      console.log(`\nCustomer [${customer}]:`);
      console.log('  ' + diffs.join('\n  '));
    }
  }
  
  console.log(`\nTotal customers with differences: ${diffCount}`);
}

findCustomerDiffs().catch(err => console.error(err));
