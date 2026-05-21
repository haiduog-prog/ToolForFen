import ExcelJS from 'exceljs';

async function inspectCustomerSales() {
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
  
  // 1. Sum sales by customer in QEC review (Row 29 to end)
  const qecCustomerSales = {};
  for (const m of months) qecCustomerSales[m] = 0;
  
  const customersInQec = new Set();
  
  for (let r = 29; r <= qecReview.rowCount; r++) {
    const row = qecReview.getRow(r);
    const customer = row.getCell(2).value ? String(row.getCell(2).value).trim() : '';
    if (!customer || customer === 'Total' || customer === 'TOTAL') continue;
    
    customersInQec.add(customer);
    
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const val = Number(row.getCell(7 + i).value || 0);
      qecCustomerSales[m] += val;
    }
  }
  
  // 2. Find unique customers in Data nguồn
  const customersInDataNguon = new Set();
  const customerMonthlySales = {};
  
  for (let r = 2; r <= dataNguon.rowCount; r++) {
    const row = dataNguon.getRow(r);
    const customer = row.getCell(10).value ? String(row.getCell(10).value).trim() : '';
    const month = row.getCell(9).value ? String(row.getCell(9).value).trim() : '';
    const revenue = Number(row.getCell(18).value || 0);
    
    if (customer) {
      customersInDataNguon.add(customer);
      if (!customerMonthlySales[customer]) {
        customerMonthlySales[customer] = {};
        for (const m of months) customerMonthlySales[customer][m] = 0;
      }
      if (month && customerMonthlySales[customer][month] !== undefined) {
        customerMonthlySales[customer][month] += revenue;
      }
    }
  }
  
  console.log(`Unique customers in QEC review: ${customersInQec.size}`);
  console.log(`Unique customers in Data nguồn: ${customersInDataNguon.size}`);
  
  // Find customers in Data nguồn but not in QEC
  const missingInQec = [];
  for (const cust of customersInDataNguon) {
    if (!customersInQec.has(cust)) {
      missingInQec.push(cust);
    }
  }
  
  console.log(`\nCustomers in Data nguồn but missing in QEC review (${missingInQec.length}):`);
  console.log(missingInQec.slice(0, 20));
  
  // Check if missing customers account for the month 2024-01 difference (7,995,000)
  if (missingInQec.length > 0) {
    let missingSales01 = 0;
    for (const cust of missingInQec) {
      missingSales01 += customerMonthlySales[cust]['2024 -01'] || 0;
    }
    console.log(`\nTotal 2024-01 sales from missing customers: ${missingSales01}`);
  }
  
  console.log('\nMonth     | Seg Total   | Customer Sum');
  console.log('--------------------------------------');
  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    const segTot = Number(qecReview.getRow(11).getCell(7 + i).value || 0);
    console.log(`${m} | ${Math.round(segTot).toString().padStart(11)} | ${Math.round(qecCustomerSales[m]).toString().padStart(12)}`);
  }
}

inspectCustomerSales().catch(err => console.error(err));
