import ExcelJS from 'exceljs';

async function inspectSourceData() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) {
    console.error('Data nguồn sheet not found');
    return;
  }
  
  const dsrs = new Set();
  const addressSamples = [];
  const dsrToAddress = {};
  const customerToDsr = {};
  
  console.log('Total rows in Data nguồn:', sheet.rowCount);
  
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const customer = row.getCell(10).value ? String(row.getCell(10).value).trim() : '';
    const dsr = row.getCell(21).value ? String(row.getCell(21).value).trim() : '';
    const address = row.getCell(24).value ? String(row.getCell(24).value).trim() : '';
    const city = row.getCell(23).value ? String(row.getCell(23).value).trim() : ''; // Đơn vị giao
    
    if (dsr) dsrs.add(dsr);
    
    if (dsr && address) {
      if (!dsrToAddress[dsr]) dsrToAddress[dsr] = new Set();
      dsrToAddress[dsr].add(address.substring(0, 30)); // Lấy 30 ký tự đầu của địa chỉ
    }
    
    if (customer && dsr) {
      if (!customerToDsr[customer]) customerToDsr[customer] = new Set();
      customerToDsr[customer].add(dsr);
    }
  }
  
  console.log('\n=== UNIQUE DSRs IN DATA NGUON ===');
  console.log(Array.from(dsrs));
  
  console.log('\n=== DSR TO ADDRESS SAMPLES ===');
  for (const dsr of Object.keys(dsrToAddress)) {
    console.log(`DSR: ${dsr}, unique address prefixes (first 5):`, Array.from(dsrToAddress[dsr]).slice(0, 5));
  }
  
  console.log('\n=== CUSTOMER TO DSR SAMPLES (first 10) ===');
  const customers = Object.keys(customerToDsr);
  for (const cust of customers.slice(0, 10)) {
    console.log(`Customer: ${cust} -> DSRs:`, Array.from(customerToDsr[cust]));
  }
}

inspectSourceData().catch(err => console.error(err));
