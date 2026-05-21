import ExcelJS from 'exceljs';

async function inspectSpecificCustomers() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  const dataNguon = wb.getWorksheet('Data nguồn');
  
  const targetNames = ['Khách lẻ (Cảnh)', 'Quầy thuốc Hoài Tính', 'Chị Châu', 'Bác sĩ Nguyễn Văn Hậu'];
  
  console.log('=== SPECIFIC CUSTOMERS IN QEC REVIEW ===');
  for (let r = 29; r <= qecReview.rowCount; r++) {
    const row = qecReview.getRow(r);
    const customer = row.getCell(2).value ? String(row.getCell(2).value).trim() : '';
    if (targetNames.includes(customer)) {
      console.log(`Row ${r}: Customer='${customer}', Segment='${row.getCell(3).value}', Tên DSR='${row.getCell(4).value}', Region='${row.getCell(6).value}'`);
    }
  }
  
  console.log('\n=== TRANSACTIONS IN DATA NGUON FOR TARGET CUSTOMERS ===');
  for (let r = 2; r <= dataNguon.rowCount; r++) {
    const row = dataNguon.getRow(r);
    const customer = row.getCell(10).value ? String(row.getCell(10).value).trim() : '';
    if (targetNames.includes(customer) || customer.includes('Vũ Thị Hòa') || customer.includes('Như Trần Ngọc')) {
      console.log(`Row ${r}: Customer='${customer}', Date='${row.getCell(9).value}', Product='${row.getCell(11).value}', Revenue=${row.getCell(18).value}, DSR='${row.getCell(21).value}', Status='${row.getCell(20).value}'`);
    }
  }
}

inspectSpecificCustomers().catch(err => console.error(err));
