import ExcelJS from 'exceljs';
import fs from 'fs';

async function findMtAkDiffs() {
  const mappings = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\customer_mappings.json', 'utf8'));
  const segmentMap = new Map();
  for (const entry of mappings) {
    if (entry.customer && entry.customer !== 'TOTAL' && entry.segment) {
      segmentMap.set(entry.customer, entry.segment);
    }
  }

  // Read chuan.xlsx QEC customer details
  const wbChuan = new ExcelJS.Workbook();
  await wbChuan.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  const qec = wbChuan.getWorksheet('QEC review');
  
  const chuanCustomerSegments = new Map();
  for (let r = 29; r <= qec.rowCount; r++) {
    const row = qec.getRow(r);
    const customer = row.getCell(2).value ? String(row.getCell(2).value).trim() : '';
    const segment = row.getCell(3).value ? String(row.getCell(3).value).trim() : '';
    if (customer && customer !== 'Total' && customer !== 'TOTAL') {
      chuanCustomerSegments.set(customer, segment);
    }
  }

  // Find customers where our mapping differs from chuan QEC detail
  console.log('=== SEGMENT MAPPING DIFFERENCES ===');
  let diffCount = 0;
  for (const [customer, chuanSegment] of chuanCustomerSegments) {
    const ourSegment = segmentMap.get(customer) ?? 'Unmapped';
    if (ourSegment !== chuanSegment && chuanSegment) {
      diffCount++;
      console.log(`  "${customer}": Ours=${ourSegment}, Chuan=${chuanSegment}`);
    }
  }
  console.log(`Total differences: ${diffCount}`);
  
  // Check specifically for AK/MT customers
  console.log('\n=== AK/MT CUSTOMERS IN MAPPING ===');
  const akCustomers = [];
  const mtCustomers = [];
  for (const [cust, seg] of segmentMap) {
    if (seg === 'AK') akCustomers.push(cust);
    if (seg === 'MT') mtCustomers.push(cust);
  }
  console.log('AK customers:', akCustomers);
  console.log('MT customers:', mtCustomers);
}

findMtAkDiffs().catch(err => console.error(err));
