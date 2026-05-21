import ExcelJS from 'exceljs';
import fs from 'fs';

async function verify() {
  // Load our mapping from JSON
  const mappings = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\customer_mappings.json', 'utf8'));
  const segmentMap = new Map();
  for (const entry of mappings) {
    if (entry.customer && entry.customer !== 'TOTAL' && entry.segment) {
      segmentMap.set(entry.customer, entry.segment);
    }
  }
  
  const nameCorrections = new Map([
    ['Hộ kinh doanh Vũ Thị Hòa (NPP Ngọc Trâm)', 'Hộ kinh doanh Vũ Thị Hòa (WS Ngọc Trâm)'],
    ['NPP Như Trần Ngọc (cô Xuyến)', 'WS Như Trần Ngọc (cô Xuyến)'],
  ]);
  
  const dsrMap = new Map([
    ['Trương Minh Cảnh', 'TRUONG MINH CANH'],
    ['Nguyễn Tùng', 'TRUONG MINH CANH'],
    ['Kha Văn Long', 'TRUONG MINH CANH'],
    ['Nguyễn Văn Khải', 'TRUONG MINH CANH'],
    ['Đỗ Văn Chính', 'TRUONG MINH CANH'],
    ['Ngô Sỹ Hiệp', 'TRUONG MINH CANH'],
    ['Lê Quang Quân', 'TRUONG MINH CANH'],
    ['Tôn Long Gia Bảo', 'TRUONG MINH CANH'],
    ['Nguyễn Đăng Đạt', 'TRUONG MINH CANH'],
    ['Đỗ Quang Long', 'TRUONG MINH CANH'],
    ['Nguyễn Thị Triết Minh', 'TRUONG MINH CANH'],
    ['Lê Quang Chơn', 'TRUONG MINH CANH'],
    ['Trịnh Ngọc Hà', 'TRINH NGOC HA'],
    ['chị Trinh', 'TRINH NGOC HA'],
    ['Chị Trinh', 'TRINH NGOC HA'],
    ['Tạ Chiêu An', 'CHIEU AN'],
  ]);
  
  // Parse datamau.xlsx  
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\datamau.xlsx');
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) { console.error('No sheet'); return; }

  const monthRevenue = {};
  const dsrRevenue = {};
  let unmappedCount = 0;
  const unmappedCustomers = new Set();
  
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const month = row.getCell(9).value ? String(row.getCell(9).value).trim() : '';
    let customer = row.getCell(10).value ? String(row.getCell(10).value).trim() : '';
    const product = row.getCell(11).value ? String(row.getCell(11).value).trim() : '';
    const revenue = Number(row.getCell(18).value || 0);
    const rawDsr = row.getCell(21).value ? String(row.getCell(21).value).trim() : '';
    
    if (!month || !customer || !product) continue;
    
    customer = nameCorrections.get(customer) ?? customer;
    const segment = segmentMap.get(customer) ?? 'Unmapped';
    const dsr = dsrMap.get(rawDsr) ?? 'TRUONG MINH CANH';
    
    let segmentForQec = segment;
    if (customer === 'Nhà thuốc An Khang (Kho AK Đà Nẵng)' || customer === 'Nhà thuốc An Khang (Hub)') {
      segmentForQec = 'MT';
    } else if (customer === 'Nhà thuốc An Khang (Kho AK Bến Tre)') {
      segmentForQec = 'OTC';
    }
    
    const excludedFromQec = [
      'Hộ kinh doanh Vũ Thị Hòa (WS Ngọc Trâm)',
      'WS Như Trần Ngọc (cô Xuyến)',
      'Khách lẻ (Cảnh)',
      'Quầy thuốc Hoài Tính',
      'Chị Châu',
      'Bác sĩ Nguyễn Văn Hậu'
    ];
    
    const isExcluded = excludedFromQec.includes(customer);
    
    if (segment === 'Unmapped') {
      unmappedCount++;
      unmappedCustomers.add(customer);
    }
    
    if (!isExcluded) {
      if (!monthRevenue[segmentForQec]) monthRevenue[segmentForQec] = {};
      monthRevenue[segmentForQec][month] = (monthRevenue[segmentForQec][month] || 0) + revenue;
      
      if (!dsrRevenue[dsr]) dsrRevenue[dsr] = {};
      dsrRevenue[dsr][month] = (dsrRevenue[dsr][month] || 0) + revenue;
    }
  }

  // Read chuan.xlsx
  const wbChuan = new ExcelJS.Workbook();
  await wbChuan.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  const qec = wbChuan.getWorksheet('QEC review');
  if (!qec) { console.error('No QEC review'); return; }

  const months = [];
  for (let c = 7; c <= 22; c++) {
    const val = qec.getRow(1).getCell(c).value;
    if (val) months.push(String(val).trim());
  }

  console.log('=== SEGMENT COMPARISON ===');
  let matchCount = 0;
  let diffCount = 0;
  
  for (let segRow = 2; segRow <= 10; segRow++) {
    const segName = String(qec.getRow(segRow).getCell(6).value).trim();
    const ourValues = monthRevenue[segName] || {};
    
    for (let i = 0; i < months.length; i++) {
      const m = months[i];
      const chuanVal = Number(qec.getRow(segRow).getCell(7 + i).value || 0);
      const ourVal = ourValues[m] || 0;
      const diff = Math.abs(ourVal - chuanVal);
      if (diff > 1) {
        diffCount++;
        if (diffCount <= 20) {
          console.log(`  DIFF ${segName} ${m}: Ours=${Math.round(ourVal)}, Chuan=${Math.round(chuanVal)}, Diff=${Math.round(ourVal - chuanVal)}`);
        }
      } else {
        matchCount++;
      }
    }
  }

  console.log(`\nMatching: ${matchCount}, Differences: ${diffCount}`);
  console.log(`Unmapped transactions: ${unmappedCount}`);
  if (unmappedCustomers.size > 0) {
    console.log(`Unmapped customers (${unmappedCustomers.size}):`, [...unmappedCustomers].slice(0, 10).join(', '));
  }
  
  // Print DSR totals
  console.log('\n=== DSR TOTALS ===');
  for (const dsr of ['TRUONG MINH CANH', 'TRINH NGOC HA', 'CHIEU AN']) {
    const vals = dsrRevenue[dsr] || {};
    const total = Object.values(vals).reduce((s, v) => s + v, 0);
    console.log(`  ${dsr}: ${Math.round(total)}`);
  }
}

verify().catch(err => console.error(err));
