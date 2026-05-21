import ExcelJS from 'exceljs';
import fs from 'fs';

async function solveExactMapping() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  if (!qecReview) {
    console.error('QEC review sheet not found');
    return;
  }
  
  const months = ['2024 -01', '2024 -02', '2024 -03', '2024 -04', '2024 -05', '2024 -06', '2024 -07'];
  
  // 1. Read target Region sales (months 1-7)
  const targetRegions = {};
  const regionNames = [
    '1. HA NOI',
    '2. N-PROVINCE',
    '3. CENTRAL',
    '4. S-EAST',
    '5. HCM',
    '6. MKD'
  ];
  for (let r = 14; r <= 19; r++) {
    const row = qecReview.getRow(r);
    const name = String(row.getCell(6).value).trim();
    if (regionNames.includes(name)) {
      const sales = [];
      for (let c = 7; c <= 13; c++) {
        sales.push(Number(row.getCell(c).value || 0));
      }
      targetRegions[name] = sales;
    }
  }
  
  // 2. Read target DSR sales (months 1-7)
  const targetDsrs = {};
  const dsrNames = [
    'TRUONG MINH CANH',
    'TRINH NGOC HA',
    'CHIEU AN'
  ];
  for (let r = 23; r <= 25; r++) {
    const row = qecReview.getRow(r);
    const name = String(row.getCell(6).value).trim();
    if (dsrNames.includes(name)) {
      const sales = [];
      for (let c = 7; c <= 13; c++) {
        sales.push(Number(row.getCell(c).value || 0));
      }
      targetDsrs[name] = sales;
    }
  }
  
  // 3. Read customer sales (months 1-7)
  const customers = [];
  for (let r = 29; r <= qecReview.rowCount; r++) {
    const row = qecReview.getRow(r);
    const name = row.getCell(2).value ? String(row.getCell(2).value).trim() : '';
    const segment = row.getCell(3).value ? String(row.getCell(3).value).trim() : '';
    if (!name || name === 'Total' || name === 'TOTAL') continue;
    
    const sales = [];
    for (let c = 7; c <= 13; c++) {
      sales.push(Number(row.getCell(c).value || 0));
    }
    
    customers.push({ name, segment, sales });
  }
  
  console.log(`Loaded ${customers.length} customers.`);
  
  // 4. Solve for Region and DSR mappings.
  // We can solve this as a constrained integer programming or matching problem.
  // Since the monthly sales of customers are very specific, we can try to find an assignment.
  // Actually, we can also look at the DSR thô and Address of each customer in Data nguồn to build a heuristic,
  // then verify if the resulting sums match the targets!
  // Let's analyze DSR first:
  // In Data nguồn, which raw DSRs are associated with which customer?
  // Let's read Data nguồn to build the raw mapping:
  const dataNguon = wb.getWorksheet('Data nguồn');
  const customerRawInfo = {}; // customerName -> { rawDsrs: Set, addresses: Set, đơnVịGiaos: Set }
  
  for (let r = 2; r <= dataNguon.rowCount; r++) {
    const row = dataNguon.getRow(r);
    let customer = row.getCell(10).value ? String(row.getCell(10).value).trim() : '';
    
    // Apply name corrections to align with QEC review names
    if (customer === 'Hộ kinh doanh Vũ Thị Hòa (NPP Ngọc Trâm)') {
      customer = 'Hộ kinh doanh Vũ Thị Hòa (WS Ngọc Trâm)';
    } else if (customer === 'NPP Như Trần Ngọc (cô Xuyến)') {
      customer = 'WS Như Trần Ngọc (cô Xuyến)';
    }
    
    const rawDsr = row.getCell(21).value ? String(row.getCell(21).value).trim() : '';
    const address = row.getCell(24).value ? String(row.getCell(24).value).trim() : '';
    const donViGiao = row.getCell(23).value ? String(row.getCell(23).value).trim() : '';
    
    if (!customer) continue;
    
    if (!customerRawInfo[customer]) {
      customerRawInfo[customer] = {
        rawDsrs: new Set(),
        addresses: new Set(),
        donViGiaos: new Set()
      };
    }
    if (rawDsr) customerRawInfo[customer].rawDsrs.add(rawDsr);
    if (address) customerRawInfo[customer].addresses.add(address);
    if (donViGiao) customerRawInfo[customer].donViGiaos.add(donViGiao);
  }
  
  // Let's define the rule to map Customer to DSR and Region:
  // Rule for DSR:
  // - If the customer has raw DSR 'Trịnh Ngọc Hà' or 'chị Trinh' or 'Chị Trinh' -> 'TRINH NGOC HA'
  // - If the customer has raw DSR 'Tạ Chiêu An' -> 'CHIEU AN' (unless they also have Trương Minh Cảnh? Let's check)
  // - If the customer has raw DSR 'Trương Minh Cảnh' or others -> 'TRUONG MINH CANH'
  // Wait, let's write a function to map a customer and verify if the sum of monthly sales of the mapped customers
  // perfectly matches the target DSR sales!
  
  // Let's test this mapping logic:
  const dsrMap = {};
  const regionMap = {};
  
  for (const cust of customers) {
    const info = customerRawInfo[cust.name] || { rawDsrs: new Set(), addresses: new Set(), donViGiaos: new Set() };
    const rawDsrs = Array.from(info.rawDsrs);
    
    // Heuristic for DSR
    let mappedDsr = 'TRUONG MINH CANH'; // Default
    if (rawDsrs.includes('Trịnh Ngọc Hà') || rawDsrs.includes('chị Trinh') || rawDsrs.includes('Chị Trinh')) {
      mappedDsr = 'TRINH NGOC HA';
    } else if (rawDsrs.includes('Tạ Chiêu An') && !rawDsrs.includes('Trương Minh Cảnh')) {
      mappedDsr = 'CHIEU AN';
    }
    
    // Wait, let's look at the actual target sales. CHIEU AN target sales are ALL 0.
    // So CHIEU AN mapped customers must have ALL 0 sales in the first 7 months!
    // Let's check if Tạ Chiêu An thô has 0 sales in the first 7 months?
    // In our previous run: Raw DSR [Tạ Chiêu An] monthly sales: [0, 0, 5192000, 27250000, 0, 12480000, 8490000].
    // Wait! If Tạ Chiêu An thô has sales in month 3, 4, 6, 7 but CHIEU AN target DSR is 0,
    // then the transactions of Tạ Chiêu An thô MUST be mapped to TRUONG MINH CANH or TRINH NGOC HA!
    // Yes! Let's check who handles 'Tạ Chiêu An' transactions.
    // If CHIEU AN has 0 target sales, then NO customers with non-zero sales can be mapped to CHIEU AN in the first 7 months!
    // Actually, let's write a solver that tries to assign each customer to a Region and DSR to match the target exactly!
    
    // Since we have 295 customers and only 7 months, we can formulate this as a subset-matching problem.
    // Let's find for each customer: which Region do they belong to?
    // Let's look at the customer's addresses and cities:
    // Region candidates: '1. HA NOI', '2. N-PROVINCE', '3. CENTRAL', '4. S-EAST', '5. HCM', '6. MKD'
    // Let's check:
    // - If city / address contains 'Hà Nội' or northern provinces -> '1. HA NOI' or '2. N-PROVINCE'
    // - If city / address contains 'Hồ Chí Minh', 'HCM', 'Quận' -> '5. HCM'
    // - Let's see if we can deduce this from address.
    // Let's write an algorithm that assigns them and checks if it matches!
  }
  
  // Let's write a script to search for the perfect mapping by solving month-by-month equations.
  // Actually, we can do this easily. We can write a script that does:
  // For each customer, check which Region's target sales can absorb the customer's sales.
  // If a customer has sales in a month where Region R is 0, then this customer CANNOT be in Region R.
  // This drastically reduces the candidates!
  
  console.log('Target Region sales for 7 months:');
  for (const r of regionNames) {
    console.log(`${r}:`, targetRegions[r]);
  }
  
  // Let's find candidates for each customer
  const customerRegionCandidates = {};
  for (const cust of customers) {
    const candidates = [];
    for (const r of regionNames) {
      let fit = true;
      for (let i = 0; i < 7; i++) {
        if (cust.sales[i] > 0 && targetRegions[r][i] === 0) {
          fit = false;
          break;
        }
      }
      if (fit) candidates.push(r);
    }
    customerRegionCandidates[cust.name] = candidates;
  }
  
  // Let's print out how many candidates each customer has
  const candidateCounts = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
  for (const cust of customers) {
    const cnt = customerRegionCandidates[cust.name].length;
    candidateCounts[cnt]++;
  }
  console.log('Customer Region candidate counts:', candidateCounts);
  
  // Let's see if we can resolve the mapping uniquely!
}

solveExactMapping().catch(err => console.error(err));
