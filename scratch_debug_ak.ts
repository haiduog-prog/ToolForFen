import fs from 'fs';
import ExcelJS from 'exceljs';
import { readSourceWorkbook } from './src/infrastructure/excel/readSourceWorkbook';
import { lookupSegment } from './src/domain/customerMapping';

async function main() {
  const buffer = fs.readFileSync("d:\\Work\\ToolsForFen\\datamau.xlsx");
  const mockFile = {
    name: "datamau.xlsx",
    arrayBuffer: async () => {
      const ab = new ArrayBuffer(buffer.length);
      const view = new Uint8Array(ab);
      for (let i = 0; i < buffer.length; ++i) {
        view[i] = buffer[i];
      }
      return ab;
    }
  };

  const parseResult = await readSourceWorkbook(mockFile as any);
  
  // Find transactions in 2025-02
  const feb2025Tx = parseResult.transactions.filter(t => t.month === '2025-02');
  console.log(`Total transactions in 2025-02: ${feb2025Tx.length}`);
  
  // Excluded list
  const EXCLUDED = [
    "Hộ kinh doanh Vũ Thị Hòa (WS Ngọc Trâm)",
    "WS Như Trần Ngọc (cô Xuyến)",
    "Khách lẻ (Cảnh)",
    "Quầy thuốc Hoài Tính",
    "Chị Châu",
    "Bác sĩ Nguyễn Văn Hậu"
  ];
  
  const mappings = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\customer_mappings.json', 'utf8'));
  const segmentMap = new Map();
  for (const entry of mappings) {
    if (entry.customer && entry.customer !== 'TOTAL' && entry.segment) {
      segmentMap.set(entry.customer, entry.segment);
    }
  }
  
  console.log(`segmentMap size: ${segmentMap.size}`);
  console.log(`First 5 keys in segmentMap:`, [...segmentMap.keys()].slice(0, 5));
  console.log(`AK Long Hậu in segmentMap:`, segmentMap.get("Nhà thuốc An Khang (Kho AK Long Hậu)"));
  console.log(`AK Trà Nóc in segmentMap:`, segmentMap.get("Nhà thuốc An Khang (Kho AK Trà Nóc)"));
  
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

  const wbData = new ExcelJS.Workbook();
  await wbData.xlsx.readFile('d:\\Work\\ToolsForFen\\datamau.xlsx');
  const sheet = wbData.getWorksheet('Data nguồn');
  if (!sheet) { console.error('No sheet'); return; }

  let totalAk = 0;
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
    
  const r403 = sheet.getRow(403);
  console.log(`=== ROW 403 CELLS ===`);
  for (let c = 1; c <= 30; c++) {
    console.log(`Col ${c}: ${JSON.stringify(r403.getCell(c).value)}`);
  }
  }
  console.log(`Sum of AK in February 2025 = ${totalAk}`);
}

main().catch(err => console.error(err));
