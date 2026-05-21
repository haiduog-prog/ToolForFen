import fs from 'fs';
import ExcelJS from 'exceljs';
import { readSourceWorkbook } from './src/infrastructure/excel/readSourceWorkbook.js';
import { lookupSegment } from './src/domain/customerMapping.js';

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
  
  for (const t of feb2025Tx) {
    if (EXCLUDED.includes(t.customer)) continue;
    
    let segForQec = lookupSegment(t.customer);
    if (t.customer === "Nhà thuốc An Khang (Kho AK Đà Nẵng)" || t.customer === "Nhà thuốc An Khang (Hub)") {
      segForQec = "MT";
    } else if (t.customer === "Nhà thuốc An Khang (Kho AK Bến Tre)") {
      segForQec = "OTC";
    }
    
    // We are looking for something that is in "AK" in chuan.xlsx but maybe mapped to somewhere else here,
    // or maybe t.revenue = 53295000
    if (t.revenue === 53295000 || segForQec === 'AK' || t.customer.includes('An Khang')) {
      console.log(`Customer: ${t.customer} | Revenue: ${t.revenue} | Segment lookup: ${lookupSegment(t.customer)} | SegForQec: ${segForQec} | Row: ${t.rowNumber}`);
    }
  }
}

main().catch(err => console.error(err));
