import ExcelJS from 'exceljs';
import fs from 'fs';

async function extractFullMappings() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  if (!qecReview) {
    console.error('QEC review sheet not found');
    return;
  }
  
  const mappings = [];
  
  for (let r = 29; r <= qecReview.rowCount; r++) {
    const row = qecReview.getRow(r);
    const customer = row.getCell(2).value ? String(row.getCell(2).value).trim() : '';
    const segment = row.getCell(3).value ? String(row.getCell(3).value).trim() : '';
    const dsr = row.getCell(4).value ? String(row.getCell(4).value).trim() : '';
    const city = row.getCell(5).value ? String(row.getCell(5).value).trim() : '';
    const region = row.getCell(6).value ? String(row.getCell(6).value).trim() : '';
    
    if (!customer || customer === 'Total' || customer === 'TOTAL') continue;
    
    mappings.push({
      customer,
      segment,
      dsr: dsr === 'null' ? '' : dsr,
      city: city === 'null' ? '' : city,
      region: region === 'null' ? '' : region
    });
  }
  
  console.log(`Extracted ${mappings.length} mappings.`);
  
  const dsrCounts = {};
  const regionCounts = {};
  for (const entry of mappings) {
    dsrCounts[entry.dsr] = (dsrCounts[entry.dsr] || 0) + 1;
    regionCounts[entry.region] = (regionCounts[entry.region] || 0) + 1;
  }
  
  console.log('DSR counts in QEC:', dsrCounts);
  console.log('Region counts in QEC:', regionCounts);
  
  fs.writeFileSync('d:\\Work\\ToolsForFen\\customer_mappings_full.json', JSON.stringify(mappings, null, 2), 'utf8');
  console.log('Saved to customer_mappings_full.json');
}

extractFullMappings().catch(err => console.error(err));
