import ExcelJS from 'exceljs';
import fs from 'fs';

async function extractMapping() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  if (!qecReview) {
    console.error('QEC review sheet not found');
    return;
  }
  
  console.log('QEC Review sheet total rows:', qecReview.rowCount);
  
  const mappingList = [];
  // Read from row 28 onwards
  for (let r = 28; r <= qecReview.rowCount; r++) {
    const row = qecReview.getRow(r);
    
    // Check if the row is empty or doesn't have a customer name
    const customer = row.getCell(2).value;
    const segment = row.getCell(3).value;
    const dsr = row.getCell(4).value;
    const city = row.getCell(5).value;
    const region = row.getCell(6).value;
    
    if (customer && customer !== 'CUSTOMER_CHANGE' && customer !== 'Total') {
      mappingList.push({
        customer: String(customer).trim(),
        segment: segment ? String(segment).trim() : '',
        dsr: dsr ? String(dsr).trim() : '',
        city: city ? String(city).trim() : '',
        region: region ? String(region).trim() : ''
      });
    }
  }
  
  console.log(`Extracted ${mappingList.length} customer mapping entries.`);
  if (mappingList.length > 0) {
    console.log('Sample entry:', mappingList[0]);
    console.log('Last entry:', mappingList[mappingList.length - 1]);
    
    // Save to a json file in scratch
    fs.writeFileSync('d:\\Work\\ToolsForFen\\customer_mappings.json', JSON.stringify(mappingList, null, 2));
    console.log('Saved mappings to d:\\Work\\ToolsForFen\\customer_mappings.json');
  }
}

extractMapping().catch(err => console.error(err));
