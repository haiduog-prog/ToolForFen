import * as xlsx from 'xlsx';
import * as fs from 'fs';

async function run() {
  const fileBuffer = fs.readFileSync('d:\\work\\ToolForFen\\input.xlsx');
  const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
  
  const sheetName = workbook.SheetNames[0];
  console.log('Sheet name:', sheetName);
  const worksheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  
  console.log('Total rows:', rows.length);
  for (let r = 0; r < Math.min(20, rows.length); r++) {
    console.log(`Row ${r + 1}:`, rows[r]);
  }
}

run().catch(console.error);
