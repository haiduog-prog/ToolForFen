import ExcelJS from 'exceljs';

async function inspectMonths() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const sheet = wb.getWorksheet('Data nguồn');
  if (!sheet) {
    console.error('Data nguồn sheet not found');
    return;
  }
  
  const months = new Set();
  const dateSamples = [];
  
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const mVal = row.getCell(9).value;
    const dVal = row.getCell(8).value;
    if (mVal) months.add(String(mVal).trim());
    if (r <= 10) {
      dateSamples.push(`Row ${r}: Col 8=${dVal}, Col 9=${mVal}`);
    }
  }
  
  console.log('=== DATA NGUON MONTHS ===');
  console.log('Unique months in Col 9:', Array.from(months).sort());
  console.log('\n=== DATE & MONTH SAMPLES (first 9 rows) ===');
  console.log(dateSamples.join('\n'));
}

inspectMonths().catch(err => console.error(err));
