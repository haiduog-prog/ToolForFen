import ExcelJS from 'exceljs';

async function inspectSpecificRows() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qecReview = wb.getWorksheet('QEC review');
  if (!qecReview) {
    console.error('QEC review sheet not found');
    return;
  }
  
  const rows = [320, 321, 322, 323];
  
  // Header row 1 months
  const months = [];
  const headerRow = qecReview.getRow(1);
  for (let c = 7; c <= 22; c++) {
    months.push(headerRow.getCell(c).value || `Col ${c}`);
  }
  
  console.log('Months in Header:', months.join(' | '));
  
  for (const rNum of rows) {
    const row = qecReview.getRow(rNum);
    const customer = row.getCell(2).value;
    const monthlyVals = [];
    for (let c = 7; c <= 22; c++) {
      monthlyVals.push(`${c}: ${row.getCell(c).value}`);
    }
    console.log(`\nRow ${rNum} [${customer}]:`);
    console.log(monthlyVals.join(' | '));
  }
}

inspectSpecificRows().catch(err => console.error(err));
