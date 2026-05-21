import ExcelJS from 'exceljs';

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('d:\\Work\\ToolsForFen\\chuan.xlsx');
  
  const qec = wb.getWorksheet('QEC review');
  console.log('=== ROW LABELS (1-20) ===');
  for (let i = 1; i <= 20; i++) {
    const row = qec.getRow(i);
    const label = row.getCell(6).value;
    console.log(`Row ${String(i).padStart(2, ' ')}: Label=${JSON.stringify(label)}`);
  }
}

main().catch(err => console.error(err));
