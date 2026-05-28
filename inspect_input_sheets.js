import fs from "fs";
import xlsx from "xlsx";

async function inspect(file) {
  try {
    const buf = fs.readFileSync(`d:\\Work\\ToolsForFen\\${file}`);
    const workbook = xlsx.read(buf, { type: "buffer" });
    
    console.log(`\n=== SHEETS IN ${file} ===`);
    workbook.SheetNames.forEach((name, idx) => {
      const sheet = workbook.Sheets[name];
      const ref = sheet["!ref"] || "";
      console.log(`  Sheet ${idx + 1}: "${name}" (Ref: ${ref})`);
    });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    console.log(`=== ROWS 1-5 IN ${sheetName} ===`);
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      console.log(`Row ${r + 1}: ${rows[r].slice(0, 15).join(" | ")}`);
    }
  } catch (err) {
    console.error(`Error inspecting ${file}:`, err);
  }
}

async function run() {
  await inspect("input.xlsx");
  await inspect("test_input_export.xlsx");
}

run().catch(console.error);
