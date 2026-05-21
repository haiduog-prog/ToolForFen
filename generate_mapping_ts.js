import fs from 'fs';

const raw = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\customer_mappings.json', 'utf8'));
const pcsProds = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\scratch_pcs_products.json', 'utf8'));
const vndProds = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\scratch_vnd_products.json', 'utf8'));

// Filter out TOTAL and empty entries
const entries = raw.filter(e => e.customer && e.customer !== 'TOTAL' && e.segment);

const lines = [];
lines.push('import { UNMAPPED_SEGMENT } from "./entities";');
lines.push('');
lines.push('/**');
lines.push(' * Static customer list in exact order as chuan.xlsx QEC review worksheet.');
lines.push(' */');
lines.push('export const STATIC_CUSTOMERS: string[] = [');
for (const entry of entries) {
  const escaped = entry.customer.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  lines.push(`  "${escaped}",`);
}
lines.push('];');
lines.push('');
lines.push('/**');
lines.push(' * Static PCS product list in exact order as chuan.xlsx SKU review worksheet.');
lines.push(' */');
lines.push('export const STATIC_PCS_PRODUCTS: (string | null)[] = [');
for (const p of pcsProds) {
  if (p === null) {
    lines.push('  null,');
  } else {
    const escaped = p.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    lines.push(`  "${escaped}",`);
  }
}
lines.push('];');
lines.push('');
lines.push('/**');
lines.push(' * Static VND product list in exact order as chuan.xlsx SKU review worksheet.');
lines.push(' */');
lines.push('export const STATIC_VND_PRODUCTS: (string | null)[] = [');
for (const p of vndProds) {
  if (p === null) {
    lines.push('  null,');
  } else {
    const escaped = p.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    lines.push(`  "${escaped}",`);
  }
}
lines.push('];');
lines.push('');
lines.push('/**');
lines.push(' * Static customer → segment mapping extracted from chuan.xlsx QEC review (row 29+).');
lines.push(' * 296 unique customers mapped to their business segment.');
lines.push(' */');
lines.push('const CUSTOMER_SEGMENT_MAP = new Map<string, string>([');

for (const entry of entries) {
  const escaped = entry.customer.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  lines.push(`  ["${escaped}", "${entry.segment}"],`);
}

lines.push(']);');
lines.push('');
lines.push('/**');
lines.push(' * Customer name corrections: raw Data nguồn names that differ from QEC review mapping.');
lines.push(' */');
lines.push('const NAME_CORRECTIONS = new Map<string, string>([');
lines.push('  ["Hộ kinh doanh Vũ Thị Hòa (NPP Ngọc Trâm)", "Hộ kinh doanh Vũ Thị Hòa (WS Ngọc Trâm)"],');
lines.push('  ["NPP Như Trần Ngọc (cô Xuyến)", "WS Như Trần Ngọc (cô Xuyến)"],');
lines.push(']);');
lines.push('');
lines.push('/**');
lines.push(' * Normalize customer name by applying known corrections.');
lines.push(' */');
lines.push('export function normalizeCustomerName(name: string): string {');
lines.push('  return NAME_CORRECTIONS.get(name) ?? name;');
lines.push('}');
lines.push('');
lines.push('/**');
lines.push(' * Look up segment for a customer. Returns UNMAPPED_SEGMENT if not found.');
lines.push(' */');
lines.push('export function lookupSegment(customerName: string): string {');
lines.push('  return CUSTOMER_SEGMENT_MAP.get(customerName) ?? UNMAPPED_SEGMENT;');
lines.push('}');
lines.push('');
lines.push('/**');
lines.push(' * Map raw DSR name ("Người thực hiện") to standardized DSR group.');
lines.push(' */');
lines.push('const DSR_MAP = new Map<string, string>([');
lines.push('  ["Trương Minh Cảnh", "TRUONG MINH CANH"],');
lines.push('  ["Nguyễn Tùng", "TRUONG MINH CANH"],');
lines.push('  ["Kha Văn Long", "TRUONG MINH CANH"],');
lines.push('  ["Nguyễn Văn Khải", "TRUONG MINH CANH"],');
lines.push('  ["Đỗ Văn Chính", "TRUONG MINH CANH"],');
lines.push('  ["Ngô Sỹ Hiệp", "TRUONG MINH CANH"],');
lines.push('  ["Lê Quang Quân", "TRUONG MINH CANH"],');
lines.push('  ["Tôn Long Gia Bảo", "TRUONG MINH CANH"],');
lines.push('  ["Nguyễn Đăng Đạt", "TRUONG MINH CANH"],');
lines.push('  ["Đỗ Quang Long", "TRUONG MINH CANH"],');
lines.push('  ["Nguyễn Thị Triết Minh", "TRUONG MINH CANH"],');
lines.push('  ["Lê Quang Chơn", "TRUONG MINH CANH"],');
lines.push('  ["Trịnh Ngọc Hà", "TRINH NGOC HA"],');
lines.push('  ["chị Trinh", "TRINH NGOC HA"],');
lines.push('  ["Chị Trinh", "TRINH NGOC HA"],');
lines.push('  ["Tạ Chiêu An", "CHIEU AN"],');
lines.push(']);');
lines.push('');
lines.push('export const DSR_ORDER = ["TRUONG MINH CANH", "TRINH NGOC HA", "CHIEU AN"];');
lines.push('');
lines.push('export function lookupDsr(rawDsr: string): string {');
lines.push('  const trimmed = rawDsr.trim();');
lines.push('  return DSR_MAP.get(trimmed) ?? "TRUONG MINH CANH";');
lines.push('}');
lines.push('');

fs.writeFileSync('d:\\Work\\ToolsForFen\\src\\domain\\customerMapping.ts', lines.join('\n') + '\n', 'utf8');
console.log(`Generated customerMapping.ts with customers, PCS products, and VND products lists.`);
