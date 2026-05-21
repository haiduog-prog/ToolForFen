import fs from 'fs';

const mappings = JSON.parse(fs.readFileSync('d:\\Work\\ToolsForFen\\customer_mappings.json', 'utf8'));

console.log('Total entries:', mappings.length);

const segmentCounts = {};
const dsrCounts = {};
const regionCounts = {};
const validDsrOrRegion = [];

for (const entry of mappings) {
  segmentCounts[entry.segment] = (segmentCounts[entry.segment] || 0) + 1;
  dsrCounts[entry.dsr] = (dsrCounts[entry.dsr] || 0) + 1;
  regionCounts[entry.region] = (regionCounts[entry.region] || 0) + 1;
  
  if (entry.dsr !== '' || entry.region !== '') {
    validDsrOrRegion.push(entry);
  }
}

console.log('\nSegment counts:', segmentCounts);
console.log('DSR counts:', dsrCounts);
console.log('Region counts:', regionCounts);
console.log('\nEntries with valid DSR or Region (count):', validDsrOrRegion.length);
if (validDsrOrRegion.length > 0) {
  console.log('Sample valid entries:', validDsrOrRegion.slice(0, 10));
}
