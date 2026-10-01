import * as fs from 'fs';
const text = fs.readFileSync('output/ContractExport (79).csv', 'utf8');
const lines = text.split('\n');
let sum = 0;
for (const line of lines) {
  if (line.includes('Power Motors East') || line.includes('MS370')) {
    // The CSV has WrittenAmount.RESERVE.ADMIN
    const match = line.match(/1360/);
    if (match) {
      console.log('Found 1360 in line!', line.substring(0, 100));
    }
  }
}
