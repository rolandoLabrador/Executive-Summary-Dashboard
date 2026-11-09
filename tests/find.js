const ExcelJS = require('exceljs');
async function run() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('output/OMNISHIELD_Executive_Report_2026-08-31.xlsx');
  const sheet = workbook.getWorksheet('Executive Dashboard');
  for (let i = 1; i <= 60; i++) {
    const val = String(sheet.getCell('A' + i).value);
    if (val.includes('FULL YEAR 2025')) console.log('FULL YEAR at', i);
  }
}
run();
