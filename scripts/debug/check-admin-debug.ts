import * as ExcelJS from 'exceljs';
async function run() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('output/OMNISHIELD_Executive_Report_2026-08-31_v2.xlsx');
  const ws = workbook.getWorksheet('Debug Math Breakdown');
  for (let i = 1; i < 500; i++) {
    if (String(ws.getCell('C' + i).value) === 'TOTAL NET RESERVE:') {
      console.log('Total Net Reserve in Debug:', ws.getCell('F' + i).value);
      break;
    }
  }
}
run().catch(console.error);
