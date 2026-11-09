import * as ExcelJS from 'exceljs';
async function run() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('output/OMNISHIELD_Executive_Report_2026-08-31.xlsx');
  const ws = workbook.getWorksheet('Dealer Dashboard');
  for (let i = 4; i < 20; i++) {
    const dealerName = String(ws.getCell('C' + i).value);
    if (dealerName.includes('Power Motors East')) {
      console.log('Row values:', ws.getRow(i).values);
    }
  }
}
run().catch(console.error);
