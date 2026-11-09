import * as ExcelJS from 'exceljs';
async function run() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('output/OMNISHIELD_Executive_Report_2026-08-31_v2.xlsx');
  const ws = workbook.getWorksheet('Dealer Dashboard');
  console.log('Headers Row 5:', ws.getRow(5).values);
  console.log('Row 6 (MS370?):', ws.getRow(6).values);
  console.log('Row 7 (Another?):', ws.getRow(7).values);
}
run().catch(console.error);
