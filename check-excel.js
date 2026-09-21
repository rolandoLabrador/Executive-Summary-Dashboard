const ExcelJS = require('exceljs');

async function check() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('output/OMNISHIELD_Executive_Report_2026-08-31.xlsx');
  
  const ws = workbook.getWorksheet('Dealer Dashboard');
  let count = 0;
  
  ws.eachRow((row, rowNumber) => {
    if (count++ > 10) return;
    const rowValues = [];
    row.eachCell({ includeEmpty: true }, (cell) => {
      rowValues.push(cell.value);
    });
    console.log(JSON.stringify(rowValues));
  });
}

check().catch(console.error);

