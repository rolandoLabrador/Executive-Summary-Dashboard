import * as ExcelJS from 'exceljs';
async function run() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('output/OMNISHIELD_Executive_Report_2026-08-31_v2.xlsx');
  const ws = workbook.getWorksheet('Debug Math Breakdown');
  
  let inPowerMotors = false;
  for (let i = 1; i < 5000; i++) {
    const valA = String(ws.getCell('A' + i).value || '');
    const valB = String(ws.getCell('B' + i).value || '');
    const valC = String(ws.getCell('C' + i).value || '');
    if (valA.includes('Power Motors East')) {
      inPowerMotors = true;
      console.log(`Found Power Motors East at row ${i}`);
    } else if (inPowerMotors && valC.includes('TOTAL NET RESERVE:')) {
       console.log('TOTAL NET RESERVE:', ws.getCell('F' + i).value); 
       console.log(`Row ${i}:`, ws.getRow(i).values);
    } else if (inPowerMotors && valC.includes('TOTAL NET ADMIN:')) {
       console.log('TOTAL NET ADMIN:', ws.getCell('F' + i).value);
       console.log(`Row ${i}:`, ws.getRow(i).values);
       break; // Found what we need
    } else if (inPowerMotors && valA.startsWith('---')) {
       break;
    }
    
    if (inPowerMotors) {
       if (valC.includes('ADMIN') || valB.includes('ADMIN')) {
          console.log(`Row ${i}:`, ws.getRow(i).values);
       }
    }
  }
}
run().catch(console.error);
