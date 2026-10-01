import * as fs from 'fs';
import * as path from 'path';

const file = path.join(__dirname, '../../src/services/excel.service.ts');
let content = fs.readFileSync(file, 'utf8');

const startMarker = '  private buildDebugMath(workbook: ExcelJS.Workbook, model: ReportModel): void {';
const endMarker = '  private buildDefinitions(workbook: ExcelJS.Workbook): void {';

const startIndex = content.indexOf(startMarker);
const endIndex = content.indexOf(endMarker);

if (startIndex === -1 || endIndex === -1) {
  throw new Error('Markers not found');
}

const newMethod = `  private buildDebugMath(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    if (debugDealers.length === 0) return;

    const ws = workbook.addWorksheet('Debug Math Breakdown');
    configureWorksheet(ws);
    title(
      ws,
      'DEBUG MATH: COMPONENT & KPI BREAKDOWN',
      'Shows exact component sums used to calculate Admin and Reserve for debugged dealers',
    );

    ws.columns = [
      { key: 'dealer', width: 45 },
      { key: 'category', width: 20 },
      { key: 'component', width: 40 },
      { key: 'written', width: 20 },
      { key: 'cancelled', width: 20 },
      { key: 'net', width: 20 },
    ];

    ws.getRow(4).values = [
      'Dealer',
      'Category',
      'Component',
      'Written Amount',
      'Cancelled Amount',
      'Net Amount',
    ];
    styleHeader(ws.getRow(4));

    // Calculate components per dealer
    for (const dealerName of debugDealers) {
      // Find ITD transactions for this dealer
      const dealerTx = model.contractTransactions.filter(
        (t) =>
          t.dealerName.toLowerCase() === dealerName ||
          t.dealerNumber.toLowerCase() === dealerName
      );

      if (dealerTx.length === 0) continue;

      const sums: Record<string, { written: number; cancelled: number; net: number }> = {};

      dealerTx.forEach((t) => {
        if (!t.components) return;
        
        const isCancel = t.transactionType === 'Cancellation';
        
        ['RESERVE', 'ADMIN'].forEach(cat => {
          if (!t.components[cat]) return;
          
          Object.entries(t.components[cat]).forEach(([comp, amt]) => {
            const key = \`\${cat}.\${comp}\`;
            if (!sums[key]) sums[key] = { written: 0, cancelled: 0, net: 0 };
            
            if (isCancel) {
              sums[key].cancelled -= Math.abs(amt);
              sums[key].net -= Math.abs(amt);
            } else {
              sums[key].written += amt;
              sums[key].net += amt;
            }
          });
        });
      });

      let startRow = ws.lastRow ? ws.lastRow.number + 2 : 5;
      
      const realDealerName = dealerTx[0].dealerName || dealerName;
      
      // Print RESERVE components
      let reserveNetTotal = 0;
      Object.keys(sums).filter(k => k.startsWith('RESERVE.')).sort().forEach(k => {
        if (sums[k].net === 0 && sums[k].written === 0) return;
        reserveNetTotal += sums[k].net;
        ws.addRow({
          dealer: realDealerName,
          category: 'RESERVE',
          component: k.replace('RESERVE.', ''),
          written: sums[k].written,
          cancelled: sums[k].cancelled,
          net: sums[k].net
        });
      });
      
      const reserveTotalRow = ws.addRow({
        dealer: '', category: '', component: 'TOTAL NET RESERVE:', net: reserveNetTotal
      });
      reserveTotalRow.font = { bold: true };
      ws.addRow({});

      // Print ADMIN components
      let adminNetTotal = 0;
      Object.keys(sums).filter(k => k.startsWith('ADMIN.')).sort().forEach(k => {
        if (sums[k].net === 0 && sums[k].written === 0) return;
        adminNetTotal += sums[k].net;
        ws.addRow({
          dealer: realDealerName,
          category: 'ADMIN',
          component: k.replace('ADMIN.', ''),
          written: sums[k].written,
          cancelled: sums[k].cancelled,
          net: sums[k].net
        });
      });

      const adminTotalRow = ws.addRow({
        dealer: '', category: '', component: 'TOTAL NET ADMIN:', net: adminNetTotal
      });
      adminTotalRow.font = { bold: true };
      ws.addRow({});
    }

    // Format currency
    ws.eachRow((row, rowNumber) => {
      if (rowNumber > 4) {
        row.getCell('written').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
        row.getCell('cancelled').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
        row.getCell('net').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
      }
    });
  }

`;

const newContent = content.substring(0, startIndex) + newMethod + content.substring(endIndex);
fs.writeFileSync(file, newContent, 'utf8');
console.log('Successfully updated excel.service.ts');
