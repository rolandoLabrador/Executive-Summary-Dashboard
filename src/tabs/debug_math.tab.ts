import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, configureWorksheet } from '../utils/excel.utils';
import { isGrossIncomeExcluded } from '../utils/report.transformer';

export class DebugMathTab implements IDashboardTab {
  readonly id = 'tab_debug_math';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, _tabConfig: unknown): void {

    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    if (debugDealers.length === 0) return;

    const asOf = config.asOfDate;

    function elapsedMonths(start: Date, end: Date): number {
      if (start > end) return 0;
      const years = end.getFullYear() - start.getFullYear();
      const months = end.getMonth() - start.getMonth();
      return years * 12 + months + 1;
    }

    const ws = workbook.addWorksheet('Debug Math Breakdown');
    configureWorksheet(ws);
    title(
      ws,
      'DEBUG MATH: COMPONENT & KPI BREAKDOWN',
      'Shows exact component sums used to calculate Admin, Reserve, and Earned Amounts for debugged dealers',
    );

    ws.columns = [
      { key: 'dealer',     width: 45 },
      { key: 'category',   width: 22 },
      { key: 'component',  width: 40 },
      { key: 'written',    width: 20 },
      { key: 'cancelled',  width: 20 },
      { key: 'net',        width: 20 },
      { key: 'earnedAmt',  width: 22 },
      { key: 'earnFactor', width: 18 },
      { key: 'elapsed',    width: 16 },
      { key: 'hasSchedule',width: 18 },
    ];

    ws.getRow(4).values = [
      'Dealer',
      'Category',
      'Component / Contract #',
      'Written Amount',
      'Cancelled Amount',
      'Net Amount',
      'Earned Amt / Contract',
      'Earn Factor (Ratio)',
      'Elapsed Months',
      'Has Schedule?',
    ];
    styleHeader(ws.getRow(4));

    const ORANGE_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
    const ORANGE_FONT: Partial<ExcelJS.Font> = { bold: true, italic: true, color: { argb: 'FF833C00' } };

    // ─────────────────────────────────────────────────────────────────────────
    // Per dealer: Earned Reserve + Component breakdown
    // ─────────────────────────────────────────────────────────────────────────
    for (const dealerName of debugDealers) {
      const dealerTx = model.contractTransactions.filter(
        (t) =>
          t.dealerName.toLowerCase() === dealerName ||
          t.dealerNumber.toLowerCase() === dealerName
      );

      if (dealerTx.length === 0) continue;

      const realDealerName = dealerTx[0]?.dealerName || dealerName;

      // ── SECTION 1: Earned Reserve Per-Contract Breakdown ─────────────────
      const earnedHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── EARNED RESERVE ──',
        component: 'Per-contract earning breakdown (effectiveDate → as-of date)',
      });
      earnedHeaderRow.font = ORANGE_FONT;
      earnedHeaderRow.fill = ORANGE_FILL;

      // Column sub-header for clarity
      const earnedSubHeader = ws.addRow({
        dealer: '',
        category: 'Tx Type',
        component: 'Contract #',
        written: 'Written Reserve',
        cancelled: 'Cancelled Reserve',
        net: 'Net Reserve',
        earnedAmt: 'Earned Reserve',
        earnFactor: 'Earn Factor',
        elapsed: 'Elapsed Mo.',
        hasSchedule: 'Has Schedule?',
      });
      earnedSubHeader.font = { bold: true, size: 9 };
      earnedSubHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DED8' } };

      let totalEarned = 0;
      let totalWrittenReserve = 0;
      let writtenCount = 0;

      dealerTx.forEach((t) => {
        const isCancel = t.transactionType === 'Cancellation';
        const elapsed = t.effectiveDate ? elapsedMonths(t.effectiveDate, asOf) : null;
        const earnFactor = t.reserveAmount !== 0
          ? t.earnedReserveAmount / Math.abs(t.reserveAmount)
          : null;

        if (!isCancel) {
          totalWrittenReserve += t.reserveAmount;
          writtenCount++;
        }
        totalEarned += t.earnedReserveAmount;

        const row = ws.addRow({
          dealer: '',
          category: t.transactionType,
          component: t.contractNumber,
          written: isCancel ? 0 : t.reserveAmount,
          cancelled: isCancel ? t.reserveAmount : 0,
          net: t.reserveAmount,
          earnedAmt: t.earnedReserveAmount,
          earnFactor: earnFactor,
          elapsed: elapsed,
          hasSchedule: t.earningSchedule ? 'YES' : 'NO (uses 1.0)',
        });

        // If there's a schedule, add a detail sub-row showing the curve at the current elapsed month
        if (t.earningSchedule && elapsed !== null) {
          const factor = t.earningSchedule[elapsed] ?? 1.0;
          const scheduleRow = ws.addRow({
            dealer: '',
            category: '  └─ Schedule',
            component: `Elapsed ${elapsed} mo → schedule factor = ${factor.toFixed(6)} → earned = ${t.reserveAmount.toFixed(2)} × ${factor.toFixed(6)} = ${(t.reserveAmount * factor).toFixed(2)}`,
          });
          scheduleRow.font = { italic: true, size: 9, color: { argb: 'FF595959' } };
        }

        // Style cancellations in red
        if (isCancel) {
          row.getCell('category').font = { color: { argb: 'FFCC0000' } };
        }
      });

      // Totals summary row
      const overallFactor = totalWrittenReserve !== 0 ? totalEarned / totalWrittenReserve : null;
      const summaryRow = ws.addRow({
        dealer: '',
        category: '',
        component: `TOTALS (${writtenCount} written contracts, ${dealerTx.length - writtenCount} cancellations):`,
        written: totalWrittenReserve,
        net: totalWrittenReserve,
        earnedAmt: totalEarned,
        earnFactor: overallFactor,
      });
      summaryRow.font = { bold: true };
      summaryRow.getCell('earnedAmt').note = 'Sum of earnedReserveAmount across all transactions for this dealer';
      summaryRow.getCell('earnFactor').note = 'totalEarned / totalWrittenReserve — overall blended earn factor';

      ws.addRow({});

      // Determine Active Contracts for this dealer
      const latestContractState = new Map<string, (typeof dealerTx)[0]>();
      dealerTx.forEach((t) => {
        const existing = latestContractState.get(t.contractNumber);
        if (!existing || t.snapshotDate >= existing.snapshotDate) {
          latestContractState.set(t.contractNumber, t);
        }
      });
      const activeContractNumbers = new Set(
        [...latestContractState.values()]
          .filter(t => t.contractStatus === 'A' && t.transactionType !== 'Cancellation')
          .map(t => t.contractNumber)
      );

      // ── SECTION 2: RESERVE Component Breakdown ────────────────────────────
      const sums: Record<string, { written: number; cancelled: number; net: number }> = {};

      dealerTx.forEach((t) => {
        if (!t.components) return;
        const isCancel = t.transactionType === 'Cancellation';
        const isActive = activeContractNumbers.has(t.contractNumber);

        ['RESERVE', 'ADMIN'].forEach(cat => {
          if (!t.components![cat]) return;
          
          // Only aggregate ADMIN components for ACTIVE contracts
          if (cat === 'ADMIN' && !isActive) return;

          Object.entries(t.components![cat]).forEach(([comp, amt]) => {
            const key = `${cat}.${comp}`;
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

      const reserveHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── RESERVE COMPONENTS ──',
        component: 'Bucket-level breakdown',
      });
      reserveHeaderRow.font = { bold: true, italic: true, color: { argb: 'FF1F4E79' } };
      reserveHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDAE3F3' } };

      let reserveNetTotal = 0;
      Object.keys(sums).filter(k => k.startsWith('RESERVE.')).sort().forEach(k => {
        const sum = sums[k];
        if (!sum || (sum.net === 0 && sum.written === 0)) return;
        reserveNetTotal += sum.net;
        ws.addRow({
          dealer: realDealerName,
          category: 'RESERVE',
          component: k.replace('RESERVE.', ''),
          written: sum.written,
          cancelled: sum.cancelled,
          net: sum.net,
        });
      });

      const reserveTotalRow = ws.addRow({
        dealer: '', category: '', component: 'TOTAL NET RESERVE:', net: reserveNetTotal,
      });
      reserveTotalRow.font = { bold: true };
      ws.addRow({});

      // ── SECTION 3: ADMIN Component Breakdown ─────────────────────────────
      const adminHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── ADMIN COMPONENTS ──',
        component: 'Bucket-level breakdown (ONLY FOR ACTIVE CONTRACTS)',
      });
      adminHeaderRow.font = { bold: true, italic: true, color: { argb: 'FF375623' } };
      adminHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };

      let adminNetTotal = 0;
      Object.keys(sums).filter(k => k.startsWith('ADMIN.')).sort().forEach(k => {
        const sum = sums[k];
        if (!sum || (sum.net === 0 && sum.written === 0)) return;
        adminNetTotal += sum.net;
        ws.addRow({
          dealer: realDealerName,
          category: 'ADMIN',
          component: k.replace('ADMIN.', ''),
          written: sum.written,
          cancelled: sum.cancelled,
          net: sum.net,
        });
      });

      const adminTotalRow = ws.addRow({
        dealer: '', category: '', component: 'TOTAL NET ADMIN:', net: adminNetTotal,
      });
      adminTotalRow.font = { bold: true };
      ws.addRow({});

      // ── SECTION 4: GROSS INCOME Component Breakdown ──────────────────────
      const grossHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── GROSS INCOME COMPONENTS ──',
        component: 'Bucket-level breakdown',
      });
      grossHeaderRow.font = { bold: true, italic: true, color: { argb: 'FF5E1E78' } };
      grossHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };

      let grossNetTotal = 0;
      
      const grossSums: Record<string, { written: number; cancelled: number; net: number }> = {};
      
      dealerTx.forEach((t) => {
        if (!t.components) return;
        const isCancel = t.transactionType === 'Cancellation';
        
        Object.keys(t.components).forEach(cat => {
          const catData = t.components![cat];
          if (!catData) return;
          Object.entries(catData).forEach(([comp, amt]) => {
            if (isGrossIncomeExcluded(cat, comp, config)) return;
            const key = `${cat}.${comp}`;
            if (!grossSums[key]) grossSums[key] = { written: 0, cancelled: 0, net: 0 };
            if (isCancel) {
              grossSums[key].cancelled -= Math.abs(amt);
              grossSums[key].net -= Math.abs(amt);
            } else {
              grossSums[key].written += amt;
              grossSums[key].net += amt;
            }
          });
        });
      });

      Object.keys(grossSums).sort().forEach(k => {
        const sum = grossSums[k];
        if (!sum || (sum.net === 0 && sum.written === 0)) return;
        grossNetTotal += sum.net;
        const [cat, ...compParts] = k.split('.');
        ws.addRow({
          dealer: realDealerName,
          category: cat,
          component: compParts.join('.'),
          written: sum.written,
          cancelled: sum.cancelled,
          net: sum.net,
        });
      });

      const grossTotalRow = ws.addRow({
        dealer: '', category: '', component: 'TOTAL NET GROSS INCOME:', net: grossNetTotal,
      });
      grossTotalRow.font = { bold: true };
      ws.addRow({});

      // ── SECTION 5: EXCLUDED COMPONENTS ───────────────────────────────────
      const excludedHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── EXCLUDED COMPONENTS ──',
        component: 'Buckets explicitly excluded from Gross Income',
      });
      excludedHeaderRow.font = { italic: true, color: { argb: 'FF808080' } };
      excludedHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

      let excludedTotal = 0;
      const excludedSums: Record<string, { written: number; cancelled: number; net: number }> = {};
      
      dealerTx.forEach((t) => {
        if (!t.components) return;
        const isCancel = t.transactionType === 'Cancellation';
        
        Object.keys(t.components).forEach(cat => {
          const catData = t.components![cat];
          if (!catData) return;
          Object.entries(catData).forEach(([comp, amt]) => {
            if (!isGrossIncomeExcluded(cat, comp, config)) return;
            const key = `${cat}.${comp}`;
            if (!excludedSums[key]) excludedSums[key] = { written: 0, cancelled: 0, net: 0 };
            if (isCancel) {
              excludedSums[key].cancelled -= Math.abs(amt);
              excludedSums[key].net -= Math.abs(amt);
            } else {
              excludedSums[key].written += amt;
              excludedSums[key].net += amt;
            }
          });
        });
      });

      Object.keys(excludedSums).sort().forEach(k => {
        const sum = excludedSums[k];
        if (!sum || (sum.net === 0 && sum.written === 0)) return;
        excludedTotal += sum.net;
        const [cat, ...compParts] = k.split('.');
        const r = ws.addRow({
          dealer: realDealerName,
          category: cat,
          component: compParts.join('.'),
          written: sum.written,
          cancelled: sum.cancelled,
          net: sum.net,
        });
        r.font = { color: { argb: 'FF808080' } }; // gray out excluded rows
      });

      const excludedTotalRow = ws.addRow({
        dealer: '', category: '', component: 'TOTAL EXCLUDED:', net: excludedTotal,
      });
      excludedTotalRow.font = { bold: true, color: { argb: 'FF808080' } };
      ws.addRow({});
      ws.addRow({});
    }

    // ── Format currency columns ───────────────────────────────────────────────
    const MONEY_FMT = '"$"#,##0.00;[Red]-"$"#,##0.00';
    const PCT_FMT = '0.00%';
    ws.eachRow((row, rowNumber) => {
      if (rowNumber > 4) {
        row.getCell('written').numFmt = MONEY_FMT;
        row.getCell('cancelled').numFmt = MONEY_FMT;
        row.getCell('net').numFmt = MONEY_FMT;
        row.getCell('earnedAmt').numFmt = MONEY_FMT;
        row.getCell('earnFactor').numFmt = PCT_FMT;
      }
    });

  }
}
