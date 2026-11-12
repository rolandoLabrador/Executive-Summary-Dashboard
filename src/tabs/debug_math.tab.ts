import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, configureWorksheet } from '../utils/excel.utils';
import { isGrossIncomeExcluded, excludedComponent } from '../utils/report.transformer';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

export class DebugMathTab implements IDashboardTab {
  readonly id = 'tab_debug_math';

  constructor() {}

  render(
    workbook: ExcelJS.Workbook,
    model: ReportModel,
    config: ReportConfig,
    _tabConfig?: unknown,
  ): void {
    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const debugAgents = process.env.DEBUG_AGENT_NAME
      ? process.env.DEBUG_AGENT_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    if (debugDealers.length === 0 && debugAgents.length === 0) return;

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
      'Shows exact component sums used to calculate Admin, Reserve, and Earned Amounts for debugged dealers and agents',
    );

    ws.columns = [
      { key: 'dealer', width: 34 },
      { key: 'product', width: 22 },
      { key: 'category', width: 20 },
      { key: 'component', width: 32 },
      { key: 'written', width: 18 },
      { key: 'cancelled', width: 18 },
      { key: 'net', width: 18 },
      { key: 'earnedAmt', width: 20 },
      { key: 'earnFactor', width: 16 },
      { key: 'elapsed', width: 14 },
      { key: 'hasSchedule', width: 16 },
    ];

    ws.getRow(4).values = [
      'Dealer / Agent',
      'Contract / Product Type',
      'Category / Tx Type',
      'Component / Contract #',
      'Written Amount',
      'Cancelled Amount',
      'Net Amount',
      'Earned Amount',
      'Earn Factor / % Earned',
      'Elapsed Months',
      'Has Schedule?',
    ];
    styleHeader(ws.getRow(4));

    const MONEY_FMT = '"$"#,##0.00;[Red]-"$"#,##0.00';
    const PCT_FMT = '0.00%';

    const ORANGE_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFCE4D6' },
    };
    const ORANGE_FONT: Partial<ExcelJS.Font> = {
      bold: true,
      italic: true,
      color: { argb: 'FF833C00' },
    };

    const NAVY_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F4E79' },
    };
    const NAVY_FONT: Partial<ExcelJS.Font> = {
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };

    const TEAL_FILL: ExcelJS.Fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD9E1F2' },
    };

    // ─────────────────────────────────────────────────────────────────────────
    // Build contract lookup for claim enrichment
    // ─────────────────────────────────────────────────────────────────────────
    const contractLookup = new Map<string, (typeof model.contractTransactions)[0]>();
    model.contractTransactions.forEach((t) => {
      if (t.contractNumber && t.transactionType !== 'Cancellation') {
        contractLookup.set(t.contractNumber, t);
      }
    });

    const claimProduct = (c: (typeof model.claims)[0]): string => {
      const orig = contractLookup.get(c.contractNumber);
      return orig?.product || c.product || 'Standard';
    };

    function matchesAgent(
      agent?: string,
      agentName?: string,
      debugList: string[] = debugAgents,
    ): boolean {
      if (debugList.length === 0) return false;
      const code = agent?.trim().toLowerCase() || '';
      const name = agentName?.trim().toLowerCase() || '';
      return debugList.some((d) => {
        if (d === code || d === name) return true;
        if (
          (d === 'exw' || d === 'executive warranties') &&
          (code === 'exw' || name === 'executive warranties')
        ) {
          return true;
        }
        return false;
      });
    }

    interface TargetInfo {
      type: 'agent' | 'dealer';
      rawKey: string;
      label: string;
      tx: typeof model.contractTransactions;
      claims: typeof model.claims;
    }

    const targets: TargetInfo[] = [];

    for (const dealerName of debugDealers) {
      const dealerTx = model.contractTransactions.filter(
        (t) =>
          t.dealerName.toLowerCase() === dealerName || t.dealerNumber.toLowerCase() === dealerName,
      );
      if (dealerTx.length === 0) continue;
      const realName = dealerTx[0]?.dealerName || dealerName;
      const dealerClaims = model.claims.filter((c) => {
        const match =
          c.dealer.toLowerCase() === dealerName || c.dealerName.toLowerCase() === dealerName;
        if (match) return true;
        const orig = contractLookup.get(c.contractNumber);
        return (
          orig &&
          (orig.dealer.toLowerCase() === dealerName || orig.dealerName.toLowerCase() === dealerName)
        );
      });
      targets.push({
        type: 'dealer',
        rawKey: dealerName,
        label: `DEALER: ${realName}`,
        tx: dealerTx,
        claims: dealerClaims,
      });
    }

    for (const agentInput of debugAgents) {
      const agentTx = model.contractTransactions.filter((t) =>
        matchesAgent(t.agent, t.agentName, [agentInput]),
      );
      if (agentTx.length === 0) continue;
      const realName = agentTx[0]?.agentName || agentTx[0]?.agent || agentInput.toUpperCase();
      const agentClaims = model.claims.filter((c) => {
        if (matchesAgent(c.agent, c.agentName, [agentInput])) return true;
        const orig = contractLookup.get(c.contractNumber);
        return orig && matchesAgent(orig.agent, orig.agentName, [agentInput]);
      });
      targets.push({
        type: 'agent',
        rawKey: agentInput,
        label: `AGENT: ${realName}`,
        tx: agentTx,
        claims: agentClaims,
      });
    }

    for (const target of targets) {
      const dealerTx = target.tx;
      const targetClaims = target.claims;
      const realDealerName = target.label;

      // ── TOTALS CALCULATION ACROSS ALL TRANSACTIONS & CLAIMS ────────────────
      const writtenContractsSet = new Set<string>();
      const cancelledContractsSet = new Set<string>();
      let totalWrittenReserve = 0;
      let totalCancelledReserve = 0;
      let totalEarnedReserve = 0;

      dealerTx.forEach((t) => {
        const isCancel = t.transactionType === 'Cancellation';
        if (!isCancel) {
          writtenContractsSet.add(t.contractNumber);
          totalWrittenReserve += t.reserveAmount;
          totalEarnedReserve += t.earnedReserveAmount;
        } else {
          cancelledContractsSet.add(t.contractNumber);
          totalCancelledReserve += Math.abs(t.reserveAmount);
        }
      });

      const totalNetReserve = totalWrittenReserve - totalCancelledReserve;
      const overallPercentEarned =
        totalNetReserve > 0 ? totalEarnedReserve / totalNetReserve : null;

      const claimNumbersSet = new Set<string>();
      let totalClaimsPaid = 0;
      targetClaims.forEach((c) => {
        totalClaimsPaid += c.paidAmount;
        claimNumbersSet.add(c.claimNumber || c.sourceId);
      });

      const underwritingProfit = totalEarnedReserve - totalClaimsPaid;
      const earnedLossRatio = totalEarnedReserve > 0 ? totalClaimsPaid / totalEarnedReserve : null;

      // ── SECTION 0: ACTUARIAL RECONCILIATION & LOSS RATIO MECHANICS ─────────
      const reconHeader = ws.addRow({
        dealer: realDealerName,
        product: '── ACTUARIAL FORMULA & LOSS RATIO RECONCILIATION ──',
        category: 'Exact Calculation Mechanics',
      });
      reconHeader.font = NAVY_FONT;
      reconHeader.fill = NAVY_FILL;

      const r1 = ws.addRow({
        dealer: 'STEP 1: NET RESERVE',
        product: 'Written Reserve - Cancelled Reserve',
        category: 'Equation:',
        component: `$${totalWrittenReserve.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - $${totalCancelledReserve.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
        written: totalWrittenReserve,
        cancelled: -totalCancelledReserve,
        net: totalNetReserve,
        earnedAmt: null,
        earnFactor: null,
      });
      r1.font = { bold: true };
      r1.fill = TEAL_FILL;

      const r2 = ws.addRow({
        dealer: 'STEP 2: EARNED RESERVE',
        product: 'Sum of (Contract Net Reserve × Elapsed Months Earning Factor)',
        category: 'Earned Amount & % Earned:',
        component: `Unearned Reserve = $${Math.max(0, totalNetReserve - totalEarnedReserve).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
        written: null,
        cancelled: null,
        net: totalNetReserve,
        earnedAmt: totalEarnedReserve,
        earnFactor: overallPercentEarned,
      });
      r2.font = { bold: true };
      r2.fill = TEAL_FILL;

      const r3 = ws.addRow({
        dealer: 'STEP 3: CLAIMS PAID',
        product: `Incurred & Paid Claims to Date (${claimNumbersSet.size} distinct claims)`,
        category: 'Total Claims Paid:',
        component: `${claimNumbersSet.size} claims paid across all products`,
        written: null,
        cancelled: null,
        net: -totalClaimsPaid,
        earnedAmt: totalClaimsPaid,
        earnFactor: null,
      });
      r3.font = { bold: true };
      r3.fill = TEAL_FILL;

      const r4 = ws.addRow({
        dealer: 'STEP 4: UNDERWRITING PROFIT',
        product: 'Earned Reserve - Claims Paid',
        category: 'Equation:',
        component: `$${totalEarnedReserve.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - $${totalClaimsPaid.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
        written: null,
        cancelled: null,
        net: underwritingProfit,
        earnedAmt: underwritingProfit,
        earnFactor: null,
      });
      r4.font = { bold: true, color: { argb: underwritingProfit >= 0 ? 'FF006100' : 'FF9C0006' } };
      r4.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: underwritingProfit >= 0 ? 'FFC6EFCE' : 'FFFFC7CE' },
      };

      const r5 = ws.addRow({
        dealer: 'STEP 5: EARNED LOSS RATIO',
        product: 'Claims Paid / Earned Reserve',
        category: 'Equation:',
        component: `$${totalClaimsPaid.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} / $${totalEarnedReserve.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
        written: null,
        cancelled: null,
        net: null,
        earnedAmt: totalClaimsPaid,
        earnFactor: earnedLossRatio,
      });
      r5.font = {
        bold: true,
        color: { argb: (earnedLossRatio || 0) > 0.8 ? 'FF9C0006' : 'FF006100' },
      };
      r5.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: (earnedLossRatio || 0) > 0.8 ? 'FFFFC7CE' : 'FFC6EFCE' },
      };

      ws.addRow({});

      // ── SECTION 0.5: CONTRACT / PRODUCT TYPE SUMMARY BREAKDOWN ─────────────
      const productSummaryHeader = ws.addRow({
        dealer: realDealerName,
        product: '── CONTRACT / PRODUCT TYPE BREAKDOWN ──',
        category: 'Summary by Product Type',
        component: 'Contracts Written, Cancelled, Net Reserve, Earned Reserve, Claims & Loss Ratio',
      });
      productSummaryHeader.font = { bold: true, italic: true, color: { argb: 'FF203764' } };
      productSummaryHeader.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' },
      };

      const productSubHeader = ws.addRow({
        dealer: 'Contract / Product Type',
        product: 'Written Contracts',
        category: 'Cancelled Contracts',
        component: 'Net Contracts',
        written: 'Written Reserve',
        cancelled: 'Cancelled Reserve',
        net: 'Net Reserve',
        earnedAmt: 'Earned Reserve',
        earnFactor: '% Earned',
        elapsed: 'Claims Paid',
        hasSchedule: 'Loss Ratio',
      });
      productSubHeader.font = { bold: true, size: 9 };
      productSubHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };

      const productTypes = new Set<string>();
      dealerTx.forEach((t) => productTypes.add(t.product || 'Standard'));
      targetClaims.forEach((c) => productTypes.add(claimProduct(c)));

      const sortedProducts = Array.from(productTypes).sort((a, b) => a.localeCompare(b));

      sortedProducts.forEach((prod) => {
        const pTx = dealerTx.filter((t) => (t.product || 'Standard') === prod);
        const pClaims = targetClaims.filter((c) => claimProduct(c) === prod);

        const pWrittenSet = new Set<string>();
        const pCancelledSet = new Set<string>();
        let pWrittenRes = 0;
        let pCancelledRes = 0;
        let pEarnedRes = 0;

        pTx.forEach((t) => {
          const isCancel = t.transactionType === 'Cancellation';
          if (!isCancel) {
            pWrittenSet.add(t.contractNumber);
            pWrittenRes += t.reserveAmount;
            pEarnedRes += t.earnedReserveAmount;
          } else {
            pCancelledSet.add(t.contractNumber);
            pCancelledRes += Math.abs(t.reserveAmount);
          }
        });

        const pNetRes = pWrittenRes - pCancelledRes;
        const pPctEarned = pNetRes > 0 ? pEarnedRes / pNetRes : null;
        const pClaimsPaid = pClaims.reduce((sum, c) => sum + c.paidAmount, 0);
        const pLossRatio = pEarnedRes > 0 ? pClaimsPaid / pEarnedRes : null;

        const pRow = ws.addRow({
          dealer: prod,
          product: pWrittenSet.size,
          category: pCancelledSet.size,
          component: `${pWrittenSet.size - pCancelledSet.size} net`,
          written: pWrittenRes,
          cancelled: -pCancelledRes,
          net: pNetRes,
          earnedAmt: pEarnedRes,
          earnFactor: pPctEarned,
          elapsed: pClaimsPaid,
          hasSchedule: pLossRatio !== null ? `${(pLossRatio * 100).toFixed(1)}%` : '0.0%',
        });
        pRow.getCell('product').numFmt = '#,##0';
        pRow.getCell('category').numFmt = '#,##0';
        pRow.getCell('elapsed').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';
      });

      const pTotalRow = ws.addRow({
        dealer: 'TOTAL (ALL PRODUCTS)',
        product: writtenContractsSet.size,
        category: cancelledContractsSet.size,
        component: `${writtenContractsSet.size - cancelledContractsSet.size} net`,
        written: totalWrittenReserve,
        cancelled: -totalCancelledReserve,
        net: totalNetReserve,
        earnedAmt: totalEarnedReserve,
        earnFactor: overallPercentEarned,
        elapsed: totalClaimsPaid,
        hasSchedule: earnedLossRatio !== null ? `${(earnedLossRatio * 100).toFixed(1)}%` : '0.0%',
      });
      pTotalRow.font = { bold: true };
      pTotalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      pTotalRow.getCell('product').numFmt = '#,##0';
      pTotalRow.getCell('category').numFmt = '#,##0';
      pTotalRow.getCell('elapsed').numFmt = '"$"#,##0.00;[Red]-"$"#,##0.00';

      ws.addRow({});

      // ── SECTION 0.6: PRODUCT LOSS ANALYSIS & CLAIMS DISTRIBUTION CHART ────
      const productLossRows: Array<{
        name: string;
        writtenRes: number;
        cancelledRes: number;
        netRes: number;
        earnedRes: number;
        claimsPaid: number;
        profit: number;
        lossRatio: number | null;
        claimShare: number;
        isLossDriver: boolean;
        color: string;
      }> = [];

      // Color palette for loss drivers vs profitable products
      // Severe losses: Crimson (#C00000), Coral Red (#ED7D31), Dark Orange (#C55A11)
      // Moderate/Unmapped: Purple (#7030A0)
      // Profitable products: Steel Blue (#41719C), Slate Blue (#2F5597), Navy (#1F4E79), Teal (#2E75B6)
      const lossColors = ['#C00000', '#ED7D31', '#C55A11', '#7030A0'];
      const profitColors = ['#41719C', '#2F5597', '#1F4E79', '#2E75B6'];
      let lossColorIdx = 0;
      let profitColorIdx = 0;

      sortedProducts.forEach((prod) => {
        const pTx = dealerTx.filter((t) => (t.product || 'Standard') === prod);
        const pClaims = targetClaims.filter((c) => claimProduct(c) === prod);

        let pWrittenRes = 0;
        let pCancelledRes = 0;
        let pEarnedRes = 0;

        pTx.forEach((t) => {
          if (t.transactionType !== 'Cancellation') {
            pWrittenRes += t.reserveAmount;
            pEarnedRes += t.earnedReserveAmount;
          } else {
            pCancelledRes += Math.abs(t.reserveAmount);
          }
        });

        const pNetRes = pWrittenRes - pCancelledRes;
        const pClaimsPaid = pClaims.reduce((sum, c) => sum + c.paidAmount, 0);
        const pProfit = pEarnedRes - pClaimsPaid;
        const pLossRatio = pEarnedRes > 0 ? pClaimsPaid / pEarnedRes : null;
        const claimShare = totalClaimsPaid > 0 ? pClaimsPaid / totalClaimsPaid : 0;
        const isLossDriver = pProfit < 0 || (pLossRatio !== null && pLossRatio > 1.0);

        if (pClaimsPaid > 0 || pWrittenRes > 0) {
          const color = isLossDriver
            ? lossColors[lossColorIdx++ % lossColors.length]!
            : profitColors[profitColorIdx++ % profitColors.length]!;

          productLossRows.push({
            name: prod,
            writtenRes: pWrittenRes,
            cancelledRes: pCancelledRes,
            netRes: pNetRes,
            earnedRes: pEarnedRes,
            claimsPaid: pClaimsPaid,
            profit: pProfit,
            lossRatio: pLossRatio,
            claimShare,
            isLossDriver,
            color,
          });
        }
      });

      // Sort by largest loss (most negative profit first)
      productLossRows.sort((a, b) => a.profit - b.profit);

      const chartSectionHeader = ws.addRow({
        dealer: realDealerName,
        product: '── SECTION 0.6: PRODUCT LOSS DRIVERS & UNDERWRITING SCORECARD ──',
        category: 'Underwriting Deficit & Claims Concentration',
        component:
          'Ranks products by underwriting deficit (Earned Reserve - Claims Paid) to pinpoint loss culprits',
      });
      chartSectionHeader.font = { bold: true, size: 11, color: { argb: 'FF9C0006' } };
      chartSectionHeader.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC7CE' },
      };
      chartSectionHeader.height = 24;

      // 1. Loss Driver Scorecard Table (Columns A to K)
      const scoreSubHeader = ws.addRow({
        dealer: 'Underwriting Status',
        product: 'Product / Coverage Line',
        category: 'Chart Key',
        component: 'Claims Paid',
        written: 'Earned Reserve',
        cancelled: 'UW Profit / (Deficit)',
        net: 'Earned Loss Ratio',
        earnedAmt: '% of Total Claims',
        earnFactor: 'Written Reserve',
        elapsed: 'Net Reserve',
        hasSchedule: 'Underwriting Finding',
      });
      scoreSubHeader.font = { bold: true, size: 9 };
      scoreSubHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      scoreSubHeader.height = 22;

      productLossRows.forEach((r) => {
        const isLoss = r.profit < 0;
        const statusLabel =
          r.profit < -10000
            ? '🔴 CRITICAL LOSS'
            : r.profit < 0
              ? '🔴 LOSS DEFICIT'
              : r.profit === 0
                ? '⚪ BREAK-EVEN'
                : '🟢 PROFIT SURPLUS';

        const lossRow = ws.addRow({
          dealer: statusLabel,
          product: r.name,
          category: `● ${r.name}`,
          component: r.claimsPaid,
          written: r.earnedRes,
          cancelled: r.profit,
          net: r.lossRatio,
          earnedAmt: r.claimShare,
          earnFactor: r.writtenRes,
          elapsed: r.netRes,
          hasSchedule: isLoss
            ? `Deficit -$${Math.abs(Math.round(r.profit)).toLocaleString()} (${((r.lossRatio || 0) * 100).toFixed(1)}% LR)`
            : `Surplus +$${Math.round(r.profit).toLocaleString()} (${((r.lossRatio || 0) * 100).toFixed(1)}% LR)`,
        });

        lossRow.height = 22;

        lossRow.getCell('component').numFmt = MONEY_FMT;
        lossRow.getCell('written').numFmt = MONEY_FMT;
        lossRow.getCell('cancelled').numFmt = MONEY_FMT;
        lossRow.getCell('net').numFmt = PCT_FMT;
        lossRow.getCell('earnedAmt').numFmt = PCT_FMT;
        lossRow.getCell('earnFactor').numFmt = MONEY_FMT;
        lossRow.getCell('elapsed').numFmt = MONEY_FMT;

        // Color dot in category column
        const colorHex = r.color.replace('#', '');
        lossRow.getCell('category').font = { bold: true, color: { argb: `FF${colorHex}` } };

        if (isLoss) {
          lossRow.getCell('dealer').font = { bold: true, color: { argb: 'FF9C0006' } };
          lossRow.getCell('cancelled').font = { bold: true, color: { argb: 'FF9C0006' } };
          lossRow.getCell('net').font = { bold: true, color: { argb: 'FF9C0006' } };
          lossRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } };
        } else {
          lossRow.getCell('dealer').font = { bold: true, color: { argb: 'FF006100' } };
          lossRow.getCell('cancelled').font = { bold: true, color: { argb: 'FF006100' } };
          lossRow.getCell('net').font = { bold: true, color: { argb: 'FF006100' } };
          lossRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
        }
      });

      // 2. Dedicated Visual Chart & Actuarial Attribution Section
      const spacer1 = ws.addRow({});
      spacer1.height = 14;

      const chartSectionBanner = ws.addRow({
        dealer: '── CLAIMS CONCENTRATION PIE CHART ──',
        product: '(Slices proportional to Claims Paid)',
        category: '',
        component: '',
        written: '── ACTUARIAL SUMMARY & LOSS DRIVER DIAGNOSIS ──',
      });
      chartSectionBanner.font = { bold: true, size: 10, color: { argb: 'FF203764' } };
      chartSectionBanner.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' },
      };
      chartSectionBanner.height = 24;

      const topLossProduct = productLossRows.find((r) => r.profit < 0);
      const totalDeficitFromLossDrivers = productLossRows
        .filter((r) => r.profit < 0)
        .reduce((sum, r) => sum + r.profit, 0);

      const chartStartRowNumber = ws.lastRow!.number + 1;
      const CHART_ROW_COUNT = 14;

      // Allocate 14 dedicated rows with locked height (total ~280px)
      // Columns A-C are left empty for the pie chart image
      // Columns E-K contain the actuarial summary and legend
      for (let i = 0; i < CHART_ROW_COUNT; i++) {
        const row = ws.addRow({});
        row.height = 20;

        if (i === 0) {
          row.getCell('written').value = 'Top Loss Driver:';
          row.getCell('written').font = { bold: true, size: 10 };
          row.getCell('cancelled').value = topLossProduct
            ? `${topLossProduct.name}  |  Underwriting Deficit: -$${Math.abs(Math.round(topLossProduct.profit)).toLocaleString()}  |  Loss Ratio: ${((topLossProduct.lossRatio || 0) * 100).toFixed(1)}%`
            : 'None (All products profitable)';
          row.getCell('cancelled').font = { bold: true, color: { argb: 'FF9C0006' } };
        } else if (i === 1) {
          row.getCell('written').value = 'Total Loss Deficit:';
          row.getCell('written').font = { bold: true, size: 10 };
          row.getCell('cancelled').value =
            `-$${Math.abs(Math.round(totalDeficitFromLossDrivers)).toLocaleString()} combined across all losing coverage lines`;
          row.getCell('cancelled').font = { bold: true, color: { argb: 'FF9C0006' } };
        } else if (i === 2) {
          row.getCell('written').value = 'Profitable Offsets:';
          row.getCell('written').font = { bold: true, size: 10 };
          const profitableLines = productLossRows.filter((r) => r.profit > 0);
          row.getCell('cancelled').value =
            profitableLines.length > 0
              ? profitableLines
                  .map((p) => `${p.name} (+$${Math.round(p.profit).toLocaleString()})`)
                  .join(', ')
              : 'None';
          row.getCell('cancelled').font = { bold: true, color: { argb: 'FF006100' } };
        } else if (i === 4) {
          row.getCell('written').value = '── CLAIMS DISTRIBUTION LEGEND ──';
          row.getCell('written').font = { bold: true, size: 9, color: { argb: 'FF203764' } };
        } else if (i >= 5 && i - 5 < productLossRows.length) {
          const sliceItem = productLossRows[i - 5]!;
          const colorHex = sliceItem.color.replace('#', '');
          row.getCell('written').value = `● ${sliceItem.name}`;
          row.getCell('written').font = { bold: true, color: { argb: `FF${colorHex}` } };
          row.getCell('cancelled').value = sliceItem.claimsPaid;
          row.getCell('cancelled').numFmt = MONEY_FMT;
          row.getCell('net').value = sliceItem.claimShare;
          row.getCell('net').numFmt = PCT_FMT;
          row.getCell('earnedAmt').value = sliceItem.profit < 0 ? '🔴 Deficit' : '🟢 Surplus';
          row.getCell('earnedAmt').font = {
            bold: true,
            color: { argb: sliceItem.profit < 0 ? 'FF9C0006' : 'FF006100' },
          };
        }
      }

      // Embed the pie chart cleanly within the dedicated rows (col 0.3, width 240, height 240)
      const pieSlices = productLossRows
        .filter((r) => r.claimsPaid > 0)
        .map((r) => ({
          value: r.claimsPaid,
          color: r.color,
        }));

      if (pieSlices.length > 0) {
        try {
          const chartBuffer = renderPieChartPng(pieSlices, 480, 480);
          const imageId = workbook.addImage({
            base64: `data:image/png;base64,${chartBuffer.toString('base64')}`,
            extension: 'png',
          });
          ws.addImage(imageId, {
            tl: { col: 0.3, row: chartStartRowNumber - 0.8 },
            ext: { width: 240, height: 240 },
          });
        } catch (err) {
          console.warn('Could not render debug pie chart:', err);
        }
      }

      const spacer2 = ws.addRow({});
      spacer2.height = 14;

      // ── SECTION 1: Earned Reserve Per-Contract Breakdown ─────────────────
      const earnedHeaderRow = ws.addRow({
        dealer: realDealerName,
        product: '── EARNED RESERVE PER CONTRACT ──',
        category: 'Per-contract earning breakdown (effectiveDate → as-of date)',
        component: 'Filter by Product Type, Tx Type, or Contract # below',
      });
      earnedHeaderRow.font = ORANGE_FONT;
      earnedHeaderRow.fill = ORANGE_FILL;

      // Column sub-header for clarity
      const earnedSubHeader = ws.addRow({
        dealer: 'Dealer / Agent',
        product: 'Contract / Product Type',
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

      let section1Earned = 0;
      let section1WrittenReserve = 0;
      let writtenCount = 0;

      dealerTx.forEach((t) => {
        const isCancel = t.transactionType === 'Cancellation';
        const elapsed = t.effectiveDate ? elapsedMonths(t.effectiveDate, asOf) : null;
        const earnFactor =
          t.reserveAmount !== 0 ? t.earnedReserveAmount / Math.abs(t.reserveAmount) : null;

        if (!isCancel) {
          section1WrittenReserve += t.reserveAmount;
          writtenCount++;
        }
        section1Earned += t.earnedReserveAmount;

        const row = ws.addRow({
          dealer: '',
          product: t.product || 'Standard',
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

        // If expired in PCMI, show status E note, else if there's a schedule, show the elapsed month schedule calculation
        if (t.contractStatus === 'E') {
          const expiredRow = ws.addRow({
            dealer: '',
            product: '',
            category: '  └─ Status E',
            component: `Contract marked EXPIRED (Status: E) in PCMI → 100% Earned (${t.reserveAmount.toFixed(2)})`,
          });
          expiredRow.font = { italic: true, size: 9, color: { argb: 'FF006100' } };
        } else if (t.earningSchedule && elapsed !== null) {
          const factor = t.earningSchedule[elapsed] ?? 1.0;
          const scheduleRow = ws.addRow({
            dealer: '',
            product: '',
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
      const overallFactor =
        section1WrittenReserve !== 0 ? section1Earned / section1WrittenReserve : null;
      const summaryRow = ws.addRow({
        dealer: '',
        product: '',
        category: '',
        component: `TOTALS (${writtenCount} written contracts, ${dealerTx.length - writtenCount} cancellations):`,
        written: section1WrittenReserve,
        net: section1WrittenReserve,
        earnedAmt: section1Earned,
        earnFactor: overallFactor,
      });
      summaryRow.font = { bold: true };
      summaryRow.getCell('earnedAmt').note =
        'Sum of earnedReserveAmount across all transactions for this dealer';
      summaryRow.getCell('earnFactor').note =
        'totalEarned / totalWrittenReserve — overall blended earn factor';

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
          .filter((t) => t.contractStatus === 'A' && t.transactionType !== 'Cancellation')
          .map((t) => t.contractNumber),
      );

      // ── SECTION 2: RESERVE Component Breakdown ────────────────────────────
      const sums: Record<string, { written: number; cancelled: number; net: number }> = {};

      dealerTx.forEach((t) => {
        if (!t.components) return;
        const isCancel = t.transactionType === 'Cancellation';
        const isActive = activeContractNumbers.has(t.contractNumber);

        ['RESERVE', 'ADMIN'].forEach((cat) => {
          if (!t.components![cat]) return;

          // Only aggregate ADMIN components for ACTIVE contracts
          if (cat === 'ADMIN' && !isActive) return;

          Object.entries(t.components![cat]).forEach(([comp, amt]) => {
            if (excludedComponent(comp, cat as 'ADMIN' | 'RESERVE', config)) return;

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
      Object.keys(sums)
        .filter((k) => k.startsWith('RESERVE.'))
        .sort()
        .forEach((k) => {
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
        dealer: '',
        category: '',
        component: 'TOTAL NET RESERVE:',
        net: reserveNetTotal,
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
      Object.keys(sums)
        .filter((k) => k.startsWith('ADMIN.'))
        .sort()
        .forEach((k) => {
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
        dealer: '',
        category: '',
        component: 'TOTAL NET ADMIN:',
        net: adminNetTotal,
      });
      adminTotalRow.font = { bold: true };
      ws.addRow({});

      // ── SECTION 4: NET INCOME Component Breakdown ──────────────────────
      const grossHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── NET INCOME COMPONENTS ──',
        component: 'Bucket-level breakdown',
      });
      grossHeaderRow.font = { bold: true, italic: true, color: { argb: 'FF5E1E78' } };
      grossHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } };

      let grossNetTotal = 0;

      const grossSums: Record<string, { written: number; cancelled: number; net: number }> = {};

      dealerTx.forEach((t) => {
        if (!t.components) return;
        const isCancel = t.transactionType === 'Cancellation';

        Object.keys(t.components).forEach((cat) => {
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

      Object.keys(grossSums)
        .sort()
        .forEach((k) => {
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
        dealer: '',
        category: '',
        component: 'TOTAL NET INCOME:',
        net: grossNetTotal,
      });
      grossTotalRow.font = { bold: true };
      ws.addRow({});

      // ── SECTION 5: EXCLUDED COMPONENTS ───────────────────────────────────
      const excludedHeaderRow = ws.addRow({
        dealer: realDealerName,
        category: '── EXCLUDED COMPONENTS ──',
        component: 'Buckets explicitly excluded from Net Income',
      });
      excludedHeaderRow.font = { italic: true, color: { argb: 'FF808080' } };
      excludedHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

      let excludedTotal = 0;
      const excludedSums: Record<string, { written: number; cancelled: number; net: number }> = {};

      dealerTx.forEach((t) => {
        if (!t.components) return;
        const isCancel = t.transactionType === 'Cancellation';

        Object.keys(t.components).forEach((cat) => {
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

      Object.keys(excludedSums)
        .sort()
        .forEach((k) => {
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
        dealer: '',
        category: '',
        component: 'TOTAL EXCLUDED:',
        net: excludedTotal,
      });
      excludedTotalRow.font = { bold: true, color: { argb: 'FF808080' } };
      ws.addRow({});
      ws.addRow({});
    }

    // ── Format currency columns ───────────────────────────────────────────────
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
