import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import {
  COLORS,
  MONEY,
  INTEGER,
  PERCENT,
  configureWorksheet,
  title,
  formatDate,
} from '../utils/excel.utils';
import {
  AiAnalystService,
  type TargetSummaryData,
  type ProductLineSummary,
} from '../services/ai_analyst.service';

export class AiAgentDebugTab implements IDashboardTab {
  readonly id = 'tab_ai_agent_debug';
  private aiService = new AiAnalystService();

  constructor() {}

  async render(
    workbook: ExcelJS.Workbook,
    model: ReportModel,
    config: ReportConfig,
    _tabConfig?: unknown,
  ): Promise<void> {
    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const debugAgents = process.env.DEBUG_AGENT_NAME
      ? process.env.DEBUG_AGENT_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];

    if (debugDealers.length === 0 && debugAgents.length === 0) {
      return;
    }

    const ws = workbook.addWorksheet('AI Analyst Lab');
    configureWorksheet(ws);

    ws.columns = [
      { width: 26 }, // A: Target / Code
      { width: 16 }, // B: Dimension
      { width: 22 }, // C: Period / Tenure
      { width: 16, numFmt: INTEGER }, // D: Contracts
      { width: 18, numFmt: MONEY }, // E: Net Reserve
      { width: 18, numFmt: MONEY }, // F: Earned Reserve
      { width: 18, numFmt: MONEY }, // G: Claims Paid
      { width: 16, numFmt: PERCENT }, // H: Loss Ratio
      { width: 16, numFmt: PERCENT }, // I: Frequency of Claims
      { width: 14, numFmt: PERCENT }, // J: % Earned
      { width: 28 }, // K: Underwriting Credibility & Status
    ];

    // 1. Header Title
    title(
      ws,
      '🤖 AI EXECUTIVE ANALYST — AGENT & DEALER TRAINING LAB',
      `Targeted Underwriting & Actuarial Intelligence | As of ${formatDate(model.asOfDate)}`,
      9,
    );

    // 2. WORK IN PROGRESS NOTICE (Prominent at top as requested)
    ws.addRow([]);
    const wipRowStart = 3;
    ws.mergeCells(`A${wipRowStart}:I${wipRowStart}`);
    const wipTitle = ws.getCell(`A${wipRowStart}`);
    wipTitle.value = '⚠️ WORK IN PROGRESS — UNDERWRITING AGENT TRAINING';
    wipTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: 'FF833C00' } };
    wipTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
    wipTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

    const wipDescRow = wipRowStart + 1;
    ws.mergeCells(`A${wipDescRow}:I${wipDescRow + 1}`);
    const wipDesc = ws.getCell(`A${wipDescRow}`);
    wipDesc.value =
      'This is a work in progress. We are actively training this AI Agent to automatically inspect and point out targeted agents and dealers (detecting loss ratios, tenure/maturity, claim spikes, and reserve health). In subsequent iterations, the agent will drill down directly to component bucket levels.';
    wipDesc.font = { name: 'Aptos', size: 10, italic: true, color: { argb: 'FF595959' } };
    wipDesc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDF3EB' } };
    wipDesc.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true, indent: 1 };

    // Border around WIP box
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach((col) => {
      ws.getCell(`${col}${wipRowStart}`).border = {
        top: { style: 'medium', color: { argb: 'FFE07030' } },
        left: col === 'A' ? { style: 'medium', color: { argb: 'FFE07030' } } : undefined,
        right: col === 'I' ? { style: 'medium', color: { argb: 'FFE07030' } } : undefined,
      };
      ws.getCell(`${col}${wipDescRow + 1}`).border = {
        bottom: { style: 'medium', color: { argb: 'FFE07030' } },
        left: col === 'A' ? { style: 'medium', color: { argb: 'FFE07030' } } : undefined,
        right: col === 'I' ? { style: 'medium', color: { argb: 'FFE07030' } } : undefined,
      };
    });

    // 3. Extract Summary Data for all Targets
    const targets: TargetSummaryData[] = [];

    function matchesAgent(
      code?: string,
      name?: string,
      debugList: string[] = debugAgents,
    ): boolean {
      const c = code?.trim().toLowerCase() || '';
      const n = name?.trim().toLowerCase() || '';
      return debugList.some((d) => {
        if (d === c || d === n) return true;
        if (
          (d === 'exw' || d === 'executive warranties') &&
          (c === 'exw' || n === 'executive warranties')
        ) {
          return true;
        }
        return false;
      });
    }

    // Helper to calculate tenure and seasoning
    function calculateTenure(firstDate: Date | null, contractsCount: number) {
      let tenureMonths = 0;
      let tenureYears = 0;
      if (firstDate && !isNaN(firstDate.getTime())) {
        tenureMonths = Math.max(
          1,
          Math.round(
            (model.asOfDate.getTime() - firstDate.getTime()) / (1000 * 60 * 60 * 24 * 30.4375),
          ),
        );
        tenureYears = +(tenureMonths / 12).toFixed(1);
      }

      let seasoningClassification:
        | 'Seasoned Program (≥2 Years)'
        | 'Developing Program (1-2 Years)'
        | 'Emerging / Unseasoned (<1 Year)';
      if (tenureYears >= 2 && contractsCount >= 50) {
        seasoningClassification = 'Seasoned Program (≥2 Years)';
      } else if (tenureYears >= 1) {
        seasoningClassification = 'Developing Program (1-2 Years)';
      } else {
        seasoningClassification = 'Emerging / Unseasoned (<1 Year)';
      }

      return { tenureMonths, tenureYears, seasoningClassification };
    }

    // Helper to extract product lines and detect cross-subsidization
    function buildProductLines(
      targetContracts: typeof model.contractTransactions,
      targetClaims: typeof model.claims,
    ): ProductLineSummary[] {
      const contractProductMap = new Map<string, string>();
      for (const t of targetContracts) {
        if (t.contractNumber && t.product) {
          contractProductMap.set(t.contractNumber, t.product);
        }
      }

      const productNames = new Set<string>();
      targetContracts.forEach((t) => productNames.add(t.product || 'Standard'));
      targetClaims.forEach((c) => {
        const p = contractProductMap.get(c.contractNumber) || c.product || 'Standard';
        productNames.add(p);
      });

      const lines: ProductLineSummary[] = [];
      for (const prod of productNames) {
        let pWritten = 0;
        let pEarned = 0;
        targetContracts
          .filter((t) => (t.product || 'Standard') === prod)
          .forEach((t) => {
            if (t.transactionType !== 'Cancellation') {
              pWritten += t.reserveAmount;
              pEarned += t.earnedReserveAmount;
            } else {
              pWritten -= Math.abs(t.reserveAmount);
            }
          });

        const pClaims = targetClaims
          .filter(
            (c) => (contractProductMap.get(c.contractNumber) || c.product || 'Standard') === prod,
          )
          .reduce((sum, c) => sum + c.paidAmount, 0);

        if (pWritten > 0 || pEarned > 0 || pClaims > 0) {
          const pProfit = pEarned - pClaims;
          const pLr = pEarned > 0 ? pClaims / pEarned : null;
          lines.push({
            name: prod,
            writtenReserve: pWritten,
            earnedReserve: pEarned,
            claimsPaid: pClaims,
            underwritingProfit: pProfit,
            lossRatio: pLr,
          });
        }
      }

      // Sort by largest loss first (most negative profit)
      lines.sort((a, b) => a.underwritingProfit - b.underwritingProfit);
      return lines;
    }

    // Check Agents
    for (const agentInput of debugAgents) {
      const rolling = model.agents.find((a) => matchesAgent(a.name, a.displayName, [agentInput]));
      const itd = model.itdAgents.find((a) => matchesAgent(a.name, a.displayName, [agentInput]));

      // Extract matching contracts to determine earliest active date
      const targetContracts = model.contractTransactions.filter((t) =>
        matchesAgent(t.agent, t.agentName, [agentInput]),
      );
      let firstContractDate: Date | null = null;
      for (const t of targetContracts) {
        const d = t.effectiveDate || t.activityDate;
        if (d && !isNaN(d.getTime())) {
          if (!firstContractDate || d.getTime() < firstContractDate.getTime()) {
            firstContractDate = d;
          }
        }
      }

      // Extract matching claims for top loss codes
      const targetClaims = model.claims.filter((cl) =>
        matchesAgent(cl.agent, cl.agentName, [agentInput]),
      );
      const lossMap = new Map<string, { desc: string; amount: number }>();
      for (const cl of targetClaims) {
        const code = cl.lossCode || 'UNMAPPED';
        const entry = lossMap.get(code) || { desc: cl.lossCodeDescription || '', amount: 0 };
        entry.amount += cl.paidAmount;
        lossMap.set(code, entry);
      }
      const topLossCodes = Array.from(lossMap.entries())
        .map(([code, data]) => ({ code, description: data.desc, amount: data.amount }))
        .sort((a, b) => b.amount - a.amount);

      const itdReserve = itd?.netReserve || 0;
      const itdEarnedReserve = itd?.earnedReserve || 0;
      const unearnedReserve = Math.max(0, itdReserve - itdEarnedReserve);
      const percentReserveEarned = itdReserve > 0 ? (itdEarnedReserve / itdReserve) * 100 : 0;
      const itdContracts = itd?.contractsWritten || 0;

      const { tenureMonths, tenureYears, seasoningClassification } = calculateTenure(
        firstContractDate,
        itdContracts,
      );

      const productLines = buildProductLines(targetContracts, targetClaims);

      targets.push({
        type: 'Agent',
        name: itd?.displayName || rolling?.displayName || agentInput.toUpperCase(),
        code:
          itd?.name ||
          rolling?.name ||
          (agentInput.toLowerCase() === 'executive warranties' ? 'EXW' : agentInput.toUpperCase()),
        tenureYears,
        tenureMonths,
        firstContractDate,
        seasoningClassification,
        unearnedReserve,
        percentReserveEarned,
        rollingContracts: rolling?.contractsWritten || 0,
        rollingReserve: rolling?.netReserve || 0,
        rollingClaims: rolling?.claimsPaid || 0,
        rollingLossRatio: rolling?.earnedLossRatio ?? 0,
        itdContracts,
        itdReserve,
        itdEarnedReserve,
        itdClaims: itd?.claimsPaid || 0,
        itdLossRatio: itd?.earnedLossRatio ?? 0,
        itdClaimFrequency:
          itd?.claimFrequency ?? (itdContracts > 0 ? (itd?.claimCount || 0) / itdContracts : 0),
        rollingClaimFrequency:
          rolling?.claimFrequency ??
          ((rolling?.contractsWritten || 0) > 0
            ? (rolling?.claimCount || 0) / (rolling?.contractsWritten || 1)
            : 0),
        topLossCodes,
        productLines,
      });
    }

    // Check Dealers
    for (const dealerInput of debugDealers) {
      const rolling = model.dealers.find(
        (d) => d.name.toLowerCase() === dealerInput || d.displayName?.toLowerCase() === dealerInput,
      );
      const itd = model.itdDealers.find(
        (d) => d.name.toLowerCase() === dealerInput || d.displayName?.toLowerCase() === dealerInput,
      );

      const targetContracts = model.contractTransactions.filter(
        (t) =>
          t.dealerNumber.toLowerCase() === dealerInput ||
          (t.dealerName && t.dealerName.toLowerCase() === dealerInput) ||
          t.dealer.toLowerCase() === dealerInput,
      );
      let firstContractDate: Date | null = null;
      for (const t of targetContracts) {
        const d = t.effectiveDate || t.activityDate;
        if (d && !isNaN(d.getTime())) {
          if (!firstContractDate || d.getTime() < firstContractDate.getTime()) {
            firstContractDate = d;
          }
        }
      }

      const targetClaims = model.claims.filter(
        (cl) =>
          cl.dealer.toLowerCase() === dealerInput ||
          (cl.dealerName && cl.dealerName.toLowerCase() === dealerInput),
      );
      const lossMap = new Map<string, { desc: string; amount: number }>();
      for (const cl of targetClaims) {
        const code = cl.lossCode || 'UNMAPPED';
        const entry = lossMap.get(code) || { desc: cl.lossCodeDescription || '', amount: 0 };
        entry.amount += cl.paidAmount;
        lossMap.set(code, entry);
      }
      const topLossCodes = Array.from(lossMap.entries())
        .map(([code, data]) => ({ code, description: data.desc, amount: data.amount }))
        .sort((a, b) => b.amount - a.amount);

      const itdReserve = itd?.netReserve || 0;
      const itdEarnedReserve = itd?.earnedReserve || 0;
      const unearnedReserve = Math.max(0, itdReserve - itdEarnedReserve);
      const percentReserveEarned = itdReserve > 0 ? (itdEarnedReserve / itdReserve) * 100 : 0;
      const itdContracts = itd?.contractsWritten || 0;

      const { tenureMonths, tenureYears, seasoningClassification } = calculateTenure(
        firstContractDate,
        itdContracts,
      );

      const productLines = buildProductLines(targetContracts, targetClaims);

      targets.push({
        type: 'Dealer',
        name: itd?.displayName || rolling?.displayName || dealerInput.toUpperCase(),
        code: itd?.name || rolling?.name || dealerInput.toUpperCase(),
        tenureYears,
        tenureMonths,
        firstContractDate,
        seasoningClassification,
        unearnedReserve,
        percentReserveEarned,
        rollingContracts: rolling?.contractsWritten || 0,
        rollingReserve: rolling?.netReserve || 0,
        rollingClaims: rolling?.claimsPaid || 0,
        rollingLossRatio: rolling?.earnedLossRatio ?? 0,
        itdContracts,
        itdReserve,
        itdEarnedReserve,
        itdClaims: itd?.claimsPaid || 0,
        itdLossRatio: itd?.earnedLossRatio ?? 0,
        itdClaimFrequency:
          itd?.claimFrequency ?? (itdContracts > 0 ? (itd?.claimCount || 0) / itdContracts : 0),
        rollingClaimFrequency:
          rolling?.claimFrequency ??
          ((rolling?.contractsWritten || 0) > 0
            ? (rolling?.claimCount || 0) / (rolling?.contractsWritten || 1)
            : 0),
        topLossCodes,
        productLines,
      });
    }

    // 4. Generate AI Overview for Primary Target (or merged overview)
    let overviewText = 'No targets found for debug analysis.';
    if (targets.length > 0) {
      const primary = targets[0]!;
      overviewText = await this.aiService.generateExecutiveOverview(primary);
    }

    let currentRow = wipDescRow + 3;

    // Overview Section Header
    ws.mergeCells(`A${currentRow}:I${currentRow}`);
    const overviewHeader = ws.getCell(`A${currentRow}`);
    overviewHeader.value = '💡 AI EXECUTIVE OVERVIEW (< 1 PARAGRAPH)';
    overviewHeader.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.white } };
    overviewHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    overviewHeader.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    currentRow++;

    // Overview Text Box (Merged, less than 1 paragraph)
    const textStartRow = currentRow;
    const textEndRow = currentRow + 2;
    ws.mergeCells(`A${textStartRow}:I${textEndRow}`);
    const overviewCell = ws.getCell(`A${textStartRow}`);
    overviewCell.value = overviewText;
    overviewCell.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FF1F4E79' } };
    overviewCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F7FC' } };
    overviewCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true, indent: 1 };

    for (let r = textStartRow; r <= textEndRow; r++) {
      ws.getRow(r).height = 24;
    }

    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach((col) => {
      ws.getCell(`${col}${textEndRow}`).border = {
        bottom: { style: 'thin', color: { argb: COLORS.blue } },
      };
    });

    currentRow = textEndRow + 2;

    // 5. Target Performance Profile Table
    ws.mergeCells(`A${currentRow}:K${currentRow}`);
    const tableTitle = ws.getCell(`A${currentRow}`);
    tableTitle.value = '📊 TARGET PERFORMANCE METRICS (ROLLING 12 vs. INCEPTION-TO-DATE)';
    tableTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.navy } };
    currentRow++;

    // Table Headers
    const headers = [
      'Target Name',
      'Dimension',
      'Period / Tenure',
      'Contracts',
      'Net Reserve ($)',
      'Earned Reserve ($)',
      'Claims Paid ($)',
      'Loss Ratio',
      'Frequency of Claims',
      '% Earned',
      'Actuarial Credibility & Status',
    ];
    const headerRow = ws.getRow(currentRow);
    headerRow.values = headers;
    headerRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    // Table Data
    for (const t of targets) {
      const getCredibilityStatus = (ratio: number, isItd: boolean) => {
        if (!isItd) {
          if (ratio <= config.warningLossRatio) return 'Rolling: Favorable (<65%)';
          if (ratio <= config.highLossRatio) return 'Rolling: Warning (65-80%)';
          return 'Rolling: Elevated (>80%)';
        }

        if (t.seasoningClassification === 'Seasoned Program (≥2 Years)') {
          if (ratio > config.highLossRatio || ratio > 0.75) {
            return 'Seasoned Adverse Experience (Pay Attention)';
          }
          if (ratio <= config.warningLossRatio) {
            return 'Seasoned Profitable (<65%)';
          }
          return 'Seasoned Moderate Margin (65-75%)';
        }

        if (t.seasoningClassification === 'Developing Program (1-2 Years)') {
          return `Developing Program (${t.percentReserveEarned.toFixed(0)}% Earned)`;
        }

        return `Immature Cohort (${t.tenureMonths} Mo, Sample Volatility)`;
      };

      const dateStr = t.firstContractDate
        ? t.firstContractDate.toLocaleDateString('en-US', { month: '2-digit', year: 'numeric' })
        : 'N/A';

      // Rolling 12 row
      const r12PctEarned = t.rollingReserve > 0 ? 1 : 0;
      const r12Freq = t.rollingClaimFrequency ?? 0;
      const r12Row = ws.addRow([
        `${t.name} (${t.code})`,
        t.type,
        'Rolling 12-Month',
        t.rollingContracts,
        t.rollingReserve,
        t.rollingReserve, // Proxy for rolling earned
        t.rollingClaims,
        t.rollingLossRatio,
        r12Freq,
        r12PctEarned,
        getCredibilityStatus(t.rollingLossRatio, false),
      ]);
      r12Row.font = { name: 'Aptos', size: 10 };
      r12Row.alignment = { vertical: 'middle' };
      r12Row.getCell(8).font = { bold: true };
      r12Row.getCell(8).numFmt = PERCENT;
      r12Row.getCell(9).numFmt = PERCENT;
      r12Row.getCell(10).numFmt = PERCENT;
      r12Row.getCell(11).font = {
        bold: true,
        color: {
          argb:
            t.rollingLossRatio > config.highLossRatio
              ? COLORS.red
              : t.rollingLossRatio > config.warningLossRatio
                ? COLORS.amber
                : COLORS.green,
        },
      };
      currentRow++;

      // ITD row
      const itdPctEarned = t.itdReserve > 0 ? t.itdEarnedReserve / t.itdReserve : 0;
      const itdPeriodStr = `ITD (${t.tenureYears} Yrs, Since ${dateStr})`;
      const itdFreq = t.itdClaimFrequency ?? 0;
      const itdRow = ws.addRow([
        `${t.name} (${t.code})`,
        t.type,
        itdPeriodStr,
        t.itdContracts,
        t.itdReserve,
        t.itdEarnedReserve,
        t.itdClaims,
        t.itdLossRatio,
        itdFreq,
        itdPctEarned,
        getCredibilityStatus(t.itdLossRatio, true),
      ]);
      itdRow.font = { name: 'Aptos', size: 10, bold: true };
      itdRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      itdRow.alignment = { vertical: 'middle' };
      itdRow.getCell(8).font = { bold: true };
      itdRow.getCell(8).numFmt = PERCENT;
      itdRow.getCell(9).numFmt = PERCENT;
      itdRow.getCell(10).numFmt = PERCENT;
      itdRow.getCell(11).font = {
        bold: true,
        color: {
          argb:
            t.itdLossRatio > config.highLossRatio &&
            t.seasoningClassification === 'Seasoned Program (≥2 Years)'
              ? COLORS.red
              : t.itdLossRatio > config.warningLossRatio
                ? COLORS.amber
                : COLORS.green,
        },
      };
      currentRow++;
    }

    currentRow += 2;

    // 5.5 Product Line Underwriting & Cross-Subsidization Analysis
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const prodTableTitle = ws.getCell(`A${currentRow}`);
    prodTableTitle.value = '🎯 PRODUCT LINE UNDERWRITING & CROSS-SUBSIDIZATION ANALYSIS';
    prodTableTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.navy } };
    currentRow++;

    const prodHeaders = [
      'Target Context',
      'Product / Coverage Line',
      'Underwriting Role',
      'Written Reserve ($)',
      'Earned Reserve ($)',
      'Claims Paid ($)',
      'UW Profit / (Deficit)',
      'Loss Ratio',
      '% of Claims',
      'Cross-Subsidization Assessment',
    ];
    const prodHeaderRow = ws.getRow(currentRow);
    prodHeaderRow.values = prodHeaders;
    prodHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    prodHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    prodHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    for (const t of targets) {
      if (!t.productLines || t.productLines.length === 0) continue;
      for (const p of t.productLines) {
        const isLoss = p.underwritingProfit < 0;
        const claimShare = t.itdClaims > 0 ? p.claimsPaid / t.itdClaims : 0;
        const roleLabel =
          p.underwritingProfit < -10000
            ? '🔴 CRITICAL LOSS DRIVER'
            : p.underwritingProfit < 0
              ? '🔴 LOSS DEFICIT'
              : p.underwritingProfit === 0
                ? '⚪ BREAK-EVEN'
                : '🟢 PROFITABLE SUBSIDIZER';

        const assessment = isLoss
          ? `Bleeding -$${Math.abs(Math.round(p.underwritingProfit)).toLocaleString()} (${((p.lossRatio || 0) * 100).toFixed(1)}% LR)`
          : `Subsidizing with +$${Math.round(p.underwritingProfit).toLocaleString()} surplus (${((p.lossRatio || 0) * 100).toFixed(1)}% LR)`;

        const row = ws.addRow([
          `${t.name} (${t.code})`,
          p.name,
          roleLabel,
          p.writtenReserve,
          p.earnedReserve,
          p.claimsPaid,
          p.underwritingProfit,
          p.lossRatio,
          claimShare,
          assessment,
        ]);

        row.font = { name: 'Aptos', size: 9 };
        row.getCell(4).numFmt = MONEY;
        row.getCell(5).numFmt = MONEY;
        row.getCell(6).numFmt = MONEY;
        row.getCell(7).numFmt = MONEY;
        row.getCell(8).numFmt = PERCENT;
        row.getCell(9).numFmt = PERCENT;

        if (isLoss) {
          row.getCell(3).font = { bold: true, color: { argb: 'FF9C0006' } };
          row.getCell(7).font = { bold: true, color: { argb: 'FF9C0006' } };
          row.getCell(8).font = { bold: true, color: { argb: 'FF9C0006' } };
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } };
        } else {
          row.getCell(3).font = { bold: true, color: { argb: 'FF006100' } };
          row.getCell(7).font = { bold: true, color: { argb: 'FF006100' } };
          row.getCell(8).font = { bold: true, color: { argb: 'FF006100' } };
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
        }
        currentRow++;
      }
    }

    currentRow += 2;

    // 6. Top Loss Code Drivers
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const lossTitle = ws.getCell(`A${currentRow}`);
    lossTitle.value = '🔍 TOP LOSS CODE EXPOSURES (PRIMARY CLAIM DRIVERS)';
    lossTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.navy } };
    currentRow++;

    const lossHeaderRow = ws.getRow(currentRow);
    lossHeaderRow.values = [
      'Loss Code',
      'Description',
      'Target Context',
      'Paid Claims ($)',
      '% of Target Claims',
      '',
      '',
      '',
      '',
    ];
    lossHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    lossHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF595959' } };
    currentRow++;

    for (const t of targets) {
      const topCodes = t.topLossCodes.slice(0, 5);
      for (const lc of topCodes) {
        const pct = t.itdClaims > 0 ? lc.amount / t.itdClaims : 0;
        const row = ws.addRow([
          lc.code,
          lc.description || 'General Component Failure',
          `${t.name} (${t.code})`,
          lc.amount,
          pct,
        ]);
        row.font = { name: 'Aptos', size: 9 };
        row.getCell(4).numFmt = MONEY;
        row.getCell(5).numFmt = PERCENT;
        currentRow++;
      }
    }

    currentRow += 2;

    // 7. Upcoming Bucket-Level Drill-Down Note
    ws.mergeCells(`A${currentRow}:I${currentRow}`);
    const roadmapHeader = ws.getCell(`A${currentRow}`);
    roadmapHeader.value = '🎯 TRAINING ROADMAP: UPCOMING COMPONENT BUCKET DRILL-DOWN';
    roadmapHeader.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FF595959' } };
    currentRow++;

    const roadmapDescRow = ws.addRow([
      'Next Agent Capabilities: (1) Sub-ledger bucket decomposition (Admin vs. Reserve vs. Surcharges), (2) Dynamic Earning Curve deviation tracking per contract, (3) Automated Part & Labor benchmark flags.',
    ]);
    roadmapDescRow.font = { name: 'Aptos', size: 9, italic: true, color: { argb: 'FF7F7F7F' } };
    ws.mergeCells(`A${currentRow}:I${currentRow}`);
  }
}
