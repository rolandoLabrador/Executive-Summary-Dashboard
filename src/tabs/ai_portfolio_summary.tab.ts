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
import { AiAnalystService, type PortfolioSummaryData } from '../services/ai_analyst.service';

export class AiPortfolioSummaryTab implements IDashboardTab {
  readonly id = 'tab_ai_portfolio_summary';
  private aiService = new AiAnalystService();

  constructor() {}

  async render(
    workbook: ExcelJS.Workbook,
    model: ReportModel,
    config: ReportConfig,
    _tabConfig?: unknown,
  ): Promise<void> {
    const ws = workbook.addWorksheet('AI Portfolio Summary');
    configureWorksheet(ws);

    ws.columns = [
      { width: 28 }, // A: Period / Agent Name
      { width: 18 }, // B: Tenure & Start
      { width: 16 }, // C: Program Maturity
      { width: 14, numFmt: INTEGER }, // D: Contracts Written
      { width: 18, numFmt: MONEY }, // E: Net Reserve
      { width: 18, numFmt: MONEY }, // F: Earned Reserve
      { width: 18, numFmt: MONEY }, // G: Claims Paid
      { width: 16, numFmt: PERCENT }, // H: Loss Ratio
      { width: 14, numFmt: PERCENT }, // I: % Earned
      { width: 28 }, // J: Credibility & Status
    ];

    // 1. Header Title
    title(
      ws,
      '🌐 AI PORTFOLIO EXECUTIVE SUMMARY — MACRO AUDIT LAB',
      `Enterprise Underwriting & Actuarial Intelligence | As of ${formatDate(model.asOfDate)}`,
      9,
    );

    // 2. WORK IN PROGRESS NOTICE (Prominently at the top)
    ws.addRow([]);
    const wipRowStart = 3;
    ws.mergeCells(`A${wipRowStart}:I${wipRowStart}`);
    const wipTitle = ws.getCell(`A${wipRowStart}`);
    wipTitle.value = '⚠️ WORK IN PROGRESS — PORTFOLIO UNDERWRITING AI';
    wipTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: 'FF833C00' } };
    wipTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
    wipTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

    const wipDescRow = wipRowStart + 1;
    ws.mergeCells(`A${wipDescRow}:I${wipDescRow + 1}`);
    const wipDesc = ws.getCell(`A${wipDescRow}`);
    wipDesc.value =
      'This is a work in progress. We are actively training this macro AI Agent to evaluate enterprise-wide portfolio health, loss ratio trajectory, underwriting profit, and cross-agent risk distribution across all product lines.';
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

    // 3. Prepare Macro Metrics for the AI Agent
    const rollingCurrent = model.rolling12.current;
    const rollingPrior = model.rolling12.prior;
    const itdMetrics = model.inceptionToDate.values;

    // Precompute earliest contract date for every agent
    const agentFirstDateMap = new Map<string, Date>();
    for (const t of model.contractTransactions) {
      const d = t.effectiveDate || t.activityDate;
      if (d && !isNaN(d.getTime())) {
        const key = t.agent.toLowerCase();
        const existing = agentFirstDateMap.get(key);
        if (!existing || d.getTime() < existing.getTime()) {
          agentFirstDateMap.set(key, d);
        }
        if (t.agentName) {
          const nameKey = t.agentName.toLowerCase();
          const nameExisting = agentFirstDateMap.get(nameKey);
          if (!nameExisting || d.getTime() < nameExisting.getTime()) {
            agentFirstDateMap.set(nameKey, d);
          }
        }
      }
    }

    // Identify top risk and profitable agents
    const agentsWithVolume = model.itdAgents.filter((a) => a.contractsWritten >= 50);
    const topAgentRisk = [...agentsWithVolume].sort(
      (a, b) => (b.earnedLossRatio ?? 0) - (a.earnedLossRatio ?? 0),
    )[0];
    const topAgentProfitable = [...agentsWithVolume].sort(
      (a, b) => (a.earnedLossRatio ?? 0) - (b.earnedLossRatio ?? 0),
    )[0];

    let topRiskTenureYears: number | undefined;
    let topRiskFirstDate: Date | null = null;
    if (topAgentRisk) {
      const fd =
        agentFirstDateMap.get(topAgentRisk.name.toLowerCase()) ||
        (topAgentRisk.displayName
          ? agentFirstDateMap.get(topAgentRisk.displayName.toLowerCase())
          : null);
      if (fd) {
        topRiskFirstDate = fd;
        const m = Math.max(
          1,
          Math.round((model.asOfDate.getTime() - fd.getTime()) / (1000 * 60 * 60 * 24 * 30.4375)),
        );
        topRiskTenureYears = +(m / 12).toFixed(1);
      }
    }

    let topProfitableTenureYears: number | undefined;
    let topProfitableFirstDate: Date | null = null;
    if (topAgentProfitable) {
      const fd =
        agentFirstDateMap.get(topAgentProfitable.name.toLowerCase()) ||
        (topAgentProfitable.displayName
          ? agentFirstDateMap.get(topAgentProfitable.displayName.toLowerCase())
          : null);
      if (fd) {
        topProfitableFirstDate = fd;
        const m = Math.max(
          1,
          Math.round((model.asOfDate.getTime() - fd.getTime()) / (1000 * 60 * 60 * 24 * 30.4375)),
        );
        topProfitableTenureYears = +(m / 12).toFixed(1);
      }
    }

    // Identify top portfolio loss code
    const topLossCodeEntry = model.lossCodeDashboard?.rows?.[0];

    const portfolioData: PortfolioSummaryData = {
      companyName: config.companyName,
      asOfDate: model.asOfDate,
      rollingContracts: rollingCurrent.contractsWritten,
      rollingReserve: rollingCurrent.netReserve,
      rollingEarnedReserve: rollingCurrent.earnedReserve,
      rollingClaims: rollingCurrent.claimsPaid,
      rollingLossRatio: rollingCurrent.earnedLossRatio ?? 0,
      rollingUnderwritingProfit: rollingCurrent.underwritingProfit,
      priorRollingContracts: rollingPrior.contractsWritten,
      priorRollingLossRatio: rollingPrior.earnedLossRatio ?? 0,
      itdContracts: itdMetrics.contractsWritten,
      itdEarnedReserve: itdMetrics.earnedReserve,
      itdClaims: itdMetrics.claimsPaid,
      itdLossRatio: itdMetrics.earnedLossRatio ?? 0,
      itdUnderwritingProfit: itdMetrics.underwritingProfit,
      topAgentRisk: topAgentRisk
        ? {
            name: topAgentRisk.displayName || topAgentRisk.name,
            lossRatio: topAgentRisk.earnedLossRatio ?? 0,
            paidClaims: topAgentRisk.claimsPaid,
            tenureYears: topRiskTenureYears,
            firstDate: topRiskFirstDate,
          }
        : undefined,
      topAgentProfitable: topAgentProfitable
        ? {
            name: topAgentProfitable.displayName || topAgentProfitable.name,
            lossRatio: topAgentProfitable.earnedLossRatio ?? 0,
            contracts: topAgentProfitable.contractsWritten,
            tenureYears: topProfitableTenureYears,
            firstDate: topProfitableFirstDate,
          }
        : undefined,
      topLossCode: topLossCodeEntry
        ? {
            code: topLossCodeEntry.code,
            description: topLossCodeEntry.description,
            amount: topLossCodeEntry.rolling12Paid,
          }
        : undefined,
    };

    // 4. Generate AI Portfolio Overview (< 1 paragraph)
    const overviewText = await this.aiService.generatePortfolioOverview(portfolioData);

    let currentRow = wipDescRow + 3;

    // Overview Header
    ws.mergeCells(`A${currentRow}:I${currentRow}`);
    const overviewHeader = ws.getCell(`A${currentRow}`);
    overviewHeader.value = '💡 AI PORTFOLIO OVERVIEW (< 1 PARAGRAPH)';
    overviewHeader.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.white } };
    overviewHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    overviewHeader.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    currentRow++;

    // Text Box
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

    // 5. Macro Underwriting & Financial Scorecard
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const tableTitle = ws.getCell(`A${currentRow}`);
    tableTitle.value = '📊 ENTERPRISE UNDERWRITING & FINANCIAL SCORECARD';
    tableTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.navy } };
    currentRow++;

    const scorecardHeaders = [
      'Reporting Period',
      'Benchmark Scope',
      'Volume Status',
      'Contracts Written',
      'Net Reserve ($)',
      'Earned Reserve ($)',
      'Claims Paid ($)',
      'Earned Loss Ratio',
      '% Earned',
      'Underwriting Status',
    ];
    const sHeaderRow = ws.getRow(currentRow);
    sHeaderRow.values = scorecardHeaders;
    sHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    sHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    sHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    const getStatus = (ratio: number) => {
      if (ratio <= config.warningLossRatio) return 'Profitable (<65%)';
      if (ratio <= config.highLossRatio) return 'Warning Zone (65-80%)';
      return 'Adverse (>80%)';
    };

    const periodRows = [
      {
        name: 'Current Month',
        scope: '30-Day Activity',
        status: 'Closed Accounting Period',
        vals: model.currentMonth.current,
      },
      {
        name: 'Year to Date',
        scope: 'Calendar Year 2026',
        status: 'Cumulative YTD Production',
        vals: model.yearToDate.current,
      },
      {
        name: 'Rolling 12-Month',
        scope: '12-Month Trailing Horizon',
        status: 'Seasoning Underwriting Base',
        vals: model.rolling12.current,
      },
      {
        name: 'Prior Calendar Year',
        scope: 'Full Year 2025',
        status: 'Historical Baseline',
        vals: model.priorCalendarYear.values,
      },
      {
        name: 'Inception to Date',
        scope: 'All Historical Cycles',
        status: 'Cumulative Enterprise Total',
        vals: model.inceptionToDate.values,
        isBold: true,
      },
    ];

    for (const p of periodRows) {
      const v = p.vals;
      const lr = v.earnedLossRatio ?? 0;
      const pctEarned =
        v.percentEarned ?? (v.netReserve > 0 ? v.earnedReserve / v.netReserve : null);
      const row = ws.addRow([
        p.name,
        p.scope,
        p.status,
        v.contractsWritten,
        v.netReserve,
        v.earnedReserve,
        v.claimsPaid,
        lr,
        pctEarned,
        getStatus(lr),
      ]);
      row.font = { name: 'Aptos', size: 10, bold: Boolean(p.isBold) };
      if (p.isBold) {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      }
      row.getCell(4).numFmt = INTEGER;
      row.getCell(5).numFmt = MONEY;
      row.getCell(6).numFmt = MONEY;
      row.getCell(7).numFmt = MONEY;
      row.getCell(8).numFmt = PERCENT;
      row.getCell(8).font = { bold: true };
      row.getCell(9).numFmt = PERCENT;
      row.getCell(10).font = {
        bold: true,
        color: {
          argb:
            lr > config.highLossRatio
              ? COLORS.red
              : lr > config.warningLossRatio
                ? COLORS.amber
                : COLORS.green,
        },
      };
      currentRow++;
    }

    currentRow += 2;

    // 6. Agent Portfolio Performance & Risk Distribution (With Tenure Context)
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const agentTitle = ws.getCell(`A${currentRow}`);
    agentTitle.value = '👥 AGENT RISK & VOLUME DISTRIBUTION (TENURE & SEASONED PERFORMANCE)';
    agentTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.navy } };
    currentRow++;

    const agentHeaders = [
      'Agent Name & Code',
      'Tenure & Active Since',
      'Program Maturity',
      'Contracts Written',
      'Net Reserve ($)',
      'Earned Reserve ($)',
      'Claims Paid ($)',
      'Earned Loss Ratio',
      '% Earned',
      'Actuarial Credibility & Status',
    ];
    const aHeaderRow = ws.getRow(currentRow);
    aHeaderRow.values = agentHeaders;
    aHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    aHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF595959' } };
    aHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    const topAgents = [...model.itdAgents]
      .sort((a, b) => b.contractsWritten - a.contractsWritten)
      .slice(0, 8);

    for (const a of topAgents) {
      const lr = a.earnedLossRatio ?? 0;
      const pctEarned =
        a.percentEarned ?? (a.netReserve > 0 ? a.earnedReserve / a.netReserve : null);
      const firstDate =
        agentFirstDateMap.get(a.name.toLowerCase()) ||
        (a.displayName ? agentFirstDateMap.get(a.displayName.toLowerCase()) : null);

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

      let maturity = 'New (<1 Yr)';
      if (tenureYears >= 2 && a.contractsWritten >= 50) {
        maturity = 'Seasoned (≥2 Yrs)';
      } else if (tenureYears >= 1) {
        maturity = 'Developing (1-2 Yrs)';
      }

      let status = '';
      if (maturity === 'Seasoned (≥2 Yrs)') {
        if (lr > config.highLossRatio || lr > 0.75) {
          status = 'Seasoned Adverse (Pay Attention)';
        } else if (lr <= config.warningLossRatio) {
          status = 'Seasoned Profitable (<65%)';
        } else {
          status = 'Seasoned Moderate Margin';
        }
      } else if (maturity === 'Developing (1-2 Yrs)') {
        status = 'Developing Cohort (Reserves Earning)';
      } else {
        status = 'Immature Cohort (Sample Volatility)';
      }

      const dateStr = firstDate
        ? firstDate.toLocaleDateString('en-US', { month: '2-digit', year: 'numeric' })
        : 'N/A';
      const tenureStr = `${tenureYears} Yrs (${dateStr})`;

      const row = ws.addRow([
        `${a.displayName || a.name} (${a.name})`,
        tenureStr,
        maturity,
        a.contractsWritten,
        a.netReserve,
        a.earnedReserve,
        a.claimsPaid,
        lr,
        pctEarned,
        status,
      ]);
      row.font = { name: 'Aptos', size: 9 };
      row.getCell(4).numFmt = INTEGER;
      row.getCell(5).numFmt = MONEY;
      row.getCell(6).numFmt = MONEY;
      row.getCell(7).numFmt = MONEY;
      row.getCell(8).numFmt = PERCENT;
      row.getCell(8).font = { bold: true };
      row.getCell(9).numFmt = PERCENT;

      if (maturity === 'Seasoned (≥2 Yrs)' && (lr > config.highLossRatio || lr > 0.75)) {
        row.getCell(10).font = { bold: true, color: { argb: COLORS.red } };
      } else if (lr <= config.warningLossRatio) {
        row.getCell(10).font = { bold: true, color: { argb: COLORS.green } };
      } else {
        row.getCell(10).font = { bold: true, color: { argb: COLORS.amber } };
      }
      currentRow++;
    }

    currentRow += 2;

    // 7. Top 5 Portfolio-Wide Claim Exposures
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const lossTitle = ws.getCell(`A${currentRow}`);
    lossTitle.value = '🔍 TOP 5 PORTFOLIO-WIDE CLAIM EXPOSURES (COMPANY-WIDE RISK DRIVERS)';
    lossTitle.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.navy } };
    currentRow++;

    const lossHeaders = [
      'Loss Code',
      'Component Description',
      'Coverage Lines',
      'Rolling 12M Paid ($)',
      'Claim Count',
      'Avg Paid / Claim ($)',
      '% of Company Claims',
      'Actuarial Classification',
      '',
    ];
    const lHeaderRow = ws.getRow(currentRow);
    lHeaderRow.values = lossHeaders;
    lHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    lHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    currentRow++;

    const topLossRows = (model.lossCodeDashboard?.rows || []).slice(0, 5);
    for (const lc of topLossRows) {
      const r = ws.addRow([
        lc.code,
        lc.description || 'Component Failure',
        (lc.coverageNames || []).slice(0, 2).join(', ') || 'VSC / GAP',
        lc.rolling12Paid,
        lc.rolling12ClaimCount,
        lc.rolling12AveragePaidPerClaim ?? 0,
        lc.rolling12PaidShare ?? 0,
        (lc.rolling12PaidShare ?? 0) > 0.15
          ? 'Primary Enterprise Loss Driver'
          : 'Standard Exposure',
        '',
      ]);
      r.font = { name: 'Aptos', size: 9 };
      r.getCell(4).numFmt = MONEY;
      r.getCell(5).numFmt = INTEGER;
      r.getCell(6).numFmt = MONEY;
      r.getCell(7).numFmt = PERCENT;
      currentRow++;
    }
  }
}
