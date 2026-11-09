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

export class AiUnderwritingIntelTab implements IDashboardTab {
  readonly id = 'tab_ai_underwriting_intel';

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

    if (debugDealers.length === 0 && debugAgents.length === 0) {
      return;
    }

    const ws = workbook.addWorksheet('AI Underwriting Intel');
    configureWorksheet(ws);

    ws.columns = [
      { width: 26 }, // Col A: Dealer Name
      { width: 14 }, // Col B: Dealer #
      { width: 20 }, // Col C: Tenure & Active Since
      { width: 18 }, // Col D: Program Maturity
      { width: 14 }, // Col E: Contracts Written
      { width: 18 }, // Col F: Net Reserve
      { width: 18 }, // Col G: Claims Paid
      { width: 16 }, // Col H: Loss Ratio
      { width: 14 }, // Col I: % Earned
      { width: 30 }, // Col J: Actuarial Credibility & Status
    ];

    // Title Block
    title(
      ws,
      '🔬 AI DEEP UNDERWRITING INTELLIGENCE — ACTUARIAL AUDIT LAB',
      `Granular Root-Cause Diagnostics & Cohort Breakdown | As of ${formatDate(model.asOfDate)}`,
      9,
    );

    function matchesAgent(
      code?: string,
      name?: string,
      debugList: string[] = debugAgents,
    ): boolean {
      if (debugList.length === 0) return false;
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

    // Filter relevant contracts and claims for debug targets
    const targetContracts = model.contractTransactions.filter((t) => {
      const agentMatch = matchesAgent(t.agent, t.agentName);
      const dealerMatch =
        debugDealers.includes(t.dealerNumber.toLowerCase()) ||
        debugDealers.includes((t.dealerName || t.dealer).toLowerCase());
      return agentMatch || dealerMatch;
    });

    const targetClaims = model.claims.filter((cl) => {
      const agentMatch = matchesAgent(cl.agent, cl.agentName);
      const dealerMatch =
        debugDealers.includes(cl.dealer.toLowerCase()) ||
        (cl.dealerName && debugDealers.includes(cl.dealerName.toLowerCase()));
      return agentMatch || dealerMatch;
    });

    // Map contract numbers to effective dates for time-to-loss calculation
    const contractEffectiveMap = new Map<string, Date>();
    for (const t of model.contractTransactions) {
      if (t.effectiveDate && !contractEffectiveMap.has(t.contractNumber)) {
        contractEffectiveMap.set(t.contractNumber, t.effectiveDate);
      }
    }

    let currentRow = 3;

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 1: DEALER CONCENTRATION ANALYSIS (With Tenure & Credibility)
    // ─────────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const sec1Header = ws.getCell(`A${currentRow}`);
    sec1Header.value =
      '🏢 1. DEALER COHORT CONCENTRATION (TENURE & SEASONED VS. NEW DEALER PERFORMANCE)';
    sec1Header.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.white } };
    sec1Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    currentRow++;

    const dealerTableHeaders = [
      'Dealer Name',
      'Dealer #',
      'Tenure & Active Since',
      'Program Maturity',
      'Contracts Written',
      'Net Reserve ($)',
      'Claims Paid ($)',
      'Loss Ratio',
      '% Earned',
      'Actuarial Credibility & Status',
    ];
    const dHeaderRow = ws.getRow(currentRow);
    dHeaderRow.values = dealerTableHeaders;
    dHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    dHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    dHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    // Aggregate by dealer under this target
    const dealerMap = new Map<
      string,
      {
        name: string;
        number: string;
        contracts: number;
        reserve: number;
        earnedReserve: number;
        claims: number;
        firstDate: Date | null;
      }
    >();

    for (const t of targetContracts) {
      const key = t.dealerNumber || t.dealerName || 'UNKNOWN';
      const entry = dealerMap.get(key) || {
        name: t.dealerName || t.dealer || key,
        number: t.dealerNumber || key,
        contracts: 0,
        reserve: 0,
        earnedReserve: 0,
        claims: 0,
        firstDate: null,
      };
      if (t.transactionType !== 'Cancellation') {
        entry.contracts++;
      }
      entry.reserve += t.reserveAmount;
      entry.earnedReserve += t.earnedReserveAmount || 0;

      const d = t.effectiveDate || t.activityDate;
      if (d && !isNaN(d.getTime())) {
        if (!entry.firstDate || d.getTime() < entry.firstDate.getTime()) {
          entry.firstDate = d;
        }
      }
      dealerMap.set(key, entry);
    }

    for (const cl of targetClaims) {
      const key = cl.dealer || cl.dealerName || 'UNKNOWN';
      const entry = dealerMap.get(key) || {
        name: cl.dealerName || cl.dealer || key,
        number: cl.dealer || key,
        contracts: 0,
        reserve: 0,
        earnedReserve: 0,
        claims: 0,
        firstDate: null,
      };
      entry.claims += cl.paidAmount;
      dealerMap.set(key, entry);
    }

    const sortedDealers = Array.from(dealerMap.values()).sort((a, b) => b.claims - a.claims);

    for (const d of sortedDealers) {
      const lr = d.reserve > 0 ? d.claims / d.reserve : d.claims > 0 ? 9.99 : 0;

      let tenureMonths = 0;
      let tenureYears = 0;
      if (d.firstDate && !isNaN(d.firstDate.getTime())) {
        tenureMonths = Math.max(
          1,
          Math.round(
            (model.asOfDate.getTime() - d.firstDate.getTime()) / (1000 * 60 * 60 * 24 * 30.4375),
          ),
        );
        tenureYears = +(tenureMonths / 12).toFixed(1);
      }

      let maturity = 'New Dealer (<1 Yr)';
      if (tenureYears >= 2 && d.contracts >= 30) {
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

      const dateStr = d.firstDate
        ? d.firstDate.toLocaleDateString('en-US', { month: '2-digit', year: 'numeric' })
        : 'N/A';
      const tenureStr = `${tenureYears} Yrs (${dateStr})`;

      const pctEarned = d.reserve > 0 ? d.earnedReserve / d.reserve : 0;
      const row = ws.addRow([
        d.name,
        d.number,
        tenureStr,
        maturity,
        d.contracts,
        d.reserve,
        d.claims,
        lr,
        pctEarned,
        status,
      ]);
      row.font = { name: 'Aptos', size: 9 };
      row.getCell(5).numFmt = INTEGER;
      row.getCell(6).numFmt = MONEY;
      row.getCell(7).numFmt = MONEY;
      row.getCell(8).numFmt = PERCENT;
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

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 2: TIME-TO-LOSS VELOCITY (EARLY CLAIM EMERGENCE)
    // ─────────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const sec2Header = ws.getCell(`A${currentRow}`);
    sec2Header.value = '⏱️ 2. TIME-TO-LOSS VELOCITY (EARLY IN-SERVICE CLAIM EMERGENCE)';
    sec2Header.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.white } };
    sec2Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    currentRow++;

    const ttlHeaders = [
      'Time-to-Loss Cohort',
      'Days Since Activation',
      'Claim Count',
      'Total Paid ($)',
      'Avg Paid / Claim ($)',
      '% of Target Claims',
      'In-Service Classification',
      'Actuarial Indicator',
      '',
    ];
    const ttlHeaderRow = ws.getRow(currentRow);
    ttlHeaderRow.values = ttlHeaders;
    ttlHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    ttlHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF595959' } };
    ttlHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    const buckets = [
      {
        label: 'Early In-Service (0–90 Days)',
        days: '0–90 Days',
        min: 0,
        max: 90,
        risk: 'Immediate In-Service Emergence',
        diag: 'Claims occurring within 90 days of activation; reflects early defect frequency.',
      },
      {
        label: 'Early Intermediate (91–180 Days)',
        days: '91–180 Days',
        min: 91,
        max: 180,
        risk: 'First-Year Operational Wear',
        diag: 'Break-in mechanical issues during initial vehicle operating period.',
      },
      {
        label: 'Mid-Term Seasoning (181–365 Days)',
        days: '181–365 Days',
        min: 181,
        max: 365,
        risk: 'Expected Seasoned Degradation',
        diag: 'Standard mechanical wear curve aligned with empirical earning schedule.',
      },
      {
        label: 'Mature Portfolio (365+ Days)',
        days: '365+ Days',
        min: 366,
        max: Infinity,
        risk: 'Mature Seasoned Wear',
        diag: 'Natural wear-and-tear as contracts approach full reserve earning.',
      },
    ];

    const ttlStats = buckets.map((b) => ({ ...b, count: 0, paid: 0 }));

    for (const cl of targetClaims) {
      const effDate = contractEffectiveMap.get(cl.contractNumber) || cl.snapshotDate;
      const days = Math.max(
        0,
        Math.round((cl.activityDate.getTime() - effDate.getTime()) / (1000 * 60 * 60 * 24)),
      );

      for (const stat of ttlStats) {
        if (days >= stat.min && days <= stat.max) {
          stat.count++;
          stat.paid += cl.paidAmount;
          break;
        }
      }
    }

    const totalClaimsCount = targetClaims.length;

    for (const stat of ttlStats) {
      const avgPaid = stat.count > 0 ? stat.paid / stat.count : 0;
      const pct = totalClaimsCount > 0 ? stat.count / totalClaimsCount : 0;

      const r = ws.addRow([
        stat.label,
        stat.days,
        stat.count,
        stat.paid,
        avgPaid,
        pct,
        stat.risk,
        stat.diag,
        '',
      ]);
      r.font = { name: 'Aptos', size: 9 };
      r.getCell(3).numFmt = INTEGER;
      r.getCell(4).numFmt = MONEY;
      r.getCell(5).numFmt = MONEY;
      r.getCell(6).numFmt = PERCENT;
      if (stat.min === 0 && pct > 0.25) {
        r.getCell(7).font = { bold: true, color: { argb: COLORS.amber } };
      }
      currentRow++;
    }

    currentRow += 2;

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 3: VEHICLE MAKE EXPOSURE & CONCENTRATION
    // ─────────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const sec3Header = ws.getCell(`A${currentRow}`);
    sec3Header.value = '🚗 3. VEHICLE MAKE & MECHANICAL FAILURE EXPOSURE';
    sec3Header.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.white } };
    sec3Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    currentRow++;

    const makeHeaders = [
      'Vehicle Make',
      'Total Claims ($)',
      'Claim Count',
      'Avg Paid / Claim ($)',
      '% of Target Claims',
      'Primary Loss Code',
      'Primary Component Affected',
      'Exposure Distribution',
      '',
    ];
    const mHeaderRow = ws.getRow(currentRow);
    mHeaderRow.values = makeHeaders;
    mHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    mHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    mHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    const makeMap = new Map<
      string,
      { paid: number; count: number; lossCodes: Map<string, { desc: string; amount: number }> }
    >();

    for (const cl of targetClaims) {
      const make = cl.vehicleMake || 'UNMAPPED MAKE';
      const entry = makeMap.get(make) || { paid: 0, count: 0, lossCodes: new Map() };
      entry.paid += cl.paidAmount;
      entry.count++;
      const code = cl.lossCode || 'UNMAPPED';
      const lcEntry = entry.lossCodes.get(code) || {
        desc: cl.lossCodeDescription || '',
        amount: 0,
      };
      lcEntry.amount += cl.paidAmount;
      entry.lossCodes.set(code, lcEntry);
      makeMap.set(make, entry);
    }

    const totalTargetClaims = targetClaims.reduce((sum, c) => sum + c.paidAmount, 0);
    const sortedMakes = Array.from(makeMap.entries())
      .map(([make, data]) => ({ make, ...data }))
      .sort((a, b) => b.paid - a.paid)
      .slice(0, 8);

    for (const m of sortedMakes) {
      const avgPaid = m.count > 0 ? m.paid / m.count : 0;
      const pct = totalTargetClaims > 0 ? m.paid / totalTargetClaims : 0;

      const topLoss = Array.from(m.lossCodes.entries()).sort(
        (a, b) => b[1].amount - a[1].amount,
      )[0];
      const primaryCode = topLoss ? topLoss[0] : 'N/A';
      const primaryDesc = topLoss ? topLoss[1].desc : 'N/A';

      const r = ws.addRow([
        m.make,
        m.paid,
        m.count,
        avgPaid,
        pct,
        primaryCode,
        primaryDesc,
        pct > 0.25 ? 'High Concentration (>25%)' : 'Normal distribution',
        '',
      ]);
      r.font = { name: 'Aptos', size: 9 };
      r.getCell(2).numFmt = MONEY;
      r.getCell(3).numFmt = INTEGER;
      r.getCell(4).numFmt = MONEY;
      r.getCell(5).numFmt = PERCENT;
      currentRow++;
    }

    currentRow += 2;

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 4: RATE CARD & BENCHMARK PRICING ADEQUACY
    // ─────────────────────────────────────────────────────────────────────────
    ws.mergeCells(`A${currentRow}:J${currentRow}`);
    const sec4Header = ws.getCell(`A${currentRow}`);
    sec4Header.value = '⚖️ 4. RATE CARD PRICING ADEQUACY (TARGET vs. COMPANY-WIDE BENCHMARKS)';
    sec4Header.font = { name: 'Aptos', size: 11, bold: true, color: { argb: COLORS.white } };
    sec4Header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    currentRow++;

    const benchHeaders = [
      'Underwriting Metric',
      'Target Entity Metric',
      'Company Portfolio Benchmark',
      'Variance ($ / %)',
      'Pricing Comparison',
      'Factual Finding',
      '',
      '',
      '',
    ];
    const bHeaderRow = ws.getRow(currentRow);
    bHeaderRow.values = benchHeaders;
    bHeaderRow.font = { name: 'Aptos', size: 10, bold: true, color: { argb: COLORS.white } };
    bHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF595959' } };
    bHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getCell(`A${currentRow}`).alignment = { vertical: 'middle', horizontal: 'left' };
    currentRow++;

    // Compute portfolio-wide averages
    const allContractsCount = model.rolling12.current.contractsWritten || 1;
    const portfolioAvgReserve = model.rolling12.current.netReserve / allContractsCount;
    const portfolioAvgAdmin = model.rolling12.current.netAdmin / allContractsCount;
    const portfolioAvgClaim =
      model.claims.length > 0 ? model.rolling12.current.claimsPaid / model.claims.length : 0;

    const targetContractsCount =
      targetContracts.filter((t) => t.transactionType !== 'Cancellation').length || 1;
    const targetNetReserve = targetContracts.reduce((sum, t) => sum + t.reserveAmount, 0);
    const targetAvgReserve = targetNetReserve / targetContractsCount;

    const targetNetAdmin = targetContracts.reduce((sum, t) => sum + t.adminAmount, 0);
    const targetAvgAdmin = targetNetAdmin / targetContractsCount;

    const targetAvgClaim = targetClaims.length > 0 ? totalTargetClaims / targetClaims.length : 0;

    const benchRows = [
      {
        metric: 'Average Reserve Collected / Contract',
        targetVal: targetAvgReserve,
        benchVal: portfolioAvgReserve,
        fmt: MONEY,
        assessment:
          targetAvgReserve < portfolioAvgReserve * 0.9
            ? 'Below Portfolio Benchmark'
            : 'At or Above Portfolio Benchmark',
        finding: `Target collected $${Math.round(targetAvgReserve).toLocaleString()} vs. portfolio average of $${Math.round(portfolioAvgReserve).toLocaleString()}`,
      },
      {
        metric: 'Average Net Admin / Contract',
        targetVal: targetAvgAdmin,
        benchVal: portfolioAvgAdmin,
        fmt: MONEY,
        assessment:
          targetAvgAdmin < portfolioAvgAdmin * 0.9
            ? 'Below Portfolio Benchmark'
            : 'At or Above Portfolio Benchmark',
        finding: `Target collected $${Math.round(targetAvgAdmin).toLocaleString()} vs. portfolio average of $${Math.round(portfolioAvgAdmin).toLocaleString()}`,
      },
      {
        metric: 'Average Paid Severity / Claim',
        targetVal: targetAvgClaim,
        benchVal: portfolioAvgClaim,
        fmt: MONEY,
        assessment:
          targetAvgClaim > portfolioAvgClaim * 1.15
            ? 'Elevated Severity Above Benchmark'
            : 'Severity In-Line With Benchmark',
        finding: `Average target claim is $${Math.round(targetAvgClaim).toLocaleString()} vs. portfolio average of $${Math.round(portfolioAvgClaim).toLocaleString()}`,
      },
    ];

    for (const b of benchRows) {
      const diff = b.targetVal - b.benchVal;
      const r = ws.addRow([
        b.metric,
        b.targetVal,
        b.benchVal,
        diff,
        b.assessment,
        b.finding,
        '',
        '',
        '',
      ]);
      r.font = { name: 'Aptos', size: 9 };
      r.getCell(2).numFmt = b.fmt;
      r.getCell(3).numFmt = b.fmt;
      r.getCell(4).numFmt = b.fmt;
      currentRow++;
    }
  }
}
