import * as ExcelJS from 'exceljs';
import {
  type DimensionMetric,
  type MetricValues,
  type PeriodComparison,
  type ReportConfig,
  type ReportModel,
  type ReportingPeriod,
} from '../models/report.types';

const COLORS = {
  navy: '17365D',
  blue: '2F75B5',
  lightBlue: 'D9EAF7',
  green: '70AD47',
  amber: 'FFC000',
  red: 'C00000',
  white: 'FFFFFF',
  gray: 'E7E6E6',
  darkGray: '595959',
};

const MONEY = '$#,##0;[Red]-$#,##0';
const INTEGER = '#,##0;[Red]-#,##0';
const PERCENT = '0.0%;[Red]-0.0%';

interface MetricDefinition {
  label: string;
  key: keyof MetricValues;
  format: string;
}
function formatDate(value: Date): string {
  return value.toLocaleDateString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
  });
}

function comparisonHeading(label: string, comparison: PeriodComparison): string {
  return (
    `${label} ${formatDate(comparison.currentStart)}–${formatDate(comparison.currentEnd)}` +
    ` vs ${formatDate(comparison.priorStart)}–${formatDate(comparison.priorEnd)}`
  );
}

function dataBarRule(priority: number): ExcelJS.ConditionalFormattingRule {
  // ExcelJS writes the required color element, but v4.4 omits it from DataBarRuleType.
  const rule: ExcelJS.DataBarRuleType & { color: Partial<ExcelJS.Color> } = {
    type: 'dataBar',
    priority,
    cfvo: [
      { type: 'num', value: 0 },
      { type: 'num', value: 1 },
    ],
    color: { argb: COLORS.blue },
    gradient: true,
  };
  return rule;
}

const KPI_DEFINITIONS: MetricDefinition[] = [
  { label: 'Active Contracts', key: 'activeContracts', format: INTEGER },
  { label: 'Net Admin', key: 'netAdmin', format: MONEY },
  { label: 'Net Reserve', key: 'netReserve', format: MONEY },
  { label: 'Claims Paid', key: 'claimsPaid', format: MONEY },
  { label: 'Paid Loss Ratio', key: 'paidLossRatio', format: PERCENT },
  { label: 'Cancellation Rate', key: 'cancellationRate', format: PERCENT },
];

function title(ws: ExcelJS.Worksheet, value: string, subtitle?: string, endColumn = 12): void {
  ws.mergeCells(1, 1, 1, endColumn);
  const cell = ws.getCell('A1');
  cell.value = value;
  cell.font = { name: 'Aptos Display', size: 20, bold: true, color: { argb: COLORS.white } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
  cell.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(1).height = 34;
  if (subtitle) {
    ws.mergeCells(2, 1, 2, endColumn);
    ws.getCell('A2').value = subtitle;
    ws.getCell('A2').font = { italic: true, color: { argb: COLORS.darkGray } };
  }
}

function styleHeader(row: ExcelJS.Row): void {
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = { bottom: { style: 'thin', color: { argb: COLORS.navy } } };
  });
  row.height = 24;
}

function configureWorksheet(ws: ExcelJS.Worksheet): void {
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 3 }];
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  ws.headerFooter.oddFooter = '&LConfidential&CPage &P of &N&RGenerated &D &T';
}

function value(metrics: MetricValues, key: keyof MetricValues): number {
  return Number(metrics[key] ?? 0);
}

function variance(current: number, prior: number): number | null {
  return prior === 0 ? null : (current - prior) / Math.abs(prior);
}

function writeComparison(
  ws: ExcelJS.Worksheet,
  startRow: number,
  heading: string,
  comparison: PeriodComparison,
  priorLabel = 'Prior Year',
): number {
  ws.mergeCells(startRow, 1, startRow, 6);
  ws.getCell(startRow, 1).value = heading;
  ws.getCell(startRow, 1).font = { bold: true, size: 14, color: { argb: COLORS.navy } };
  const header = ws.getRow(startRow + 1);
  header.values = ['Metric', 'Current', priorLabel, 'Change', 'Change %', 'Status'];
  styleHeader(header);

  KPI_DEFINITIONS.forEach((definition, index) => {
    const row = ws.getRow(startRow + 2 + index);
    const current = value(comparison.current, definition.key);
    const prior = value(comparison.prior, definition.key);
    const changePercent = variance(current, prior);
    row.values = [
      definition.label,
      current,
      prior,
      current - prior,
      changePercent,
      definition.key === 'paidLossRatio' && current >= 0.8
        ? 'High'
        : definition.key === 'paidLossRatio' && current >= 0.65
          ? 'Watch'
          : 'OK',
    ];
    [2, 3, 4].forEach((column) => {
      row.getCell(column).numFmt = definition.format;
    });
    row.getCell(5).numFmt = PERCENT;
  });
  return startRow + KPI_DEFINITIONS.length + 3;
}

function writePeriodSnapshot(
  ws: ExcelJS.Worksheet,
  startRow: number,
  heading: string,
  period: ReportingPeriod,
): number {
  ws.mergeCells(startRow, 1, startRow, 6);
  const titleCell = ws.getCell(startRow, 1);
  titleCell.value = `${heading} ${formatDate(period.start)}–${formatDate(period.end)}`;
  titleCell.font = { bold: true, size: 14, color: { argb: COLORS.navy } };

  const header = ws.getRow(startRow + 1);
  header.values = ['Metric', 'Full-Year Result'];
  styleHeader(header);
  KPI_DEFINITIONS.forEach((definition, index) => {
    const row = ws.getRow(startRow + 2 + index);
    row.values = [definition.label, value(period.values, definition.key)];
    row.getCell(2).numFmt = definition.format;
  });
  return startRow + KPI_DEFINITIONS.length + 3;
}

function writeDimensionSheet(
  workbook: ExcelJS.Workbook,
  name: string,
  heading: string,
  rows: DimensionMetric[],
  config: ReportConfig,
  reportingEnd: Date,
  limit?: number,
  reportingMonth?: Date,
): void {
  const ws = workbook.addWorksheet(name);
  configureWorksheet(ws);
  title(
    ws,
    heading,
    reportingMonth
      ? `Monthly results for ${reportingMonth.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })} through ${formatDate(reportingEnd)}`
      : `Rolling 12 months through ${formatDate(reportingEnd)}`,
  );
  const selected = limit ? rows.slice(0, limit) : rows;
  const isDealer = name === 'Dealer Dashboard';
  const isMonthlyProduct = name === 'Product Dashboard' && reportingMonth !== undefined;
  const headers = [
    'Rank',
    isDealer ? 'Dealer Number' : name.replace(' Dashboard', ''),
    ...(isDealer ? ['Dealer Name', 'Agents'] : []),
    ...(isMonthlyProduct ? ['Reporting Month'] : []),
    'Written',
    'Active Contracts',
    'Cancellations Processed',
    'Net Admin',
    'Net Reserve',
    'Claims Paid',
    'Claim Count',
    'Paid Loss Ratio',
    'Cancellation Rate',
    'Loss Ratio Bar',
  ];
  ws.getRow(4).values = headers;
  styleHeader(ws.getRow(4));

  selected.forEach((item, index) => {
    const row = ws.getRow(index + 5);
    row.values = [
      index + 1,
      item.name,
      ...(isDealer
        ? [item.displayName || 'Name unavailable', item.relatedAgents?.join(', ') || 'Unassigned']
        : []),
      ...(isMonthlyProduct ? [reportingMonth] : []),
      item.contractsWritten,
      item.activeContracts,
      item.contractsCancelled,
      item.netAdmin,
      item.netReserve,
      item.claimsPaid,
      item.claimCount,
      item.paidLossRatio,
      item.cancellationRate,
      item.paidLossRatio,
    ];
    if (isMonthlyProduct) row.getCell(3).numFmt = 'mmmm yyyy';
    const offset = (isDealer ? 2 : 0) + (isMonthlyProduct ? 1 : 0);
    [3, 4, 5, 9].forEach((column) => {
      row.getCell(column + offset).numFmt = INTEGER;
    });
    [6, 7, 8].forEach((column) => {
      row.getCell(column + offset).numFmt = MONEY;
    });
    [10, 11, 12].forEach((column) => {
      row.getCell(column + offset).numFmt = PERCENT;
    });
  });

  if (selected.length > 0) {
    const lastRow = selected.length + 4;
    ws.addConditionalFormatting({
      ref: `${isDealer ? 'N' : isMonthlyProduct ? 'M' : 'L'}5:${isDealer ? 'N' : isMonthlyProduct ? 'M' : 'L'}${lastRow}`,
      rules: [dataBarRule(1)],
    });
    ws.addConditionalFormatting({
      ref: `${isDealer ? 'L' : isMonthlyProduct ? 'K' : 'J'}5:${isDealer ? 'L' : isMonthlyProduct ? 'K' : 'J'}${lastRow}`,
      rules: [
        {
          type: 'cellIs',
          priority: 2,
          operator: 'greaterThan',
          formulae: [config.highLossRatio],
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F4CCCC' } } },
        },
        {
          type: 'cellIs',
          priority: 3,
          operator: 'between',
          formulae: [config.warningLossRatio, config.highLossRatio],
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2CC' } } },
        },
      ],
    });
  }
  ws.columns = [
    { width: 8 },
    { width: isDealer ? 16 : 30 },
    ...(isDealer ? [{ width: 34 }, { width: 36 }] : []),
    ...(isMonthlyProduct ? [{ width: 18 }] : []),
    { width: 12 },
    { width: 12 },
    { width: 14 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 12 },
    { width: 16 },
    { width: 17 },
    { width: 28 },
  ];
  ws.autoFilter = {
    from: 'A4',
    to: `${isDealer ? 'N' : isMonthlyProduct ? 'M' : 'L'}${Math.max(5, selected.length + 4)}`,
  };
}

export class ExcelService {
  constructor(private readonly config: ReportConfig) {}

  async generateFullReport(model: ReportModel): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'OmniShield Reporting Engine';
    workbook.company = this.config.companyName;
    workbook.created = model.generatedAt;
    workbook.modified = model.generatedAt;
    workbook.calcProperties.fullCalcOnLoad = true;

    this.buildExecutive(workbook, model);
    writeDimensionSheet(
      workbook,
      'Dealer Dashboard',
      'DEALER PERFORMANCE DASHBOARD',
      model.dealers,
      this.config,
      model.currentMonth.currentEnd,
      this.config.topDealerCount,
    );
    writeDimensionSheet(
      workbook,
      'Agent Dashboard',
      'AGENT PERFORMANCE DASHBOARD',
      model.agents,
      this.config,
      model.currentMonth.currentEnd,
    );
    writeDimensionSheet(
      workbook,
      'Product Dashboard',
      'PRODUCT PERFORMANCE DASHBOARD',
      model.products,
      this.config,
      model.currentMonth.currentEnd,
      undefined,
      model.currentMonth.currentStart,
    );
    this.buildMonthly(workbook, model);
    this.buildContractDetail(workbook, model);
    this.buildClaimDetail(workbook, model);
    this.buildDataQuality(workbook, model);
    this.buildDefinitions(workbook);

    return Buffer.from(await workbook.xlsx.writeBuffer());
  }

  private buildExecutive(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Executive Dashboard');
    configureWorksheet(ws);
    title(
      ws,
      `${this.config.companyName.toUpperCase()} EXECUTIVE SUMMARY`,
      `Reporting through ${formatDate(model.currentMonth.currentEnd)} | Paid claims / net written reserve`,
      19,
    );
    let row = writeComparison(
      ws,
      4,
      comparisonHeading('Current Month vs Same Month Prior Year', model.currentMonth),
      model.currentMonth,
      'Prior Year',
    );
    row = writeComparison(
      ws,
      row,
      comparisonHeading('Year to Date vs Prior-Year YTD', model.yearToDate),
      model.yearToDate,
    );
    row = writeComparison(
      ws,
      row,
      comparisonHeading('Rolling 12 Months vs Preceding 12 Months', model.rolling12),
      model.rolling12,
    );
    writePeriodSnapshot(
      ws,
      row,
      `FULL YEAR ${model.priorCalendarYear.start.getFullYear()}`,
      model.priorCalendarYear,
    );
    this.writeExecutiveTrend(ws, model);
    ws.columns = [
      { width: 24 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 14 },
      { width: 13 },
      { width: 3 },
      { width: 3 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
      { width: 12 },
    ];
  }

  private writeExecutiveTrend(ws: ExcelJS.Worksheet, model: ReportModel): void {
    const months = model.monthly.slice(-12);
    const firstColumn = 8;
    const lastColumn = firstColumn + months.length - 1;
    ws.mergeCells(4, firstColumn, 4, Math.max(firstColumn, lastColumn));
    const heading = ws.getCell(4, firstColumn);
    heading.value = '12-MONTH TREND — THROUGH LATEST COMPLETED MONTH';
    heading.font = { bold: true, size: 14, color: { argb: COLORS.navy } };

    months.forEach((month, index) => {
      const column = firstColumn + index;
      const header = ws.getCell(5, column);
      header.value = month.periodStart;
      header.numFmt = 'mmm-yy';
      header.font = { bold: true, color: { argb: COLORS.white } };
      header.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb:
            index === months.length - 1
              ? COLORS.green
              : index === months.length - 2
                ? COLORS.amber
                : COLORS.blue,
        },
      };
      header.alignment = { horizontal: 'center' };
    });

    const trends: Array<{ row: number; label: string; key: keyof MetricValues; format: string }> = [
      { row: 7, label: 'Net Reserve', key: 'netReserve', format: MONEY },
      { row: 9, label: 'Claims Paid', key: 'claimsPaid', format: MONEY },
      { row: 11, label: 'Paid Loss Ratio', key: 'paidLossRatio', format: PERCENT },
    ];
    trends.forEach(({ row, label, key, format }, trendIndex) => {
      ws.mergeCells(row - 1, firstColumn, row - 1, Math.max(firstColumn, lastColumn));
      const labelCell = ws.getCell(row - 1, firstColumn);
      labelCell.value = label;
      labelCell.font = { bold: true, color: { argb: COLORS.darkGray } };
      months.forEach((month, index) => {
        const cell = ws.getCell(row, firstColumn + index);
        cell.value = value(month, key);
        cell.numFmt = format;
        cell.alignment = { horizontal: 'center' };
      });
      if (months.length > 0) {
        ws.addConditionalFormatting({
          ref: `${ws.getCell(row, firstColumn).address}:${ws.getCell(row, lastColumn).address}`,
          rules: [dataBarRule(10 + trendIndex)],
        });
      }
    });

    ws.mergeCells(13, firstColumn, 13, Math.max(firstColumn, lastColumn));
    const note = ws.getCell(13, firstColumn);
    note.value =
      'Green = latest completed month  |  Amber = preceding month  |  Bars show relative monthly magnitude';
    note.font = { italic: true, color: { argb: COLORS.darkGray } };
    note.alignment = { horizontal: 'center' };
  }

  private buildMonthly(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Monthly Trends');
    configureWorksheet(ws);
    title(ws, 'MONTHLY TRENDS', 'Latest 24 reporting months');
    ws.getRow(4).values = [
      'Month',
      'Written',
      'Active Contracts',
      'Cancellations Processed',
      'Net Admin',
      'Net Reserve',
      'Claims Paid',
      'Claim Count',
      'Paid Loss Ratio',
      'Cancellation Rate',
      'Loss Ratio Bar',
    ];
    styleHeader(ws.getRow(4));
    model.monthly.forEach((item, index) => {
      const row = ws.getRow(index + 5);
      row.values = [
        item.periodStart,
        item.contractsWritten,
        item.activeContracts,
        item.contractsCancelled,
        item.netAdmin,
        item.netReserve,
        item.claimsPaid,
        item.claimCount,
        item.paidLossRatio,
        item.cancellationRate,
        item.paidLossRatio,
      ];
      row.getCell(1).numFmt = 'mmm-yy';
      [2, 3, 4, 8].forEach((column) => {
        row.getCell(column).numFmt = INTEGER;
      });
      [5, 6, 7].forEach((column) => {
        row.getCell(column).numFmt = MONEY;
      });
      [9, 10, 11].forEach((column) => {
        row.getCell(column).numFmt = PERCENT;
      });
    });
    const lastRow = model.monthly.length + 4;
    ws.addConditionalFormatting({
      ref: `K5:K${lastRow}`,
      rules: [dataBarRule(1)],
    });
    ws.columns = [
      { width: 13 },
      { width: 12 },
      { width: 12 },
      { width: 14 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 12 },
      { width: 16 },
      { width: 17 },
      { width: 30 },
    ];
  }

  private buildContractDetail(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Contract Activity');
    configureWorksheet(ws);
    title(
      ws,
      'SANITIZED CONTRACT ACTIVITY',
      'No customer names, contact details, addresses, or VINs',
    );
    ws.getRow(4).values = [
      'Source ID',
      'Snapshot Date',
      'Activity Date',
      'Contract Number',
      'Transaction',
      'Agent',
      'Dealer Number',
      'Dealer',
      'Product',
      'Coverage Code',
      'Company',
      'Risk Entity',
      'Admin',
      'Reserve',
    ];
    styleHeader(ws.getRow(4));
    model.contractTransactions.forEach((item, index) => {
      const row = ws.getRow(index + 5);
      row.values = [
        item.sourceId,
        item.snapshotDate,
        item.activityDate,
        item.contractNumber,
        item.transactionType,
        item.agent,
        item.dealerNumber,
        item.dealer,
        item.product,
        item.coverageCode,
        item.company,
        item.riskEntity,
        item.adminAmount,
        item.reserveAmount,
      ];
      row.getCell(2).numFmt = 'mm/dd/yyyy';
      row.getCell(3).numFmt = 'mm/dd/yyyy';
      row.getCell(13).numFmt = MONEY;
      row.getCell(14).numFmt = MONEY;
    });
    ws.autoFilter = { from: 'A4', to: `N${Math.max(5, model.contractTransactions.length + 4)}` };
    ws.columns = [
      { width: 26 },
      { width: 14 },
      { width: 14 },
      { width: 18 },
      { width: 14 },
      { width: 24 },
      { width: 14 },
      { width: 30 },
      { width: 25 },
      { width: 16 },
      { width: 12 },
      { width: 18 },
      { width: 14 },
      { width: 14 },
    ];
  }

  private buildClaimDetail(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Claim Activity');
    configureWorksheet(ws);
    title(
      ws,
      'SANITIZED PAID CLAIM ACTIVITY',
      'Paid claims only; no customer or vehicle identifiers',
    );
    ws.getRow(4).values = [
      'Source ID',
      'Snapshot Date',
      'Paid Date',
      'Claim Number',
      'Contract Number',
      'Status',
      'Paid Amount',
      'Agent',
      'Dealer',
      'Product',
      'Payment Key',
    ];
    styleHeader(ws.getRow(4));
    model.claims.forEach((item, index) => {
      const row = ws.getRow(index + 5);
      row.values = [
        item.sourceId,
        item.snapshotDate,
        item.activityDate,
        item.claimNumber,
        item.contractNumber,
        item.status,
        item.paidAmount,
        item.agent,
        item.dealer,
        item.product,
        item.paymentKey,
      ];
      row.getCell(2).numFmt = 'mm/dd/yyyy';
      row.getCell(3).numFmt = 'mm/dd/yyyy';
      row.getCell(7).numFmt = MONEY;
    });
    ws.autoFilter = { from: 'A4', to: `K${Math.max(5, model.claims.length + 4)}` };
    ws.columns = [
      { width: 26 },
      { width: 14 },
      { width: 14 },
      { width: 18 },
      { width: 18 },
      { width: 14 },
      { width: 16 },
      { width: 24 },
      { width: 30 },
      { width: 28 },
      { width: 65 },
    ];
  }

  private buildDataQuality(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Data Quality');
    configureWorksheet(ws);
    title(
      ws,
      'DATA QUALITY AND RECONCILIATION',
      `Generated ${model.generatedAt.toLocaleString('en-US')}`,
    );
    ws.getRow(4).values = ['Source', 'Documents'];
    styleHeader(ws.getRow(4));
    const counts: Array<[string, number]> = [
      ['Contract source documents', model.sourceCounts.contractDocuments],
      ['Cancellation source documents', model.sourceCounts.cancellationDocuments],
      ['Claim source documents', model.sourceCounts.claimDocuments],
      ['Unique normalized contract transactions', model.sourceCounts.uniqueContractTransactions],
      ['Unique paid claim transactions', model.sourceCounts.uniqueClaims],
      ['Data-quality issues', model.dataQualityIssues.length],
    ];
    counts.forEach(([label, count], index) => {
      ws.getRow(index + 5).values = [label, count];
    });

    ws.getRow(13).values = ['Severity', 'Category', 'Source ID', 'Message'];
    styleHeader(ws.getRow(13));
    model.dataQualityIssues.forEach((issue, index) => {
      ws.getRow(index + 14).values = [
        issue.severity,
        issue.category,
        issue.sourceId,
        issue.message,
      ];
    });
    ws.autoFilter = { from: 'A13', to: `D${Math.max(14, model.dataQualityIssues.length + 13)}` };
    ws.columns = [{ width: 20 }, { width: 28 }, { width: 28 }, { width: 80 }];
  }

  private buildDefinitions(workbook: ExcelJS.Workbook): void {
    const ws = workbook.addWorksheet('Definitions');
    configureWorksheet(ws);
    title(ws, 'REPORT DEFINITIONS', 'Controlled business rules used by this workbook');
    ws.getRow(4).values = ['Metric / Rule', 'Definition'];
    styleHeader(ws.getRow(4));
    const definitions = [
      ['Paid Loss Ratio', 'Claims paid divided by net written reserve.'],
      [
        'Active Contracts',
        'Distinct contracts whose latest snapshot has ContractStatus A and whose metadata.ActivationDate falls within the reporting period.',
      ],
      [
        'Cancellations Processed',
        'Distinct cancellation contracts whose metadata.CancelBillDate falls within the reporting period. This is activity, not a subtraction from the active cohort.',
      ],
      [
        'Net Written Reserve',
        'Included written reserve components less included cancelled reserve components.',
      ],
      ['Net Admin', 'Included written admin components less included cancelled admin components.'],
      [
        'Cancellation Timing',
        'Cancellation activity is recognized exclusively from metadata.CancelBillDate. Blank or invalid Cancel Bill Date records are excluded and reported as data-quality errors.',
      ],
      [
        'Contract Timing',
        'Written contract activity is recognized exclusively from metadata.ActivationDate. Blank or invalid Activation Date records are excluded and reported as data-quality errors.',
      ],
      [
        'Current Month',
        'Latest fully completed month, compared with the same calendar month in the prior year.',
      ],
      [
        'Year to Date',
        'January 1 through the latest completed month, compared with the same prior-year months.',
      ],
      [
        'Rolling 12 Months',
        'Latest completed month plus the preceding 11 months, compared with the preceding 12-month period.',
      ],
      [
        'Prior Full Calendar Year',
        'January 1 through December 31 of the calendar year immediately before the report as-of year.',
      ],
      [
        'Dealer Ranking',
        `Top ${this.config.topDealerCount} dealers ranked by rolling-12 net written reserve.`,
      ],
      [
        'Excluded Components',
        'The entire commission section plus component names containing DEALER, DLR, COMMISSION, COMM, F&I, or PACK, and configured exact exclusions.',
      ],
      ['Claims', 'Paid claim/payment records with a non-zero Total Paid Amount.'],
      [
        'Snapshot Deduplication',
        'Contract and cancellation snapshots retain the newest record per Contract# and transaction type. MongoDB _id is used only for traceability.',
      ],
      [
        'Claim Deduplication',
        'Claim count uses distinct Claim Number. Paid amounts retain the newest snapshot per payment/detail signature: claim, paid date, check, method, payee, loss code, RO, and amount.',
      ],
      [
        'Claim Attribution',
        'Dealer number, agent number, and coverage code come directly from the claim; contract data is used only when a claim dimension is missing. Names are fallback values.',
      ],
      [
        'Cancellation Written Reference',
        'A cancellation record supplies its original WrittenAmount only when the matching new-business transaction is absent; it is dated using metadata.ActivationDate.',
      ],
      [
        'Privacy',
        'Customer identity, contact, address, and VIN fields are excluded at MongoDB extraction.',
      ],
    ];
    definitions.forEach((definition, index) => {
      ws.getRow(index + 5).values = definition;
    });
    ws.columns = [{ width: 28 }, { width: 110 }];
    ws.getColumn(2).alignment = { wrapText: true, vertical: 'top' };
  }
}
