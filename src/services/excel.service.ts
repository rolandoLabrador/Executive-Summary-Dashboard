import * as ExcelJS from 'exceljs';
import {
  type DimensionMetric,
  type LossCodeMetric,
  type MetricValues,
  type PeriodComparison,
  type ReportConfig,
  type ReportModel,
  type ReportingPeriod,
} from '../models/report.types';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

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
const MIN_LOSS_CODE_SHARE = 0.02;
const PIE_COLORS = ['2F75B5', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '8064A2'];

interface MetricDefinition {
  label: string;
  key: keyof MetricValues;
  format: string;
}

interface LossCodeChartSegment {
  code: string;
  description: string;
  coverageNames: string[];
  paidAmount: number;
  share: number;
  color: string;
}

function hslToHex(hue: number, saturation: number, lightness: number): string {
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const hueSegment = hue / 60;
  const secondary = chroma * (1 - Math.abs((hueSegment % 2) - 1));
  const [red, green, blue] =
    hueSegment < 1
      ? [chroma, secondary, 0]
      : hueSegment < 2
        ? [secondary, chroma, 0]
        : hueSegment < 3
          ? [0, chroma, secondary]
          : hueSegment < 4
            ? [0, secondary, chroma]
            : hueSegment < 5
              ? [secondary, 0, chroma]
              : [chroma, 0, secondary];
  const match = lightness - chroma / 2;
  return [red, green, blue]
    .map((channel) =>
      Math.round((channel + match) * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase();
}

function pieColor(index: number): string {
  if (index < PIE_COLORS.length) return PIE_COLORS[index]!;
  const goldenAngle = 137.508;
  return hslToHex((index * goldenAngle) % 360, 0.64, 0.48);
}

function lossCodeChartSegments(rows: LossCodeMetric[]): LossCodeChartSegment[] {
  const positiveRows = rows.filter((row) => row.rolling12Paid > 0);
  const total = positiveRows.reduce((sum, row) => sum + row.rolling12Paid, 0);
  if (total <= 0) return [];

  return positiveRows.map((row, index) => ({
    code: row.code,
    description: row.description,
    coverageNames: row.coverageNames,
    paidAmount: row.rolling12Paid,
    share: row.rolling12Paid / total,
    color: pieColor(index),
  }));
}

function visibleLossCodeRows(rows: LossCodeMetric[]): LossCodeMetric[] {
  return rows.filter((row) => (row.rolling12PaidShare ?? 0) >= MIN_LOSS_CODE_SHARE);
}

function formatDate(value: Date): string {
  return value.toLocaleDateString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
  });
}

function formatDateRange(start: Date, end: Date): string {
  return `${formatDate(start)}–${formatDate(end)}`;
}

function comparisonHeading(label: string, comparison: PeriodComparison): string {
  return (
    `${label} ${formatDateRange(comparison.currentStart, comparison.currentEnd)}` +
    ` vs ${formatDateRange(comparison.priorStart, comparison.priorEnd)}`
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
  { label: 'Contracts Written', key: 'contractsWritten', format: INTEGER },
  { label: 'Active Contracts', key: 'activeContracts', format: INTEGER },
  { label: 'Cancellations Processed', key: 'contractsCancelled', format: INTEGER },
  { label: 'Gross Income', key: 'grossIncome', format: MONEY },
  { label: 'Net Admin', key: 'netAdmin', format: MONEY },
  { label: 'Avg Admin / Contract', key: 'adminPerContract', format: MONEY },
  { label: 'Net Reserve', key: 'netReserve', format: MONEY },
  { label: 'Premium', key: 'premium', format: MONEY },
  { label: 'Earned Reserve', key: 'earnedReserve', format: MONEY },
  { label: 'Claims Paid', key: 'claimsPaid', format: MONEY },
  { label: 'Underwriting Profit', key: 'underwritingProfit', format: MONEY },
  { label: 'Paid Loss Ratio', key: 'paidLossRatio', format: PERCENT },
  { label: 'Earned Loss Ratio', key: 'earnedLossRatio', format: PERCENT },
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

function styleHeaderCell(cell: ExcelJS.Cell): void {
  cell.font = { bold: true, color: { argb: COLORS.white } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
  cell.border = { bottom: { style: 'thin', color: { argb: COLORS.navy } } };
}

function styleHeader(row: ExcelJS.Row): void {
  row.eachCell(styleHeaderCell);
  row.height = 24;
}

function styleHeaderRange(
  ws: ExcelJS.Worksheet,
  rowNumber: number,
  startColumn: number,
  endColumn: number,
): void {
  for (let column = startColumn; column <= endColumn; column += 1) {
    styleHeaderCell(ws.getCell(rowNumber, column));
  }
  ws.getRow(rowNumber).height = Math.max(ws.getRow(rowNumber).height ?? 0, 24);
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
  ws.mergeCells(startRow, 1, startRow, 5);
  ws.getCell(startRow, 1).value = heading;
  ws.getCell(startRow, 1).font = { bold: true, size: 14, color: { argb: COLORS.navy } };
  const header = ws.getRow(startRow + 1);
  header.values = ['Metric', 'Current', priorLabel, 'Change', 'Change %'];
  styleHeader(header);

  KPI_DEFINITIONS.forEach((definition, index) => {
    const row = ws.getRow(startRow + 2 + index);
    const current = value(comparison.current, definition.key);
    const prior = value(comparison.prior, definition.key);
    const changePercent = variance(current, prior);
    row.values = [definition.label, current, prior, current - prior, changePercent];
    [2, 3, 4].forEach((column) => {
      row.getCell(column).numFmt = definition.format;
    });
    const cell = row.getCell(5);
    cell.numFmt = PERCENT;
    if (changePercent !== null && changePercent !== 0) {
      const isBadIncrease = [
        'paidLossRatio',
        'earnedLossRatio',
        'cancellationRate',
        'claimsPaid',
      ].includes(definition.key);
      const isPositive = changePercent > 0;
      const isGreen = isBadIncrease ? !isPositive : isPositive;
      cell.font = {
        color: { argb: isGreen ? '009900' : 'FF0000' },
      };
    }
  });
  return startRow + KPI_DEFINITIONS.length + 3;
}

function writePeriodSnapshot(
  ws: ExcelJS.Worksheet,
  startRow: number,
  heading: string,
  period: ReportingPeriod,
): number {
  ws.mergeCells(startRow, 1, startRow, 5);
  const titleCell = ws.getCell(startRow, 1);
  titleCell.value = `${heading} ${formatDateRange(period.start, period.end)}`;
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

function buildThreeTierDimensionSheet(
  workbook: ExcelJS.Workbook,
  name: string,
  heading: string,
  itdRows: DimensionMetric[],
  rollingRows: DimensionMetric[],
  monthlyRows: DimensionMetric[],
  config: ReportConfig,
  itdStart: Date,
  rollingStart: Date,
  currentStart: Date,
  currentEnd: Date,
  limit?: number,
): void {
  const ws = workbook.addWorksheet(name);
  configureWorksheet(ws);
  const isDealer = name === 'Dealer Dashboard';
  const isAgent = name === 'Agent Dashboard';
  const maxCols = isDealer || isAgent ? 20 : 18;
  title(
    ws,
    heading,
    `Inception to Date, Rolling 12-Month & Latest Month Results | Through ${formatDate(currentEnd)}`,
    maxCols,
  );

  const writeTable = (
    startRow: number,
    sectionHeading: string,
    rows: DimensionMetric[],
  ): number => {
    ws.mergeCells(startRow, 1, startRow, maxCols);
    const headingCell = ws.getCell(startRow, 1);
    headingCell.value = sectionHeading;
    headingCell.font = { bold: true, size: 12, color: { argb: COLORS.navy } };

    const selected = limit ? rows.slice(0, limit) : rows;
    const headers = [
      'Rank',
      isDealer ? 'Dealer Number' : name.replace(' Dashboard', ''),
      ...(isDealer ? ['Dealer Name', 'Agents'] : isAgent ? ['Agent Name', 'Dealers'] : []),
      'Written',
      'Active Contracts',
      'Gross Income',
      'Cancellations Processed',
      'Net Admin',
      'Avg Admin / Contract',
      'Net Reserve',
      'Premium',
      'Earned Reserve',
      'Claims Paid',
      'Underwriting Profit',
      'Claim Count',
      'Paid Loss Ratio',
      'Earned Loss Ratio',
      'Cancellation Rate',
      'Loss Ratio Bar',
    ];
    ws.getRow(startRow + 1).values = headers;
    styleHeaderRange(ws, startRow + 1, 1, maxCols);

    selected.forEach((item, index) => {
      const row = ws.getRow(startRow + 2 + index);
      row.values = [
        index + 1,
        item.name,
        ...(isDealer
          ? [item.displayName || 'Name unavailable', item.relatedAgents?.join(', ') || 'Unassigned']
          : isAgent
            ? [item.name, item.relatedDealers?.join(', ') || 'Unassigned']
            : []),
        item.contractsWritten,
        item.activeContracts,
        item.grossIncome,
        item.contractsCancelled,
        item.netAdmin,
        item.adminPerContract,
        item.netReserve,
        item.premium,
        item.earnedReserve,
        item.claimsPaid,
        item.underwritingProfit,
        item.claimCount,
        item.paidLossRatio,
        item.earnedLossRatio,
        item.cancellationRate,
        item.paidLossRatio,
      ];
      const offset = isDealer || isAgent ? 2 : 0;
      [3, 4, 6, 14].forEach((column) => {
        row.getCell(column + offset).numFmt = INTEGER;
      });
      [5, 7, 8, 9, 10, 11, 12, 13].forEach((column) => {
        row.getCell(column + offset).numFmt = MONEY;
      });
      [15, 16, 17, 18].forEach((column) => {
        row.getCell(column + offset).numFmt = PERCENT;
      });
    });

    if (selected.length > 0) {
      const firstDataRow = startRow + 2;
      const lastDataRow = startRow + 1 + selected.length;
      const barCol = isDealer || isAgent ? 'T' : 'R';
      const paidRatioCol = isDealer || isAgent ? 'Q' : 'O';
      const earnedRatioCol = isDealer || isAgent ? 'R' : 'P';

      ws.addConditionalFormatting({
        ref: `${barCol}${firstDataRow}:${barCol}${lastDataRow}`,
        rules: [dataBarRule(1)],
      });

      const ratioRules: ExcelJS.ConditionalFormattingRule[] = [
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
      ];

      ws.addConditionalFormatting({
        ref: `${paidRatioCol}${firstDataRow}:${paidRatioCol}${lastDataRow}`,
        rules: ratioRules,
      });
      ws.addConditionalFormatting({
        ref: `${earnedRatioCol}${firstDataRow}:${earnedRatioCol}${lastDataRow}`,
        rules: ratioRules,
      });
    }

    return startRow + selected.length + 4;
  };

  let nextRow = writeTable(
    4,
    `INCEPTION TO DATE RESULTS (${formatDateRange(itdStart, currentEnd)})`,
    itdRows,
  );
  nextRow = writeTable(
    nextRow,
    `ROLLING 12-MONTH RESULTS (${formatDateRange(rollingStart, currentEnd)})`,
    rollingRows,
  );
  writeTable(
    nextRow,
    `YEAR TO DATE RESULTS (${formatDateRange(currentStart, currentEnd)})`,
    monthlyRows,
  );

  ws.columns = [
    { width: 8 },
    { width: isDealer ? 16 : 30 },
    ...(isDealer ? [{ width: 34 }, { width: 36 }] : []),
    { width: 12 }, // Written
    { width: 12 }, // Active Contracts
    { width: 16 }, // Gross Income
    { width: 14 }, // Cancellations Processed
    { width: 16 }, // Net Admin
    { width: 16 }, // Net Reserve
    { width: 16 }, // Premium
    { width: 16 }, // Earned Reserve
    { width: 16 }, // Claims Paid
    { width: 16 }, // Underwriting Profit
    { width: 12 }, // Claim Count
    { width: 16 }, // Paid Loss Ratio
    { width: 16 }, // Earned Loss Ratio
    { width: 17 }, // Cancellation Rate
    { width: 28 }, // Loss Ratio Bar
  ];
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
    buildThreeTierDimensionSheet(
      workbook,
      'Dealer Dashboard',
      'DEALER PERFORMANCE DASHBOARD',
      model.itdDealers,
      model.dealers,
      model.ytdDealers,
      this.config,
      new Date(2010, 0, 1),
      model.rolling12.currentStart,
      model.yearToDate.currentStart,
      model.yearToDate.currentEnd,
      this.config.topDealerCount,
    );
    buildThreeTierDimensionSheet(
      workbook,
      'Agent Dashboard',
      'AGENT PERFORMANCE DASHBOARD',
      model.itdAgents,
      model.agents,
      model.ytdAgents,
      this.config,
      new Date(2010, 0, 1),
      model.rolling12.currentStart,
      model.yearToDate.currentStart,
      model.yearToDate.currentEnd,
    );
    buildThreeTierDimensionSheet(
      workbook,
      'Product Dashboard',
      'PRODUCT PERFORMANCE DASHBOARD',
      model.itdProducts,
      model.products,
      model.ytdProducts,
      this.config,
      new Date(2010, 0, 1),
      model.rolling12.currentStart,
      model.yearToDate.currentStart,
      model.yearToDate.currentEnd,
    );
    this.buildLossCodeDashboard(workbook, model);
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
      `Current reporting month: ${formatDateRange(model.currentMonth.currentStart, model.currentMonth.currentEnd)} | Paid claims / net written reserve`,
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
    this.writeWorstDealers(ws, model);
    ws.columns = [
      { width: 24 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 14 },
      { width: 4 },
      { width: 4 },
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
    const months = model.monthly.slice(-12).reverse();
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
          argb: index === 0 ? COLORS.green : index === 1 ? COLORS.amber : COLORS.blue,
        },
      };
      header.alignment = { horizontal: 'center' };
    });

    const trends: Array<{ row: number; label: string; key: keyof MetricValues; format: string }> = [
      { row: 7, label: 'Net Reserve', key: 'netReserve', format: MONEY },
      { row: 9, label: 'Earned Reserve', key: 'earnedReserve', format: MONEY },
      { row: 11, label: 'Claims Paid', key: 'claimsPaid', format: MONEY },
      { row: 13, label: 'Paid Loss Ratio', key: 'paidLossRatio', format: PERCENT },
      { row: 15, label: 'Earned Loss Ratio', key: 'earnedLossRatio', format: PERCENT },
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

    ws.mergeCells(17, firstColumn, 17, Math.max(firstColumn, lastColumn));
    const note = ws.getCell(17, firstColumn);
    note.value =
      'Green = latest completed month  |  Amber = preceding month  |  Bars show relative monthly magnitude';
    note.font = { italic: true, color: { argb: COLORS.darkGray } };
    note.alignment = { horizontal: 'center' };
  }

  private writeWorstDealers(ws: ExcelJS.Worksheet, model: ReportModel): void {
    const startRow = 20;
    const firstColumn = 8;
    const dealers = [...model.dealers].filter((d) => d.earnedReserve > 0 && d.claimsPaid > 0);
    dealers.sort((a, b) => (b.earnedLossRatio || 0) - (a.earnedLossRatio || 0));
    const worst = dealers.slice(0, 20);

    ws.mergeCells(startRow, firstColumn, startRow, firstColumn + 11);
    const heading = ws.getCell(startRow, firstColumn);
    heading.value = 'TOP 20 WORST PERFORMING DEALERS BY EARNED LOSS RATIO (ROLLING 12)';
    heading.font = { bold: true, size: 12, color: { argb: COLORS.navy } };

    const headerRow = ws.getRow(startRow + 1);
    
    headerRow.getCell(8).value = 'Rank';
    
    ws.mergeCells(startRow + 1, 9, startRow + 1, 13);
    headerRow.getCell(9).value = 'Dealer';
    
    ws.mergeCells(startRow + 1, 14, startRow + 1, 15);
    headerRow.getCell(14).value = 'Active Contracts';
    
    ws.mergeCells(startRow + 1, 16, startRow + 1, 17);
    headerRow.getCell(16).value = 'Premium';
    
    ws.mergeCells(startRow + 1, 18, startRow + 1, 19);
    headerRow.getCell(18).value = 'Earned Loss Ratio';

    [8, 9, 14, 16, 18].forEach(col => {
       const cell = headerRow.getCell(col);
       cell.font = { bold: true, color: { argb: COLORS.white } };
       cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
       cell.alignment = { horizontal: 'center' };
    });

    worst.forEach((dealer, index) => {
      const row = ws.getRow(startRow + 2 + index);
      
      row.getCell(8).value = index + 1;
      row.getCell(8).alignment = { horizontal: 'center' };
      
      ws.mergeCells(startRow + 2 + index, 9, startRow + 2 + index, 13);
      row.getCell(9).value = dealer.displayName || dealer.name;
      
      ws.mergeCells(startRow + 2 + index, 14, startRow + 2 + index, 15);
      row.getCell(14).value = dealer.activeContracts;
      row.getCell(14).numFmt = INTEGER;
      row.getCell(14).alignment = { horizontal: 'center' };
      
      ws.mergeCells(startRow + 2 + index, 16, startRow + 2 + index, 17);
      row.getCell(16).value = dealer.premium;
      row.getCell(16).numFmt = MONEY;
      row.getCell(16).alignment = { horizontal: 'center' };
      
      ws.mergeCells(startRow + 2 + index, 18, startRow + 2 + index, 19);
      row.getCell(18).value = dealer.earnedLossRatio;
      row.getCell(18).numFmt = PERCENT;
      row.getCell(18).alignment = { horizontal: 'center' };
    });

    if (worst.length > 0) {
      ws.addConditionalFormatting({
        ref: `R${startRow + 2}:S${startRow + 1 + worst.length}`,
        rules: [dataBarRule(20)],
      });
    }
  }

  private buildLossCodeDashboard(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Loss Code Dashboard');
    configureWorksheet(ws);
    title(
      ws,
      'LOSS CODE PERFORMANCE DASHBOARD',
      `Paid amounts by covered vehicle component | Through ${formatDate(model.currentMonth.currentEnd)}`,
      15,
    );

    ws.getRow(4).values = [
      'KPI',
      `Current Month\n${formatDateRange(model.currentMonth.currentStart, model.currentMonth.currentEnd)}`,
      `Year to Date\n${formatDateRange(model.yearToDate.currentStart, model.yearToDate.currentEnd)}`,
      `Rolling 12 Months\n${formatDateRange(model.rolling12.currentStart, model.rolling12.currentEnd)}`,
    ];
    styleHeader(ws.getRow(4));
    ws.getRow(4).eachCell((cell) => {
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    ws.getRow(4).height = 38;
    const kpis = [
      { label: 'Total Paid Amount', key: 'totalPaid' as const, format: MONEY },
      { label: 'Unique Paid Claims', key: 'claimCount' as const, format: INTEGER },
      { label: 'Loss Code Groups Paid', key: 'lossCodeCount' as const, format: INTEGER },
      {
        label: 'Average Paid per Claim',
        key: 'averagePaidPerClaim' as const,
        format: MONEY,
      },
    ];
    kpis.forEach(({ label, key, format }, index) => {
      const row = ws.getRow(index + 5);
      row.values = [
        label,
        model.lossCodeDashboard.currentMonth[key],
        model.lossCodeDashboard.yearToDate[key],
        model.lossCodeDashboard.rolling12[key],
      ];
      row.getCell(1).font = { bold: true, color: { argb: COLORS.navy } };
      [2, 3, 4].forEach((column) => {
        row.getCell(column).numFmt = format;
      });
    });

    ws.mergeCells('K4:O4');
    const vehicleMakeHeading = ws.getCell('K4');
    vehicleMakeHeading.value = `TOP 10 VEHICLE MAKES — ${formatDateRange(model.rolling12.currentStart, model.rolling12.currentEnd)}`;
    vehicleMakeHeading.font = { bold: true, size: 14, color: { argb: COLORS.navy } };
    vehicleMakeHeading.alignment = { horizontal: 'center', vertical: 'middle' };
    ['Rank', 'Vehicle Make', 'Paid Amount', 'Claim Count', '% of R12 Paid'].forEach(
      (label, index) => {
        ws.getCell(5, index + 11).value = label;
      },
    );
    styleHeaderRange(ws, 5, 11, 15);
    model.lossCodeDashboard.topVehicleMakes.forEach((item, index) => {
      const row = ws.getRow(index + 6);
      row.getCell(11).value = index + 1;
      row.getCell(12).value = item.make;
      row.getCell(13).value = item.paidAmount;
      row.getCell(13).numFmt = MONEY;
      row.getCell(14).value = item.claimCount;
      row.getCell(14).numFmt = INTEGER;
      row.getCell(15).value = item.paidShare;
      row.getCell(15).numFmt = PERCENT;
    });
    if (model.lossCodeDashboard.topVehicleMakes.length > 0) {
      ws.addConditionalFormatting({
        ref: `O6:O${model.lossCodeDashboard.topVehicleMakes.length + 5}`,
        rules: [dataBarRule(2)],
      });
    }
    if (model.lossCodeDashboard.topVehicleMakes.length === 0) {
      ws.mergeCells('K6:O6');
      const noVehicleMakes = ws.getCell('K6');
      noVehicleMakes.value = 'No paid vehicle-make data is available for this period.';
      noVehicleMakes.font = { italic: true, color: { argb: COLORS.darkGray } };
      noVehicleMakes.alignment = { horizontal: 'center' };
    }

    const topVehicleMakesEndRow = Math.max(
      15,
      model.lossCodeDashboard.topVehicleMakes.length > 0
        ? model.lossCodeDashboard.topVehicleMakes.length + 5
        : 6,
    );
    const chartHeadingRow = Math.max(18, topVehicleMakesEndRow + 3);
    const legendHeaderRow = chartHeadingRow + 1;
    const legendStartRow = legendHeaderRow + 1;
    const minimumChartEndRow = chartHeadingRow + 15;
    ws.mergeCells(chartHeadingRow, 1, chartHeadingRow, 10);
    const chartHeading = ws.getCell(chartHeadingRow, 1);
    chartHeading.value = `ROLLING 12-MONTH PAID AMOUNT MIX (≥2%) — ${formatDateRange(model.rolling12.currentStart, model.rolling12.currentEnd)}`;
    chartHeading.font = { bold: true, size: 14, color: { argb: COLORS.navy } };
    const displayedRows = visibleLossCodeRows(model.lossCodeDashboard.rows);
    const chartSegments = lossCodeChartSegments(displayedRows);
    let chartSectionEndRow = minimumChartEndRow;
    if (chartSegments.length > 0) {
      const pieChartPng = renderPieChartPng(
        chartSegments.map((segment) => ({
          value: segment.paidAmount,
          color: segment.color,
        })),
      );
      const imageId = workbook.addImage({
        base64: `data:image/png;base64,${pieChartPng.toString('base64')}`,
        extension: 'png',
      });
      ws.addImage(imageId, {
        tl: { col: 0.75, row: chartHeadingRow + 0.15 },
        ext: { width: 280, height: 280 },
      });
      [
        'Color',
        'Loss Code',
        'Component Description',
        'Product (Coverage Name)',
        'Paid Amount',
        '% of Chart',
      ].forEach((label, index) => {
        ws.getCell(legendHeaderRow, index + 4).value = label;
      });
      styleHeaderRange(ws, legendHeaderRow, 4, 9);
      chartSegments.forEach((segment, index) => {
        const row = ws.getRow(index + legendStartRow);
        const colorCell = row.getCell(4);
        colorCell.value = '●';
        colorCell.font = { bold: true, size: 18, color: { argb: segment.color } };
        colorCell.alignment = { horizontal: 'center' };
        row.getCell(5).value = segment.code;
        row.getCell(6).value = segment.description;
        row.getCell(6).alignment = { wrapText: true, vertical: 'middle' };
        row.getCell(7).value = segment.coverageNames.join(', ');
        row.getCell(7).alignment = { wrapText: true, vertical: 'middle' };
        row.getCell(8).value = segment.paidAmount;
        row.getCell(8).numFmt = MONEY;
        row.getCell(9).value = segment.share;
        row.getCell(9).numFmt = PERCENT;
        row.height = 26;
      });
      const lastLegendRow = chartSegments.length + legendHeaderRow;
      const chartNoteRow = Math.max(chartHeadingRow + 10, lastLegendRow + 2);
      ws.mergeCells(chartNoteRow, 4, chartNoteRow + 2, 10);
      const chartNote = ws.getCell(chartNoteRow, 4);
      chartNote.value =
        'Every loss code contributing at least 2% of rolling-12 paid amount is shown as its own slice. I will group similar components but in the future we should take a look at the loss codes.';
      chartNote.font = { italic: true, color: { argb: COLORS.darkGray } };
      chartNote.alignment = { wrapText: true, vertical: 'top' };
      chartSectionEndRow = Math.max(chartSectionEndRow, chartNoteRow + 2);
    } else {
      ws.mergeCells(legendHeaderRow, 1, minimumChartEndRow, 10);
      const emptyChart = ws.getCell(legendHeaderRow, 1);
      emptyChart.value = 'No loss code meets the 2% display threshold for this period.';
      emptyChart.font = { italic: true, color: { argb: COLORS.darkGray } };
      emptyChart.alignment = { horizontal: 'center', vertical: 'middle' };
      emptyChart.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.gray } };
    }

    const detailHeadingRow = chartSectionEndRow + 4;
    const detailHeaderRow = detailHeadingRow + 1;
    const detailStartRow = detailHeaderRow + 1;
    ws.mergeCells(detailHeadingRow, 1, detailHeadingRow, 9);
    const detailHeading = ws.getCell(detailHeadingRow, 1);
    detailHeading.value = `ROLLING 12-MONTH DETAIL (≥2%) — ${formatDateRange(model.rolling12.currentStart, model.rolling12.currentEnd)}`;
    detailHeading.font = { bold: true, size: 14, color: { argb: COLORS.navy } };
    ws.getRow(detailHeaderRow).values = [
      'Rank',
      'Loss Code',
      'Component Description',
      'Current Month Paid',
      'YTD Paid',
      'Rolling 12 Paid',
      '% of Rolling 12 Paid',
      'Rolling 12 Claim Count',
      'Average Paid per Claim',
    ];
    styleHeader(ws.getRow(detailHeaderRow));
    ws.getRow(detailHeaderRow).eachCell((cell) => {
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    ws.getRow(detailHeaderRow).height = 38;

    displayedRows.forEach((item, index) => {
      const row = ws.getRow(index + detailStartRow);
      row.values = [
        index + 1,
        item.code,
        item.description,
        item.currentMonthPaid,
        item.yearToDatePaid,
        item.rolling12Paid,
        item.rolling12PaidShare,
        item.rolling12ClaimCount,
        item.rolling12AveragePaidPerClaim,
      ];
      [4, 5, 6, 9].forEach((column) => {
        row.getCell(column).numFmt = MONEY;
      });
      row.getCell(7).numFmt = PERCENT;
      row.getCell(8).numFmt = INTEGER;
      if (index % 2 === 1) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F5F9FC' } };
        });
      }
    });

    if (displayedRows.length > 0) {
      ws.addConditionalFormatting({
        ref: `G${detailStartRow}:G${displayedRows.length + detailHeaderRow}`,
        rules: [dataBarRule(1)],
      });
    }
    const finalDetailRow = Math.max(detailStartRow, displayedRows.length + detailHeaderRow);
    ws.autoFilter = { from: `A${detailHeaderRow}`, to: `I${finalDetailRow}` };
    const noteRow = finalDetailRow + 2;
    ws.mergeCells(noteRow, 1, noteRow, 9);
    const note = ws.getCell(noteRow, 1);
    note.value =
      'Loss codes below 2% of rolling-12 paid amount are omitted from the pie and detail but remain in KPI totals. Claim counts are distinct by Claim Number; a multi-code claim is counted once per applicable row and once overall.';
    note.font = { italic: true, color: { argb: COLORS.darkGray } };
    note.alignment = { wrapText: true };
    ws.getRow(noteRow).height = 30;
    ws.columns = [
      { width: 8 },
      { width: 16 },
      { width: 38 },
      { width: 20 },
      { width: 16 },
      { width: 18 },
      { width: 22 },
      { width: 23 },
      { width: 22 },
      { width: 3 },
      { width: 8 },
      { width: 20 },
      { width: 16 },
      { width: 14 },
      { width: 18 },
    ];
  }

  private buildMonthly(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Monthly Trends');
    configureWorksheet(ws);
    const trendStart = model.monthly[0]?.periodStart ?? model.currentMonth.currentStart;
    title(
      ws,
      'MONTHLY TRENDS',
      `Monthly trend coverage: ${formatDateRange(trendStart, model.currentMonth.currentEnd)}`,
    );
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
    const uniqueComponents = new Set<string>();
    model.contractTransactions.forEach((item) => {
      if (item.components) {
        Object.entries(item.components).forEach(([cat, comps]) => {
          if (cat.toUpperCase().includes('COMMISSION') || cat.toUpperCase().includes('COMM')) return;
          Object.keys(comps).forEach((c) => uniqueComponents.add(`${cat} - ${c}`));
        });
      }
    });
    const componentColumns = Array.from(uniqueComponents).sort();

    const baseColumns = [
      { width: 26 },
      { width: 14, numFmt: 'mm/dd/yyyy' },
      { width: 14, numFmt: 'mm/dd/yyyy' },
      { width: 18 },
      { width: 14 },
      { width: 24 },
      { width: 14 },
      { width: 30 },
      { width: 25 },
      { width: 16 },
      { width: 12 },
      { width: 18 },
      { width: 14, numFmt: MONEY },
      { width: 14, numFmt: MONEY },
      { width: 14, numFmt: MONEY }, // Premium
      { width: 16, numFmt: MONEY },
    ];
    const dynColumns = componentColumns.map(() => ({ width: 18, numFmt: MONEY }));
    ws.columns = [...baseColumns, ...dynColumns];

    ws.getRow(4).values = [
      'Source ID',
      'Snapshot Date',
      'Activity Date',
      'Contract Number',
      'Transaction',
      'Agent',
      'Dealer Number',
      'Dealer Name',
      'Product',
      'Coverage Code',
      'Company',
      'Risk Entity',
      'Admin',
      'Reserve',
      'Premium',
      'Earned Reserve',
      ...componentColumns,
    ];
    styleHeader(ws.getRow(4));
    // SAMPLE SIZE REDUCTION
    // Writes 1 out of every 33 contracts (~3% chronological sample spread across all years).
    // This dramatically solves the Excel generation bottleneck while providing a huge representative dataset.
    let sampledCount = 0;
    model.contractTransactions.forEach((item, index) => {
      if (index % 33 !== 0) return;
      sampledCount++;

      const rowValues: ExcelJS.CellValue[] = [
        item.sourceId,
        item.snapshotDate,
        item.activityDate,
        item.contractNumber,
        item.transactionType,
        item.agent,
        item.dealerNumber,
        item.dealerName || item.dealer,
        item.product,
        item.coverageCode,
        item.company,
        item.riskEntity,
        item.adminAmount,
        item.reserveAmount,
        item.adminAmount + item.reserveAmount, // Premium
        item.earnedReserveAmount,
      ];
      componentColumns.forEach((col) => {
        const parts = col.split(' - ');
        const cat = parts[0];
        const comp = parts.slice(1).join(' - ');
        if (cat && comp) {
          rowValues.push(item.components?.[cat]?.[comp] ?? 0);
        } else {
          rowValues.push(0);
        }
      });
      ws.addRow(rowValues);
    });

    const endColLetter = ws.getColumn(15 + componentColumns.length).letter;
    ws.autoFilter = {
      from: 'A4',
      to: `${endColLetter}${Math.max(5, sampledCount + 4)}`,
    };
  }

  private buildClaimDetail(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Claim Activity');
    configureWorksheet(ws);
    title(
      ws,
      'SANITIZED PAID CLAIM ACTIVITY',
      'Paid claims only; no customer or vehicle identifiers',
    );
    ws.columns = [
      { width: 26 },
      { width: 14, numFmt: 'mm/dd/yyyy' },
      { width: 14, numFmt: 'mm/dd/yyyy' },
      { width: 18 },
      { width: 18 },
      { width: 12 },
      { width: 14, numFmt: MONEY },
      { width: 24 },
      { width: 30 },
      { width: 25 },
      { width: 12 },
      { width: 40 },
      { width: 26 },
    ];
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
      'Loss Code',
      'Component Description',
      'Payment Key',
    ];
    styleHeader(ws.getRow(4));
    model.claims.forEach((item) => {
      ws.addRow([
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
        item.lossCode,
        item.lossCodeDescription,
        item.paymentKey,
      ]);
    });
    ws.autoFilter = { from: 'A4', to: `M${Math.max(5, model.claims.length + 4)}` };
  }

  private buildDataQuality(workbook: ExcelJS.Workbook, model: ReportModel): void {
    const ws = workbook.addWorksheet('Data Quality');
    configureWorksheet(ws);
    title(
      ws,
      'DATA QUALITY AND RECONCILIATION',
      `Generated ${model.generatedAt.toLocaleString('en-US')}`,
      8,
    );

    let currentRow = 4;

    ws.mergeCells(currentRow, 1, currentRow, 8);
    const auditHeading = ws.getCell(currentRow, 1);
    auditHeading.value = 'PIPELINE INGESTION AUDIT & RECONCILIATION';
    auditHeading.font = { bold: true, size: 12, color: { argb: COLORS.navy } };
    currentRow++;

    const auditHeaderRow = currentRow;
    ws.getRow(auditHeaderRow).values = [
      'Status',
      'Job Type',
      'Portal Count',
      'Unique Units',
      'Uploaded Line Items',
      'Variance',
      'Source File',
      'Execution Timestamp (UTC)',
    ];
    styleHeaderRange(ws, auditHeaderRow, 1, 8);
    currentRow++;

    if (model.pipelineAudits && model.pipelineAudits.length > 0) {
      model.pipelineAudits.forEach((audit) => {
        const portalCount = audit.counts.portalCount;
        const uniqueUnits = audit.counts.uniqueCount ?? audit.counts.uploadedCount;
        const uploadedLines = audit.counts.uploadedCount;
        const unitVariance = uniqueUnits - portalCount;

        const isReconciled =
          (audit.reconciliation.isMatch && audit.reconciliation.status === 'PASSED') ||
          (Math.abs(unitVariance) <= Math.max(2, Math.round(portalCount * 0.001)) &&
            audit.counts.processedCount === audit.counts.uploadedCount);

        const status = isReconciled ? 'PASSED' : audit.reconciliation.status;

        const row = ws.getRow(currentRow);
        const statusCell = row.getCell(1);
        statusCell.value = status;
        statusCell.font = { bold: true };
        if (isReconciled) {
          statusCell.font = { bold: true, color: { argb: '006100' } };
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } };
        } else {
          statusCell.font = { bold: true, color: { argb: '9C0006' } };
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
        }
        statusCell.alignment = { horizontal: 'center' };

        row.getCell(2).value = audit.jobType;
        row.getCell(3).value = portalCount;
        row.getCell(3).numFmt = INTEGER;
        row.getCell(4).value = uniqueUnits;
        row.getCell(4).numFmt = INTEGER;
        row.getCell(5).value = uploadedLines;
        row.getCell(5).numFmt = INTEGER;
        row.getCell(6).value = unitVariance;
        row.getCell(6).numFmt = INTEGER;
        row.getCell(7).value = audit.fileMetadata.fileName;
        row.getCell(8).value = audit.executionTimestamp
          ? audit.executionTimestamp.toISOString().replace('T', ' ').replace(/\..+/, '')
          : audit.executionDateStr;
        currentRow++;
      });
    } else {
      ws.mergeCells(currentRow, 1, currentRow, 8);
      const noAuditCell = ws.getCell(currentRow, 1);
      noAuditCell.value = 'No pipeline audit records found in AuditDB.DataReconciliationAudit.';
      noAuditCell.font = { italic: true, color: { argb: COLORS.darkGray } };
      currentRow++;
    }

    currentRow += 2; // Blank row spacing

    const countsHeaderRow = currentRow;
    ws.getRow(countsHeaderRow).values = ['Source Document Category', 'Document Count'];
    styleHeaderRange(ws, countsHeaderRow, 1, 2);
    currentRow++;

    const counts: Array<[string, number]> = [
      ['Contract source documents', model.sourceCounts.contractDocuments],
      ['Cancellation source documents', model.sourceCounts.cancellationDocuments],
      ['Claim source documents', model.sourceCounts.claimDocuments],
      ['Unique normalized contract transactions', model.sourceCounts.uniqueContractTransactions],
      ['Unique paid claim transactions', model.sourceCounts.uniqueClaims],
      ['Data-quality issues', model.dataQualityIssues.length],
    ];
    counts.forEach(([label, count]) => {
      const row = ws.getRow(currentRow);
      row.getCell(1).value = label;
      row.getCell(2).value = count;
      row.getCell(2).numFmt = INTEGER;
      currentRow++;
    });

    currentRow += 2; // Blank row spacing

    const issuesHeaderRow = currentRow;
    ws.getRow(issuesHeaderRow).values = [
      'Severity',
      'Category',
      'Contract Number',
      'Dealer Name',
      'Source ID',
      'Message',
    ];
    styleHeaderRange(ws, issuesHeaderRow, 1, 6);
    currentRow++;

    const issuesStartRow = currentRow;
    model.dataQualityIssues.forEach((issue) => {
      ws.getRow(currentRow).values = [
        issue.severity,
        issue.category,
        issue.contractNumber,
        issue.dealerName,
        issue.sourceId,
        issue.message,
      ];
      currentRow++;
    });

    const finalIssueRow = Math.max(issuesStartRow, currentRow - 1);
    ws.autoFilter = { from: `A${issuesHeaderRow}`, to: `F${finalIssueRow}` };
    ws.columns = [
      { width: 18 },
      { width: 28 },
      { width: 20 },
      { width: 32 },
      { width: 28 },
      { width: 60 },
      { width: 34 },
      { width: 26 },
    ];
  }

  private buildDefinitions(workbook: ExcelJS.Workbook): void {
    const ws = workbook.addWorksheet('Definitions');
    configureWorksheet(ws);
    title(ws, 'REPORT DEFINITIONS', 'Controlled business rules used by this workbook');
    ws.getRow(4).values = ['Metric / Rule', 'Definition'];
    styleHeader(ws.getRow(4));
    const definitions = [
      ['Gross Income', 'Add Net admin Net reserve tax ceeding full amount'],
      ['Premium', 'Calculated as Net Admin + Net Written Reserve.'],
      ['Underwriting Profit', 'Calculated as Premium - Claims Paid.'],
      ['Paid Loss Ratio', 'Claims paid divided by Premium.'],
      ['Earned Loss Ratio', 'Claims paid divided by (Earned Reserve + Net Admin).'],
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
        'Inception to Date (ITD)',
        'All recognized activity from the very first recorded date (inception) up to the latest completed month. ITD Loss Ratio is the total claims paid since inception divided by the ITD Premium.',
      ],
      [
        'Dealer Ranking',
        `Top ${this.config.topDealerCount} dealers ranked by rolling-12 net written reserve.`,
      ],
      [
        'Excluded Components',
        'Broadly excludes commission section plus components containing DEALER, DLR, COMMISSION, COMM, F&I, or PACK. Additionally, when calculating Reserve, specifically excludes CLIPFEE, PREMIUMTAX, CEDINGFEE, and ADMIN. When calculating Admin, specifically excludes ROADSIDEADMIN and LOANPMT.',
      ],
      ['Claims', 'Paid claim/payment records with a non-zero Total Paid Amount.'],
      [
        'Loss Code Dashboard',
        'Paid claim amounts grouped by normalized Loss Code. The KPI section shows current month, year-to-date, and rolling-12 results; detail rows are ranked by rolling-12 paid amount.',
      ],
      [
        'Loss Code Display Threshold',
        'Loss codes contributing less than 2% of rolling-12 paid amount are omitted from the pie chart and rolling-12 detail. Their paid amounts and claims remain included in dashboard KPI totals.',
      ],
      [
        'Top Vehicle Makes',
        'The ten vehicle Make values with the highest deduplicated rolling-12 paid claim amount. Claim Count is distinct by Claim Number within each make; blank Make values remain visible as UNMAPPED MAKE.',
      ],
      [
        'Loss Code Product',
        'The pie-chart legend uses the claim Coverage Name as the product. Multiple applicable coverage names are listed without selecting one arbitrarily; blanks remain visible as Unmapped Coverage Name.',
      ],
      [
        'Loss Code Pie Chart',
        'Positive rolling-12 paid amounts. Every paid loss code appears as a separate slice. The color-matched legend shows its exact loss code, component description, product Coverage Name, paid amount, and share. Negative adjustments remain in the detail table.',
      ],
      [
        'Loss Code Claim Count',
        'Distinct Claim Number within each loss code. A claim with multiple loss codes appears once in each applicable row but only once in the overall KPI.',
      ],
      [
        'Unmapped Loss Code',
        'Paid claim records with a blank Loss Code are retained under UNMAPPED so dashboard totals reconcile to paid-claim totals.',
      ],
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
        'Contracts Written',
        'Gross original contracts grouped by metadata.ActivationDate. When the new-business source is missing, a cancellation may supply an auditable written reference that retains canceled status and is never active.',
      ],
      [
        'Contract Count Reconciliation',
        'For the same ActivationDate cohort and as-of cutoff: Written Contracts = Active Contracts + contracts from that cohort whose latest state is canceled. Cancellations Processed is grouped by CancelBillDate and is not generally the cancellation term in this equation.',
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
