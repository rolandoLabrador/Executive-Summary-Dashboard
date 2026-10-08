import type * as ExcelJS from 'exceljs';
import { type MetricValues, type PeriodComparison, type ReportingPeriod } from '../models/report.types';

export const COLORS = {
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

export const MONEY = '$#,##0;[Red]-$#,##0';
export const INTEGER = '#,##0;[Red]-#,##0';
export const PERCENT = '0.0%;[Red]-0.0%';

export interface MetricDefinition {
  label: string;
  key: keyof MetricValues;
  format: string;
}

export const REPORT_DEFINITIONS: Record<string, string> = {
  'Gross Income': 'Sum of Net Admin, Net Reserve, and all other non-excluded components (tax, ceding, etc.) from transactions that occurred specifically within the selected reporting period.',
  'Underwriting Profit': 'Calculated as Earned Reserve - Claims Paid (for the selected reporting period).',
  'Earned Loss Ratio': 'Claims Paid divided by Earned Reserve (for the selected reporting period).',
  'Earned Reserve': 'Calculated using written/cancelled amounts from transactions within the reporting period. For contracts ultimately active as of the run date, calculated dynamically using the earning schedule curve. For contracts ultimately cancelled as of the run date, calculated as MAX(0, Written Reserve - Cancelled Reserve).',
  'Active Contracts': 'Distinct contracts whose metadata.ActivationDate falls within the reporting period AND whose absolute latest historical snapshot has ContractStatus A (and is not a cancellation). Expired contracts (Status E) are explicitly excluded from this count.',
  'Expired Contracts': 'Contracts with ContractStatus E are treated mathematically as New Business (written contracts) that have naturally reached the end of their term. They are included in Contracts Written and Earned Reserve calculations for their respective periods, but are never counted as Active.',
  'Cancellations': 'Distinct cancellation contracts whose metadata.CancelBillDate falls within the reporting period. This is a measure of cancellation activity during the period, not a subtraction from the active cohort.',
  'Net Written Reserve': 'Sum of written reserve components minus cancelled reserve components from transactions that occurred specifically within the selected reporting period. Calculated for all contracts regardless of their current active/cancelled status.',
  'Net Admin': 'Sum of written admin components minus cancelled admin components from transactions that occurred specifically within the selected reporting period. Importantly, this sum ONLY includes amounts for contracts that are currently Active (Status A and not cancelled) as of the report run date.',
  'Cancellation Timing': 'Cancellation activity is recognized exclusively from metadata.CancelBillDate. Blank or invalid Cancel Bill Date records are excluded and reported as data-quality errors.',
  'Contract Timing': 'Written contract activity is recognized exclusively from metadata.ActivationDate. Blank or invalid Activation Date records are excluded and reported as data-quality errors.',
  'Current Month': 'Latest fully completed month, compared with the same calendar month in the prior year.',
  'Year to Date': 'January 1 through the latest completed month, compared with the same prior-year months.',
  'Rolling 12 Months': 'Latest completed month plus the preceding 11 months, compared with the preceding 12-month period.',
  'Prior Full Calendar Year': 'January 1 through December 31 of the calendar year immediately before the report as-of year.',
  'Inception to Date (ITD)': 'All recognized activity from the very first recorded date (inception) up to the latest completed month. ITD Loss Ratio is the total claims paid since inception divided by the ITD Earned Reserve.',
  'Excluded Components': 'Broadly excludes commission section plus components containing DEALER, DLR, COMMISSION, COMM, F&I, or PACK. Additionally, when calculating Reserve, specifically excludes CLIPFEE, PREMIUMTAX, CEDINGFEE, ADMIN, SLUSH, and OEM TRANSPORT RESRVE. When calculating Admin, specifically excludes ROADSIDEADMIN, LOANPMT, OTHERCOMM, PREMTAX, AGENTNCB, and ROADSIDE AKMC.',
  'Claims': 'Paid claim/payment records with a non-zero Total Paid Amount.',
  'Loss Code Dashboard': 'Paid claim amounts grouped by normalized Loss Code. The KPI section shows current month, year-to-date, and rolling-12 results; detail rows are ranked by rolling-12 paid amount.',
  'Loss Code Display Threshold': 'Loss codes contributing less than 2% of rolling-12 paid amount are omitted from the pie chart and rolling-12 detail. Their paid amounts and claims remain included in dashboard KPI totals.',
  'Top Vehicle Makes': 'The ten vehicle Make values with the highest deduplicated rolling-12 paid claim amount. Claim Count is distinct by Claim Number within each make; blank Make values remain visible as UNMAPPED MAKE.',
  'Loss Code Product': 'The pie-chart legend uses the claim Coverage Name as the product. Multiple applicable coverage names are listed without selecting one arbitrarily; blanks remain visible as Unmapped Coverage Name.',
  'Loss Code Pie Chart': 'Positive rolling-12 paid amounts. Every paid loss code appears as a separate slice. The color-matched legend shows its exact loss code, component description, product Coverage Name, paid amount, and share. Negative adjustments remain in the detail table.',
  'Loss Code Claim Count': 'Distinct Claim Number within each loss code. A claim with multiple loss codes appears once in each applicable row but only once in the overall KPI.',
  'Unmapped Loss Code': 'Paid claim records with a blank Loss Code are retained under UNMAPPED so dashboard totals reconcile to paid-claim totals.',
  'Snapshot Deduplication': 'Contract and cancellation snapshots retain the newest record per Contract# and transaction type. MongoDB _id is used only for traceability.',
  'Claim Deduplication': 'Claim count uses distinct Claim Number. Paid amounts retain the newest snapshot per payment/detail signature: claim, paid date, check, method, payee, loss code, RO, and amount.',
  'Claim Attribution': 'Dealer number, agent number, and coverage code come directly from the claim; contract data is used only when a claim dimension is missing. Names are fallback values.',
  'Contracts Written': 'Gross original contracts grouped by metadata.ActivationDate. When the new-business source is missing, a cancellation may supply an auditable written reference that retains canceled status and is never active.',
  'Contract Count Reconciliation': 'For the same ActivationDate cohort and as-of cutoff: Written Contracts = Active Contracts + contracts from that cohort whose latest state is canceled. Cancellations is grouped by CancelBillDate and is not generally the cancellation term in this equation.',
  'Privacy': 'Customer identity, contact, address, and VIN fields are excluded at MongoDB extraction.',
  'Average admin': 'Calculated as Net Admin (from the reporting period) divided by Active Contracts (from the reporting period). If a cohort has 0 Active Contracts in the reporting period (e.g. only cancelled/expired contracts or claims), this value is left intentionally blank to avoid mathematical errors.',
  'Rule: Low Admin Flag': 'If the calculated Average Admin per contract is positive but less than $20.00, the cell is highlighted in bold red font and flagged with a tooltip/note.',
  'Rule: Loss Ratio Data Bars': 'Loss ratio cells contain embedded data bars that visually indicate performance: Green (< 85%), Amber (85% - 100%), and Red (> 100%).',
  'Rule: High Loss Ratio Highlights': 'Row-level earned loss ratios exceeding the warning threshold are highlighted yellow; those exceeding the high threshold are highlighted red.',
  'Rule: Agent Display Limit': 'The Agent Dashboard displays ALL agents without any row limit.',
  'Rule: Dealer Display Limit': 'The Dealer Dashboard is limited to displaying only the top 20 dealers (configurable via TOP_DEALER_COUNT).'
};

export const KPI_DEFINITIONS: MetricDefinition[] = [
  { label: 'Contracts Written', key: 'contractsWritten', format: INTEGER },
  { label: 'Active Contracts', key: 'activeContracts', format: INTEGER },
  { label: 'cancellation', key: 'contractsCancelled', format: INTEGER },
  { label: 'Gross Income', key: 'grossIncome', format: MONEY },
  { label: 'Net Admin', key: 'netAdmin', format: MONEY },
  { label: 'Average admin', key: 'adminPerContract', format: MONEY },
  { label: 'Net Reserve', key: 'netReserve', format: MONEY },
  { label: 'Earned Reserve', key: 'earnedReserve', format: MONEY },
  { label: 'Claims Paid', key: 'claimsPaid', format: MONEY },
  { label: 'Underwriting Profit', key: 'underwritingProfit', format: MONEY },
  { label: 'Earned Loss Ratio', key: 'earnedLossRatio', format: PERCENT },
  { label: 'Cancellation Rate', key: 'cancellationRate', format: PERCENT },
];

export function formatDate(value: Date): string {
  return value.toLocaleDateString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
  });
}

export function formatDateRange(start: Date, end: Date): string {
  return `${formatDate(start)}–${formatDate(end)}`;
}

export function comparisonHeading(label: string, comparison: PeriodComparison): string {
  return (
    `${label} ${formatDateRange(comparison.currentStart, comparison.currentEnd)}` +
    ` vs ${formatDateRange(comparison.priorStart, comparison.priorEnd)}`
  );
}

export function dataBarRule(priority: number, colorHex: string = COLORS.blue): ExcelJS.ConditionalFormattingRule {
  const rule: ExcelJS.DataBarRuleType & { color: Partial<ExcelJS.Color> } = {
    type: 'dataBar',
    priority,
    cfvo: [
      { type: 'num', value: 0 },
      { type: 'num', value: 1 },
    ],
    color: { argb: colorHex },
    gradient: true,
  };
  return rule;
}

export function title(ws: ExcelJS.Worksheet, value: string, subtitle?: string, endColumn = 12): void {
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

export function styleHeaderCell(cell: ExcelJS.Cell): void {
  cell.font = { bold: true, color: { argb: COLORS.white } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
  cell.border = { bottom: { style: 'thin', color: { argb: COLORS.navy } } };
  
  if (cell.value && typeof cell.value === 'string') {
    const definition = REPORT_DEFINITIONS[cell.value];
    if (definition) {
      cell.note = definition;
    }
  }
}

export function styleHeader(row: ExcelJS.Row): void {
  row.eachCell(styleHeaderCell);
  row.height = 24;
}

export function styleHeaderRange(
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

export function configureWorksheet(ws: ExcelJS.Worksheet): void {
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 3 }];
  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  ws.headerFooter.oddFooter = '&LConfidential&CPage &P of &N&RGenerated &D &T';
}

export function value(metrics: MetricValues, key: keyof MetricValues): number {
  return Number(metrics[key] ?? 0);
}

export function variance(current: number, prior: number): number | null {
  return prior === 0 ? null : (current - prior) / Math.abs(prior);
}

export function getLossRatioBarColorHex(ratio: number | null): string {
  if (ratio === null) return COLORS.blue;
  if (ratio >= 1.0) return COLORS.red;
  if (ratio >= 0.85) return COLORS.amber;
  return COLORS.green;
}

export function writeComparison(
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
    
    // RULE 1: Visual Flag for Avg Admin / Contract < 20
    const currentCell = row.getCell(2);
    if (definition.key === 'adminPerContract' && current > 0 && current < 20) {
      currentCell.font = { color: { argb: COLORS.red }, bold: true };
      currentCell.note = 'FLAGGED: Avg Admin / Contract is below $20.00';
    }
    
    const cell = row.getCell(5);
    cell.numFmt = PERCENT;
    if (changePercent !== null && changePercent !== 0) {
      const isBadIncrease = ['earnedLossRatio', 'cancellationRate', 'claimsPaid'].includes(
        definition.key,
      );
      const isPositive = changePercent > 0;
      const isGreen = isBadIncrease ? !isPositive : isPositive;
      cell.font = {
        color: { argb: isGreen ? '009900' : 'FF0000' },
      };
    }
  });
  return startRow + KPI_DEFINITIONS.length + 3;
}

export function writePeriodSnapshot(
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
  header.values = ['Metric', 'Total', '', '', ''];
  styleHeaderRange(ws, startRow + 1, 1, 5);
  ws.mergeCells(startRow + 1, 2, startRow + 1, 5);

  KPI_DEFINITIONS.forEach((definition, index) => {
    const row = ws.getRow(startRow + 2 + index);
    const val = value(period.values, definition.key);
    row.values = [definition.label, val, '', '', ''];
    
    const currentCell = row.getCell(2);
    currentCell.numFmt = definition.format;
    ws.mergeCells(startRow + 2 + index, 2, startRow + 2 + index, 5);
    
    // RULE 1: Visual Flag for Average admin < 20
    if (definition.key === 'adminPerContract' && val > 0 && val < 20) {
      currentCell.font = { color: { argb: COLORS.red }, bold: true };
      currentCell.note = 'FLAGGED: Average admin is below $20.00';
    }
  });
  return startRow + KPI_DEFINITIONS.length + 3;
}

// Restored Loss Code helpers
import { type LossCodeMetric } from "../models/report.types";
export const MIN_LOSS_CODE_SHARE = 0.02;
export const PIE_COLORS = ["2F75B5", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "8064A2"];
export interface LossCodeChartSegment {
  code: string;
  description: string;
  coverageNames: string[];
  paidAmount: number;
  share: number;
  color: string;
}

export function hslToHex(hue: number, saturation: number, lightness: number): string {
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

export function pieColor(index: number): string {
  if (index < PIE_COLORS.length) return PIE_COLORS[index]!;
  const goldenAngle = 137.508;
  return hslToHex((index * goldenAngle) % 360, 0.64, 0.48);
}

export function lossCodeChartSegments(rows: LossCodeMetric[]): LossCodeChartSegment[] {
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

export function visibleLossCodeRows(rows: LossCodeMetric[]): LossCodeMetric[] {
  return rows.filter((row) => (row.rolling12PaidShare ?? 0) >= MIN_LOSS_CODE_SHARE);
}

