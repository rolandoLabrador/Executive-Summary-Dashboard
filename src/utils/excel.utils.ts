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

export const KPI_DEFINITIONS: MetricDefinition[] = [
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
    
    // RULE 1: Visual Flag for Avg Admin / Contract < 20
    if (definition.key === 'adminPerContract' && val > 0 && val < 20) {
      currentCell.font = { color: { argb: COLORS.red }, bold: true };
      currentCell.note = 'FLAGGED: Avg Admin / Contract is below $20.00';
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

