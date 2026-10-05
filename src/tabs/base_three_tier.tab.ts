import type * as ExcelJS from 'exceljs';
import { type IDashboardTab } from './IDashboardTab';
import { type ReportModel, type ReportConfig, type DimensionMetric } from '../models/report.types';
import {
  COLORS,
  MONEY,
  INTEGER,
  PERCENT,
  configureWorksheet,
  title,
  formatDate,
  formatDateRange,
  dataBarRule,
  getLossRatioBarColorHex
} from '../utils/excel.utils';

export abstract class BaseThreeTierTab implements IDashboardTab {
  public abstract readonly id: string;
  protected abstract readonly name: string;
  protected abstract readonly heading: string;
  protected abstract readonly isDealer: boolean;
  protected abstract readonly isAgent: boolean;

  protected abstract getItdRows(model: ReportModel): DimensionMetric[];
  protected abstract getRollingRows(model: ReportModel): DimensionMetric[];
  protected abstract getYtdRows(model: ReportModel): DimensionMetric[];
  protected abstract getLimit(config: ReportConfig): number | undefined;

  public render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, tabConfig: any): void {
    const ws = workbook.addWorksheet(this.name);
    configureWorksheet(ws);
    const maxCols = this.isDealer || this.isAgent ? 19 : 17;
    title(
      ws,
      this.heading,
      `Inception to Date, Rolling 12-Month & Latest Month Results | Through ${formatDate(model.yearToDate.currentEnd)}`,
      maxCols,
    );
    
    const tables = tabConfig?.tables || {};

    let nextRow = 4;
    
    if (tables.inceptionToDate !== false) {
      nextRow = this.writeTable(
        ws,
        nextRow,
        `INCEPTION TO DATE RESULTS (${formatDateRange(new Date(2010, 0, 1), model.yearToDate.currentEnd)})`,
        this.getItdRows(model),
        config,
        maxCols
      );
    }
    
    if (tables.rolling12Month !== false) {
      nextRow = this.writeTable(
        ws,
        nextRow,
        `ROLLING 12-MONTH RESULTS (${formatDateRange(model.rolling12.currentStart, model.yearToDate.currentEnd)})`,
        this.getRollingRows(model),
        config,
        maxCols
      );
    }
    
    if (tables.yearToDate !== false) {
      this.writeTable(
        ws,
        nextRow,
        `YEAR TO DATE RESULTS (${formatDateRange(model.yearToDate.currentStart, model.yearToDate.currentEnd)})`,
        this.getYtdRows(model),
        config,
        maxCols
      );
    }

    ws.columns = [
      { width: 8 },
      { width: this.isDealer ? 16 : 30 },
      ...(this.isDealer ? [{ width: 34 }, { width: 36 }] : []),
      { width: 12 }, // Written
      { width: 12 }, // Active Contracts
      { width: 16 }, // Gross Income
      { width: 14 }, // Cancellations
      { width: 16 }, // Net Admin
      { width: 16 }, // Average admin
      { width: 16 }, // Net Reserve
      { width: 16 }, // Premium
      { width: 16 }, // Earned Reserve
      { width: 16 }, // Claims Paid
      { width: 16 }, // Underwriting Profit
      { width: 12 }, // Claim Count
      { width: 16 }, // Earned Loss Ratio
      { width: 17 }, // Cancellation Rate
      { width: 28 }, // Loss Ratio Bar
    ];
  }

  private writeTable(
    ws: ExcelJS.Worksheet,
    startRow: number,
    sectionHeading: string,
    rows: DimensionMetric[],
    config: ReportConfig,
    maxCols: number
  ): number {
    ws.mergeCells(startRow, 1, startRow, maxCols);
    const headingCell = ws.getCell(startRow, 1);
    headingCell.value = sectionHeading;
    headingCell.font = { bold: true, size: 12, color: { argb: COLORS.navy } };

    const limit = this.getLimit(config);
    const selected = limit ? rows.slice(0, limit) : rows;
    const headers = [
      'Rank',
      this.isDealer ? 'Dealer Number' : this.name.replace(' Dashboard', ''),
      ...(this.isDealer ? ['Dealer Name', 'Agents'] : this.isAgent ? ['Agent Name', 'Dealers'] : []),
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
      'Earned Loss Ratio',
      'Cancellation Rate',
      'Loss Ratio Bar',
    ];
    ws.getRow(startRow + 1).values = headers;
    this.styleHeaderRange(ws, startRow + 1, 1, maxCols);

    selected.forEach((item, index) => {
      const row = ws.getRow(startRow + 2 + index);
      row.values = [
        index + 1,
        item.name,
        ...(this.isDealer
          ? [item.displayName || 'Name unavailable', item.relatedAgents?.join(', ') || 'Unassigned']
          : this.isAgent
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
        item.earnedLossRatio,
        item.cancellationRate,
        item.earnedLossRatio,
      ];
      const offset = this.isDealer || this.isAgent ? 2 : 0;
      [3, 4, 6, 14].forEach((column) => {
        row.getCell(column + offset).numFmt = INTEGER;
      });
      [5, 7, 8, 9, 10, 11, 12, 13].forEach((column) => {
        row.getCell(column + offset).numFmt = MONEY;
      });
      [15, 16, 17].forEach((column) => {
        row.getCell(column + offset).numFmt = PERCENT;
      });
      
      // RULE 1: Visual Flag for Avg Admin / Contract < 20
      const adminCell = row.getCell(8 + offset);
      if (item.adminPerContract !== null && item.adminPerContract > 0 && item.adminPerContract < 20) {
        adminCell.font = { color: { argb: COLORS.red }, bold: true };
        adminCell.note = 'FLAGGED: Avg Admin / Contract is below $20.00';
      }

      // RULE 2: Colored Data Bars for Loss Ratio
      const barCell = row.getCell(17 + offset);
      const colorHex = getLossRatioBarColorHex(item.earnedLossRatio);
      ws.addConditionalFormatting({
        ref: barCell.address,
        rules: [dataBarRule(1, colorHex)],
      });
    });

    if (selected.length > 0) {
      const firstDataRow = startRow + 2;
      const lastDataRow = startRow + 1 + selected.length;
      const earnedRatioCol = this.isDealer || this.isAgent ? 'Q' : 'O';

      const ratioRules: ExcelJS.ConditionalFormattingRule[] = [
        {
          type: 'cellIs',
          priority: 2,
          operator: 'greaterThan',
          formulae: [config.highLossRatio.toString()],
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F4CCCC' } } },
        },
        {
          type: 'cellIs',
          priority: 3,
          operator: 'between',
          formulae: [config.warningLossRatio.toString(), config.highLossRatio.toString()],
          style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2CC' } } },
        },
      ];

      ws.addConditionalFormatting({
        ref: `${earnedRatioCol}${firstDataRow}:${earnedRatioCol}${lastDataRow}`,
        rules: ratioRules,
      });
    }

    return startRow + selected.length + 4;
  }

  private styleHeaderRange(
    ws: ExcelJS.Worksheet,
    rowNumber: number,
    startColumn: number,
    endColumn: number,
  ): void {
    for (let column = startColumn; column <= endColumn; column += 1) {
      this.styleHeaderCell(ws.getCell(rowNumber, column));
    }
    ws.getRow(rowNumber).height = Math.max(ws.getRow(rowNumber).height ?? 0, 24);
  }

  private styleHeaderCell(cell: ExcelJS.Cell): void {
    cell.font = { bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
}
