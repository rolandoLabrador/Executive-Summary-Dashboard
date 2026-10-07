import type * as ExcelJS from 'exceljs';
import { type IDashboardTab } from './IDashboardTab';
import { type ReportModel, type ReportConfig, type MetricValues } from '../models/report.types';
import rawConfig from './dashboard.config.json';

interface DashboardConfig {
  activeTabs?: Record<string, boolean | {
    enabled?: boolean;
    tables?: Record<string, boolean>;
  }>;
}
const dashboardConfig = rawConfig as unknown as DashboardConfig;
import {
  COLORS,
  MONEY,
  INTEGER,
  PERCENT,
  configureWorksheet,
  title,
  formatDateRange,
  comparisonHeading,
  writeComparison,
  writePeriodSnapshot,
  value,
  dataBarRule
} from '../utils/excel.utils';

export class ExecutiveDashboardTab implements IDashboardTab {
  public readonly id = 'tab_executive_summary';

  public render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig): void {
    const ws = workbook.addWorksheet('Executive Dashboard');
    configureWorksheet(ws);
    title(
      ws,
      `${config.companyName.toUpperCase()} EXECUTIVE SUMMARY`,
      `Current reporting month: ${formatDateRange(model.currentMonth.currentStart, model.currentMonth.currentEnd)} | Paid claims / net written reserve`,
      19,
    );
    
    // Read granular table toggles from the JSON configuration
    const tabConfig = dashboardConfig.activeTabs?.[this.id];
    const tables = (typeof tabConfig === 'object' ? tabConfig?.tables : {}) || {};

    let row = 4;

    if (tables.currentMonthComparison !== false) {
      row = writeComparison(
        ws,
        row,
        comparisonHeading('Current Month vs Same Month Prior Year', model.currentMonth),
        model.currentMonth,
        'Prior Year',
      );
    }
    
    if (tables.yearToDateComparison !== false) {
      row = writeComparison(
        ws,
        row,
        comparisonHeading('Year to Date vs Prior-Year YTD', model.yearToDate),
        model.yearToDate,
        'Prior Year',
      );
    }
    
    if (tables.inceptionToDateSnapshot !== false) {
      row = writePeriodSnapshot(
        ws,
        row,
        'INCEPTION TO DATE',
        model.inceptionToDate,
      );
    }
    
    if (tables.priorYearSnapshot !== false) {
      writePeriodSnapshot(
        ws,
        row,
        `FULL YEAR ${model.priorCalendarYear.start.getFullYear()}`,
        model.priorCalendarYear,
      );
    }
    
    if (tables.trend12Month !== false) {
      this.writeExecutiveTrend(ws, model);
    }
    
    this.writeWorstDealers(ws, model, tables);
    
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
    // We add 1 to the length to account for the ITD column
    const lastColumn = firstColumn + months.length;
    
    ws.mergeCells(4, firstColumn, 4, lastColumn);
    const heading = ws.getCell(4, firstColumn);
    heading.value = '12-MONTH TREND & INCEPTION TO DATE';
    heading.font = { bold: true, size: 14, color: { argb: COLORS.navy } };

    // Write month headers
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

    // Write ITD header
    const itdCol = lastColumn;
    const itdHeader = ws.getCell(5, itdCol);
    itdHeader.value = 'ITD';
    itdHeader.font = { bold: true, color: { argb: COLORS.white } };
    itdHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    itdHeader.alignment = { horizontal: 'center' };

    const trends: Array<{ row: number; label: string; key: keyof MetricValues; format: string }> = [
      { row: 7, label: 'Net Reserve', key: 'netReserve', format: MONEY },
      { row: 9, label: 'Earned Reserve', key: 'earnedReserve', format: MONEY },
      { row: 11, label: 'Claims Paid', key: 'claimsPaid', format: MONEY },
      { row: 13, label: 'Earned Loss Ratio', key: 'earnedLossRatio', format: PERCENT },
      { row: 15, label: 'Net Admin', key: 'netAdmin', format: MONEY },
    ];

    trends.forEach(({ row, label, key, format }, trendIndex) => {
      ws.mergeCells(row - 1, firstColumn, row - 1, lastColumn);
      const labelCell = ws.getCell(row - 1, firstColumn);
      labelCell.value = label;
      labelCell.font = { bold: true, color: { argb: COLORS.darkGray } };
      
      // Write month values
      months.forEach((month, index) => {
        const cell = ws.getCell(row, firstColumn + index);
        const val = value(month, key);
        cell.value = val;
        cell.numFmt = format;
        cell.alignment = { horizontal: 'center' };
      });

      // Write ITD value
      const itdCell = ws.getCell(row, itdCol);
      const itdVal = value(model.inceptionToDate.values, key);
      itdCell.value = itdVal;
      itdCell.numFmt = format;
      itdCell.alignment = { horizontal: 'center' };
      itdCell.font = { bold: true };

      if (months.length > 0) {
        // Data bar only for the months (don't include ITD in the relative bar sizing because ITD is huge)
        ws.addConditionalFormatting({
          ref: `${ws.getCell(row, firstColumn).address}:${ws.getCell(row, firstColumn + months.length - 1).address}`,
          rules: [dataBarRule(10 + trendIndex)],
        });
      }
    });

    ws.mergeCells(19, firstColumn, 19, lastColumn);
    const note = ws.getCell(19, firstColumn);
    note.value =
      'Green = latest completed month  |  Amber = preceding month  |  Bars show relative monthly magnitude (excluding ITD)';
    note.font = { italic: true, color: { argb: COLORS.darkGray } };
    note.alignment = { horizontal: 'center' };
  }

  private writeWorstDealers(
    ws: ExcelJS.Worksheet,
    model: ReportModel,
    tablesConfig: Record<string, boolean>,
  ): void {
    if (tablesConfig.worstDealers === false) return;

    const startRow = 20;
    const firstColumn = 8;
    
    // Filter ITD dealers: must have earned reserve > 0, claims paid > 0, and AT LEAST 5 CONTRACTS
    const dealers = [...model.itdDealers].filter(
      (d) =>
        d.earnedReserve > 0 &&
        d.claimsPaid > 0 &&
        (d.contractsWritten >= 5 || d.activeContracts >= 5 || d.netContracts >= 5)
    );
    dealers.sort((a, b) => (b.earnedLossRatio || 0) - (a.earnedLossRatio || 0));
    const worst = dealers.slice(0, 20);

    ws.mergeCells(startRow, firstColumn, startRow, firstColumn + 11);
    const heading = ws.getCell(startRow, firstColumn);
    heading.value = 'TOP 20 WORST PERFORMING DEALERS BY EARNED LOSS RATIO (INCEPTION TO DATE)';
    heading.font = { bold: true, size: 12, color: { argb: COLORS.navy } };

    const headerRow = ws.getRow(startRow + 1);
    headerRow.height = 30; // Make room for wrapping text

    headerRow.getCell(8).value = 'Rank';

    ws.mergeCells(startRow + 1, 9, startRow + 1, 11);
    headerRow.getCell(9).value = 'Dealer';

    headerRow.getCell(12).value = 'Agent';

    headerRow.getCell(13).value = 'Active Contracts';
    headerRow.getCell(14).value = 'Claim Count';
    headerRow.getCell(15).value = 'Claims Paid';
    headerRow.getCell(16).value = 'Net Admin';
    headerRow.getCell(17).value = 'Earned Reserve';
    headerRow.getCell(18).value = 'Earned Loss Ratio';

    [8, 9, 12, 13, 14, 15, 16, 17, 18].forEach((col) => {
      const cell = headerRow.getCell(col);
      cell.font = { bold: true, color: { argb: COLORS.white } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });

    worst.forEach((dealer, index) => {
      const row = ws.getRow(startRow + 2 + index);

      row.getCell(8).value = index + 1;
      row.getCell(8).alignment = { horizontal: 'center' };

      ws.mergeCells(startRow + 2 + index, 9, startRow + 2 + index, 11);
      row.getCell(9).value = dealer.displayName || dealer.name;

      row.getCell(12).value = (dealer.relatedAgents || []).join(', ');
      row.getCell(12).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

      row.getCell(13).value = dealer.activeContracts;
      row.getCell(13).numFmt = INTEGER;
      row.getCell(13).alignment = { horizontal: 'center' };

      row.getCell(14).value = dealer.claimCount;
      row.getCell(14).numFmt = INTEGER;
      row.getCell(14).alignment = { horizontal: 'center' };

      row.getCell(15).value = dealer.claimsPaid;
      row.getCell(15).numFmt = MONEY;
      row.getCell(15).alignment = { horizontal: 'center' };

      row.getCell(16).value = dealer.netAdmin;
      row.getCell(16).numFmt = MONEY;
      row.getCell(16).alignment = { horizontal: 'center' };

      row.getCell(17).value = dealer.earnedReserve;
      row.getCell(17).numFmt = MONEY;
      row.getCell(17).alignment = { horizontal: 'center' };

      const lossRatioCell = row.getCell(18);
      lossRatioCell.value = dealer.earnedLossRatio;
      lossRatioCell.numFmt = PERCENT;
      lossRatioCell.alignment = { horizontal: 'center' };
    });
    
    if (worst.length > 0) {
      ws.addConditionalFormatting({
        ref: `R${startRow + 2}:R${startRow + 1 + worst.length}`,
        rules: [dataBarRule(20)],
      });
    }
  }
}
