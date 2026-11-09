import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import {
  title,
  styleHeader,
  MONEY,
  PERCENT,
  INTEGER,
  dataBarRule,
  configureWorksheet,
  formatDateRange,
} from '../utils/excel.utils';

export class MonthlyTrendsTab implements IDashboardTab {
  readonly id = 'tab_monthly';

  constructor() {}

  render(
    workbook: ExcelJS.Workbook,
    model: ReportModel,
    _config: ReportConfig,
    _tabConfig?: unknown,
  ): void {
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
      'Cancellations',
      'Net Admin',
      'Net Reserve',
      'Claim Severity',
      'Frequency of Claims',
      'Earned Loss Ratio',
      '% Earned',
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
        item.claimSeverity ?? (item.claimCount > 0 ? item.claimsPaid / item.claimCount : null),
        item.claimFrequency ??
          (item.contractsWritten > 0 ? item.claimCount / item.contractsWritten : null),
        item.earnedLossRatio,
        item.percentEarned ?? (item.netReserve > 0 ? item.earnedReserve / item.netReserve : null),
        item.cancellationRate,
        item.earnedLossRatio,
      ];
      row.getCell(1).numFmt = 'mmm-yy';
      [2, 3, 4].forEach((column) => {
        row.getCell(column).numFmt = INTEGER;
      });
      [5, 6, 7].forEach((column) => {
        row.getCell(column).numFmt = MONEY;
      });
      [8, 9, 10, 11, 12].forEach((column) => {
        row.getCell(column).numFmt = PERCENT;
      });
    });
    const lastRow = model.monthly.length + 4;
    ws.addConditionalFormatting({
      ref: `L5:L${lastRow}`,
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
      { width: 14 },
      { width: 17 },
      { width: 30 },
    ];
  }
}
