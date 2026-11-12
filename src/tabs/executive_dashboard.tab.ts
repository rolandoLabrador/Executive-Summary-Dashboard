import type * as ExcelJS from 'exceljs';
import { type IDashboardTab } from './IDashboardTab';
import { type ReportModel, type ReportConfig, type MetricValues } from '../models/report.types';
import rawConfig from './dashboard.config.json';

interface DashboardConfig {
  activeTabs?: Record<
    string,
    | boolean
    | {
        enabled?: boolean;
        tables?: Record<string, boolean>;
      }
  >;
}
const dashboardConfig = rawConfig as unknown as DashboardConfig;
import { renderPieChartPng } from '../utils/pie-chart.renderer';
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
  dataBarRule,
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
      row = writePeriodSnapshot(ws, row, 'INCEPTION TO DATE', model.inceptionToDate);
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

    this.writeExecutivePieCharts(ws, workbook, model);

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
      { row: 15, label: '% Earned', key: 'percentEarned', format: PERCENT },
      { row: 17, label: 'Net Admin', key: 'netAdmin', format: MONEY },
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
        (d.contractsWritten >= 5 || d.activeContracts >= 5 || d.netContracts >= 5),
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

    headerRow.getCell(13).value = 'Activated Contracts';
    headerRow.getCell(14).value = 'Frequency of Claims';
    headerRow.getCell(15).value = 'Claim Severity';
    headerRow.getCell(16).value = 'Net Admin';
    headerRow.getCell(17).value = 'Earned Reserve';
    headerRow.getCell(18).value = 'Earned Loss Ratio';
    headerRow.getCell(19).value = '% Earned';

    [8, 9, 12, 13, 14, 15, 16, 17, 18, 19].forEach((col) => {
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

      row.getCell(14).value =
        dealer.claimFrequency ??
        (dealer.contractsWritten > 0 ? dealer.claimCount / dealer.contractsWritten : null);
      row.getCell(14).numFmt = PERCENT;
      row.getCell(14).alignment = { horizontal: 'center' };

      row.getCell(15).value =
        dealer.claimSeverity ??
        (dealer.claimCount > 0 ? dealer.claimsPaid / dealer.claimCount : null);
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

      const pctEarnedCell = row.getCell(19);
      pctEarnedCell.value =
        dealer.percentEarned ??
        (dealer.netReserve > 0 ? dealer.earnedReserve / dealer.netReserve : null);
      pctEarnedCell.numFmt = PERCENT;
      pctEarnedCell.alignment = { horizontal: 'center' };
    });

    if (worst.length > 0) {
      ws.addConditionalFormatting({
        ref: `R${startRow + 2}:R${startRow + 1 + worst.length}`,
        rules: [dataBarRule(20)],
      });
    }
  }

  private writeExecutivePieCharts(
    ws: ExcelJS.Worksheet,
    workbook: ExcelJS.Workbook,
    model: ReportModel,
  ): void {
    const products = [...model.itdProducts].filter(
      (p) => p.netReserve > 0 || p.claimsPaid > 0 || p.netAdmin > 0,
    );
    if (products.length === 0) return;

    // Palette for products: consistent color per product
    const COLOR_PALETTE = [
      '#1F4E79', // Navy
      '#2CA02C', // Green
      '#D9534F', // Coral / Red
      '#FF7F0E', // Orange
      '#9467BD', // Purple
      '#17BECF', // Teal
      '#8C564B', // Brown
      '#E377C2', // Pink
      '#7F7F7F', // Gray
      '#BCBD22', // Olive
    ];

    const productColors = new Map<string, string>();
    products.forEach((p, idx) => {
      productColors.set(p.name, COLOR_PALETTE[idx % COLOR_PALETTE.length]!);
    });

    const startRow = Math.max(45, (ws.lastRow?.number ?? 42) + 3);

    // Section Banner
    ws.mergeCells(startRow, 1, startRow, 19);
    const banner = ws.getCell(startRow, 1);
    banner.value = 'PORTFOLIO PRODUCT & REVENUE DISTRIBUTION (INCEPTION TO DATE)';
    banner.font = { bold: true, size: 12, color: { argb: COLORS.navy } };
    banner.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF1F5' } };
    banner.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(startRow).height = 24;

    const chartHeaderRow = startRow + 2;
    const legendHeaderRow = chartHeaderRow + 1;
    const contentStartRow = legendHeaderRow + 1;
    const CHART_ROWS = Math.max(12, products.length + 3);

    // Lock row heights for all chart rows so the images never get squished or overlap text
    for (let r = chartHeaderRow; r <= contentStartRow + CHART_ROWS; r++) {
      ws.getRow(r).height = 20;
    }

    // ── CHART 1: PRODUCT MIX (NET WRITTEN RESERVE) ──
    // Cols 1-5
    ws.mergeCells(chartHeaderRow, 1, chartHeaderRow, 5);
    const h1 = ws.getCell(chartHeaderRow, 1);
    h1.value = 'OVERALL PRODUCT MIX (NET RESERVE)';
    h1.font = { bold: true, size: 10, color: { argb: COLORS.white } };
    h1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    h1.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell(legendHeaderRow, 3).value = 'Product';
    ws.getCell(legendHeaderRow, 4).value = 'Net Reserve';
    ws.getCell(legendHeaderRow, 5).value = 'Share';
    [3, 4, 5].forEach((c) => {
      const cell = ws.getCell(legendHeaderRow, c);
      cell.font = { bold: true, size: 9, color: { argb: COLORS.white } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
      cell.alignment = { horizontal: c === 3 ? 'left' : 'right', vertical: 'middle' };
    });

    const reserveProducts = [...products]
      .filter((p) => p.netReserve > 0)
      .sort((a, b) => b.netReserve - a.netReserve);
    const totalReserve = reserveProducts.reduce((sum, p) => sum + p.netReserve, 0);

    reserveProducts.forEach((p, idx) => {
      const r = contentStartRow + idx;
      const col3 = ws.getCell(r, 3);
      const color = productColors.get(p.name) || '#1F4E79';
      col3.value = `● ${p.displayName || p.name}`;
      col3.font = { bold: true, size: 9, color: { argb: color.replace('#', '') } };
      col3.alignment = { horizontal: 'left', vertical: 'middle' };

      const col4 = ws.getCell(r, 4);
      col4.value = p.netReserve;
      col4.numFmt = MONEY;
      col4.alignment = { horizontal: 'right', vertical: 'middle' };

      const col5 = ws.getCell(r, 5);
      col5.value = totalReserve > 0 ? p.netReserve / totalReserve : 0;
      col5.numFmt = PERCENT;
      col5.alignment = { horizontal: 'right', vertical: 'middle' };
    });

    if (reserveProducts.length > 0 && totalReserve > 0) {
      try {
        const slices = reserveProducts.map((p) => ({
          value: p.netReserve,
          color: productColors.get(p.name) || '#1F4E79',
        }));
        const chartPng = renderPieChartPng(slices, 440, 440);
        const imageId = workbook.addImage({
          base64: `data:image/png;base64,${chartPng.toString('base64')}`,
          extension: 'png',
        });
        ws.addImage(imageId, {
          tl: { col: 0.15, row: contentStartRow - 0.7 },
          ext: { width: 190, height: 190 },
        });
      } catch (err) {
        console.warn('Failed rendering Executive Product Mix Pie Chart:', err);
      }
    }

    // ── CHART 2: ADMIN REVENUE MIX ──
    // Cols 8-13
    ws.mergeCells(chartHeaderRow, 8, chartHeaderRow, 13);
    const h2 = ws.getCell(chartHeaderRow, 8);
    h2.value = 'ADMIN REVENUE MIX (NET ADMIN)';
    h2.font = { bold: true, size: 10, color: { argb: COLORS.white } };
    h2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    h2.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell(legendHeaderRow, 11).value = 'Product';
    ws.getCell(legendHeaderRow, 12).value = 'Net Admin';
    ws.getCell(legendHeaderRow, 13).value = 'Share';
    [11, 12, 13].forEach((c) => {
      const cell = ws.getCell(legendHeaderRow, c);
      cell.font = { bold: true, size: 9, color: { argb: COLORS.white } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
      cell.alignment = { horizontal: c === 11 ? 'left' : 'right', vertical: 'middle' };
    });

    const adminProducts = [...products]
      .filter((p) => p.netAdmin > 0)
      .sort((a, b) => b.netAdmin - a.netAdmin);
    const totalAdmin = adminProducts.reduce((sum, p) => sum + p.netAdmin, 0);

    adminProducts.forEach((p, idx) => {
      const r = contentStartRow + idx;
      const col11 = ws.getCell(r, 11);
      const color = productColors.get(p.name) || '#1F4E79';
      col11.value = `● ${p.displayName || p.name}`;
      col11.font = { bold: true, size: 9, color: { argb: color.replace('#', '') } };
      col11.alignment = { horizontal: 'left', vertical: 'middle' };

      const col12 = ws.getCell(r, 12);
      col12.value = p.netAdmin;
      col12.numFmt = MONEY;
      col12.alignment = { horizontal: 'right', vertical: 'middle' };

      const col13 = ws.getCell(r, 13);
      col13.value = totalAdmin > 0 ? p.netAdmin / totalAdmin : 0;
      col13.numFmt = PERCENT;
      col13.alignment = { horizontal: 'right', vertical: 'middle' };
    });

    if (adminProducts.length > 0 && totalAdmin > 0) {
      try {
        const slices = adminProducts.map((p) => ({
          value: p.netAdmin,
          color: productColors.get(p.name) || '#1F4E79',
        }));
        const chartPng = renderPieChartPng(slices, 440, 440);
        const imageId = workbook.addImage({
          base64: `data:image/png;base64,${chartPng.toString('base64')}`,
          extension: 'png',
        });
        ws.addImage(imageId, {
          tl: { col: 7.2, row: contentStartRow - 0.7 },
          ext: { width: 190, height: 190 },
        });
      } catch (err) {
        console.warn('Failed rendering Executive Admin Revenue Pie Chart:', err);
      }
    }

    // ── CHART 3: CLAIMS PAID & LOSS RATIO MIX ──
    // Cols 14-19
    ws.mergeCells(chartHeaderRow, 14, chartHeaderRow, 19);
    const h3 = ws.getCell(chartHeaderRow, 14);
    h3.value = 'CLAIMS PAID & LOSS RATIO MIX';
    h3.font = { bold: true, size: 10, color: { argb: COLORS.white } };
    h3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    h3.alignment = { horizontal: 'center', vertical: 'middle' };

    ws.getCell(legendHeaderRow, 17).value = 'Product';
    ws.getCell(legendHeaderRow, 18).value = 'Claims Paid';
    ws.getCell(legendHeaderRow, 19).value = 'Loss Ratio';
    [17, 18, 19].forEach((c) => {
      const cell = ws.getCell(legendHeaderRow, c);
      cell.font = { bold: true, size: 9, color: { argb: COLORS.white } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
      cell.alignment = { horizontal: c === 17 ? 'left' : 'right', vertical: 'middle' };
    });

    const claimsProducts = [...products]
      .filter((p) => p.claimsPaid > 0)
      .sort((a, b) => b.claimsPaid - a.claimsPaid);
    const totalClaims = claimsProducts.reduce((sum, p) => sum + p.claimsPaid, 0);

    claimsProducts.forEach((p, idx) => {
      const r = contentStartRow + idx;
      const col17 = ws.getCell(r, 17);
      const color = productColors.get(p.name) || '#1F4E79';
      col17.value = `● ${p.displayName || p.name}`;
      col17.font = { bold: true, size: 9, color: { argb: color.replace('#', '') } };
      col17.alignment = { horizontal: 'left', vertical: 'middle' };

      const col18 = ws.getCell(r, 18);
      col18.value = p.claimsPaid;
      col18.numFmt = MONEY;
      col18.alignment = { horizontal: 'right', vertical: 'middle' };

      const col19 = ws.getCell(r, 19);
      col19.value = p.earnedLossRatio;
      col19.numFmt = PERCENT;
      col19.alignment = { horizontal: 'right', vertical: 'middle' };
      if ((p.earnedLossRatio || 0) > 0.8) {
        col19.font = { bold: true, color: { argb: 'FF9C0006' } };
      } else {
        col19.font = { bold: true, color: { argb: 'FF006100' } };
      }
    });

    if (claimsProducts.length > 0 && totalClaims > 0) {
      try {
        const slices = claimsProducts.map((p) => ({
          value: p.claimsPaid,
          color: productColors.get(p.name) || '#1F4E79',
        }));
        const chartPng = renderPieChartPng(slices, 440, 440);
        const imageId = workbook.addImage({
          base64: `data:image/png;base64,${chartPng.toString('base64')}`,
          extension: 'png',
        });
        ws.addImage(imageId, {
          tl: { col: 13.2, row: contentStartRow - 0.7 },
          ext: { width: 190, height: 190 },
        });
      } catch (err) {
        console.warn('Failed rendering Executive Claims & Loss Ratio Pie Chart:', err);
      }
    }
  }
}
