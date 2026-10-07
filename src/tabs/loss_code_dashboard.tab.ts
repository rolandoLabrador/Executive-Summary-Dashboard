import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, styleHeaderRange, COLORS, MONEY, PERCENT, INTEGER, dataBarRule, configureWorksheet, visibleLossCodeRows, lossCodeChartSegments, formatDateRange, formatDate } from '../utils/excel.utils';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

export class LossCodeDashboardTab implements IDashboardTab {
  readonly id = 'tab_loss_code';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, tabConfig: unknown): void {

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

  
}
