import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, styleHeaderRange, COLORS, MONEY, PERCENT, INTEGER, dataBarRule, configureWorksheet, visibleLossCodeRows, lossCodeChartSegments, formatDateRange, formatDate } from '../utils/excel.utils';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

export class ClaimActivityTab implements IDashboardTab {
  readonly id = 'tab_claim_detail';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, model: ReportModel, _config: ReportConfig, _tabConfig: any): void {
    
    const ws = workbook.addWorksheet('Claim Activity');
    configureWorksheet(ws);
    title(
      ws,
      'SANITIZED PAID CLAIM ACTIVITY',
      'Paid claims only; no customer or vehicle identifiers',
    );
    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const debugAgents = process.env.DEBUG_AGENT_NAME
      ? process.env.DEBUG_AGENT_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const hasDebug = debugDealers.length > 0 || debugAgents.length > 0;

    const baseColumns: Partial<ExcelJS.Column>[] = [
      { width: 26 },
      { width: 14, numFmt: 'mm/dd/yyyy' },
      { width: 14, numFmt: 'mm/dd/yyyy' },
      { width: 18 },
      { width: 18 },
      { width: 12 },
      { width: 14, numFmt: MONEY }, // Paid Amount
      { width: 24 },
      { width: 30 },
      { width: 25 },
      { width: 12 },
      { width: 40 },
      { width: 26 },
    ];

    if (hasDebug) {
      baseColumns[6]!.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; // Light Orange
    }

    ws.columns = baseColumns;
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

    let sampledCount = 0;
    model.claims.forEach((item, index) => {
      if (hasDebug) {
        const dealerMatch = debugDealers.includes(item.dealer.toLowerCase());
        const agentMatch = debugAgents.includes(item.agent.toLowerCase());
        if (!dealerMatch && !agentMatch) return;
      } else {
        // SAMPLE SIZE REDUCTION
        // Writes 1 out of every 33 claims (~3% chronological sample spread across all years).
        if (index % 33 !== 0) return;
      }
      sampledCount++;

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
    ws.autoFilter = { from: 'A4', to: `M${Math.max(5, sampledCount + 4)}` };
  
  }
}
