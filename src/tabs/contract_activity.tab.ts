import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, MONEY, PERCENT, configureWorksheet } from '../utils/excel.utils';

export class ContractActivityTab implements IDashboardTab {
  readonly id = 'tab_contract_detail';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, _tabConfig: unknown): void {

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
          if (cat.toUpperCase().includes('COMMISSION') || cat.toUpperCase().includes('COMM'))
            return;
          Object.keys(comps).forEach((c) => uniqueComponents.add(`${cat} - ${c}`));
        });
      }
    });
    const componentColumns = Array.from(uniqueComponents).sort();

    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const debugAgents = process.env.DEBUG_AGENT_NAME
      ? process.env.DEBUG_AGENT_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const hasDebug = debugDealers.length > 0 || debugAgents.length > 0;

    // Helper: elapsed months from effectiveDate to as-of date (mirrors report.transformer logic)
    const asOf = config.asOfDate;
    function elapsedMonths(start: Date, end: Date): number {
      if (start > end) return 0;
      const years = end.getFullYear() - start.getFullYear();
      const months = end.getMonth() - start.getMonth();
      return years * 12 + months + 1;
    }

    const baseColumns: Partial<ExcelJS.Column>[] = [
      { width: 26 },                           // col 1  Source ID
      { width: 14, numFmt: 'mm/dd/yyyy' },     // col 2  Snapshot Date
      { width: 14, numFmt: 'mm/dd/yyyy' },     // col 3  Activity Date
      { width: 14, numFmt: 'mm/dd/yyyy' },     // col 4  Effective / Activation Date
      { width: 18 },                           // col 5  Contract Number
      { width: 14 },                           // col 6  Transaction
      { width: 12 },                           // col 7  Status
      { width: 24 },                           // col 8  Agent Name
      { width: 14 },                           // col 9  Dealer Number
      { width: 30 },                           // col 10 Dealer Name
      { width: 25 },                           // col 11 Product
      { width: 16 },                           // col 12 Coverage Code
      { width: 12 },                           // col 13 Company
      { width: 18 },                           // col 14 Risk Entity
      { width: 14, numFmt: MONEY },            // col 15 Admin
      { width: 14, numFmt: MONEY },            // col 16 Reserve
      { width: 16, numFmt: MONEY },            // col 17 Earned Reserve
    ];

    if (hasDebug) {
      // Highlight financial columns green in debug mode
      baseColumns[14]!.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }; // Admin
      baseColumns[15]!.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }; // Reserve
      baseColumns[16]!.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }; // Earned Reserve

      // Debug-only earned breakdown columns — orange highlight
      const orange = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFCE4D6' } };
      baseColumns.push({ width: 22, numFmt: MONEY, fill: orange });   // col 18 Earned Amt / Contract
      baseColumns.push({ width: 16, numFmt: PERCENT, fill: orange });  // col 19 Earn Ratio (factor)
      baseColumns.push({ width: 16, fill: orange });                   // col 20 Elapsed Months
      baseColumns.push({ width: 20, fill: orange });                   // col 21 Schedule Max Month
      baseColumns.push({ width: 18, fill: orange });                   // col 22 Factor @ Elapsed (source)
    }

    const dynColumns = componentColumns.map(() => ({ width: 18, numFmt: MONEY }));
    ws.columns = [...baseColumns, ...dynColumns];

    ws.getRow(4).values = [
      'Source ID',
      'Snapshot Date',
      'Activity Date',
      'Effective / Activation Date',
      'Contract Number',
      'Transaction',
      'Status',
      'Agent Name',
      'Dealer Number',
      'Dealer Name',
      'Product',
      'Coverage Code',
      'Company',
      'Risk Entity',
      'Admin',
      'Reserve',
      'Earned Reserve',
      ...(hasDebug ? [
        'Earned Amt / Contract',
        'Earn Factor (Ratio)',
        'Elapsed Months',
        'Schedule Max Month',
        'Factor @ Elapsed (source)',
      ] : []),
      ...componentColumns,
    ];
    styleHeader(ws.getRow(4));

    // SAMPLE SIZE REDUCTION
    // Writes 1 out of every 33 contracts (~3% chronological sample spread across all years).
    // In debug mode all matching dealer/agent contracts are written.
    let sampledCount = 0;
    model.contractTransactions.forEach((item, index) => {
      if (hasDebug) {
        const dealerMatch =
          debugDealers.includes(item.dealerNumber.toLowerCase()) ||
          debugDealers.includes((item.dealerName || item.dealer).toLowerCase());
        const agentMatch = debugAgents.includes(item.agent.toLowerCase());
        if (!dealerMatch && !agentMatch) return;
      } else {
        if (index % 33 !== 0) return;
      }
      sampledCount++;

      const rowValues: ExcelJS.CellValue[] = [
        item.sourceId,
        item.snapshotDate,
        item.activityDate,
        item.effectiveDate,                          // Activation Date (moved here)
        item.contractNumber,
        item.transactionType,
        item.contractStatus,
        item.agent,
        item.dealerNumber,
        item.dealerName || item.dealer,
        item.product,
        item.coverageCode,
        item.company,
        item.riskEntity,
        item.adminAmount,
        item.reserveAmount,
        item.earnedReserveAmount,
      ];

      if (hasDebug) {
        // Earn factor = earnedReserve / writtenReserve (shows % of reserve earned so far)
        const earnFactor = item.reserveAmount !== 0
          ? item.earnedReserveAmount / item.reserveAmount
          : null;
        // Elapsed months from effective date to as-of date
        const elapsed = item.effectiveDate
          ? elapsedMonths(item.effectiveDate, asOf)
          : null;

        // The max key in the earning schedule = full contract term in months
        // If elapsed > maxScheduleMonth → code defaults to 1.0 (fully earned)
        // If elapsed <= maxScheduleMonth → code uses schedule[elapsed] (partially earned)
        const scheduleKeys = item.earningSchedule ? Object.keys(item.earningSchedule).map(Number) : [];
        const maxScheduleMonth = scheduleKeys.length > 0 ? Math.max(...scheduleKeys) : null;

        // What factor did the schedule actually provide at this elapsed month?
        const factorAtElapsed =
          item.earningSchedule && elapsed !== null
            ? item.earningSchedule[elapsed] !== undefined
              ? `${item.earningSchedule[elapsed]} (from schedule)`
              : `1.0 (defaulted — elapsed ${elapsed} > max ${maxScheduleMonth})`
            : 'No schedule (1.0)';

        rowValues.push(item.earnedReserveAmount);                             // Earned Amt / Contract
        rowValues.push(earnFactor as ExcelJS.CellValue);                      // Earn Factor (Ratio)
        rowValues.push(elapsed as ExcelJS.CellValue);                         // Elapsed Months
        rowValues.push(maxScheduleMonth as ExcelJS.CellValue);                // Schedule Max Month
        rowValues.push(factorAtElapsed);                                      // Factor @ Elapsed (source)
      }

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

    const debugExtraCols = hasDebug ? 4 : 0;
    const endColLetter = ws.getColumn(18 + debugExtraCols + componentColumns.length).letter;
    ws.autoFilter = {
      from: 'A4',
      to: `${endColLetter}${Math.max(5, sampledCount + 4)}`,
    };

  }
}
