import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, styleHeaderRange, COLORS, MONEY, PERCENT, INTEGER, dataBarRule, configureWorksheet, visibleLossCodeRows, lossCodeChartSegments, formatDateRange, formatDate } from '../utils/excel.utils';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

export class DefinitionsTab implements IDashboardTab {
  readonly id = 'tab_definitions';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, _model: ReportModel, config: ReportConfig, _tabConfig: any): void {
    
    const ws = workbook.addWorksheet('Definitions');
    configureWorksheet(ws);
    title(ws, 'REPORT DEFINITIONS', 'Controlled business rules used by this workbook');
    ws.getRow(4).values = ['Metric / Rule', 'Definition'];
    styleHeader(ws.getRow(4));
    const definitions = [
      ['Gross Income', 'Add Net admin Net reserve tax ceeding full amount'],
      ['Premium', 'Calculated as Net Admin + Net Written Reserve.'],
      ['Underwriting Profit', 'Calculated as Premium - Claims Paid.'],
      ['Earned Loss Ratio', 'Claims paid divided by Earned Reserve.'],
      [
        'Earned Reserve',
        'For active contracts, calculated dynamically using the earning schedule curve. For cancelled contracts, calculated as MAX(0, Written Reserve - Cancelled Reserve).',
      ],
      [
        'Active Contracts',
        'Distinct contracts whose latest snapshot has ContractStatus A and whose metadata.ActivationDate falls within the reporting period. Expired contracts (Status E) are explicitly excluded from this count.',
      ],
      [
        'Expired Contracts',
        'Contracts with ContractStatus E are treated mathematically as New Business (written contracts) that have naturally reached the end of their term. They are included in Contracts Written, Premium, and Earned Reserve calculations, but are not counted as Active.',
      ],
      [
        'Cancellations',
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
        `Top ${config.topDealerCount} dealers ranked by rolling-12 net written reserve.`,
      ],
      [
        'Excluded Components',
        'Broadly excludes commission section plus components containing DEALER, DLR, COMMISSION, COMM, F&I, or PACK. Additionally, when calculating Reserve, specifically excludes CLIPFEE, PREMIUMTAX, CEDINGFEE, and ADMIN. When calculating Admin, specifically excludes ROADSIDEADMIN, LOANPMT, OTHERCOMM, PREMTAX, AGENTNCB, and ROADSIDE AKMC.',
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
        'For the same ActivationDate cohort and as-of cutoff: Written Contracts = Active Contracts + contracts from that cohort whose latest state is canceled. Cancellations is grouped by CancelBillDate and is not generally the cancellation term in this equation.',
      ],
      [
        'Privacy',
        'Customer identity, contact, address, and VIN fields are excluded at MongoDB extraction.',
      ],
    ];
    definitions.forEach((definition, index) => {
      ws.getRow(index + 5).values = definition;
    });
    
    let currentRow = (ws.lastRow ? ws.lastRow.number : 30) + 3;

    ws.mergeCells(currentRow, 1, currentRow, 3);
    const bucketTitle = ws.getCell(currentRow, 1);
    bucketTitle.value = 'COMPONENT BUCKETING LOGIC';
    bucketTitle.font = { bold: true, size: 12, color: { argb: COLORS.navy } };
    currentRow++;

    ws.getRow(currentRow).values = ['Reporting Bucket', 'Included Components (Examples)', 'Strictly Excluded Components'];
    styleHeaderRange(ws, currentRow, 1, 3);
    currentRow++;

    const bucketRules = [
      [
        'Net Admin', 
        'ADMIN, BASEADMIN, BASEADMINMS, BASEADMINOS, MARKETINGFEE', 
        'ROADSIDEADMIN, LOANPMT, OTHERCOMM, PREMTAX, AGENTNCB, ROADSIDE AKMC\n(Also broadly excluded: DEALER, COMM, F&I, PACK)'
      ],
      [
        'Net Reserve', 
        'BASERESERVE, BASERESERVEFTP, OBLIGORFEE, OVERRESERVE, OWRESERVE, SIRESERVES, SLUSH, SURCHARGE, OEM TRANSPORT RESRVE', 
        'CLIPFEE, PREMIUMTAX, CEDINGFEE, CEEDING, ADMIN\n(Also broadly excluded: DEALER, COMM, F&I, PACK)'
      ],
      [
        'Commission / Other', 
        '(Ignored entirely - not aggregated into Net Admin or Net Reserve)', 
        'AGENTCOMM, COMMISSION, COMMISSION2, DEALERCOMM, DEALERSPIFF, DLROVERREMIT, FINDMOREF&I, SUBAGENTCOMM'
      ]
    ];

    bucketRules.forEach(rule => {
      const row = ws.getRow(currentRow);
      row.values = rule;
      row.height = 40; // Ensure enough room for wrapped text
      currentRow++;
    });

    ws.columns = [{ width: 28 }, { width: 55 }, { width: 55 }];
    ws.getColumn(2).alignment = { wrapText: true, vertical: 'top' };
    ws.getColumn(3).alignment = { wrapText: true, vertical: 'top' };
  }
}
