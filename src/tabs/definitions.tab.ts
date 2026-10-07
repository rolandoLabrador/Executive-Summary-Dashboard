import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, styleHeaderRange, COLORS, REPORT_DEFINITIONS, configureWorksheet } from '../utils/excel.utils';

export class DefinitionsTab implements IDashboardTab {
  readonly id = 'tab_definitions';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, _model: ReportModel, config: ReportConfig, _tabConfig: unknown): void {
    
    const ws = workbook.addWorksheet('Definitions');
    configureWorksheet(ws);
    title(ws, 'REPORT DEFINITIONS', 'Controlled business rules used by this workbook');
    ws.getRow(4).values = ['Metric / Rule', 'Definition'];
    styleHeader(ws.getRow(4));
    const definitions = Object.entries(REPORT_DEFINITIONS);
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
        'BASERESERVE, BASERESERVEFTP, OBLIGORFEE, OVERRESERVE, OWRESERVE, SIRESERVES, SURCHARGE', 
        'CLIPFEE, PREMIUMTAX, CEDINGFEE, CEEDING, ADMIN, SLUSH, OEM TRANSPORT RESRVE\n(Also broadly excluded: DEALER, COMM, F&I, PACK)'
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
