import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, styleHeaderRange, COLORS, MONEY, PERCENT, INTEGER, dataBarRule, configureWorksheet, visibleLossCodeRows, lossCodeChartSegments, formatDateRange, formatDate } from '../utils/excel.utils';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

export class DataQualityTab implements IDashboardTab {
  readonly id = 'tab_data_quality';

  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, model: ReportModel, _config: ReportConfig, _tabConfig: any): void {
    
    const ws = workbook.addWorksheet('Data Quality');
    configureWorksheet(ws);
    title(
      ws,
      'DATA QUALITY AND RECONCILIATION',
      `Generated ${model.generatedAt.toLocaleString('en-US')}`,
      8,
    );

    let currentRow = 4;

    ws.mergeCells(currentRow, 1, currentRow, 8);
    const auditHeading = ws.getCell(currentRow, 1);
    auditHeading.value = 'PIPELINE INGESTION AUDIT & RECONCILIATION';
    auditHeading.font = { bold: true, size: 12, color: { argb: COLORS.navy } };
    currentRow++;

    const auditHeaderRow = currentRow;
    ws.getRow(auditHeaderRow).values = [
      'Status',
      'Job Type',
      'Portal Count',
      'Unique Units',
      'Uploaded Line Items',
      'Variance',
      'Source File',
      'Execution Timestamp (UTC)',
    ];
    styleHeaderRange(ws, auditHeaderRow, 1, 8);
    currentRow++;

    if (model.pipelineAudits && model.pipelineAudits.length > 0) {
      model.pipelineAudits.forEach((audit) => {
        const portalCount = audit.counts.portalCount;
        const uniqueUnits = audit.counts.uniqueCount ?? audit.counts.uploadedCount;
        const uploadedLines = audit.counts.uploadedCount;
        const unitVariance = uniqueUnits - portalCount;

        const isReconciled =
          (audit.reconciliation.isMatch && audit.reconciliation.status === 'PASSED') ||
          (Math.abs(unitVariance) <= Math.max(2, Math.round(portalCount * 0.001)) &&
            audit.counts.processedCount === audit.counts.uploadedCount);

        const status = isReconciled ? 'PASSED' : audit.reconciliation.status;

        const row = ws.getRow(currentRow);
        const statusCell = row.getCell(1);
        statusCell.value = status;
        statusCell.font = { bold: true };
        if (isReconciled) {
          statusCell.font = { bold: true, color: { argb: '006100' } };
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } };
        } else {
          statusCell.font = { bold: true, color: { argb: '9C0006' } };
          statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
        }
        statusCell.alignment = { horizontal: 'center' };

        row.getCell(2).value = audit.jobType;
        row.getCell(3).value = portalCount;
        row.getCell(3).numFmt = INTEGER;
        row.getCell(4).value = uniqueUnits;
        row.getCell(4).numFmt = INTEGER;
        row.getCell(5).value = uploadedLines;
        row.getCell(5).numFmt = INTEGER;
        row.getCell(6).value = unitVariance;
        row.getCell(6).numFmt = INTEGER;
        row.getCell(7).value = audit.fileMetadata.fileName;
        row.getCell(8).value = audit.executionTimestamp
          ? audit.executionTimestamp.toISOString().replace('T', ' ').replace(/\..+/, '')
          : audit.executionDateStr;
        currentRow++;
      });
    } else {
      ws.mergeCells(currentRow, 1, currentRow, 8);
      const noAuditCell = ws.getCell(currentRow, 1);
      noAuditCell.value = 'No pipeline audit records found in AuditDB.DataReconciliationAudit.';
      noAuditCell.font = { italic: true, color: { argb: COLORS.darkGray } };
      currentRow++;
    }

    currentRow += 2; // Blank row spacing

    const countsHeaderRow = currentRow;
    ws.getRow(countsHeaderRow).values = ['Source Document Category', 'Document Count'];
    styleHeaderRange(ws, countsHeaderRow, 1, 2);
    currentRow++;

    const counts: Array<[string, number]> = [
      ['Contract source documents', model.sourceCounts.contractDocuments],
      ['Cancellation source documents', model.sourceCounts.cancellationDocuments],
      ['Claim source documents', model.sourceCounts.claimDocuments],
      ['Unique normalized contract transactions', model.sourceCounts.uniqueContractTransactions],
      ['Unique paid claim transactions', model.sourceCounts.uniqueClaims],
      ['Data-quality issues', model.dataQualityIssues.length],
    ];
    counts.forEach(([label, count]) => {
      const row = ws.getRow(currentRow);
      row.getCell(1).value = label;
      row.getCell(2).value = count;
      row.getCell(2).numFmt = INTEGER;
      currentRow++;
    });

    currentRow += 2; // Blank row spacing

    // --- ISSUE BUCKETS SUMMARY TABLE ---
    ws.mergeCells(currentRow, 1, currentRow, 4);
    const bucketTitle = ws.getCell(currentRow, 1);
    bucketTitle.value = 'DATA QUALITY ISSUES BY CATEGORY (BUCKETS)';
    bucketTitle.font = { bold: true, size: 12, color: { argb: COLORS.navy } };
    currentRow++;

    const bucketsHeaderRow = currentRow;
    ws.getRow(bucketsHeaderRow).values = ['Category (Bucket)', 'Errors', 'Warnings', 'Total Issues'];
    styleHeaderRange(ws, bucketsHeaderRow, 1, 4);
    currentRow++;

    const buckets = new Map<string, { errors: number; warnings: number }>();
    model.dataQualityIssues.forEach((issue) => {
      const cat = issue.category || 'Uncategorized';
      const b = buckets.get(cat) || { errors: 0, warnings: 0 };
      if (issue.severity === 'Error') b.errors++;
      if (issue.severity === 'Warning') b.warnings++;
      buckets.set(cat, b);
    });

    const sortedBuckets = Array.from(buckets.entries()).sort(
      (a, b) => (b[1].errors + b[1].warnings) - (a[1].errors + a[1].warnings)
    );
    
    if (sortedBuckets.length === 0) {
      ws.getCell(currentRow, 1).value = 'No data quality issues found.';
      ws.getCell(currentRow, 1).font = { italic: true };
      currentRow++;
    } else {
      sortedBuckets.forEach(([category, bCounts]) => {
        const row = ws.getRow(currentRow);
        row.getCell(1).value = category;
        row.getCell(2).value = bCounts.errors;
        row.getCell(3).value = bCounts.warnings;
        row.getCell(4).value = bCounts.errors + bCounts.warnings;
        
        [2, 3, 4].forEach(col => { 
          row.getCell(col).numFmt = INTEGER; 
          row.getCell(col).alignment = { horizontal: 'center' }; 
        });
        
        // Highlight errors in red if > 0
        if (bCounts.errors > 0) {
          row.getCell(2).font = { color: { argb: '9C0006' }, bold: true };
        }
        currentRow++;
      });
    }

    currentRow += 2; // Blank row spacing

    ws.mergeCells(currentRow, 1, currentRow, 6);
    const detailTitle = ws.getCell(currentRow, 1);
    detailTitle.value = 'DETAILED ISSUE LOG';
    detailTitle.font = { bold: true, size: 12, color: { argb: COLORS.navy } };
    currentRow++;

    const issuesHeaderRow = currentRow;
    ws.getRow(issuesHeaderRow).values = [
      'Severity',
      'Category',
      'Contract Number',
      'Dealer Name',
      'Source ID',
      'Message',
    ];
    styleHeaderRange(ws, issuesHeaderRow, 1, 6);
    currentRow++;

    const issuesStartRow = currentRow;
    model.dataQualityIssues.forEach((issue) => {
      ws.getRow(currentRow).values = [
        issue.severity,
        issue.category,
        issue.contractNumber,
        issue.dealerName,
        issue.sourceId,
        issue.message,
      ];
      currentRow++;
    });

    const finalIssueRow = Math.max(issuesStartRow, currentRow - 1);
    ws.autoFilter = { from: `A${issuesHeaderRow}`, to: `F${finalIssueRow}` };
    ws.columns = [
      { width: 18 },
      { width: 28 },
      { width: 20 },
      { width: 32 },
      { width: 28 },
      { width: 60 },
      { width: 34 },
      { width: 26 },
    ];
  
  }
}
