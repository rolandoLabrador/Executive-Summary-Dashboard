import type * as ExcelJS from 'exceljs';

export function buildAgentDashboard(workbook: ExcelJS.Workbook, data: unknown): void {
  void data;
  const ws = workbook.addWorksheet('Agent Dashboard', { views: [{ showGridLines: false }] });

  // TODO: Implement agent specific layout here
  ws.getCell('A1').value = 'Agent Performance Dashboard';
}
