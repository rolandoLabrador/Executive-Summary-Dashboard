const fs = require('fs');

const tabCode = `import * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { IDashboardTab } from './IDashboardTab';
import { configureWorksheet, title, styleHeader, formatDate, formatDateRange, COLORS, MONEY, INTEGER, formatPercent } from '../utils/excel.utils';

export class LossCodeDashboardTab implements IDashboardTab {
  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, tabConfig: any): void {
` + fs.readFileSync('temp_loss_code.txt', 'utf8').replace('private buildLossCodeDashboard(workbook: ExcelJS.Workbook, model: ReportModel): void {', '').replace(/private buildMonthly[\s\S]*/, '') + `
}
`;

fs.writeFileSync('src/tabs/loss_code_dashboard.tab.ts', tabCode);
console.log('Saved to src/tabs/loss_code_dashboard.tab.ts');
