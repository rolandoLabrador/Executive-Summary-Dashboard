import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';

export interface IDashboardTab {
  readonly id: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  render(workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, tabConfig?: any): void | Promise<void>;
}
