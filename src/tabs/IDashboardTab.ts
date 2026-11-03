import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';

export interface IDashboardTab {
  readonly id: string;
  render(
    workbook: ExcelJS.Workbook,
    model: ReportModel,
    config: ReportConfig,
    tabConfig?: unknown,
  ): void | Promise<void>;
}
