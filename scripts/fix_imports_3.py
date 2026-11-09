import os

tabs = [
    'src/tabs/claim_activity.tab.ts',
    'src/tabs/contract_activity.tab.ts',
    'src/tabs/data_quality.tab.ts',
    'src/tabs/debug_math.tab.ts',
    'src/tabs/definitions.tab.ts',
    'src/tabs/loss_code_dashboard.tab.ts',
    'src/tabs/monthly_trends.tab.ts'
]

for tab in tabs:
    with open(tab, 'r') as f:
        lines = f.readlines()
    
    # Extract only lines starting from "export class"
    out_lines = []
    in_class = False
    for line in lines:
        if line.startswith('export class'):
            in_class = True
        if in_class:
            out_lines.append(line)
            
    # Write the exact imports
    header = """import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
import { title, styleHeader, styleHeaderRange, COLORS, MONEY, PERCENT, INTEGER, dataBarRule, configureWorksheet, visibleLossCodeRows, lossCodeChartSegments, formatDateRange, formatDate } from '../utils/excel.utils';
import { renderPieChartPng } from '../utils/pie-chart.renderer';

"""
    with open(tab, 'w') as f:
        f.write(header + ''.join(out_lines))
