tabs = [
    'src/tabs/claim_activity.tab.ts',
    'src/tabs/contract_activity.tab.ts',
    'src/tabs/data_quality.tab.ts',
    'src/tabs/debug_math.tab.ts',
    'src/tabs/definitions.tab.ts',
    'src/tabs/loss_code_dashboard.tab.ts',
    'src/tabs/monthly_trends.tab.ts'
]

tab_ids = [
    'tab_claim_detail',
    'tab_contract_detail',
    'tab_data_quality',
    'tab_debug_math',
    'tab_definitions',
    'tab_loss_code',
    'tab_monthly'
]

for tab, tab_id in zip(tabs, tab_ids):
    with open(tab, 'r') as f:
        code = f.read()
    
    # Restore the missing imports
    restored_imports = """import type * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { type IDashboardTab } from './IDashboardTab';
"""
    code = code.replace("import type * as ExcelJS from 'exceljs';\n", restored_imports)

    # Add readonly id
    code = code.replace("constructor() {}", f"readonly id = '{tab_id}';\n\n  constructor() {{}}")
    
    with open(tab, 'w') as f:
        f.write(code)

print('Restored missing imports and added ids!')
