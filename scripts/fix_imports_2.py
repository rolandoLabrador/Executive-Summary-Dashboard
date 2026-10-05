import re

def fix_imports(tab_file):
    with open(tab_file, 'r') as f:
        code = f.read()

    replacement = """import { title, styleHeader, styleHeaderRange, COLORS, MONEY, PERCENT, INTEGER, dataBarRule, configureWorksheet, visibleLossCodeRows, lossCodeChartSegments, formatDateRange, formatDate } from '../utils/excel.utils';
import { renderPieChartPng } from '../utils/pie-chart.renderer';"""

    code = re.sub(
        r'import \{.*?\} from \'../utils/excel.utils\';',
        replacement,
        code,
        flags=re.DOTALL
    )

    with open(tab_file, 'w') as f:
        f.write(code)

fix_imports('src/tabs/loss_code_dashboard.tab.ts')
fix_imports('src/tabs/monthly_trends.tab.ts')
fix_imports('src/tabs/contract_activity.tab.ts')
fix_imports('src/tabs/claim_activity.tab.ts')
fix_imports('src/tabs/data_quality.tab.ts')
fix_imports('src/tabs/debug_math.tab.ts')
fix_imports('src/tabs/definitions.tab.ts')
