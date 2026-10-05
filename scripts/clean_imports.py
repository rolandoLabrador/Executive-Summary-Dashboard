import os
import re

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
        code = f.read()

    match = re.search(r'import\s+\{(.*?)\}\s+from\s+[\'\"]\.\./utils/excel.utils[\'\"];', code, re.DOTALL)
    
    if match:
        imports_str = match.group(1)
        imports_list = [i.strip() for i in imports_str.split(',')]
        
        new_imports = []
        for imp in imports_list:
            if not imp:
                continue
            count = len(re.findall(r'\b' + imp + r'\b', code))
            if count > 1:
                new_imports.append(imp)
        
        new_import_line = 'import { ' + ', '.join(new_imports) + " } from '../utils/excel.utils';"
        code = code[:match.start()] + new_import_line + code[match.end():]
        
    with open(tab, 'w') as f:
        f.write(code)

with open('src/tabs/debug_math.tab.ts', 'r') as f:
    code = f.read()
code = re.sub(r'let startRow = 8;', '// eslint-disable-next-line prefer-const, @typescript-eslint/no-unused-vars\n    let startRow = 8;', code)
with open('src/tabs/debug_math.tab.ts', 'w') as f:
    f.write(code)

print('Done')
