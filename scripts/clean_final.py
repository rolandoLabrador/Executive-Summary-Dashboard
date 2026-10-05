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

    # Clean the excel.utils import
    match = re.search(r'import\s+\{(.*?)\}\s+from\s+[\'\"]\.\./utils/excel.utils[\'\"];', code, re.DOTALL)
    if match:
        imports_str = match.group(1)
        imports_list = [i.strip() for i in imports_str.split(',')]
        
        new_imports = []
        for imp in imports_list:
            if not imp:
                continue
            # If it's used elsewhere (more than 1 time in the whole file, including the import statement)
            count = len(re.findall(r'\b' + imp + r'\b', code))
            if count > 1:
                new_imports.append(imp)
        
        new_import_line = 'import { ' + ', '.join(new_imports) + " } from '../utils/excel.utils';"
        code = code[:match.start()] + new_import_line + code[match.end():]
        
    # Clean the pie-chart.renderer import
    match2 = re.search(r'import\s+\{(.*?)\}\s+from\s+[\'\"]\.\./utils/pie-chart.renderer[\'\"];', code, re.DOTALL)
    if match2:
        imports_str = match2.group(1)
        imports_list = [i.strip() for i in imports_str.split(',')]
        
        new_imports = []
        for imp in imports_list:
            if not imp:
                continue
            count = len(re.findall(r'\b' + imp + r'\b', code))
            if count > 1:
                new_imports.append(imp)
        
        if len(new_imports) > 0:
            new_import_line = 'import { ' + ', '.join(new_imports) + " } from '../utils/pie-chart.renderer';"
            code = code[:match2.start()] + new_import_line + code[match2.end():]
        else:
            # remove entire line
            code = code[:match2.start()] + code[match2.end():]
            code = code.replace('\n\n\n', '\n\n')

    with open(tab, 'w') as f:
        f.write(code)

print('Cleaned unused imports')
