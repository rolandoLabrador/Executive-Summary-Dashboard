const fs = require('fs');

let excelCode = fs.readFileSync('src/services/excel.service.ts', 'utf8');

function extractMethod(methodName, className, fileName, argsSig, imports) {
    const startPattern = `private ${methodName}(${argsSig}): void {`;
    let startIdx = excelCode.indexOf(startPattern);
    
    // Fallback without void
    if (startIdx === -1) {
        startIdx = excelCode.indexOf(`private ${methodName}(${argsSig}) {`);
    }

    if (startIdx === -1) {
        console.log(`Could not find ${methodName}`);
        return;
    }

    let endIdx = -1;
    let braceCount = 0;
    let foundFirstBrace = false;

    for (let i = startIdx; i < excelCode.length; i++) {
        if (excelCode[i] === '{') {
            braceCount++;
            foundFirstBrace = true;
        } else if (excelCode[i] === '}') {
            braceCount--;
        }

        if (foundFirstBrace && braceCount === 0) {
            endIdx = i;
            break;
        }
    }

    if (endIdx === -1) {
        console.log(`Could not find end of ${methodName}`);
        return;
    }

    const methodBody = excelCode.substring(startIdx, endIdx + 1)
        .replace(new RegExp(`private ${methodName}\\([\\s\\S]*?\\)(: void)? \\{`), '');
    
    let renderSig = `workbook: ExcelJS.Workbook, model: ReportModel, config: ReportConfig, tabConfig: any`;
    if (argsSig.indexOf('model') === -1) {
        renderSig = `workbook: ExcelJS.Workbook, _model: ReportModel, _config: ReportConfig, _tabConfig: any`;
    } else {
        renderSig = `workbook: ExcelJS.Workbook, model: ReportModel, _config: ReportConfig, _tabConfig: any`;
    }

    const tabCode = `${imports}
import { IDashboardTab } from './IDashboardTab';
import { configureWorksheet, title, styleHeader, formatDate, formatDateRange, COLORS, MONEY, INTEGER, formatPercent } from '../utils/excel.utils';

export class ${className} implements IDashboardTab {
  constructor() {}

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  render(${renderSig}): void {
    ${methodBody.substring(0, methodBody.length - 1)}
  }
}
`;

    fs.writeFileSync(`src/tabs/${fileName}`, tabCode);
    console.log(`Saved ${className} to ${fileName}`);
    
    // Remove from excel.service.ts
    excelCode = excelCode.substring(0, startIdx) + excelCode.substring(endIdx + 1);
}

const standardImports = `import * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';`;

extractMethod('buildMonthly', 'MonthlyTrendsTab', 'monthly_trends.tab.ts', 'workbook: ExcelJS.Workbook, model: ReportModel', standardImports);
extractMethod('buildContractDetail', 'ContractActivityTab', 'contract_activity.tab.ts', 'workbook: ExcelJS.Workbook, model: ReportModel', standardImports);
extractMethod('buildClaimDetail', 'ClaimActivityTab', 'claim_activity.tab.ts', 'workbook: ExcelJS.Workbook, model: ReportModel', standardImports);
extractMethod('buildDataQuality', 'DataQualityTab', 'data_quality.tab.ts', 'workbook: ExcelJS.Workbook, model: ReportModel', standardImports);
extractMethod('buildDebugMath', 'DebugMathTab', 'debug_math.tab.ts', 'workbook: ExcelJS.Workbook, model: ReportModel', standardImports);
extractMethod('buildDefinitions', 'DefinitionsTab', 'definitions.tab.ts', 'workbook: ExcelJS.Workbook', standardImports);

// Also manually remove the hardcoded buildLossCodeDashboard since it's already extracted
const lossStart = excelCode.indexOf('private buildLossCodeDashboard(workbook: ExcelJS.Workbook, model: ReportModel): void {');
if (lossStart > -1) {
    let lossEnd = -1;
    let braceCount = 0;
    let foundFirstBrace = false;
    for (let i = lossStart; i < excelCode.length; i++) {
        if (excelCode[i] === '{') {
            braceCount++;
            foundFirstBrace = true;
        } else if (excelCode[i] === '}') {
            braceCount--;
        }
        if (foundFirstBrace && braceCount === 0) {
            lossEnd = i;
            break;
        }
    }
    if (lossEnd > -1) {
        excelCode = excelCode.substring(0, lossStart) + excelCode.substring(lossEnd + 1);
    }
}

// Ensure the new registry block is added ONLY ONCE
const newRegistryBlock = `
      registry.register(new LossCodeDashboardTab());
      registry.register(new MonthlyTrendsTab());
      registry.register(new ContractActivityTab());
      registry.register(new ClaimActivityTab());
      registry.register(new DataQualityTab());
      registry.register(new DebugMathTab());
      registry.register(new DefinitionsTab());

      const lossCodeConfig: any = dashboardConfig.activeTabs['tab_loss_code_dashboard'];
      if (lossCodeConfig === true || lossCodeConfig?.enabled === true) {
        registry.getTab('tab_loss_code_dashboard')?.render(workbook, model, this.config, lossCodeConfig);
      }

      const monthlyConfig: any = dashboardConfig.activeTabs['tab_monthly_trends'];
      if (monthlyConfig === true || monthlyConfig?.enabled === true) {
        registry.getTab('tab_monthly_trends')?.render(workbook, model, this.config, monthlyConfig);
      }

      const contractConfig: any = dashboardConfig.activeTabs['tab_contract_activity'];
      if (contractConfig === true || contractConfig?.enabled === true) {
        registry.getTab('tab_contract_activity')?.render(workbook, model, this.config, contractConfig);
      }

      const claimConfig: any = dashboardConfig.activeTabs['tab_claim_activity'];
      if (claimConfig === true || claimConfig?.enabled === true) {
        registry.getTab('tab_claim_activity')?.render(workbook, model, this.config, claimConfig);
      }

      const dqConfig: any = dashboardConfig.activeTabs['tab_data_quality'];
      if (dqConfig === true || dqConfig?.enabled === true) {
        registry.getTab('tab_data_quality')?.render(workbook, model, this.config, dqConfig);
      }

      const debugConfig: any = dashboardConfig.activeTabs['tab_debug_math'];
      if (debugConfig === true || debugConfig?.enabled === true) {
        registry.getTab('tab_debug_math')?.render(workbook, model, this.config, debugConfig);
      }

      const defConfig: any = dashboardConfig.activeTabs['tab_definitions'];
      if (defConfig === true || defConfig?.enabled === true) {
        registry.getTab('tab_definitions')?.render(workbook, model, this.config, defConfig);
      }
`;

if (excelCode.indexOf('registry.register(new LossCodeDashboardTab());') === -1) {
    excelCode = excelCode.replace("registry.getTab('tab_product_dashboard')?.render(workbook, model, this.config, productConfig);\n      }", "registry.getTab('tab_product_dashboard')?.render(workbook, model, this.config, productConfig);\n      }" + newRegistryBlock);
}

const newImportsBlock = `import { LossCodeDashboardTab } from '../tabs/loss_code_dashboard.tab';
import { MonthlyTrendsTab } from '../tabs/monthly_trends.tab';
import { ContractActivityTab } from '../tabs/contract_activity.tab';
import { ClaimActivityTab } from '../tabs/claim_activity.tab';
import { DataQualityTab } from '../tabs/data_quality.tab';
import { DebugMathTab } from '../tabs/debug_math.tab';
import { DefinitionsTab } from '../tabs/definitions.tab';
`;

if (excelCode.indexOf('import { LossCodeDashboardTab }') === -1) {
    excelCode = excelCode.replace("import { ProductDashboardTab } from '../tabs/product_dashboard.tab';", "import { ProductDashboardTab } from '../tabs/product_dashboard.tab';\n" + newImportsBlock);
}

// Remove hardcoded calls
excelCode = excelCode.replace(/this\.buildLossCodeDashboard\(workbook, model\);\r?\n/g, '');
excelCode = excelCode.replace(/this\.buildMonthly\(workbook, model\);\r?\n/g, '');
excelCode = excelCode.replace(/this\.buildContractDetail\(workbook, model\);\r?\n/g, '');
excelCode = excelCode.replace(/this\.buildClaimDetail\(workbook, model\);\r?\n/g, '');
excelCode = excelCode.replace(/this\.buildDataQuality\(workbook, model\);\r?\n/g, '');
excelCode = excelCode.replace(/this\.buildDebugMath\(workbook, model\);\r?\n/g, '');
excelCode = excelCode.replace(/this\.buildDefinitions\(workbook\);\r?\n/g, '');

fs.writeFileSync('src/services/excel.service.ts', excelCode);
console.log('Successfully extracted ALL tabs!');
