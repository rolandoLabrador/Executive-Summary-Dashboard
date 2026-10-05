const fs = require('fs');

let content = fs.readFileSync('src/services/excel.service.ts', 'utf8');

// 1. Remove buildThreeTierDimensionSheet completely
let start = content.indexOf('function buildThreeTierDimensionSheet(');
if (start > -1) {
    let end = content.indexOf('];\n}', start);
    if (end === -1) end = content.indexOf('];\r\n}', start);
    if (end > -1) content = content.substring(0, start) + content.substring(end + 5);
}

// 2. Remove this.buildExecutive
start = content.indexOf('private buildExecutive(');
if (start > -1) {
    let end = content.indexOf('ws.columns = [', start);
    if (end > -1) {
        end = content.indexOf('];\n  }', end);
        if (end === -1) end = content.indexOf('];\r\n  }', end);
        if (end > -1) content = content.substring(0, start) + content.substring(end + 7);
    }
}

// 3. Remove styleHeaderRange and styleHeaderCell
start = content.indexOf('function styleHeaderRange(');
if (start > -1) {
    let end = content.indexOf('}', content.indexOf('}', start) + 1) + 1;
    content = content.substring(0, start) + content.substring(end);
}

start = content.indexOf('function styleHeaderCell(');
if (start > -1) {
    let end = content.indexOf('}', start) + 1;
    content = content.substring(0, start) + content.substring(end);
}

// 4. Add imports
const imports = `import rawConfig from '../tabs/dashboard.config.json';
const dashboardConfig: any = rawConfig;
import { TabRegistry } from '../tabs/TabRegistry';
import { ExecutiveDashboardTab } from '../tabs/executive_dashboard.tab';
import { DealerDashboardTab } from '../tabs/dealer_dashboard.tab';
import { AgentDashboardTab } from '../tabs/agent_dashboard.tab';
import { ProductDashboardTab } from '../tabs/product_dashboard.tab';
`;
content = imports + content;

// 5. Replace `this.buildExecutive(workbook, model);` and `buildThreeTier...` calls with registry logic
const registryLogic = `
      const registry = new TabRegistry();
      registry.register(new ExecutiveDashboardTab());
      registry.register(new DealerDashboardTab());
      registry.register(new AgentDashboardTab());
      registry.register(new ProductDashboardTab());

      const execConfig: any = dashboardConfig.activeTabs['tab_executive_summary'];
      if (execConfig === true || execConfig?.enabled === true) {
        registry.getTab('tab_executive_summary')?.render(workbook, model, this.config, execConfig);
      }

      const dealerConfig: any = dashboardConfig.activeTabs['tab_dealer_dashboard'];
      if (dealerConfig === true || dealerConfig?.enabled === true) {
        registry.getTab('tab_dealer_dashboard')?.render(workbook, model, this.config, dealerConfig);
      }
      
      const agentConfig: any = dashboardConfig.activeTabs['tab_agent_dashboard'];
      if (agentConfig === true || agentConfig?.enabled === true) {
        registry.getTab('tab_agent_dashboard')?.render(workbook, model, this.config, agentConfig);
      }
      
      const productConfig: any = dashboardConfig.activeTabs['tab_product_dashboard'];
      if (productConfig === true || productConfig?.enabled === true) {
        registry.getTab('tab_product_dashboard')?.render(workbook, model, this.config, productConfig);
      }
`;

const buildExecutiveCall = 'this.buildExecutive(workbook, model);';
content = content.replace(buildExecutiveCall, registryLogic);

// Remove the remaining buildThreeTier calls
const regex = /buildThreeTierDimensionSheet\([\s\S]*?,\s*model\.yearToDate\.currentEnd[\s\S]*?\);/g;
content = content.replace(regex, '');

fs.writeFileSync('src/services/excel.service.ts', content);
console.log('Fully patched excel.service.ts!');
