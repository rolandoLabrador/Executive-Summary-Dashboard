const fs = require('fs');
let content = fs.readFileSync('src/services/excel.service.ts', 'utf8');

// Remove styleHeaderRange
let start = content.indexOf('function styleHeaderRange(');
if (start > -1) {
    let end = content.indexOf('}', content.indexOf('}', start) + 1) + 1;
    content = content.substring(0, start) + content.substring(end);
}

// Remove styleHeaderCell
start = content.indexOf('function styleHeaderCell(');
if (start > -1) {
    let end = content.indexOf('}', start) + 1;
    content = content.substring(0, start) + content.substring(end);
}

// Remove buildThreeTierDimensionSheet
start = content.indexOf('function buildThreeTierDimensionSheet(');
if (start > -1) {
    let end = content.indexOf('];\n}', start);
    if (end === -1) end = content.indexOf('];\r\n}', start);
    if (end > -1) content = content.substring(0, start) + content.substring(end + 5);
}

// Add imports
const imports = `import { DealerDashboardTab } from '../tabs/dealer_dashboard.tab';
import { AgentDashboardTab } from '../tabs/agent_dashboard.tab';
import { ProductDashboardTab } from '../tabs/product_dashboard.tab';
`;
content = content.replace("import { ExecutiveDashboardTab } from '../tabs/executive_dashboard.tab';", "import { ExecutiveDashboardTab } from '../tabs/executive_dashboard.tab';\n" + imports);

// Add to registry
const registryCode = `    registry.register(new DealerDashboardTab());
    registry.register(new AgentDashboardTab());
    registry.register(new ProductDashboardTab());
`;
content = content.replace('    registry.register(new ExecutiveDashboardTab());', '    registry.register(new ExecutiveDashboardTab());\n' + registryCode);

// Add execution blocks
const execBlocks = `
    const dealerConfig: any = dashboardConfig.activeTabs['tab_dealer_dashboard'];
    if (dealerConfig === true || dealerConfig?.enabled === true) {
      const dealerTab = registry.getTab('tab_dealer_dashboard');
      if (dealerTab) dealerTab.render(workbook, model, this.config, dealerConfig);
    }
    
    const agentConfig: any = dashboardConfig.activeTabs['tab_agent_dashboard'];
    if (agentConfig === true || agentConfig?.enabled === true) {
      const agentTab = registry.getTab('tab_agent_dashboard');
      if (agentTab) agentTab.render(workbook, model, this.config, agentConfig);
    }
    
    const productConfig: any = dashboardConfig.activeTabs['tab_product_dashboard'];
    if (productConfig === true || productConfig?.enabled === true) {
      const productTab = registry.getTab('tab_product_dashboard');
      if (productTab) productTab.render(workbook, model, this.config, productConfig);
    }
`;
content = content.replace('      if (execTab) execTab.render(workbook, model, this.config, execConfig);\n    }', '      if (execTab) execTab.render(workbook, model, this.config, execConfig);\n    }\n' + execBlocks);
content = content.replace('      if (execTab) execTab.render(workbook, model, this.config, execConfig);\r\n    }', '      if (execTab) execTab.render(workbook, model, this.config, execConfig);\r\n    }\r\n' + execBlocks);
content = content.replace('      if (execTab) execTab.render(workbook, model, this.config);\n    }', '      if (execTab) execTab.render(workbook, model, this.config, execConfig);\n    }\n' + execBlocks);
content = content.replace('      if (execTab) execTab.render(workbook, model, this.config);\r\n    }', '      if (execTab) execTab.render(workbook, model, this.config, execConfig);\r\n    }\r\n' + execBlocks);


// Remove legacy method calls
const legacyStart = content.indexOf('    // Legacy method calls (to be migrated next to individual classes)');
const legacyEnd = content.indexOf('    this.buildLossCodeDashboard(workbook, model);');
if (legacyStart > -1 && legacyEnd > -1) {
    content = content.substring(0, legacyStart) + content.substring(legacyEnd);
}

fs.writeFileSync('src/services/excel.service.ts', content);
console.log('Successfully patched excel.service.ts!');
