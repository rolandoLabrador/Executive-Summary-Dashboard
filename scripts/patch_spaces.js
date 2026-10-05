const fs = require('fs');
let content = fs.readFileSync('src/services/excel.service.ts', 'utf8');

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

content = content.replace(/    this\.buildExecutive\(workbook, model\);/, registryLogic);

// Remove the unused legacy dimension variables
const legacyVars = "import { DimensionMetric } from '../models/report.types';\n";
content = content.replace(legacyVars, '');

fs.writeFileSync('src/services/excel.service.ts', content);
console.log('Fixed spaces in patch!');
