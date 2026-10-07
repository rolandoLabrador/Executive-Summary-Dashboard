import * as ExcelJS from 'exceljs';
import { type ReportModel, type ReportConfig } from '../models/report.types';
import { TabRegistry } from '../tabs/TabRegistry';
import rawConfig from '../tabs/dashboard.config.json';
export type TabConfig = boolean | { enabled?: boolean; tables?: Record<string, boolean> };
interface DashboardConfig {
  activeTabs: Record<string, TabConfig>;
}
const dashboardConfig: DashboardConfig = rawConfig as DashboardConfig;

import { ExecutiveDashboardTab } from '../tabs/executive_dashboard.tab';
import { DealerDashboardTab } from '../tabs/dealer_dashboard.tab';
import { AgentDashboardTab } from '../tabs/agent_dashboard.tab';
import { ProductDashboardTab } from '../tabs/product_dashboard.tab';
import { LossCodeDashboardTab } from '../tabs/loss_code_dashboard.tab';
import { MonthlyTrendsTab } from '../tabs/monthly_trends.tab';
import { ContractActivityTab } from '../tabs/contract_activity.tab';
import { ClaimActivityTab } from '../tabs/claim_activity.tab';
import { DataQualityTab } from '../tabs/data_quality.tab';
import { DebugMathTab } from '../tabs/debug_math.tab';
import { DefinitionsTab } from '../tabs/definitions.tab';

export class ExcelService {
  constructor(private readonly config: ReportConfig) {}

  public async generateFullReport(model: ReportModel): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Executive Summary System';
    workbook.created = new Date();
    workbook.calcProperties.fullCalcOnLoad = true;

    const registry = new TabRegistry();
    registry.register(new ExecutiveDashboardTab());
    registry.register(new DealerDashboardTab());
    registry.register(new AgentDashboardTab());
    registry.register(new ProductDashboardTab());
    registry.register(new LossCodeDashboardTab());
    registry.register(new MonthlyTrendsTab());
    registry.register(new ContractActivityTab());
    registry.register(new ClaimActivityTab());
    registry.register(new DataQualityTab());
    registry.register(new DebugMathTab());
    registry.register(new DefinitionsTab());

    // Check debug mode
    const debugDealers = process.env.DEBUG_DEALER_NAME
      ? process.env.DEBUG_DEALER_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const debugAgents = process.env.DEBUG_AGENT_NAME
      ? process.env.DEBUG_AGENT_NAME.split(',').map((s) => s.trim().toLowerCase())
      : [];
    const hasDebug = debugDealers.length > 0 || debugAgents.length > 0;

    // Render tabs based on config and debug mode
    const activeTabs: Record<string, TabConfig> = dashboardConfig.activeTabs as Record<string, TabConfig>;

    const tabsToRender = [
      { key: 'tab_executive_summary', id: 'tab_executive_summary' },
      { key: 'tab_agent_dashboard', id: 'tab_agent_dashboard' },
      { key: 'tab_dealer_dashboard', id: 'tab_dealer_dashboard' },
      { key: 'tab_product_dashboard', id: 'tab_product_dashboard' },
      { key: 'tab_loss_code', id: 'tab_loss_code' },
      { key: 'tab_monthly', id: 'tab_monthly' },
      { key: 'tab_contract_detail', id: 'tab_contract_detail' },
      { key: 'tab_claim_detail', id: 'tab_claim_detail' },
      { key: 'tab_data_quality', id: 'tab_data_quality' },
      { key: 'tab_debug_math', id: 'tab_debug_math' },
      { key: 'tab_definitions', id: 'tab_definitions' }
    ];

    const forceDebugTabs = ['tab_loss_code', 'tab_monthly', 'tab_contract_detail', 'tab_claim_detail', 'tab_data_quality', 'tab_debug_math', 'tab_definitions'];

    for (const tabInfo of tabsToRender) {
      const tabConfig = activeTabs[tabInfo.key];
      let shouldRender = false;

      if (tabConfig === true || (typeof tabConfig === 'object' && tabConfig !== null && 'enabled' in tabConfig && tabConfig.enabled === true)) {
        shouldRender = true;
      }

      // If in debug mode, override and force these detail tabs to render
      if (hasDebug && forceDebugTabs.includes(tabInfo.key)) {
        shouldRender = true;
      }

      if (shouldRender) {
        registry.getTab(tabInfo.id)?.render(workbook, model, this.config, tabConfig);
      }
    }

    return Buffer.from(await workbook.xlsx.writeBuffer());
  }
}

