import { BaseThreeTierTab } from './base_three_tier.tab';
import { type ReportModel, type ReportConfig, type DimensionMetric } from '../models/report.types';

export class AgentDashboardTab extends BaseThreeTierTab {
  public readonly id = 'tab_agent_dashboard';
  protected readonly name = 'Agent Dashboard';
  protected readonly heading = 'AGENT PERFORMANCE DASHBOARD';
  protected readonly isDealer = false;
  protected readonly isAgent = true;

  protected getItdRows(model: ReportModel): DimensionMetric[] {
    return model.itdAgents;
  }

  protected getRollingRows(model: ReportModel): DimensionMetric[] {
    return model.agents;
  }

  protected getYtdRows(model: ReportModel): DimensionMetric[] {
    return model.ytdAgents;
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  protected getLimit(_config: ReportConfig): number | undefined {
    return undefined; // No limit for agents
  }
}
