import { BaseThreeTierTab } from './base_three_tier.tab';
import { type ReportModel, type ReportConfig, type DimensionMetric } from '../models/report.types';

export class DealerDashboardTab extends BaseThreeTierTab {
  public readonly id = 'tab_dealer_dashboard';
  protected readonly name = 'Dealer Dashboard';
  protected readonly heading = 'DEALER PERFORMANCE DASHBOARD';
  protected readonly isDealer = true;
  protected readonly isAgent = false;

  protected getItdRows(model: ReportModel): DimensionMetric[] {
    return model.itdDealers;
  }

  protected getRollingRows(model: ReportModel): DimensionMetric[] {
    return model.dealers;
  }

  protected getYtdRows(model: ReportModel): DimensionMetric[] {
    return model.ytdDealers;
  }

  protected getLimit(config: ReportConfig): number | undefined {
    return config.topDealerCount;
  }
}
