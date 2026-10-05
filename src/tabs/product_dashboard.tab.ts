import { BaseThreeTierTab } from './base_three_tier.tab';
import { type ReportModel, type ReportConfig, type DimensionMetric } from '../models/report.types';

export class ProductDashboardTab extends BaseThreeTierTab {
  public readonly id = 'tab_product_dashboard';
  protected readonly name = 'Product Dashboard';
  protected readonly heading = 'PRODUCT PERFORMANCE DASHBOARD';
  protected readonly isDealer = false;
  protected readonly isAgent = false;

  protected getItdRows(model: ReportModel): DimensionMetric[] {
    return model.itdProducts;
  }

  protected getRollingRows(model: ReportModel): DimensionMetric[] {
    return model.products;
  }

  protected getYtdRows(model: ReportModel): DimensionMetric[] {
    return model.ytdProducts;
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  protected getLimit(_config: ReportConfig): number | undefined {
    return undefined; // No limit for products
  }
}
