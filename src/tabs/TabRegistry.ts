import { type IDashboardTab } from './IDashboardTab';

export class TabRegistry {
  private availableTabs = new Map<string, IDashboardTab>();

  public register(tab: IDashboardTab): void {
    this.availableTabs.set(tab.id, tab);
  }

  public getTab(id: string): IDashboardTab | undefined {
    return this.availableTabs.get(id);
  }
}
