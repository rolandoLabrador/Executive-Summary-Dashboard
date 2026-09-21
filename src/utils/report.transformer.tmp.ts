import { ReportConfig } from '../models/report.types';

export function excludedComponent(
  name: string,
  category: 'ADMIN' | 'RESERVE',
  config: ReportConfig,
): boolean {
  const upper = name.trim().toUpperCase();

  if (
    config.excludedComponentCodes.has(upper) ||
    upper.includes('DEALER') ||
    upper.includes('DLR') ||
    upper.includes('COMMISSION') ||
    upper.includes('COMM') ||
    upper.includes('F&I') ||
    upper.includes('PACK')
  ) {
    return true;
  }

  if (category === 'RESERVE') {
    if (
      upper.includes('PREMIUM TAX') ||
      upper === 'CEEDINGFEE' ||
      upper === 'CEDINGFEE' ||
      upper.includes('CEEDING') ||
      upper.includes('CEDING') ||
      upper === 'ADMIN'
    ) {
      return true;
    }
  }

  return false;
}

