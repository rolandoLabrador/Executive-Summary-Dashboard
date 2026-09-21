import { ReportConfig } from './src/models/report.types';

function excludedComponent(
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
      ['CLIPFEE', 'PREMIUMTAX', 'CEDINGFEE', 'ADMIN'].includes(upper) ||
      upper.includes('PREMIUM TAX') ||
      upper.includes('CEEDING') ||
      upper.includes('CEDING') ||
      upper.includes('CLIP FEE')
    ) {
      return true;
    }
  }

  if (category === 'ADMIN') {
    if (['ROADSIDEADMIN', 'LOANPMT'].includes(upper)) {
      return true;
    }
  }

  return false;
}

const config = { excludedComponentCodes: new Set<string>() } as unknown as ReportConfig;

console.log('CLIPFEE in RESERVE:', excludedComponent('CLIPFEE', 'RESERVE', config));
console.log('PREMIUMTAX in RESERVE:', excludedComponent('PREMIUMTAX', 'RESERVE', config));
console.log('ROADSIDEADMIN in ADMIN:', excludedComponent('ROADSIDEADMIN', 'ADMIN', config));
console.log('BASERESERVE in RESERVE:', excludedComponent('BASERESERVE', 'RESERVE', config));

