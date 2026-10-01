import { MongoClient } from 'mongodb';
import { loadMongoUri } from '../../src/config';
import * as dotenv from 'dotenv';
dotenv.config();

function getAmount(container: any, category: string, sub: string): number {
  if (!container || !container[category] || !container[category][sub]) return 0;
  let val = container[category][sub];
  if (typeof val === 'object' && val !== null) {
    if (val.value !== undefined) return Number(val.value); 
    val = val.toString();
  }
  return Number(val) || 0;
}

async function main() {
  const args = process.argv.slice(2);
  const dealerNumber = args.length > 0 ? args[0] : 'MS370';
  console.log(`\n======================================================`);
  console.log(`   MAXIMUM RECONCILIATION BREAKDOWN FOR ${dealerNumber}`);
  console.log(`======================================================\n`);

  const uri = loadMongoUri();
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');
  const cancelDb = client.db('CancelDataDB');
  
  const cutoffDate = new Date('2026-08-31T23:59:59.999Z');
  
  const contracts = await db.collection('ContractData').find({
    $or: [{ 'metadata.DealerNumber': dealerNumber }]
  }).toArray();
  
  const cancellations = await cancelDb.collection('CancelData').find({
    $or: [{ 'metadata.DealerNumber': dealerNumber }]
  }).toArray();

  const RESERVE_COMPONENTS = new Set<string>();
  const ADMIN_COMPONENTS = new Set<string>();

  const extractKeys = (doc: any, cat: string, set: Set<string>) => {
    if (doc && doc[cat]) {
      Object.keys(doc[cat]).forEach(k => set.add(k));
    }
  };

  contracts.forEach(d => {
    extractKeys(d.WrittenAmount, 'RESERVE', RESERVE_COMPONENTS);
    extractKeys(d.WrittenAmount, 'ADMIN', ADMIN_COMPONENTS);
  });
  cancellations.forEach(d => {
    extractKeys(d.CancelledAmount, 'RESERVE', RESERVE_COMPONENTS);
    extractKeys(d.CancelledAmount, 'ADMIN', ADMIN_COMPONENTS);
  });

  const results: Record<string, any> = {};

  // Initialize
  [...RESERVE_COMPONENTS].forEach(c => {
    results[`RESERVE.${c}`] = { Written: 0, Cancelled: 0, Net: 0, Voided_Drops: 0, MissingDate_Drops: 0, Future_Drops: 0 };
  });
  [...ADMIN_COMPONENTS].forEach(c => {
    results[`ADMIN.${c}`] = { Written: 0, Cancelled: 0, Net: 0, Voided_Drops: 0, MissingDate_Drops: 0, Future_Drops: 0 };
  });
  results[`RESERVE.ADMIN`] = { Written: 0, Cancelled: 0, Net: 0, Voided_Drops: 0, MissingDate_Drops: 0, Future_Drops: 0 };

  for (const doc of contracts) {
    const wa = doc.WrittenAmount;
    const actDate = doc.metadata.ActivationDate;
    const hasDate = actDate && actDate.toString().length > 0;
    const isFuture = hasDate && new Date(actDate) > cutoffDate;
    const isVoid = doc.metadata.ContractStatus === 'V';

    const processComponent = (cat: string, sub: string, bucket: string) => {
      const amt = getAmount(wa, cat, sub);
      if (amt === 0) return;
      
      results[bucket].Written += amt;
      results[bucket].Net += amt;
      
      if (isVoid) results[bucket].Voided_Drops += amt;
      else if (!hasDate) results[bucket].MissingDate_Drops += amt;
      else if (isFuture) results[bucket].Future_Drops += amt;
    };

    RESERVE_COMPONENTS.forEach(c => processComponent('RESERVE', c, `RESERVE.${c}`));
    ADMIN_COMPONENTS.forEach(c => processComponent('ADMIN', c, `ADMIN.${c}`));
    processComponent('RESERVE', 'ADMIN', 'RESERVE.ADMIN');
  }

  for (const doc of cancellations) {
    const ca = doc.CancelledAmount;
    const cancelDate = doc.metadata.CancelBillDate;
    const hasDate = cancelDate && cancelDate.toString().length > 0;
    const isFuture = hasDate && new Date(cancelDate) > cutoffDate;
    const isVoid = doc.metadata.ContractStatus === 'V';

    const processComponent = (cat: string, sub: string, bucket: string) => {
      const amt = getAmount(ca, cat, sub);
      if (amt === 0) return;
      
      // In MongoDB, cancelled amounts might be positive or negative depending on upstream.
      // We force them to be a positive deduction amount for the math
      const deduction = Math.abs(amt);
      
      results[bucket].Cancelled -= deduction;
      results[bucket].Net -= deduction;
      
      // Note: If a cancellation is voided/missing date, the engine ignores the refund, 
      // which means the dashboard keeps the money (so it's technically a positive variance drop)
      if (isVoid) results[bucket].Voided_Drops -= deduction;
      else if (!hasDate) results[bucket].MissingDate_Drops -= deduction;
      else if (isFuture) results[bucket].Future_Drops -= deduction;
    };

    RESERVE_COMPONENTS.forEach(c => processComponent('RESERVE', c, `RESERVE.${c}`));
    ADMIN_COMPONENTS.forEach(c => processComponent('ADMIN', c, `ADMIN.${c}`));
    processComponent('RESERVE', 'ADMIN', 'RESERVE.ADMIN');
  }

  console.log(`--- RAW EXCEL PIVOT (Before Dashboard Exclusions) ---`);
  console.table(
    Object.keys(results).reduce((acc, key) => {
      acc[key] = {
        Written: results[key].Written.toFixed(2),
        Cancelled: results[key].Cancelled.toFixed(2),
        'Raw Net (Excel)': results[key].Net.toFixed(2),
      };
      return acc;
    }, {} as any)
  );

  console.log(`\n--- DASHBOARD EXCLUSIONS (Why the dashboard is slightly lower) ---`);
  console.table(
    Object.keys(results).reduce((acc, key) => {
      acc[key] = {
        'Raw Net': results[key].Net.toFixed(2),
        '- Voids': results[key].Voided_Drops.toFixed(2),
        '- Missing Dates': results[key].MissingDate_Drops.toFixed(2),
        '- Sept 2026': results[key].Future_Drops.toFixed(2),
        '= True Dashboard KPI': (results[key].Net - results[key].Voided_Drops - results[key].MissingDate_Drops - results[key].Future_Drops).toFixed(2),
      };
      return acc;
    }, {} as any)
  );

  let manualReserve = 0;
  let trueDashboardReserve = 0;
  RESERVE_COMPONENTS.forEach(c => {
    manualReserve += results[`RESERVE.${c}`].Net;
    trueDashboardReserve += (results[`RESERVE.${c}`].Net - results[`RESERVE.${c}`].Voided_Drops - results[`RESERVE.${c}`].MissingDate_Drops - results[`RESERVE.${c}`].Future_Drops);
  });

  console.log(`\nTOTAL RESERVE (Manual Excel): $${manualReserve.toFixed(2)}`);
  console.log(`TOTAL RESERVE (True Dashboard): $${trueDashboardReserve.toFixed(2)}`);
  console.log(`Difference (Purged Garbage Data): $${(manualReserve - trueDashboardReserve).toFixed(2)}`);

  await client.close();
}

main().catch(console.error);
