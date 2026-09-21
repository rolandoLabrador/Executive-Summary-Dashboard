import { MongoService } from './src/services/mongo.service';
import { loadMongoUri, loadMongoSourceConfig, loadReportConfig } from './src/config';
import { DataRepository } from './src/services/data.repository';
import { ReportTransformer } from './src/utils/report.transformer';
import * as fs from 'fs';
import { parse } from 'csv-parse/sync';
// npx ts-node debug-memory.ts

async function run() {
  const uri = loadMongoUri();
  const mongo = new MongoService(uri);
  await mongo.connect();
  const repo = new DataRepository(mongo, loadMongoSourceConfig());

  const [contracts, cancellations, claims, audits] = await Promise.all([
    repo.getContracts(),
    repo.getCancellations(),
    repo.getClaims(),
    repo.getLatestReconciliationAudits(),
  ]);

  // If you run this script in a different month, you can specify the target cutoff
  // date like: loadReportConfig(['--as-of', '2026-08-31'])
  const config = loadReportConfig([]);
  const transformer = new ReportTransformer(config);
  const report = transformer.transform(contracts, cancellations, claims, audits);

  // Adjust the year and month to match the claims you want to inspect
  // Note: getMonth() is 0-indexed (7 = August)
  const targetClaims = report.claims.filter(
    (c) => c.activityDate.getFullYear() === 2026 && c.activityDate.getMonth() === 7,
  );

  const dbAmounts = new Map<string, number>();
  for (const c of targetClaims) {
    dbAmounts.set(c.claimNumber, (dbAmounts.get(c.claimNumber) || 0) + c.paidAmount);
  }

  // Make sure this points to the exact export file you want to compare against
  const csvContent = fs.readFileSync('./output/ClaimExport (16).csv', 'utf8');
  const csvRecords = parse(csvContent, { columns: true, skip_empty_lines: true });

  const csvAmounts = new Map<string, number>();
  for (const r of csvRecords) {
    const row = r as any;
    const row = r as Record<string, string>;
    const claimNo = row['Claim Number'].trim().toUpperCase().replace(/\s+/g, '');
    const paid = parseFloat(row['Total Paid Amount']);
    if (!isNaN(paid)) {
      csvAmounts.set(claimNo, (csvAmounts.get(claimNo) || 0) + paid);
    }
  }

  let discrepancy = 0;
  for (const [claimNo, csvPaid] of csvAmounts.entries()) {
    const dbPaid = dbAmounts.get(claimNo) || 0;

    // Check for any significant floating point differences
    if (Math.abs(csvPaid - dbPaid) > 0.01) {
      console.log(
        `Discrepancy for ${claimNo}: CSV = ${csvPaid}, DB = ${dbPaid}, Diff = ${csvPaid - dbPaid}`,
      );
      discrepancy += csvPaid - dbPaid;
    }
  }

  console.log('\nTotal difference (CSV sum - DB sum for these claims):', discrepancy);
  await mongo.close();
}

run().catch(console.error);
