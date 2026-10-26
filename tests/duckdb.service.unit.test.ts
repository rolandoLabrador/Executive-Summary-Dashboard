import * as fs from 'fs';
import * as path from 'path';
import { DuckDBService } from '../src/services/duckdb.service';
import {
  type NormalizedContractTransaction,
  type NormalizedClaim,
  type PipelineAuditRecord,
} from '../src/models/report.types';

async function runTest(): Promise<void> {
  console.log('🧪 Starting DuckDBService Comprehensive Verification Test...');
  const duckdb = new DuckDBService(':memory:');
  await duckdb.init();
  console.log('✅ DuckDB in-memory database initialized successfully.');

  const sampleContracts: NormalizedContractTransaction[] = [
    {
      sourceId: 'src_01',
      snapshotDate: new Date('2026-07-01T10:00:00Z'),
      contractNumber: 'CNT-1001',
      contractStatus: 'A',
      transactionType: 'NewBusiness',
      activityDate: new Date('2026-07-01T00:00:00Z'),
      effectiveDate: new Date('2026-07-01T00:00:00Z'),
      dealer: 'D100',
      dealerNumber: 'D100',
      dealerName: 'Precision Motors',
      agent: 'AG01',
      agentName: 'Apex Agency',
      product: 'VSC Gold',
      coverageCode: 'GOLD',
      company: 'OmniShield',
      riskEntity: 'Entity A',
      adminAmount: 100.0,
      reserveAmount: 400.0,
      grossIncomeAmount: 500.0,
      earnedReserveAmount: 0,
      earningSchedule: { 1: 0.05, 2: 0.10 },
      components: { ADMIN: { BASEADMIN: 100 }, RESERVE: { BASERESERVE: 400 } },
    },
    // Snapshot 2 of CNT-1001 with updated reserve ($450)
    {
      sourceId: 'src_02',
      snapshotDate: new Date('2026-07-05T15:00:00Z'),
      contractNumber: 'CNT-1001',
      contractStatus: 'A',
      transactionType: 'NewBusiness',
      activityDate: new Date('2026-07-01T00:00:00Z'),
      effectiveDate: new Date('2026-07-01T00:00:00Z'),
      dealer: 'D100',
      dealerNumber: 'D100',
      dealerName: 'Precision Motors',
      agent: 'AG01',
      agentName: 'Apex Agency',
      product: 'VSC Gold',
      coverageCode: 'GOLD',
      company: 'OmniShield',
      riskEntity: 'Entity A',
      adminAmount: 100.0,
      reserveAmount: 450.0,
      grossIncomeAmount: 550.0,
      earnedReserveAmount: 0,
      earningSchedule: { 1: 0.05, 2: 0.10 },
      components: { ADMIN: { BASEADMIN: 100 }, RESERVE: { BASERESERVE: 450 } },
    },
    {
      sourceId: 'src_03',
      snapshotDate: new Date('2026-07-02T10:00:00Z'),
      contractNumber: 'CNT-1002',
      contractStatus: 'A',
      transactionType: 'NewBusiness',
      activityDate: new Date('2026-07-02T00:00:00Z'),
      effectiveDate: new Date('2026-07-02T00:00:00Z'),
      dealer: 'D200',
      dealerNumber: 'D200',
      dealerName: 'Summit Auto',
      agent: 'AG02',
      agentName: 'Premier Agency',
      product: 'GAP Standard',
      coverageCode: 'GAP',
      company: 'OmniShield',
      riskEntity: 'Entity B',
      adminAmount: 75.0,
      reserveAmount: 225.0,
      grossIncomeAmount: 300.0,
      earnedReserveAmount: 0,
      earningSchedule: null,
      components: { ADMIN: { BASEADMIN: 75 }, RESERVE: { BASERESERVE: 225 } },
    },
  ];

  const sampleClaims: NormalizedClaim[] = [
    {
      sourceId: 'clm_src_01',
      snapshotDate: new Date('2026-07-10T12:00:00Z'),
      paymentKey: 'PAY-5001',
      claimNumber: 'CLM-9001',
      contractNumber: 'CNT-1001',
      activityDate: new Date('2026-07-10T00:00:00Z'),
      status: 'Paid',
      paidAmount: 250.0,
      agent: 'AG01',
      agentName: 'Apex Agency',
      dealer: 'D100',
      dealerName: 'Precision Motors',
      product: 'VSC Gold',
      coverageName: 'Transmission',
      vehicleMake: 'TOYOTA',
      lossCode: 'TRANS_01',
      lossCodeDescription: 'Transmission repair',
    },
  ];

  const sampleAudits: PipelineAuditRecord[] = [
    {
      jobType: 'Contract',
      executionTimestamp: new Date('2026-07-31T23:59:59Z'),
      executionDateStr: '2026-07-31',
      dateRange: { startDate: '2026-07-01', endDate: '2026-07-31' },
      counts: {
        portalCount: 3,
        uniqueCount: 2,
        processedCount: 3,
        uploadedCount: 3, // Matches raw_contracts (3 written snapshots)
      },
      reconciliation: {
        isMatch: true,
        portalVsProcessedDiff: 0,
        processedVsUploadedDiff: 0,
        status: 'PASSED',
        summary: '100% volumetric parity',
      },
      fileMetadata: { fileName: 'contract_export_07312026.csv' },
    },
    {
      jobType: 'Claim',
      executionTimestamp: new Date('2026-07-31T23:59:59Z'),
      executionDateStr: '2026-07-31',
      dateRange: { startDate: '2026-07-01', endDate: '2026-07-31' },
      counts: {
        portalCount: 1,
        uniqueCount: 1,
        processedCount: 1,
        uploadedCount: 1, // Matches raw_claims (1 claim)
      },
      reconciliation: {
        isMatch: true,
        portalVsProcessedDiff: 0,
        processedVsUploadedDiff: 0,
        status: 'PASSED',
        summary: '100% volumetric parity',
      },
      fileMetadata: { fileName: 'claim_export_07312026.csv' },
    },
  ];

  await duckdb.insertContractsBatch(sampleContracts);
  await duckdb.insertClaimsBatch(sampleClaims);
  await duckdb.insertPipelineAudits(sampleAudits);
  await duckdb.createDeduplicationViews();

  console.log('✅ Ingested contracts, claims, audits, and created deduplication views.');

  // 1. Test Volumetric Integrity Verification
  console.log('🔍 Running verifyVolumetricIntegrity()...');
  const auditResult = await duckdb.verifyVolumetricIntegrity(sampleAudits);
  console.log('📊 Audit Verification Result:', auditResult);
  if (!auditResult.isClean) {
    throw new Error('Audit verification failed: expected isClean = true');
  }
  console.log('✅ Volumetric reconciliation passed: 0 variance, 0 missing units.');

  // 2. Test Executive Period Metrics (Current Month: 2026-07-01 to 2026-07-31)
  console.log('🔍 Running getPeriodMetrics()...');
  const metrics = await duckdb.getPeriodMetrics(
    new Date('2026-07-01T00:00:00Z'),
    new Date('2026-07-31T23:59:59Z'),
  );
  console.log('📊 Computed Executive Period Metrics:', metrics);

  // Expected calculations:
  // CNT-1001 latest reserve = 450, admin = 100
  // CNT-1002 latest reserve = 225, admin = 75
  // Total Net Reserve = 450 + 225 = 675
  // Total Net Admin = 100 + 75 = 175
  // Claims Paid = 250
  // Underwriting Profit = 675 - 250 = 425
  if (metrics.contractsWritten !== 2) {
    throw new Error(`Expected 2 written contracts, got ${metrics.contractsWritten}`);
  }
  if (metrics.netReserve !== 675) {
    throw new Error(`Expected 675 net reserve, got ${metrics.netReserve}`);
  }
  if (metrics.netAdmin !== 175) {
    throw new Error(`Expected 175 net admin, got ${metrics.netAdmin}`);
  }
  if (metrics.claimsPaid !== 250) {
    throw new Error(`Expected 250 claims paid, got ${metrics.claimsPaid}`);
  }
  if (metrics.underwritingProfit !== 425) {
    throw new Error(`Expected 425 underwriting profit, got ${metrics.underwritingProfit}`);
  }
  console.log('✅ Period metrics mathematically match expected values.');

  // 3. Test Top Dealers Ranking
  console.log('🔍 Running getTopDealers()...');
  const topDealers = await duckdb.getTopDealers(
    new Date('2026-07-01T00:00:00Z'),
    new Date('2026-07-31T23:59:59Z'),
  );
  console.log('📊 Top Dealers Ranking:', topDealers);
  if (topDealers.length !== 2) {
    throw new Error(`Expected 2 dealers, got ${topDealers.length}`);
  }
  // Precision Motors (D100) has $450 net reserve; Summit Auto (D200) has $225. D100 must be first.
  if (topDealers[0]?.name !== 'D100') {
    throw new Error(`Expected D100 to be #1 ranked dealer, got ${topDealers[0]?.name}`);
  }
  console.log('✅ Top Dealers correctly ranked in descending order by Net Reserve.');

  // 4. Test Parquet Export
  const tempOutputDir = path.join(process.cwd(), 'output', 'test_parquet');
  if (!fs.existsSync(tempOutputDir)) {
    fs.mkdirSync(tempOutputDir, { recursive: true });
  }
  const exportResult = await duckdb.exportToParquet(tempOutputDir);
  console.log('📦 Parquet Export Result:', exportResult);
  if (!fs.existsSync(exportResult.contractsPath) || !fs.existsSync(exportResult.claimsPath)) {
    throw new Error('Parquet files were not found on disk after export.');
  }
  console.log('✅ Parquet export successfully created compressed files on disk.');

  // Clean up test export
  fs.rmSync(tempOutputDir, { recursive: true, force: true });

  await duckdb.close();
  console.log('🎉 ALL DUCKDB ENTERPRISE VERIFICATION TESTS PASSED (100%)!');
}

runTest().catch((err) => {
  console.error('❌ Test failed:', err);
  process.exit(1);
});
