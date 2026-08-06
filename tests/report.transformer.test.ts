import * as assert from 'assert';
import * as ExcelJS from 'exceljs';
import { type ReportConfig, type UnknownDocument } from '../src/models/report.types';
import { ExcelService } from '../src/services/excel.service';
import { ReportTransformer } from '../src/utils/report.transformer';

const config: ReportConfig = {
  companyName: 'TEST',
  asOfDate: new Date('2026-08-01T23:59:59.999'),
  outputDirectory: 'output',
  topDealerCount: 20,
  warningLossRatio: 0.65,
  highLossRatio: 0.8,
  excludedComponentCodes: new Set(),
};

function contract(
  id: string,
  number: string,
  type: 'NewBusiness' | 'Cancellation',
  activityDate: string,
): UnknownDocument {
  return {
    _id: id,
    WrittenAmount: {
      ADMIN: { BASEADMIN: 100, DEALERCOMM: 900, PACK: 400 },
      RESERVE: { BASERESERVE: 500, DEALERRESERVE: 100 },
    },
    CancelledAmount: {
      ADMIN: { BASEADMIN: 100, DEALERCOMM: 900 },
      RESERVE: { BASERESERVE: 500 },
    },
    metadata: {
      'Contract#': number,
      'Status(NewBusiness,Cancellation,Upgrade,Adjustment)': type,
      ContractStatus: type === 'Cancellation' ? 'C' : 'A',
      BillingDate: activityDate,
      ActivationDate: activityDate,
      CancelBillDate: type === 'Cancellation' ? activityDate : '',
      CancellationEffectiveDate: type === 'Cancellation' ? activityDate : '',
      Agent: 'Agent One',
      DealerName: 'Dealer One',
      DealerNumber: 'D1',
      ProductType: 'VSC',
      CoverageCode: 'VSC-A',
    },
  };
}

const contracts = [
  contract('sale-1', 'C1', 'NewBusiness', '07/05/2026'),
  contract('sale-1-copy', 'C1', 'NewBusiness', '07/05/2026'), // same business contract, different MongoDB _id
  contract('sale-2', 'C2', 'NewBusiness', '07/06/2026'),
  contract('august-sale', 'C-AUG', 'NewBusiness', '08/02/2026'),
];
const cancellations = [contract('cancel-1', 'C2', 'Cancellation', '07/20/2026')];
const claims: UnknownDocument[] = [
  {
    _id: 'claim-1',
    Activity: 'Payment Issued',
    'Claim Status': 'Paid',
    'Claim Number': 'CL1',
    'Contract Number': 'C1',
    'Agent Number': 'Agent One',
    'Agent Name': 'Different Agent Display Name',
    'Selling Dealer Number': 'D1',
    'Selling Dealer Name': 'Different Dealer Display Name',
    'Coverage Code': 'VSC-A',
    'Date Paid': '07/25/2026',
    'Total Paid Amount': '250.00',
  },
];

const report = new ReportTransformer(config).transform(contracts, cancellations, claims);
assert.strictEqual(report.asOfDate.getTime(), new Date(2026, 6, 31, 23, 59, 59, 999).getTime());
assert.strictEqual(
  report.currentMonth.currentStart.getMonth(),
  6,
  'An August run must report the fully completed month of July.',
);
assert.strictEqual(report.currentMonth.currentStart.getFullYear(), 2026);
assert.strictEqual(report.currentMonth.currentEnd.getDate(), 31);
assert.strictEqual(
  report.currentMonth.priorStart.getMonth(),
  6,
  'Current month must compare with the same month in the prior year.',
);
assert.strictEqual(report.currentMonth.priorStart.getFullYear(), 2025);
assert.strictEqual(report.currentMonth.priorEnd.getDate(), 31);
assert.strictEqual(report.yearToDate.currentStart.getTime(), new Date(2026, 0, 1).getTime());
assert.strictEqual(
  report.yearToDate.currentEnd.getTime(),
  new Date(2026, 6, 31, 23, 59, 59, 999).getTime(),
);
assert.strictEqual(report.yearToDate.priorStart.getTime(), new Date(2025, 0, 1).getTime());
assert.strictEqual(
  report.yearToDate.priorEnd.getTime(),
  new Date(2025, 6, 31, 23, 59, 59, 999).getTime(),
);
assert.strictEqual(report.rolling12.currentStart.getTime(), new Date(2025, 7, 1).getTime());
assert.strictEqual(
  report.rolling12.currentEnd.getTime(),
  new Date(2026, 6, 31, 23, 59, 59, 999).getTime(),
);
assert.strictEqual(report.rolling12.priorStart.getTime(), new Date(2024, 7, 1).getTime());
assert.strictEqual(
  report.rolling12.priorEnd.getTime(),
  new Date(2025, 6, 31, 23, 59, 59, 999).getTime(),
);
assert.strictEqual(report.priorCalendarYear.start.getFullYear(), 2025);
assert.strictEqual(report.priorCalendarYear.start.getMonth(), 0);
assert.strictEqual(report.priorCalendarYear.start.getDate(), 1);
assert.strictEqual(report.priorCalendarYear.end.getFullYear(), 2025);
assert.strictEqual(report.priorCalendarYear.end.getMonth(), 11);
assert.strictEqual(report.priorCalendarYear.end.getDate(), 31);
assert.strictEqual(
  report.sourceCounts.uniqueContractTransactions,
  3,
  'Open-month activity must be excluded from every report tab.',
);
assert.ok(!report.contractTransactions.some((item) => item.contractNumber === 'C-AUG'));
assert.strictEqual(report.currentMonth.current.contractsWritten, 2);
assert.strictEqual(report.currentMonth.current.contractsCancelled, 1);
assert.strictEqual(report.currentMonth.current.netContracts, 1);
assert.strictEqual(report.currentMonth.current.activeContracts, 1);
assert.strictEqual(report.currentMonth.current.netAdmin, 100);
assert.strictEqual(report.currentMonth.current.netReserve, 500);
assert.strictEqual(report.currentMonth.current.claimsPaid, 250);
assert.strictEqual(report.currentMonth.current.paidLossRatio, 0.5);
assert.strictEqual(
  report.products.length,
  1,
  'Coverage codes must consolidate into their contract ProductType.',
);
assert.strictEqual(report.products[0]?.name, 'VSC');
assert.strictEqual(report.products[0]?.claimsPaid, 250);
assert.strictEqual(report.dealers[0]?.name, 'D1');
assert.deepStrictEqual(report.dealers[0]?.relatedAgents, ['Agent One']);
assert.strictEqual(report.dealers[0]?.paidLossRatio, 0.5);
assert.ok(!report.dataQualityIssues.some((issue) => issue.category === 'Unmatched Claim'));

const testAgentContract = contract('test-agent-sale', 'C-TEST', 'NewBusiness', '07/10/2026');
(testAgentContract.metadata as Record<string, unknown>).Agent = ' Test ';
const agentFilterReport = new ReportTransformer(config).transform([testAgentContract], [], []);
assert.strictEqual(
  agentFilterReport.agents.length,
  0,
  'The Test agent must not appear on the Agent Dashboard.',
);
assert.strictEqual(
  agentFilterReport.currentMonth.current.contractsWritten,
  1,
  'Filtering the Agent Dashboard must not alter executive totals.',
);

const cancellationOnly = contract('cancel-only', 'C3', 'Cancellation', '07/20/2026');
const cancellationMetadata = cancellationOnly.metadata as Record<string, unknown>;
cancellationMetadata.BillingDate = '06/05/2026';
cancellationMetadata.ActivationDate = '06/05/2026';
cancellationMetadata.SaleDate = '06/01/2026';
const fallbackReport = new ReportTransformer(config).transform([], [cancellationOnly], []);
const june = fallbackReport.monthly.find((item) => item.period === '2026-06');
assert.strictEqual(
  june?.contractsWritten,
  1,
  'Cancellation WrittenAmount should restore missing original new business.',
);
assert.strictEqual(june?.reserveWritten, 500);
assert.strictEqual(
  fallbackReport.currentMonth.current.contractsWritten,
  0,
  'Original written amount must not be added in the cancellation month.',
);
assert.strictEqual(fallbackReport.currentMonth.current.contractsCancelled, 1);
assert.strictEqual(fallbackReport.currentMonth.current.activeContracts, 0);

async function verifyWorkbook(): Promise<void> {
  const buffer = await new ExcelService(config).generateFullReport(report);
  assert.ok(buffer.length > 10_000, 'Generated workbook should contain formatted report content.');
  assert.strictEqual(
    buffer.subarray(0, 2).toString(),
    'PK',
    'Generated report must be a valid XLSX/ZIP container.',
  );
  const workbook = new ExcelJS.Workbook();
  // ExcelJS declares its input as ArrayBuffer, while report generation returns
  // Node's Buffer. Copying provides the precise portable type expected here.
  await workbook.xlsx.load(Uint8Array.from(buffer).buffer);
  const dashboard = workbook.getWorksheet('Executive Dashboard');
  assert.strictEqual(dashboard?.getCell('C5').value, 'Prior Year');
  assert.match(
    String(dashboard?.getCell('A4').value),
    /07\/01\/2026.*07\/31\/2026.*07\/01\/2025.*07\/31\/2025/,
  );
  assert.strictEqual(
    dashboard?.getCell('H4').value,
    '12-MONTH TREND — THROUGH LATEST COMPLETED MONTH',
  );
  assert.match(String(dashboard?.getCell('A31').value), /^FULL YEAR 2025/);
  assert.strictEqual(dashboard?.getCell('A32').value, 'Metric');
  assert.strictEqual(dashboard?.getCell('B32').value, 'Full-Year Result');
  const dealerDashboard = workbook.getWorksheet('Dealer Dashboard');
  assert.strictEqual(dealerDashboard?.getCell('D4').value, 'Agents');
  assert.strictEqual(dealerDashboard?.getCell('D5').value, 'Agent One');
  assert.strictEqual(dealerDashboard?.getCell('A2').value, 'Rolling 12 months through 07/31/2026');
  const productDashboard = workbook.getWorksheet('Product Dashboard');
  assert.strictEqual(
    productDashboard?.getCell('A2').value,
    'Monthly results for July 2026 through 07/31/2026',
  );
  assert.strictEqual(productDashboard?.getCell('C4').value, 'Reporting Month');
  assert.strictEqual(
    (productDashboard?.getCell('C5').value as Date).getTime(),
    new Date(2026, 6, 1).getTime(),
  );
  assert.strictEqual(productDashboard?.getCell('B5').value, 'VSC');
  console.log('Report transformation and workbook-generation tests passed.');
}

verifyWorkbook().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
