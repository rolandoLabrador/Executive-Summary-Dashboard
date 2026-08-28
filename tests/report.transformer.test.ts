import * as assert from 'assert';
import * as ExcelJS from 'exceljs';
import { type ReportConfig, type UnknownDocument } from '../src/models/report.types';
import { ExcelService } from '../src/services/excel.service';
import { renderPieChartPng } from '../src/utils/pie-chart.renderer';
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
  contract('invalid-date', 'C-INVALID-DATE', 'NewBusiness', 'not-a-date'),
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
    'Coverage Name': 'Executive Protection VSC Max EE',
    Make: 'BMW',
    'Loss Code': 'EL016',
    'Loss Code Description': 'Miscellaneous Covered',
    'Date Paid': '07/25/2026',
    'Total Paid Amount': '250.00',
  },
];

const sampleAudits = [
  {
    jobType: 'Cancel',
    executionTimestamp: new Date('2026-08-28T11:46:07.570Z'),
    executionDateStr: '2026-08-28',
    dateRange: { startDate: '08/28/2024', endDate: '08/28/2026' },
    counts: { portalCount: 3952, processedCount: 3952, uploadedCount: 3952 },
    reconciliation: {
      isMatch: true,
      portalVsProcessedDiff: 0,
      processedVsUploadedDiff: 0,
      status: 'PASSED',
      summary: '100% Volumetric Integrity: All 3 stages match exactly (3,952 records).',
    },
    fileMetadata: { fileName: 'CancelExport.csv' },
    systemInfo: { environment: 'GitHub_Actions_CI', source: 'Playwright_ETL_Pipeline' },
  },
  {
    jobType: 'Claim',
    executionTimestamp: new Date('2026-08-28T11:47:42.087Z'),
    executionDateStr: '2026-08-28',
    dateRange: { startDate: '08/28/2024', endDate: '08/28/2026' },
    counts: { portalCount: 2861, processedCount: 14400, uploadedCount: 14400 },
    reconciliation: {
      isMatch: false,
      portalVsProcessedDiff: -11539,
      processedVsUploadedDiff: 0,
      status: 'DISCREPANCY',
      summary: 'Discrepancy Detected! Portal: 2,861, Processed: 14,400, Uploaded: 14,400',
    },
    fileMetadata: { fileName: 'ClaimExport_PROCESSED.csv' },
    systemInfo: { environment: 'GitHub_Actions_CI', source: 'Playwright_ETL_Pipeline' },
  },
];

const report = new ReportTransformer(config).transform(
  contracts,
  cancellations,
  claims,
  sampleAudits,
);
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
assert.strictEqual(report.claims[0]?.lossCode, 'EL016');
assert.strictEqual(report.claims[0]?.lossCodeDescription, 'Miscellaneous Covered');
assert.strictEqual(report.claims[0]?.coverageName, 'Executive Protection VSC Max EE');
assert.strictEqual(report.claims[0]?.vehicleMake, 'BMW');
assert.deepStrictEqual(report.lossCodeDashboard.currentMonth, {
  totalPaid: 250,
  claimCount: 1,
  lossCodeCount: 1,
  averagePaidPerClaim: 250,
});
assert.strictEqual(report.lossCodeDashboard.yearToDate.totalPaid, 250);
assert.strictEqual(report.lossCodeDashboard.rolling12.totalPaid, 250);
assert.deepStrictEqual(report.lossCodeDashboard.topVehicleMakes, [
  { make: 'BMW', paidAmount: 250, claimCount: 1, paidShare: 1 },
]);
assert.deepStrictEqual(report.lossCodeDashboard.rows, [
  {
    code: 'EL016',
    description: 'Miscellaneous Covered',
    coverageNames: ['Executive Protection VSC Max EE'],
    currentMonthPaid: 250,
    yearToDatePaid: 250,
    rolling12Paid: 250,
    rolling12PaidShare: 1,
    rolling12ClaimCount: 1,
    rolling12AveragePaidPerClaim: 250,
  },
]);
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
cancellationMetadata.ContractStatus = 'A';
const cancellationOnlyReport = new ReportTransformer(config).transform([], [cancellationOnly], []);
const june = cancellationOnlyReport.monthly.find((item) => item.period === '2026-06');
assert.strictEqual(
  june?.contractsWritten,
  1,
  'A cancellation should restore its missing original contract to the ActivationDate cohort.',
);
assert.strictEqual(june?.reserveWritten, 500);
assert.strictEqual(
  cancellationOnlyReport.currentMonth.current.contractsWritten,
  0,
  'The original written contract must not be recognized in the cancellation month.',
);
assert.strictEqual(cancellationOnlyReport.currentMonth.current.contractsCancelled, 1);
assert.strictEqual(
  cancellationOnlyReport.currentMonth.current.activeContracts,
  0,
  'An explicitly classified cancellation must never be active, even when ContractStatus is inconsistent.',
);
assert.strictEqual(cancellationOnlyReport.contractTransactions.length, 2);
const writtenReference = cancellationOnlyReport.contractTransactions.find((item) =>
  item.sourceId.endsWith(':written-reference'),
);
assert.strictEqual(writtenReference?.transactionType, 'NewBusiness');
assert.strictEqual(writtenReference?.contractStatus, 'C');

const componentPayments = [
  { code: 'TR001', description: 'Transmission', amount: 200 },
  { code: 'EN001', description: 'Engine', amount: 100 },
  { code: 'BR001', description: 'Brakes', amount: 80 },
  { code: 'AC001', description: 'Air Conditioning', amount: 70 },
  { code: 'EL001', description: 'Electrical', amount: 60 },
  { code: 'ST001', description: 'Steering', amount: 50 },
  { code: 'SU001', description: 'Suspension', amount: 40 },
  { code: 'WH001', description: 'Wheel Bearing', amount: 30 },
  { code: 'CL001', description: 'Cooling System', amount: 20 },
  { code: 'EX001', description: 'Exhaust', amount: 10 },
  { code: 'FS001', description: 'Fuel System', amount: 9 },
  { code: 'CV001', description: 'CV Axle', amount: 8 },
  { code: 'DR001', description: 'Drive Shaft', amount: 7 },
  { code: 'SE001', description: 'Seals and Gaskets', amount: 6 },
  { code: 'MI001', description: 'Miscellaneous', amount: 5 },
  { code: 'TN001', description: 'Below Display Threshold', amount: 0.1 },
];
const multiComponentClaimReport = new ReportTransformer(config).transform(
  [],
  [],
  componentPayments.map(({ code, description, amount }, index) => ({
    _id: `shared-claim-component-${index + 1}`,
    Activity: 'Payment Issued',
    'Claim Status': 'Paid',
    'Claim Number': 'CL-SHARED',
    'Contract Number': 'C-SHARED',
    'Loss Code': code,
    'Loss Code Description': description,
    'Product Type': 'VSC',
    'Coverage Name': 'Executive Protection VSC Max EE',
    Make: 'BMW',
    'Date Paid': '07/10/2026',
    'Total Paid Amount': amount.toFixed(2),
  })),
);
assert.ok(
  Math.abs(multiComponentClaimReport.lossCodeDashboard.currentMonth.totalPaid - 695.1) < 0.0001,
);
assert.strictEqual(
  multiComponentClaimReport.lossCodeDashboard.currentMonth.claimCount,
  1,
  'The overall KPI must count one claim even when it contains multiple paid loss codes.',
);
assert.strictEqual(multiComponentClaimReport.lossCodeDashboard.currentMonth.lossCodeCount, 16);
assert.ok(
  multiComponentClaimReport.lossCodeDashboard.rows.every((row) => row.rolling12ClaimCount === 1),
  'Each loss-code row must independently count the shared claim.',
);
const renderedPie = renderPieChartPng([
  { value: 100, color: '2F75B5' },
  { value: 200, color: 'ED7D31' },
]);
assert.deepStrictEqual(
  [...renderedPie.subarray(0, 8)],
  [137, 80, 78, 71, 13, 10, 26, 10],
  'The loss-code visualization must be rendered as a valid PNG container.',
);
assert.ok(renderedPie.length > 1_000);

const vehicleMakeReport = new ReportTransformer(config).transform(
  [],
  [],
  Array.from({ length: 12 }, (_, index) => ({
    _id: `vehicle-make-claim-${index + 1}`,
    Activity: 'Payment Issued',
    'Claim Status': 'Paid',
    'Claim Number': `CL-MAKE-${index + 1}`,
    'Contract Number': `C-MAKE-${index + 1}`,
    Make: `Make ${String(index + 1).padStart(2, '0')}`,
    'Loss Code': 'MK001',
    'Loss Code Description': 'Make Ranking Test',
    'Date Paid': '07/10/2026',
    'Total Paid Amount': String((12 - index) * 10),
  })),
);
assert.strictEqual(vehicleMakeReport.lossCodeDashboard.topVehicleMakes.length, 10);
assert.strictEqual(vehicleMakeReport.lossCodeDashboard.topVehicleMakes[0]?.make, 'MAKE 01');
assert.strictEqual(vehicleMakeReport.lossCodeDashboard.topVehicleMakes[0]?.paidAmount, 120);
assert.strictEqual(vehicleMakeReport.lossCodeDashboard.topVehicleMakes[9]?.make, 'MAKE 10');
assert.strictEqual(vehicleMakeReport.lossCodeDashboard.topVehicleMakes[9]?.paidAmount, 30);

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
  assert.strictEqual(
    dashboard?.getCell('A2').value,
    'Current reporting month: 07/01/2026–07/31/2026 | Paid claims / net written reserve',
  );
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
  assert.strictEqual(
    dealerDashboard?.getCell('A2').value,
    'Rolling 12-month results: 08/01/2025–07/31/2026',
  );
  const agentDashboard = workbook.getWorksheet('Agent Dashboard');
  assert.strictEqual(
    agentDashboard?.getCell('A2').value,
    'Rolling 12-month results: 08/01/2025–07/31/2026',
  );
  const productDashboard = workbook.getWorksheet('Product Dashboard');
  assert.strictEqual(
    productDashboard?.getCell('A2').value,
    'Monthly results: 07/01/2026–07/31/2026',
  );
  assert.strictEqual(productDashboard?.getCell('C4').value, 'Reporting Month');
  assert.strictEqual(
    (productDashboard?.getCell('C5').value as Date).getTime(),
    new Date(2026, 6, 1).getTime(),
  );
  assert.strictEqual(productDashboard?.getCell('B5').value, 'VSC');
  const lossCodeDashboard = workbook.getWorksheet('Loss Code Dashboard');
  assert.strictEqual(
    lossCodeDashboard?.getCell('A2').value,
    'Paid amounts by covered vehicle component | Through 07/31/2026',
  );
  assert.strictEqual(lossCodeDashboard?.getCell('A5').value, 'Total Paid Amount');
  assert.strictEqual(lossCodeDashboard?.getCell('B5').value, 250);
  assert.match(String(lossCodeDashboard?.getCell('K4').value), /TOP 10 VEHICLE MAKES/);
  assert.strictEqual(lossCodeDashboard?.getCell('L5').value, 'Vehicle Make');
  assert.strictEqual(lossCodeDashboard?.getCell('L6').value, 'BMW');
  assert.strictEqual(lossCodeDashboard?.getCell('M6').value, 250);
  assert.strictEqual(lossCodeDashboard?.getCell('N6').value, 1);
  assert.strictEqual(lossCodeDashboard?.getCell('O6').value, 1);
  const lossCodeConditionalFormatting = (
    lossCodeDashboard as ExcelJS.Worksheet & {
      conditionalFormattings: Array<{ ref: string }>;
    }
  ).conditionalFormattings;
  assert.ok(
    lossCodeConditionalFormatting.some(({ ref }) => ref === 'O6:O6'),
    'Top vehicle-make paid shares should use the original percentage data bar.',
  );
  assert.strictEqual(lossCodeDashboard?.getImages().length, 1);
  assert.match(String(lossCodeDashboard?.getCell('A18').value), /PAID AMOUNT MIX/);
  assert.strictEqual(lossCodeDashboard?.getCell('D19').value, 'Color');
  assert.strictEqual(lossCodeDashboard?.getCell('E20').value, 'EL016');
  assert.strictEqual(lossCodeDashboard?.getCell('F20').value, 'Miscellaneous Covered');
  assert.strictEqual(lossCodeDashboard?.getCell('G20').value, 'Executive Protection VSC Max EE');
  assert.strictEqual(lossCodeDashboard?.getCell('H20').value, 250);
  assert.strictEqual(lossCodeDashboard?.getCell('I20').value, 1);
  assert.strictEqual(lossCodeDashboard?.getCell('B39').value, 'EL016');
  assert.strictEqual(lossCodeDashboard?.getCell('C39').value, 'Miscellaneous Covered');
  assert.strictEqual(lossCodeDashboard?.getCell('D38').value, 'Current Month Paid');
  assert.strictEqual(lossCodeDashboard?.getCell('F39').value, 250);
  assert.strictEqual(lossCodeDashboard?.getCell('G39').value, 1);
  assert.strictEqual(lossCodeDashboard?.getCell('H39').value, 1);
  assert.strictEqual(lossCodeDashboard?.getCell('I39').value, 250);

  const allComponentsBuffer = await new ExcelService(config).generateFullReport(
    multiComponentClaimReport,
  );
  const allComponentsWorkbook = new ExcelJS.Workbook();
  await allComponentsWorkbook.xlsx.load(Uint8Array.from(allComponentsBuffer).buffer);
  const allComponentsDashboard = allComponentsWorkbook.getWorksheet('Loss Code Dashboard');
  const displayedComponents = componentPayments.filter(
    (component) => component.amount / 695.1 >= 0.02,
  );
  const legendCodes =
    allComponentsDashboard
      ?.getRows(20, displayedComponents.length)
      ?.map((row) => String(row.getCell(5).value)) ?? [];
  assert.deepStrictEqual(
    legendCodes,
    displayedComponents.map((component) => component.code),
    'Every loss code meeting the 2% threshold must retain its own pie slice and legend row.',
  );
  assert.ok(!legendCodes.includes('OTHER'), 'The pie chart must not synthesize an OTHER group.');
  assert.ok(
    !allComponentsDashboard
      ?.getColumn(5)
      .values.map((value) => String(value))
      .includes('TN001'),
    'Loss codes below 2% must be omitted from both the legend and rolling-12 detail.',
  );
  assert.ok(
    !allComponentsDashboard
      ?.getColumn(5)
      .values.map((value) => String(value))
      .includes('EX001'),
    'Loss codes below 2% must be omitted from both the legend and rolling-12 detail.',
  );
  assert.strictEqual(allComponentsDashboard?.getImages().length, 1);
  assert.strictEqual(allComponentsDashboard?.getCell('E28').value, 'CL001');
  assert.strictEqual(allComponentsDashboard?.getCell('F28').value, 'Cooling System');
  assert.strictEqual(
    allComponentsDashboard?.getCell('G28').value,
    'Executive Protection VSC Max EE',
  );
  assert.strictEqual(allComponentsDashboard?.getCell('H28').value, 20);
  assert.match(String(allComponentsDashboard?.getCell('A37').value), /ROLLING 12-MONTH DETAIL/);
  assert.strictEqual(allComponentsDashboard?.getCell('B38').value, 'Loss Code');
  assert.strictEqual(allComponentsDashboard?.getCell('D38').value, 'Current Month Paid');
  assert.strictEqual(allComponentsDashboard?.getCell('B39').value, 'TR001');
  const monthlyTrends = workbook.getWorksheet('Monthly Trends');
  assert.strictEqual(
    monthlyTrends?.getCell('A2').value,
    'Monthly trend coverage: 08/01/2024–07/31/2026',
  );
  assert.strictEqual(
    workbook.getWorksheet('Contract Activity')?.getCell('A2').value,
    'No customer names, contact details, addresses, or VINs',
  );
  assert.strictEqual(
    workbook.getWorksheet('Claim Activity')?.getCell('A2').value,
    'Paid claims only; no customer or vehicle identifiers',
  );
  assert.strictEqual(workbook.getWorksheet('Claim Activity')?.getCell('K4').value, 'Loss Code');
  assert.strictEqual(
    workbook.getWorksheet('Claim Activity')?.getCell('L4').value,
    'Component Description',
  );
  assert.strictEqual(workbook.getWorksheet('Claim Activity')?.getCell('K5').value, 'EL016');
  const dataQuality = workbook.getWorksheet('Data Quality');
  assert.strictEqual(dataQuality?.getCell('A6').value, 'PASSED');
  assert.strictEqual(dataQuality?.getCell('B6').value, 'Cancel');
  assert.strictEqual(dataQuality?.getCell('C6').value, 3952);
  assert.strictEqual(dataQuality?.getCell('D6').value, 3952);
  assert.strictEqual(dataQuality?.getCell('E6').value, 0);
  assert.strictEqual(dataQuality?.getCell('A7').value, 'DISCREPANCY');
  assert.strictEqual(dataQuality?.getCell('B7').value, 'Claim');
  assert.strictEqual(dataQuality?.getCell('C7').value, 2861);
  assert.strictEqual(dataQuality?.getCell('D7').value, 14400);
  assert.strictEqual(dataQuality?.getCell('E7').value, -11539);
  assert.strictEqual(dataQuality?.getCell('C19').value, 'Contract Number');
  assert.strictEqual(dataQuality?.getCell('D19').value, 'Dealer Name');
  assert.strictEqual(dataQuality?.getCell('C20').value, 'C-INVALID-DATE');
  assert.strictEqual(dataQuality?.getCell('D20').value, 'Dealer One');
  assert.strictEqual(dataQuality?.getCell('E20').value, 'invalid-date');
  const definitions = workbook.getWorksheet('Definitions');
  const reconciliationDefinition = definitions
    ? definitions
        .getRows(1, definitions.rowCount)
        ?.find((row) => row.getCell(1).value === 'Contract Count Reconciliation')
    : undefined;
  assert.match(
    String(reconciliationDefinition?.getCell(2).value),
    /Written Contracts = Active Contracts/,
  );
  console.log('Report transformation and workbook-generation tests passed.');
}

verifyWorkbook().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
