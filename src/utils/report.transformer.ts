import {
  type DataQualityIssue,
  type DimensionMetric,
  type LossCodeDashboard,
  type LossCodeKpis,
  type LossCodeMetric,
  type MetricValues,
  type MonthlyMetric,
  type NormalizedClaim,
  type NormalizedContractTransaction,
  type PipelineAuditRecord,
  type ReportConfig,
  type ReportModel,
  type TransactionType,
  type UnknownDocument,
  type VehicleMakeMetric,
} from '../models/report.types';

const UNMAPPED_LOSS_CODE = 'UNMAPPED';
const UNAVAILABLE_LOSS_DESCRIPTION = 'Description unavailable';

const EMPTY_METRICS: MetricValues = {
  contractsWritten: 0,
  activeContracts: 0,
  contractsCancelled: 0,
  netContracts: 0,
  adminWritten: 0,
  adminCancelled: 0,
  netAdmin: 0,
  reserveWritten: 0,
  reserveCancelled: 0,
  netReserve: 0,
  claimsPaid: 0,
  claimCount: 0,
  paidLossRatio: null,
  cancellationRate: null,
  adminPerContract: null,
};

type Dimension = 'dealer' | 'agent' | 'product';

function record(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' ? (value as Record<string, unknown>) : {};
}

function text(value: unknown): string {
  return value === null || value === undefined ? '' : String(value).trim();
}

function businessId(value: unknown): string {
  return text(value).replace(/\s+/g, '').toUpperCase();
}

function number(value: unknown): number {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? '').replace(/[$,%\s,]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

function date(value: unknown): Date | null {
  const raw = text(value);
  if (!raw) return null;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function idOf(document: UnknownDocument): string {
  const id = document._id;
  if (typeof id === 'string') return id;
  if (id && typeof id === 'object' && '$oid' in id) return text(id.$oid);
  return id?.toString() || '';
}

function objectIdDate(value: unknown): Date | null {
  if (value && typeof value === 'object' && 'getTimestamp' in value) {
    const getter = (value as { getTimestamp?: () => Date }).getTimestamp;
    if (typeof getter === 'function') return getter.call(value);
  }
  const identifier =
    typeof value === 'string'
      ? value
      : value && typeof value === 'object' && '$oid' in value
        ? text((value as { $oid?: unknown }).$oid)
        : '';
  if (!/^[a-f\d]{24}$/i.test(identifier)) return null;
  return new Date(Number.parseInt(identifier.slice(0, 8), 16) * 1000);
}

function snapshotDate(document: UnknownDocument, metadata: Record<string, unknown> = {}): Date {
  return (
    firstDate(
      metadata.ExtractionDate,
      document.ExtractionDate,
      document.extractionDate,
      document.ingestedAt,
      document.createdAt,
    ) ||
    objectIdDate(document._id) ||
    new Date(0)
  );
}

function firstDate(...values: unknown[]): Date | null {
  for (const value of values) {
    const parsed = date(value);
    if (parsed) return parsed;
  }
  return null;
}

function classify(value: unknown, status: unknown): TransactionType {
  const normalized = text(value).replace(/\s/g, '').toLowerCase();
  if (normalized === 'newbusiness') return 'NewBusiness';
  if (normalized === 'cancellation') return 'Cancellation';
  if (normalized === 'adjustment') return 'Adjustment';
  if (normalized === 'upgrade') return 'Upgrade';
  if (text(status).toUpperCase() === 'C') return 'Cancellation';
  if (text(status).toUpperCase() === 'A') return 'NewBusiness';
  return 'Unknown';
}

function excludedComponent(name: string, config: ReportConfig): boolean {
  const upper = name.trim().toUpperCase();
  return (
    config.excludedComponentCodes.has(upper) ||
    upper.includes('DEALER') ||
    upper.includes('DLR') ||
    upper.includes('COMMISSION') ||
    upper.includes('COMM') ||
    upper.includes('F&I') ||
    upper.includes('PACK')
  );
}

function sumComponents(
  container: unknown,
  category: 'ADMIN' | 'RESERVE',
  config: ReportConfig,
): number {
  const values = record(record(container)[category]);
  return Object.entries(values).reduce(
    (sum, [name, value]) => sum + (excludedComponent(name, config) ? 0 : number(value)),
    0,
  );
}

function monthStart(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), 1);
}

function addMonths(value: Date, amount: number): Date {
  return new Date(value.getFullYear(), value.getMonth() + amount, 1);
}

function endOfMonth(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth() + 1, 0, 23, 59, 59, 999);
}

function inRange(value: Date, start: Date, end: Date): boolean {
  return value >= start && value <= end;
}

function periodKey(value: Date): string {
  return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}`;
}

function finalize(metrics: MetricValues): MetricValues {
  const result = { ...metrics };
  result.netContracts = result.contractsWritten - result.contractsCancelled;
  result.netAdmin = result.adminWritten - result.adminCancelled;
  result.netReserve = result.reserveWritten - result.reserveCancelled;
  result.paidLossRatio = result.netReserve > 0 ? result.claimsPaid / result.netReserve : null;
  result.cancellationRate =
    result.contractsWritten > 0 ? result.contractsCancelled / result.contractsWritten : null;
  result.adminPerContract =
    result.activeContracts > 0 ? result.netAdmin / result.activeContracts : null;
  return result;
}

function aggregate(
  transactions: NormalizedContractTransaction[],
  claims: NormalizedClaim[],
  start: Date,
  end: Date,
): MetricValues {
  const result = { ...EMPTY_METRICS };
  const latestState = new Map<string, NormalizedContractTransaction>();
  for (const transaction of transactions) {
    if (!transaction.contractNumber) continue;
    const existing = latestState.get(transaction.contractNumber);
    if (!existing || transaction.snapshotDate >= existing.snapshotDate) {
      latestState.set(transaction.contractNumber, transaction);
    }
  }
  result.activeContracts = [...latestState.values()].filter(
    (transaction) =>
      transaction.contractStatus === 'A' &&
      transaction.transactionType !== 'Cancellation' &&
      inRange(transaction.activityDate, start, end),
  ).length;

  for (const transaction of transactions) {
    if (!inRange(transaction.activityDate, start, end)) continue;
    if (
      transaction.transactionType !== 'Cancellation' &&
      transaction.transactionType !== 'Unknown'
    ) {
      result.contractsWritten += 1;
      result.adminWritten += transaction.adminAmount;
      result.reserveWritten += transaction.reserveAmount;
    } else if (transaction.transactionType === 'Cancellation') {
      result.contractsCancelled += 1;
      result.adminCancelled += Math.abs(transaction.adminAmount);
      result.reserveCancelled += Math.abs(transaction.reserveAmount);
    }
  }

  const claimNumbers = new Set<string>();
  for (const claim of claims) {
    if (!inRange(claim.activityDate, start, end)) continue;
    result.claimsPaid += claim.paidAmount;
    claimNumbers.add(claim.claimNumber || claim.sourceId);
  }
  result.claimCount = claimNumbers.size;
  return finalize(result);
}

function normalizeContract(
  document: UnknownDocument,
  config: ReportConfig,
  issues: DataQualityIssue[],
): NormalizedContractTransaction | null {
  const metadata = record(document.metadata);
  const sourceId = idOf(document);
  const contractNumber = businessId(metadata['Contract#']);
  const dealerName = text(metadata.DealerName);
  const transactionType = classify(
    metadata['Status(NewBusiness,Cancellation,Upgrade,Adjustment)'],
    metadata.ContractStatus,
  );
  const isCancellation = transactionType === 'Cancellation';
  const activityDate = isCancellation
    ? firstDate(metadata.CancelBillDate)
    : firstDate(metadata.ActivationDate);

  if (!activityDate) {
    issues.push({
      severity: 'Error',
      category: isCancellation ? 'Missing Cancel Bill Date' : 'Missing Activation Date',
      contractNumber,
      dealerName,
      sourceId,
      message: isCancellation
        ? 'Cancellation is excluded because metadata.CancelBillDate is blank or invalid.'
        : 'Contract transaction is excluded because metadata.ActivationDate is blank or invalid.',
    });
    return null;
  }
  if (activityDate > config.asOfDate) return null;

  if (!contractNumber) {
    issues.push({
      severity: 'Error',
      category: 'Missing Contract Number',
      contractNumber,
      dealerName,
      sourceId,
      message: 'Contract transaction has no contract number.',
    });
  }
  if (transactionType === 'Unknown') {
    issues.push({
      severity: 'Warning',
      category: 'Unknown Transaction',
      contractNumber,
      dealerName,
      sourceId,
      message: 'Transaction type could not be classified and is excluded from KPIs.',
    });
  }

  const amountContainer = isCancellation ? document.CancelledAmount : document.WrittenAmount;
  return {
    sourceId,
    snapshotDate: snapshotDate(document, metadata),
    contractNumber,
    contractStatus: text(metadata.ContractStatus).toUpperCase(),
    transactionType,
    activityDate,
    agent: text(metadata.Agent) || 'Unassigned',
    dealer: text(metadata.DealerNumber) || text(metadata.DealerName) || 'Unknown Dealer',
    dealerNumber: text(metadata.DealerNumber),
    dealerName: text(metadata.DealerName),
    product: text(metadata.ProductType) || 'Unmapped Product Type',
    coverageCode: text(metadata.CoverageCode),
    company: text(metadata.Company),
    riskEntity: text(metadata.RiskEntity),
    adminAmount: sumComponents(amountContainer, 'ADMIN', config),
    reserveAmount: sumComponents(amountContainer, 'RESERVE', config),
  };
}

function normalizeWrittenReferenceFromCancellation(
  document: UnknownDocument,
  config: ReportConfig,
): NormalizedContractTransaction | null {
  const metadata = record(document.metadata);
  if (
    classify(
      metadata['Status(NewBusiness,Cancellation,Upgrade,Adjustment)'],
      metadata.ContractStatus,
    ) !== 'Cancellation'
  )
    return null;

  const activityDate = firstDate(metadata.ActivationDate);
  if (!activityDate || activityDate > config.asOfDate) return null;

  const sourceId = idOf(document);
  return {
    sourceId: `${sourceId}:written-reference`,
    snapshotDate: snapshotDate(document, metadata),
    contractNumber: businessId(metadata['Contract#']),
    contractStatus: 'C',
    transactionType: 'NewBusiness',
    activityDate,
    agent: text(metadata.Agent) || 'Unassigned',
    dealer: text(metadata.DealerNumber) || text(metadata.DealerName) || 'Unknown Dealer',
    dealerNumber: text(metadata.DealerNumber),
    dealerName: text(metadata.DealerName),
    product: text(metadata.ProductType) || 'Unmapped Product Type',
    coverageCode: text(metadata.CoverageCode),
    company: text(metadata.Company),
    riskEntity: text(metadata.RiskEntity),
    adminAmount: sumComponents(document.WrittenAmount, 'ADMIN', config),
    reserveAmount: sumComponents(document.WrittenAmount, 'RESERVE', config),
  };
}

function normalizeClaim(
  document: UnknownDocument,
  asOfDate: Date,
  issues: DataQualityIssue[],
): NormalizedClaim | null {
  const sourceId = idOf(document);
  const contractNumber = businessId(document['Contract Number']);
  const dealerName = text(document['Selling Dealer Name'] ?? document.DealerName);
  const status = text(document['Claim Status'] ?? document['Claim Detail Status']);
  const activity = text(document.Activity);
  const paid = number(document['Total Paid Amount']);
  const activityDate = firstDate(document['Date Paid'], document['Claim Date Claim is Reported']);

  if (status.toLowerCase() !== 'paid' && activity.toLowerCase() !== 'payment issued') return null;
  if (paid === 0) return null;
  if (!activityDate) {
    issues.push({
      severity: 'Error',
      category: 'Invalid Claim Date',
      contractNumber,
      dealerName,
      sourceId,
      message: 'Paid claim has no usable paid or reported date.',
    });
    return null;
  }
  if (activityDate > asOfDate) return null;

  const claimNumber = businessId(document['Claim Number']);
  const paidDateKey = text(document['Date Paid'] ?? document['Claim Date Claim is Reported']);
  const paymentParts = [
    claimNumber,
    paidDateKey,
    text(document['Check Number']),
    text(document['Pay Method'] ?? document['Payment Method']),
    text(document['Payee Name']),
    text(document['Loss Code']),
    text(document['RO Number']),
    paid.toFixed(4),
  ];
  return {
    sourceId,
    snapshotDate: snapshotDate(document),
    paymentKey: paymentParts.some(Boolean) ? paymentParts.join('|').toUpperCase() : sourceId,
    claimNumber,
    contractNumber,
    activityDate,
    status,
    paidAmount: paid,
    // Use stable business identifiers so claim dimensions align with contract
    // metadata.Agent and metadata.DealerNumber. Names remain fallbacks only.
    agent:
      text(document['Agent Number'] ?? document.Agent ?? document['Agent Name']) || 'Unassigned',
    dealer:
      text(
        document['Selling Dealer Number'] ??
          document.DealerNumber ??
          document['Selling Dealer Name'],
      ) || 'Unknown Dealer',
    dealerName: text(document['Selling Dealer Name'] ?? document.DealerName),
    product: text(document['Product Type'] ?? document.ProductType) || 'Unmapped Product Type',
    coverageName: text(document['Coverage Name'] ?? document.CoverageName),
    vehicleMake:
      text(document.Make ?? document['Vehicle Make'])
        .replace(/\s+/g, ' ')
        .toUpperCase() || 'UNMAPPED MAKE',
    lossCode: businessId(document['Loss Code']) || UNMAPPED_LOSS_CODE,
    lossCodeDescription: text(document['Loss Code Description']) || UNAVAILABLE_LOSS_DESCRIPTION,
  };
}

function latestSnapshotBy<T>(
  values: T[],
  key: (value: T) => string,
  timestamp: (value: T) => Date,
): T[] {
  const latest = new Map<string, T>();
  for (const value of values) {
    const businessKey = key(value);
    const existing = latest.get(businessKey);
    if (!existing || timestamp(value) >= timestamp(existing)) latest.set(businessKey, value);
  }
  return [...latest.values()];
}

function dimensionMetrics(
  dimension: Dimension,
  transactions: NormalizedContractTransaction[],
  claims: NormalizedClaim[],
  start: Date,
  end: Date,
): DimensionMetric[] {
  const contractLookup = new Map<string, NormalizedContractTransaction>();
  for (const transaction of transactions) {
    if (transaction.transactionType === 'NewBusiness') {
      contractLookup.set(transaction.contractNumber, transaction);
    }
  }

  const claimDimension = (claim: NormalizedClaim): string => {
    // Claims often contain only a coverage code. For product reporting, the
    // originating contract is authoritative so coverage variants consolidate
    // under their shared ProductType.
    if (dimension === 'product') {
      const contractValue = contractLookup.get(claim.contractNumber)?.product;
      if (contractValue) return contractValue;
    }
    return claim[dimension];
  };

  const names = new Set<string>();
  for (const transaction of transactions) {
    if (inRange(transaction.activityDate, start, end)) names.add(transaction[dimension]);
  }
  for (const claim of claims) {
    if (!inRange(claim.activityDate, start, end)) continue;
    const claimName = claimDimension(claim);
    if (
      claimName &&
      !['Unassigned', 'Unknown Dealer', 'Unmapped', 'Unmapped Product Type'].includes(claimName)
    )
      names.add(claimName);
  }
  const rows: DimensionMetric[] = [];
  for (const name of names) {
    const matchingTransactions = transactions.filter((item) => item[dimension] === name);
    const matchingClaims = claims.filter((claim) => {
      if (!inRange(claim.activityDate, start, end)) return false;
      const directValue = claimDimension(claim);
      if (
        directValue &&
        !['Unassigned', 'Unknown Dealer', 'Unmapped', 'Unmapped Product Type'].includes(directValue)
      ) {
        return directValue === name;
      }
      return contractLookup.get(claim.contractNumber)?.[dimension] === name;
    });
    const displayName =
      dimension === 'dealer'
        ? matchingTransactions.find((item) => item.dealerName)?.dealerName ||
          matchingClaims.find((item) => item.dealerName)?.dealerName
        : undefined;
    const relatedAgents =
      dimension === 'dealer'
        ? [
            ...new Set(
              matchingTransactions
                .filter((item) => inRange(item.activityDate, start, end))
                .map((item) => item.agent)
                .filter((agent) => agent && agent !== 'Unassigned'),
            ),
          ].sort((a, b) => a.localeCompare(b))
        : undefined;
    rows.push({
      name,
      displayName,
      relatedAgents,
      ...aggregate(matchingTransactions, matchingClaims, start, end),
    });
  }
  return rows.sort((a, b) => b.netReserve - a.netReserve || a.name.localeCompare(b.name));
}

function claimsInRange(claims: NormalizedClaim[], start: Date, end: Date): NormalizedClaim[] {
  return claims.filter((claim) => inRange(claim.activityDate, start, end));
}

function distinctClaimCount(claims: NormalizedClaim[]): number {
  return new Set(claims.map((claim) => claim.claimNumber || claim.sourceId)).size;
}

function lossCodeKpis(claims: NormalizedClaim[], start: Date, end: Date): LossCodeKpis {
  const selected = claimsInRange(claims, start, end);
  const totalPaid = selected.reduce((sum, claim) => sum + claim.paidAmount, 0);
  const claimCount = distinctClaimCount(selected);
  return {
    totalPaid,
    claimCount,
    lossCodeCount: new Set(selected.map((claim) => claim.lossCode)).size,
    averagePaidPerClaim: claimCount > 0 ? totalPaid / claimCount : null,
  };
}

function preferredLossCodeDescription(claims: NormalizedClaim[]): string {
  const frequencies = new Map<string, number>();
  for (const claim of claims) {
    if (claim.lossCodeDescription === UNAVAILABLE_LOSS_DESCRIPTION) continue;
    frequencies.set(
      claim.lossCodeDescription,
      (frequencies.get(claim.lossCodeDescription) ?? 0) + 1,
    );
  }
  return (
    [...frequencies.entries()].sort(
      ([leftDescription, leftCount], [rightDescription, rightCount]) =>
        rightCount - leftCount || leftDescription.localeCompare(rightDescription),
    )[0]?.[0] ?? UNAVAILABLE_LOSS_DESCRIPTION
  );
}

function topVehicleMakes(
  claims: NormalizedClaim[],
  start: Date,
  end: Date,
  limit = 10,
): VehicleMakeMetric[] {
  const selected = claimsInRange(claims, start, end);
  const totalPaid = selected.reduce((sum, claim) => sum + claim.paidAmount, 0);
  const aggregates = new Map<string, { paidAmount: number; claimNumbers: Set<string> }>();
  for (const claim of selected) {
    const aggregate = aggregates.get(claim.vehicleMake) ?? {
      paidAmount: 0,
      claimNumbers: new Set<string>(),
    };
    aggregate.paidAmount += claim.paidAmount;
    aggregate.claimNumbers.add(claim.claimNumber || claim.sourceId);
    aggregates.set(claim.vehicleMake, aggregate);
  }
  return [...aggregates.entries()]
    .map(([make, aggregate]) => ({
      make,
      paidAmount: aggregate.paidAmount,
      claimCount: aggregate.claimNumbers.size,
      paidShare: totalPaid !== 0 ? aggregate.paidAmount / totalPaid : null,
    }))
    .filter((row) => row.paidAmount > 0)
    .sort(
      (left, right) => right.paidAmount - left.paidAmount || left.make.localeCompare(right.make),
    )
    .slice(0, Math.max(0, limit));
}

function buildLossCodeDashboard(
  claims: NormalizedClaim[],
  currentStart: Date,
  currentEnd: Date,
  yearToDateStart: Date,
  rolling12Start: Date,
): LossCodeDashboard {
  const currentMonth = lossCodeKpis(claims, currentStart, currentEnd);
  const yearToDate = lossCodeKpis(claims, yearToDateStart, currentEnd);
  const rolling12 = lossCodeKpis(claims, rolling12Start, currentEnd);
  const rollingClaims = claimsInRange(claims, rolling12Start, currentEnd);
  const claimsByLossCode = new Map<string, NormalizedClaim[]>();
  for (const claim of rollingClaims) {
    const matching = claimsByLossCode.get(claim.lossCode) ?? [];
    matching.push(claim);
    claimsByLossCode.set(claim.lossCode, matching);
  }

  const rows: LossCodeMetric[] = [...claimsByLossCode.entries()].map(([code, matching]) => {
    const currentMonthPaid = claimsInRange(matching, currentStart, currentEnd).reduce(
      (sum, claim) => sum + claim.paidAmount,
      0,
    );
    const yearToDatePaid = claimsInRange(matching, yearToDateStart, currentEnd).reduce(
      (sum, claim) => sum + claim.paidAmount,
      0,
    );
    const rolling12Paid = matching.reduce((sum, claim) => sum + claim.paidAmount, 0);
    const rolling12ClaimCount = distinctClaimCount(matching);
    return {
      code,
      description: preferredLossCodeDescription(matching),
      coverageNames: [
        ...new Set(matching.map((claim) => claim.coverageName || 'Unmapped Coverage Name')),
      ].sort((left, right) => left.localeCompare(right)),
      currentMonthPaid,
      yearToDatePaid,
      rolling12Paid,
      rolling12PaidShare: rolling12.totalPaid !== 0 ? rolling12Paid / rolling12.totalPaid : null,
      rolling12ClaimCount,
      rolling12AveragePaidPerClaim:
        rolling12ClaimCount > 0 ? rolling12Paid / rolling12ClaimCount : null,
    };
  });
  rows.sort((left, right) =>
    right.rolling12Paid !== left.rolling12Paid
      ? right.rolling12Paid - left.rolling12Paid
      : left.code.localeCompare(right.code),
  );

  return {
    currentMonth,
    yearToDate,
    rolling12,
    rows,
    topVehicleMakes: topVehicleMakes(claims, rolling12Start, currentEnd),
  };
}

export class ReportTransformer {
  constructor(private readonly config: ReportConfig) {}

  transform(
    contractDocuments: UnknownDocument[],
    cancellationDocuments: UnknownDocument[],
    claimDocuments: UnknownDocument[],
    pipelineAudits: PipelineAuditRecord[] = [],
  ): ReportModel {
    // A report run during an open month is always cut off at the end of the
    // previous month. Apply that cutoff during normalization so current-month
    // activity cannot leak into detail, quality, or dashboard tabs.
    const runMonthStart = monthStart(this.config.asOfDate);
    const currentStart = addMonths(runMonthStart, -1);
    const currentEnd = endOfMonth(currentStart);
    const reportingConfig: ReportConfig = { ...this.config, asOfDate: currentEnd };
    const issues: DataQualityIssue[] = [];
    const normalizedContractSnapshots = contractDocuments
      .map((item) => normalizeContract(item, reportingConfig, issues))
      .filter((item): item is NormalizedContractTransaction => item !== null);
    const normalizedCancellationSnapshots = cancellationDocuments
      .map((item) => normalizeContract(item, reportingConfig, issues))
      .filter((item): item is NormalizedContractTransaction => item !== null);

    const contractTransactions = latestSnapshotBy(
      normalizedContractSnapshots,
      (item) => item.contractNumber || item.sourceId,
      (item) => item.snapshotDate,
    );
    const cancellationTransactions = latestSnapshotBy(
      normalizedCancellationSnapshots,
      (item) => item.contractNumber || item.sourceId,
      (item) => item.snapshotDate,
    );
    const normalizedTransactions = [...contractTransactions, ...cancellationTransactions];

    // Preserve the original written side of a canceled contract when the
    // new-business source is absent. This reference belongs to the original
    // ActivationDate cohort and retains status C, so it can never be active.
    const writtenContracts = new Set(
      normalizedTransactions
        .filter((item) => item.transactionType !== 'Cancellation' && item.contractNumber)
        .map((item) => item.contractNumber),
    );
    for (const document of cancellationDocuments) {
      const reference = normalizeWrittenReferenceFromCancellation(document, reportingConfig);
      if (reference?.contractNumber && !writtenContracts.has(reference.contractNumber)) {
        normalizedTransactions.push(reference);
        writtenContracts.add(reference.contractNumber);
      }
    }

    const transactions = latestSnapshotBy(
      normalizedTransactions,
      (item) =>
        item.contractNumber
          ? `${item.contractNumber.trim().toUpperCase()}|${item.transactionType}`
          : item.sourceId,
      (item) => item.snapshotDate,
    );
    const normalizedClaims = claimDocuments
      .map((item) => normalizeClaim(item, currentEnd, issues))
      .filter((item): item is NormalizedClaim => item !== null);
    const claims = latestSnapshotBy(
      normalizedClaims,
      (item) => item.paymentKey || item.sourceId,
      (item) => item.snapshotDate,
    );

    // Reports run at the beginning of a month, so every executive KPI must use
    // the latest fully closed month. The monthly comparison is year-over-year,
    // not a comparison with the immediately preceding month.
    const priorMonthStart = addMonths(currentStart, -12);
    const priorMonthEnd = endOfMonth(priorMonthStart);
    const ytdStart = new Date(currentStart.getFullYear(), 0, 1);
    const priorYtdStart = new Date(currentStart.getFullYear() - 1, 0, 1);
    const priorYtdEnd = endOfMonth(addMonths(currentStart, -12));
    const rollingStart = addMonths(currentStart, -11);
    const priorRollingStart = addMonths(rollingStart, -12);
    const priorRollingEnd = endOfMonth(addMonths(currentStart, -12));
    const priorCalendarYearStart = new Date(currentStart.getFullYear() - 1, 0, 1);
    const priorCalendarYearEnd = new Date(currentStart.getFullYear() - 1, 11, 31, 23, 59, 59, 999);

    const monthly: MonthlyMetric[] = [];
    for (
      let cursor = addMonths(currentStart, -23);
      cursor <= currentStart;
      cursor = addMonths(cursor, 1)
    ) {
      const end = periodKey(cursor) === periodKey(currentStart) ? currentEnd : endOfMonth(cursor);
      monthly.push({
        period: periodKey(cursor),
        periodStart: cursor,
        ...aggregate(transactions, claims, cursor, end),
      });
    }

    return {
      generatedAt: new Date(),
      asOfDate: currentEnd,
      currentMonth: {
        current: aggregate(transactions, claims, currentStart, currentEnd),
        prior: aggregate(transactions, claims, priorMonthStart, priorMonthEnd),
        currentStart,
        currentEnd,
        priorStart: priorMonthStart,
        priorEnd: priorMonthEnd,
      },
      yearToDate: {
        current: aggregate(transactions, claims, ytdStart, currentEnd),
        prior: aggregate(transactions, claims, priorYtdStart, priorYtdEnd),
        currentStart: ytdStart,
        currentEnd,
        priorStart: priorYtdStart,
        priorEnd: priorYtdEnd,
      },
      rolling12: {
        current: aggregate(transactions, claims, rollingStart, currentEnd),
        prior: aggregate(transactions, claims, priorRollingStart, priorRollingEnd),
        currentStart: rollingStart,
        currentEnd,
        priorStart: priorRollingStart,
        priorEnd: priorRollingEnd,
      },
      priorCalendarYear: {
        values: aggregate(transactions, claims, priorCalendarYearStart, priorCalendarYearEnd),
        start: priorCalendarYearStart,
        end: priorCalendarYearEnd,
      },
      monthly,
      dealers: dimensionMetrics('dealer', transactions, claims, rollingStart, currentEnd),
      agents: dimensionMetrics('agent', transactions, claims, rollingStart, currentEnd).filter(
        (agent) => agent.name.trim().toLowerCase() !== 'test',
      ),
      products: dimensionMetrics('product', transactions, claims, currentStart, currentEnd),
      lossCodeDashboard: buildLossCodeDashboard(
        claims,
        currentStart,
        currentEnd,
        ytdStart,
        rollingStart,
      ),
      contractTransactions: transactions,
      claims,
      dataQualityIssues: issues,
      pipelineAudits,
      sourceCounts: {
        contractDocuments: contractDocuments.length,
        cancellationDocuments: cancellationDocuments.length,
        claimDocuments: claimDocuments.length,
        uniqueContractTransactions: transactions.length,
        uniqueClaims: claims.length,
      },
    };
  }
}
