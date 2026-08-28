import { type ObjectId } from 'mongodb';

export type MongoIdentifier = ObjectId | string | { $oid?: string };
export type UnknownDocument = Record<string, unknown> & { _id?: MongoIdentifier };

export interface ReportConfig {
  companyName: string;
  asOfDate: Date;
  outputDirectory: string;
  topDealerCount: number;
  warningLossRatio: number;
  highLossRatio: number;
  excludedComponentCodes: ReadonlySet<string>;
}

export interface MongoSourceConfig {
  contractDb: string;
  contractCollection: string;
  cancellationDb: string;
  cancellationCollection: string;
  claimDb: string;
  claimCollection: string;
  auditDb?: string;
  auditCollection?: string;
}

export interface EmailConfig {
  enabled: boolean;
  apiKey: string;
  from: string;
  to: string[];
  cc: string[];
}

export type TransactionType = 'NewBusiness' | 'Cancellation' | 'Adjustment' | 'Upgrade' | 'Unknown';

export interface NormalizedContractTransaction {
  sourceId: string;
  snapshotDate: Date;
  contractNumber: string;
  contractStatus: string;
  transactionType: TransactionType;
  activityDate: Date;
  agent: string;
  dealer: string;
  dealerNumber: string;
  dealerName: string;
  product: string;
  coverageCode: string;
  company: string;
  riskEntity: string;
  adminAmount: number;
  reserveAmount: number;
}

export interface NormalizedClaim {
  sourceId: string;
  snapshotDate: Date;
  paymentKey: string;
  claimNumber: string;
  contractNumber: string;
  activityDate: Date;
  status: string;
  paidAmount: number;
  agent: string;
  dealer: string;
  dealerName: string;
  product: string;
  coverageName: string;
  vehicleMake: string;
  lossCode: string;
  lossCodeDescription: string;
}

export interface MetricValues {
  contractsWritten: number;
  activeContracts: number;
  contractsCancelled: number;
  netContracts: number;
  adminWritten: number;
  adminCancelled: number;
  netAdmin: number;
  reserveWritten: number;
  reserveCancelled: number;
  netReserve: number;
  claimsPaid: number;
  claimCount: number;
  paidLossRatio: number | null;
  cancellationRate: number | null;
  adminPerContract: number | null;
}

export interface PeriodComparison {
  current: MetricValues;
  prior: MetricValues;
  currentStart: Date;
  currentEnd: Date;
  priorStart: Date;
  priorEnd: Date;
}

export interface ReportingPeriod {
  values: MetricValues;
  start: Date;
  end: Date;
}

export interface MonthlyMetric extends MetricValues {
  period: string;
  periodStart: Date;
}

export interface DimensionMetric extends MetricValues {
  name: string;
  displayName?: string;
  relatedAgents?: string[];
}

export interface LossCodeKpis {
  totalPaid: number;
  claimCount: number;
  lossCodeCount: number;
  averagePaidPerClaim: number | null;
}

export interface LossCodeMetric {
  code: string;
  description: string;
  coverageNames: string[];
  currentMonthPaid: number;
  yearToDatePaid: number;
  rolling12Paid: number;
  rolling12PaidShare: number | null;
  rolling12ClaimCount: number;
  rolling12AveragePaidPerClaim: number | null;
}

export interface VehicleMakeMetric {
  make: string;
  paidAmount: number;
  claimCount: number;
  paidShare: number | null;
}

export interface LossCodeDashboard {
  currentMonth: LossCodeKpis;
  yearToDate: LossCodeKpis;
  rolling12: LossCodeKpis;
  rows: LossCodeMetric[];
  topVehicleMakes: VehicleMakeMetric[];
}

export interface DataQualityIssue {
  severity: 'Warning' | 'Error';
  category: string;
  contractNumber: string;
  dealerName: string;
  sourceId: string;
  message: string;
}

export interface PipelineAuditRecord {
  jobType: 'Contract' | 'Cancel' | 'Claim' | string;
  executionTimestamp: Date;
  executionDateStr: string;
  dateRange: {
    startDate: string;
    endDate: string;
  };
  counts: {
    portalCount: number;
    processedCount: number;
    uploadedCount: number;
  };
  reconciliation: {
    isMatch: boolean;
    portalVsProcessedDiff: number;
    processedVsUploadedDiff: number;
    status: 'PASSED' | 'DISCREPANCY' | 'FAILED' | string;
    summary: string;
  };
  fileMetadata: {
    fileName: string;
  };
  systemInfo?: {
    environment?: string;
    source?: string;
  };
}

export interface ReportModel {
  generatedAt: Date;
  asOfDate: Date;
  currentMonth: PeriodComparison;
  yearToDate: PeriodComparison;
  rolling12: PeriodComparison;
  priorCalendarYear: ReportingPeriod;
  monthly: MonthlyMetric[];
  dealers: DimensionMetric[];
  agents: DimensionMetric[];
  products: DimensionMetric[];
  lossCodeDashboard: LossCodeDashboard;
  contractTransactions: NormalizedContractTransaction[];
  claims: NormalizedClaim[];
  dataQualityIssues: DataQualityIssue[];
  pipelineAudits: PipelineAuditRecord[];
  sourceCounts: {
    contractDocuments: number;
    cancellationDocuments: number;
    claimDocuments: number;
    uniqueContractTransactions: number;
    uniqueClaims: number;
  };
}
