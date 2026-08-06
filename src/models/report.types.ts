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

export interface DataQualityIssue {
  severity: 'Warning' | 'Error';
  category: string;
  sourceId: string;
  message: string;
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
  contractTransactions: NormalizedContractTransaction[];
  claims: NormalizedClaim[];
  dataQualityIssues: DataQualityIssue[];
  sourceCounts: {
    contractDocuments: number;
    cancellationDocuments: number;
    claimDocuments: number;
    uniqueContractTransactions: number;
    uniqueClaims: number;
  };
}
