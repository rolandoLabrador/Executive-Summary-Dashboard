import { type MongoService } from '../services/mongo.service';
import {
  type MongoSourceConfig,
  type PipelineAuditRecord,
  type UnknownDocument,
} from '../models/report.types';

function toDate(val: unknown): Date {
  if (val instanceof Date) return val;
  if (typeof val === 'string' || typeof val === 'number') {
    const parsed = new Date(val);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  return new Date();
}

function toNumber(val: unknown): number {
  if (typeof val === 'number') return val;
  if (typeof val === 'string') {
    const parsed = Number(val.replace(/,/g, ''));
    if (!Number.isNaN(parsed)) return parsed;
  }
  if (typeof val === 'object' && val !== null && '$numberInt' in val) {
    return Number((val as { $numberInt: string }).$numberInt);
  }
  return 0;
}

function normalizeAuditRecord(doc: Record<string, unknown>): PipelineAuditRecord {
  const counts = (doc.counts as Record<string, unknown>) || {};
  const reconciliation = (doc.reconciliation as Record<string, unknown>) || {};
  const dateRange = (doc.dateRange as Record<string, unknown>) || {};
  const fileMetadata = (doc.fileMetadata as Record<string, unknown>) || {};
  const systemInfo = (doc.systemInfo as Record<string, unknown>) || {};

  return {
    jobType: String(doc.jobType || 'Unknown'),
    executionTimestamp: toDate(doc.executionTimestamp),
    executionDateStr: String(doc.executionDateStr || ''),
    dateRange: {
      startDate: String(dateRange.startDate || ''),
      endDate: String(dateRange.endDate || ''),
    },
    counts: {
      portalCount: toNumber(counts.portalCount),
      processedCount: toNumber(counts.processedCount),
      uploadedCount: toNumber(counts.uploadedCount),
    },
    reconciliation: {
      isMatch: Boolean(reconciliation.isMatch),
      portalVsProcessedDiff: toNumber(reconciliation.portalVsProcessedDiff),
      processedVsUploadedDiff: toNumber(reconciliation.processedVsUploadedDiff),
      status: String(reconciliation.status || 'UNKNOWN'),
      summary: String(reconciliation.summary || ''),
    },
    fileMetadata: {
      fileName: String(fileMetadata.fileName || ''),
    },
    systemInfo: {
      environment: systemInfo.environment ? String(systemInfo.environment) : undefined,
      source: systemInfo.source ? String(systemInfo.source) : undefined,
    },
  };
}

export class DataRepository {
  constructor(
    private readonly mongoService: MongoService,
    private readonly sources: MongoSourceConfig,
  ) {}

  /**
   * Extracts raw contract data from ContractDataDB
   */
  async getContracts(): Promise<UnknownDocument[]> {
    return this.readCollection(this.sources.contractDb, this.sources.contractCollection);
  }

  /**
   * Extracts raw claim data from ClaimDataDB
   */
  async getClaims(): Promise<UnknownDocument[]> {
    return this.readCollection(this.sources.claimDb, this.sources.claimCollection);
  }

  /**
   * Extracts raw cancellation data from CancelDataDB
   */
  async getCancellations(): Promise<UnknownDocument[]> {
    return this.readCollection(this.sources.cancellationDb, this.sources.cancellationCollection);
  }

  /**
   * Extracts the most recent reconciliation audit records per jobType from AuditDB.DataReconciliationAudit
   */
  async getLatestReconciliationAudits(): Promise<PipelineAuditRecord[]> {
    const database = this.sources.auditDb || 'AuditDB';
    const collectionName = this.sources.auditCollection || 'DataReconciliationAudit';
    try {
      const collection = this.mongoService
        .getDb(database)
        .collection<Record<string, unknown>>(collectionName);

      const jobTypes = ['Contract', 'Cancel', 'Claim'];
      const auditPromises = jobTypes.map(async (jobType) => {
        const doc = await collection
          .find({ jobType })
          .sort({ executionTimestamp: -1, _id: -1 })
          .limit(1)
          .next();
        return doc ? normalizeAuditRecord(doc) : null;
      });

      const results = await Promise.all(auditPromises);
      return results.filter((item): item is PipelineAuditRecord => item !== null);
    } catch (error) {
      console.warn(`Could not load reconciliation audits from ${database}.${collectionName}:`, error);
      return [];
    }
  }

  private async readCollection(
    database: string,
    collectionName: string,
  ): Promise<UnknownDocument[]> {
    const collection = this.mongoService
      .getDb(database)
      .collection<UnknownDocument>(collectionName);
    // Project out known customer PII at the source. Field variants are intentionally included.
    const projection = {
      CustomerAddress1: 0,
      CustomerAddress2: 0,
      CustomerCity: 0,
      CustomerState: 0,
      CustomerZipCode: 0,
      CustomerCountry: 0,
      CustomerEmail: 0,
      CustomerFirstName: 0,
      CustomerLastName: 0,
      CustomerPhoneNumber: 0,
      CustomerSecondaryPhone: 0,
      'Customer Address 1': 0,
      'Customer Address 2': 0,
      'Customer City': 0,
      'Customer State': 0,
      'Customer Zip': 0,
      'Customer Country': 0,
      'Customer Email': 0,
      'Customer First Name': 0,
      'Customer Last name': 0,
      'Customer Phone': 0,
      VIN: 0,
    };
    return collection.find({}, { projection }).toArray();
  }
}
