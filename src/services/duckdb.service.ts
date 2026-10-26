import * as duckdb from 'duckdb';
import {
  type DimensionMetric,
  type MetricValues,
  type NormalizedClaim,
  type NormalizedContractTransaction,
  type PipelineAuditRecord,
} from '../models/report.types';

export type DuckDBParam = string | number | boolean | Date | null;

export interface DuckDBQueryRow {
  [columnName: string]: string | number | boolean | Date | null | bigint;
}

export interface ParquetExportResult {
  contractsPath: string;
  claimsPath: string;
}

export interface AuditVerificationResult {
  isClean: boolean;
  contractCount: number;
  cancellationCount: number;
  claimCount: number;
  auditRecords: PipelineAuditRecord[];
  discrepancies: Array<{
    jobType: string;
    portalCount: number;
    uploadedCount: number;
    duckdbCount: number;
    variance: number;
    status: string;
  }>;
}

/**
 * Enterprise DuckDB Analytical Service
 * Provides high-throughput streaming ingestion, vectorized snapshot deduplication,
 * volumetric audit reconciliation, and sub-second analytical aggregations.
 */
export class DuckDBService {
  private db: duckdb.Database | null = null;
  private connection: duckdb.Connection | null = null;

  constructor(private readonly dbPath: string = ':memory:') {}

  /**
   * Initializes the DuckDB instance, establishes connection, and configures memory bounds.
   */
  async init(): Promise<void> {
    return new Promise((resolve, reject) => {
      this.db = new duckdb.Database(this.dbPath, (err) => {
        if (err) {
          return reject(new Error(`Failed to initialize DuckDB at ${this.dbPath}: ${err.message}`));
        }
        if (!this.db) {
          return reject(new Error('DuckDB database handle is null after initialization.'));
        }

        this.connection = this.db.connect();
        this.run('SET threads TO 4;')
          .then(() => this.run("SET max_memory = '4GB';"))
          .then(() => this.createSchema())
          .then(() => resolve())
          .catch(reject);
      });
    });
  }

  /**
   * Defines the typed columnar tables for raw contract snapshots, claim events, and audit logs.
   */
  private async createSchema(): Promise<void> {
    await this.run(`
      CREATE TABLE IF NOT EXISTS raw_contracts (
        sourceId VARCHAR,
        snapshotDate TIMESTAMP,
        contractNumber VARCHAR,
        contractStatus VARCHAR,
        transactionType VARCHAR,
        activityDate DATE,
        effectiveDate DATE,
        dealerNumber VARCHAR,
        dealerName VARCHAR,
        agent VARCHAR,
        agentName VARCHAR,
        product VARCHAR,
        coverageCode VARCHAR,
        company VARCHAR,
        riskEntity VARCHAR,
        adminAmount DOUBLE,
        reserveAmount DOUBLE,
        grossIncomeAmount DOUBLE,
        scheduleJson VARCHAR,
        componentsJson VARCHAR
      );

      CREATE TABLE IF NOT EXISTS raw_claims (
        sourceId VARCHAR,
        snapshotDate TIMESTAMP,
        paymentKey VARCHAR,
        claimNumber VARCHAR,
        contractNumber VARCHAR,
        activityDate DATE,
        status VARCHAR,
        paidAmount DOUBLE,
        agent VARCHAR,
        agentName VARCHAR,
        dealer VARCHAR,
        dealerName VARCHAR,
        product VARCHAR,
        coverageName VARCHAR,
        vehicleMake VARCHAR,
        lossCode VARCHAR,
        lossCodeDescription VARCHAR
      );

      CREATE TABLE IF NOT EXISTS pipeline_audits (
        jobType VARCHAR,
        executionTimestamp TIMESTAMP,
        executionDateStr VARCHAR,
        startDate VARCHAR,
        endDate VARCHAR,
        portalCount BIGINT,
        uniqueCount BIGINT,
        processedCount BIGINT,
        uploadedCount BIGINT,
        isMatch BOOLEAN,
        variance BIGINT,
        status VARCHAR,
        summary VARCHAR,
        sourceFile VARCHAR
      );

      CREATE INDEX IF NOT EXISTS idx_contracts_lookup ON raw_contracts (contractNumber, transactionType);
      CREATE INDEX IF NOT EXISTS idx_claims_lookup ON raw_claims (claimNumber, paymentKey);
    `);
  }

  /**
   * Executes a parameter-free DDL/DML statement.
   */
  async run(sql: string): Promise<void> {
    const conn = this.ensureConnection();
    return new Promise((resolve, reject) => {
      conn.exec(sql, (err) => {
        if (err) return reject(new Error(`DuckDB exec error [${sql}]: ${err.message}`));
        resolve();
      });
    });
  }

  /**
   * Executes a parameterized query returning typed row objects.
   */
  async query<T extends DuckDBQueryRow = DuckDBQueryRow>(
    sql: string,
    params: DuckDBParam[] = [],
  ): Promise<T[]> {
    const conn = this.ensureConnection();
    return new Promise((resolve, reject) => {
      conn.all(sql, ...params, (err: duckdb.DuckDbError | null, res: duckdb.TableData) => {
        if (err) return reject(new Error(`DuckDB query error [${sql}]: ${err.message}`));
        resolve((res ?? []) as unknown as T[]);
      });
    });
  }

  /**
   * Batch inserts streaming normalized contract transactions using a prepared statement
   * wrapped in a fast single transaction.
   */
  async insertContractsBatch(contracts: readonly NormalizedContractTransaction[]): Promise<void> {
    if (contracts.length === 0) return;
    const conn = this.ensureConnection();

    await this.run('BEGIN TRANSACTION;');
    try {
      const stmt = conn.prepare(`
        INSERT INTO raw_contracts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
      `);

      for (const c of contracts) {
        stmt.run(
          c.sourceId,
          c.snapshotDate,
          c.contractNumber,
          c.contractStatus,
          c.transactionType,
          c.activityDate ? c.activityDate.toISOString().slice(0, 10) : null,
          c.effectiveDate ? c.effectiveDate.toISOString().slice(0, 10) : null,
          c.dealerNumber,
          c.dealerName,
          c.agent,
          c.agentName ?? '',
          c.product,
          c.coverageCode,
          c.company,
          c.riskEntity,
          c.adminAmount,
          c.reserveAmount,
          c.grossIncomeAmount,
          c.earningSchedule ? JSON.stringify(c.earningSchedule) : null,
          c.components ? JSON.stringify(c.components) : null,
        );
      }

      await new Promise<void>((resolve, reject) => {
        stmt.finalize((err) => {
          if (err) reject(err);
          else resolve();
        });
      });

      await this.run('COMMIT;');
    } catch (error) {
      await this.run('ROLLBACK;');
      throw error;
    }
  }

  /**
   * Batch inserts streaming normalized claims using a prepared statement.
   */
  async insertClaimsBatch(claims: readonly NormalizedClaim[]): Promise<void> {
    if (claims.length === 0) return;
    const conn = this.ensureConnection();

    await this.run('BEGIN TRANSACTION;');
    try {
      const stmt = conn.prepare(`
        INSERT INTO raw_claims VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
      `);

      for (const cl of claims) {
        stmt.run(
          cl.sourceId,
          cl.snapshotDate,
          cl.paymentKey,
          cl.claimNumber,
          cl.contractNumber,
          cl.activityDate ? cl.activityDate.toISOString().slice(0, 10) : null,
          cl.status,
          cl.paidAmount,
          cl.agent,
          cl.agentName ?? '',
          cl.dealer,
          cl.dealerName,
          cl.product,
          cl.coverageName,
          cl.vehicleMake,
          cl.lossCode,
          cl.lossCodeDescription,
        );
      }

      await new Promise<void>((resolve, reject) => {
        stmt.finalize((err) => {
          if (err) reject(err);
          else resolve();
        });
      });

      await this.run('COMMIT;');
    } catch (error) {
      await this.run('ROLLBACK;');
      throw error;
    }
  }

  /**
   * Ingests latest reconciliation records from AuditDB.DataReconciliationAudit.
   */
  async insertPipelineAudits(audits: readonly PipelineAuditRecord[]): Promise<void> {
    if (audits.length === 0) return;
    const conn = this.ensureConnection();

    await this.run('BEGIN TRANSACTION;');
    try {
      const stmt = conn.prepare(`
        INSERT INTO pipeline_audits VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
      `);

      for (const a of audits) {
        stmt.run(
          a.jobType,
          a.executionTimestamp,
          a.executionDateStr,
          a.dateRange.startDate,
          a.dateRange.endDate,
          a.counts.portalCount,
          a.counts.uniqueCount ?? a.counts.uploadedCount,
          a.counts.processedCount,
          a.counts.uploadedCount,
          a.reconciliation.isMatch,
          a.reconciliation.portalVsProcessedDiff,
          a.reconciliation.status,
          a.reconciliation.summary,
          a.fileMetadata.fileName,
        );
      }

      await new Promise<void>((resolve, reject) => {
        stmt.finalize((err) => {
          if (err) reject(err);
          else resolve();
        });
      });

      await this.run('COMMIT;');
    } catch (error) {
      await this.run('ROLLBACK;');
      throw error;
    }
  }

  /**
   * Creates the deduplication views using vectorized analytical window functions.
   */
  async createDeduplicationViews(): Promise<void> {
    await this.run(`
      CREATE OR REPLACE VIEW v_contracts_deduped AS
      SELECT * EXCLUDE (rn) FROM (
        SELECT *,
          ROW_NUMBER() OVER (
            PARTITION BY contractNumber, transactionType
            ORDER BY snapshotDate DESC
          ) as rn
        FROM raw_contracts
        WHERE contractNumber IS NOT NULL AND contractNumber != ''
      ) WHERE rn = 1;

      CREATE OR REPLACE VIEW v_claims_deduped AS
      SELECT * EXCLUDE (rn) FROM (
        SELECT *,
          ROW_NUMBER() OVER (
            PARTITION BY COALESCE(NULLIF(paymentKey, ''), claimNumber, sourceId)
            ORDER BY snapshotDate DESC
          ) as rn
        FROM raw_claims
      ) WHERE rn = 1;
    `);
  }

  /**
   * Cross-checks DuckDB ingested counts against upstream AuditDB counts.
   * Proves zero volumetric data leaks across the entire pipeline.
   */
  async verifyVolumetricIntegrity(audits: readonly PipelineAuditRecord[]): Promise<AuditVerificationResult> {
    const [contractStats] = await this.query<{ writtenCount: bigint; cancelCount: bigint }>(`
      SELECT 
        COUNT(CASE WHEN transactionType != 'Cancellation' THEN 1 END) AS writtenCount,
        COUNT(CASE WHEN transactionType = 'Cancellation' THEN 1 END) AS cancelCount
      FROM raw_contracts;
    `);

    const [claimStats] = await this.query<{ claimCount: bigint }>(`
      SELECT COUNT(*) AS claimCount FROM raw_claims;
    `);

    const contractCount = Number(contractStats?.writtenCount ?? BigInt(0));
    const cancellationCount = Number(contractStats?.cancelCount ?? BigInt(0));
    const claimCount = Number(claimStats?.claimCount ?? BigInt(0));

    const discrepancies: AuditVerificationResult['discrepancies'] = [];
    let isClean = true;

    for (const audit of audits) {
      let duckdbCount = 0;
      if (audit.jobType === 'Contract') duckdbCount = contractCount;
      else if (audit.jobType === 'Cancel') duckdbCount = cancellationCount;
      else if (audit.jobType === 'Claim') duckdbCount = claimCount;

      const variance = duckdbCount - audit.counts.uploadedCount;
      const isMatch = variance === 0 && (audit.reconciliation.status === 'PASSED' || audit.reconciliation.isMatch);

      if (!isMatch) isClean = false;

      discrepancies.push({
        jobType: audit.jobType,
        portalCount: audit.counts.portalCount,
        uploadedCount: audit.counts.uploadedCount,
        duckdbCount,
        variance,
        status: isMatch ? 'PASSED' : 'DISCREPANCY',
      });
    }

    return {
      isClean,
      contractCount,
      cancellationCount,
      claimCount,
      auditRecords: [...audits],
      discrepancies,
    };
  }

  /**
   * Computes executive financial metrics for a specific date range.
   */
  async getPeriodMetrics(startDate: Date, endDate: Date): Promise<MetricValues> {
    const startIso = startDate.toISOString().slice(0, 10);
    const endIso = endDate.toISOString().slice(0, 10);

    const [contractSums] = await this.query<{
      writtenCount: bigint;
      cancelCount: bigint;
      adminWritten: number;
      adminCancelled: number;
      reserveWritten: number;
      reserveCancelled: number;
      grossIncomeWritten: number;
      grossIncomeCancelled: number;
    }>(`
      SELECT
        COUNT(DISTINCT CASE WHEN transactionType != 'Cancellation' THEN contractNumber END) AS writtenCount,
        COUNT(DISTINCT CASE WHEN transactionType = 'Cancellation' THEN contractNumber END) AS cancelCount,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN adminAmount ELSE 0 END), 0) AS adminWritten,
        COALESCE(SUM(CASE WHEN transactionType = 'Cancellation' THEN ABS(adminAmount) ELSE 0 END), 0) AS adminCancelled,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN reserveAmount ELSE 0 END), 0) AS reserveWritten,
        COALESCE(SUM(CASE WHEN transactionType = 'Cancellation' THEN ABS(reserveAmount) ELSE 0 END), 0) AS reserveCancelled,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN grossIncomeAmount ELSE 0 END), 0) AS grossIncomeWritten,
        COALESCE(SUM(CASE WHEN transactionType = 'Cancellation' THEN ABS(grossIncomeAmount) ELSE 0 END), 0) AS grossIncomeCancelled
      FROM v_contracts_deduped
      WHERE activityDate >= '${startIso}' AND activityDate <= '${endIso}';
    `);

    const [activeRow] = await this.query<{ activeCount: bigint; netAdminActive: number }>(`
      WITH latest_contract AS (
        SELECT * EXCLUDE (rn) FROM (
          SELECT *,
            ROW_NUMBER() OVER (PARTITION BY contractNumber ORDER BY snapshotDate DESC) as rn
          FROM raw_contracts
          WHERE contractNumber IS NOT NULL AND contractNumber != ''
        ) WHERE rn = 1
      )
      SELECT 
        COUNT(*) AS activeCount,
        COALESCE(SUM(adminAmount), 0) AS netAdminActive
      FROM latest_contract
      WHERE contractStatus = 'A' AND transactionType != 'Cancellation'
        AND activityDate >= '${startIso}' AND activityDate <= '${endIso}';
    `);

    const [claimSums] = await this.query<{ totalPaid: number; uniqueClaims: bigint }>(`
      SELECT
        COALESCE(SUM(paidAmount), 0) AS totalPaid,
        COUNT(DISTINCT COALESCE(NULLIF(claimNumber, ''), sourceId)) AS uniqueClaims
      FROM v_claims_deduped
      WHERE activityDate >= '${startIso}' AND activityDate <= '${endIso}';
    `);

    const contractsWritten = Number(contractSums?.writtenCount ?? BigInt(0));
    const contractsCancelled = Number(contractSums?.cancelCount ?? BigInt(0));
    const netContracts = contractsWritten - contractsCancelled;
    const activeContracts = Number(activeRow?.activeCount ?? BigInt(0));
    const adminWritten = contractSums?.adminWritten ?? 0;
    const adminCancelled = contractSums?.adminCancelled ?? 0;
    const netAdmin = activeRow?.netAdminActive ?? (adminWritten - adminCancelled);
    const reserveWritten = contractSums?.reserveWritten ?? 0;
    const reserveCancelled = contractSums?.reserveCancelled ?? 0;
    const netReserve = reserveWritten - reserveCancelled;
    const claimsPaid = claimSums?.totalPaid ?? 0;
    const claimCount = Number(claimSums?.uniqueClaims ?? BigInt(0));
    const grossIncomeWritten = contractSums?.grossIncomeWritten ?? 0;
    const grossIncomeCancelled = contractSums?.grossIncomeCancelled ?? 0;
    const grossIncome = grossIncomeWritten - grossIncomeCancelled;

    const earnedReserve = netReserve;
    const underwritingProfit = earnedReserve - claimsPaid;
    const earnedLossRatio = earnedReserve > 0 ? claimsPaid / earnedReserve : null;
    const cancellationRate = contractsWritten > 0 ? contractsCancelled / contractsWritten : null;
    const adminPerContract = activeContracts > 0 ? netAdmin / activeContracts : null;

    return {
      contractsWritten,
      activeContracts,
      contractsCancelled,
      netContracts,
      adminWritten,
      adminCancelled,
      netAdmin,
      reserveWritten,
      reserveCancelled,
      netReserve,
      earnedReserve,
      claimsPaid,
      claimCount,
      underwritingProfit,
      grossIncomeWritten,
      grossIncomeCancelled,
      grossIncome,
      earnedLossRatio,
      cancellationRate,
      adminPerContract,
    };
  }

  /**
   * Ranks top dealers by Net Written Reserve in a date range.
   */
  async getTopDealers(startDate: Date, endDate: Date, limit = 20): Promise<DimensionMetric[]> {
    const startIso = startDate.toISOString().slice(0, 10);
    const endIso = endDate.toISOString().slice(0, 10);

    const rows = await this.query<{
      dealerNumber: string;
      dealerName: string;
      writtenReserve: number;
      cancelledReserve: number;
      netReserve: number;
      writtenAdmin: number;
      cancelledAdmin: number;
      netAdmin: number;
    }>(`
      SELECT
        dealerNumber,
        dealerName,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN reserveAmount ELSE 0 END), 0) AS writtenReserve,
        COALESCE(SUM(CASE WHEN transactionType = 'Cancellation' THEN ABS(reserveAmount) ELSE 0 END), 0) AS cancelledReserve,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN reserveAmount ELSE -ABS(reserveAmount) END), 0) AS netReserve,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN adminAmount ELSE 0 END), 0) AS writtenAdmin,
        COALESCE(SUM(CASE WHEN transactionType = 'Cancellation' THEN ABS(adminAmount) ELSE 0 END), 0) AS cancelledAdmin,
        COALESCE(SUM(CASE WHEN transactionType != 'Cancellation' THEN adminAmount ELSE -ABS(adminAmount) END), 0) AS netAdmin
      FROM v_contracts_deduped
      WHERE activityDate >= '${startIso}' AND activityDate <= '${endIso}'
        AND dealerNumber IS NOT NULL AND dealerNumber != ''
      GROUP BY dealerNumber, dealerName
      ORDER BY netReserve DESC
      LIMIT ${Math.max(1, limit)};
    `);

    return rows.map((r) => ({
      name: r.dealerNumber,
      displayName: r.dealerName,
      contractsWritten: 0,
      activeContracts: 0,
      contractsCancelled: 0,
      netContracts: 0,
      adminWritten: r.writtenAdmin,
      adminCancelled: r.cancelledAdmin,
      netAdmin: r.netAdmin,
      reserveWritten: r.writtenReserve,
      reserveCancelled: r.cancelledReserve,
      netReserve: r.netReserve,
      earnedReserve: r.netReserve,
      claimsPaid: 0,
      claimCount: 0,
      underwritingProfit: r.netReserve,
      grossIncomeWritten: 0,
      grossIncomeCancelled: 0,
      grossIncome: 0,
      earnedLossRatio: null,
      cancellationRate: null,
      adminPerContract: null,
    }));
  }

  /**
   * Exports deduplicated tables to compressed Parquet files for Vercel / serverless storage.
   */
  async exportToParquet(outputDir: string): Promise<ParquetExportResult> {
    const normalizedDir = outputDir.replace(/\\/g, '/');
    const contractsPath = `${normalizedDir}/contracts_deduped.parquet`;
    const claimsPath = `${normalizedDir}/claims_deduped.parquet`;

    await this.run(`
      COPY (SELECT * FROM v_contracts_deduped) TO '${contractsPath}' (FORMAT PARQUET, COMPRESSION ZSTD);
      COPY (SELECT * FROM v_claims_deduped) TO '${claimsPath}' (FORMAT PARQUET, COMPRESSION ZSTD);
    `);

    return { contractsPath, claimsPath };
  }

  /**
   * Safely closes the database connection.
   */
  async close(): Promise<void> {
    return new Promise((resolve, reject) => {
      if (this.connection) {
        this.connection.close((err) => {
          if (err) return reject(err);
          this.connection = null;
          if (this.db) {
            this.db.close((dbErr) => {
              if (dbErr) return reject(dbErr);
              this.db = null;
              resolve();
            });
          } else {
            resolve();
          }
        });
      } else {
        resolve();
      }
    });
  }

  private ensureConnection(): duckdb.Connection {
    if (!this.connection) {
      throw new Error('DuckDBService has not been initialized. Call init() first.');
    }
    return this.connection;
  }
}
