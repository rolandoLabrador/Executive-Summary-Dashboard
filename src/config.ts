import * as path from 'path';
import { type EmailConfig, type MongoSourceConfig, type ReportConfig } from './models/report.types';

function required(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function positiveInteger(name: string, fallback: number): number {
  const raw = process.env[name];
  if (!raw) return fallback;
  const value = Number.parseInt(raw, 10);
  if (!Number.isInteger(value) || value <= 0) {
    throw new Error(`${name} must be a positive integer`);
  }
  return value;
}

function ratio(name: string, fallback: number): number {
  const raw = process.env[name];
  if (!raw) return fallback;
  const value = Number(raw);
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`${name} must be a non-negative number`);
  }
  return value;
}

export function parseAsOfDate(args: string[]): Date {
  const index = args.indexOf('--as-of');
  const raw = index >= 0 ? args[index + 1] : undefined;
  const date = raw ? new Date(`${raw}T23:59:59.999`) : new Date();
  if (Number.isNaN(date.getTime())) {
    throw new Error('--as-of must use YYYY-MM-DD format');
  }
  return date;
}

export function loadMongoSourceConfig(): MongoSourceConfig {
  return {
    contractDb: process.env.CONTRACT_DB?.trim() || 'ContractDataDB',
    contractCollection: process.env.CONTRACT_COLLECTION?.trim() || 'ContractDatd',
    cancellationDb: process.env.CANCEL_DB?.trim() || 'CancelDataDB',
    cancellationCollection: process.env.CANCEL_COLLECTION?.trim() || 'CancelData',
    claimDb: process.env.CLAIM_DB?.trim() || 'ClaimDataDB',
    claimCollection: process.env.CLAIM_COLLECTION?.trim() || 'ClaimData_Claim',
    auditDb: process.env.AUDIT_DB?.trim() || 'AuditDB',
    auditCollection: process.env.AUDIT_COLLECTION?.trim() || 'DataReconciliationAudit',
  };
}

export function loadReportConfig(args: string[]): ReportConfig {
  const excluded = (process.env.EXCLUDED_COMPONENT_CODES || '')
    .split(',')
    .map((value) => value.trim().toUpperCase())
    .filter(Boolean);

  return {
    companyName: process.env.REPORT_COMPANY_NAME?.trim() || 'OMNISHIELD',
    asOfDate: parseAsOfDate(args),
    outputDirectory: path.resolve(process.env.REPORT_OUTPUT_DIR?.trim() || 'output'),
    topDealerCount: positiveInteger('TOP_DEALER_COUNT', 20),
    warningLossRatio: ratio('LOSS_RATIO_WARNING', 0.65),
    highLossRatio: ratio('LOSS_RATIO_HIGH', 0.8),
    excludedComponentCodes: new Set(excluded),
  };
}

export function loadMongoUri(): string {
  return required('MONGO_URI');
}

function emailList(name: string): string[] {
  return (process.env[name] || '')
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean);
}

export function loadEmailConfig(): EmailConfig {
  const enabled = process.env.EMAIL_ENABLED?.trim().toLowerCase() === 'true';
  const config: EmailConfig = {
    enabled,
    apiKey: process.env.SENDGRID_API_KEY?.trim() || '',
    from: process.env.EMAIL_FROM?.trim() || '',
    to: emailList('EMAIL_TO'),
    cc: emailList('EMAIL_CC'),
  };

  if (enabled) {
    if (!config.apiKey || config.apiKey.includes('REPLACE_WITH')) {
      throw new Error('SENDGRID_API_KEY is required when EMAIL_ENABLED=true');
    }
    if (!config.from) throw new Error('EMAIL_FROM is required when EMAIL_ENABLED=true');
    if (config.to.length === 0)
      throw new Error('EMAIL_TO requires at least one recipient when EMAIL_ENABLED=true');
  }
  return config;
}
