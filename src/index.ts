import * as dotenv from 'dotenv';
import { mkdir, writeFile } from 'fs/promises';
import * as path from 'path';
import { loadEmailConfig, loadMongoSourceConfig, loadMongoUri, loadReportConfig } from './config';
import { type ReportModel } from './models/report.types';
import { DataRepository } from './services/data.repository';
import { EmailService } from './services/email.service';
import { ExcelService } from './services/excel.service';
import { MongoService } from './services/mongo.service';
import { ReportTransformer } from './utils/report.transformer';

dotenv.config();

function fileDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function logMemory(checkpoint: string): void {
  const mem = process.memoryUsage();
  const heapUsedMB = Math.round(mem.heapUsed / 1024 / 1024);
  const heapTotalMB = Math.round(mem.heapTotal / 1024 / 1024);
  const rssMB = Math.round(mem.rss / 1024 / 1024);
  console.log(`⏱️ [Memory Checkpoint] ${checkpoint} -> Heap: ${heapUsedMB} MB / ${heapTotalMB} MB (RSS: ${rssMB} MB)`);
}

async function extractAndTransformModel(
  repository: DataRepository,
  reportConfig: ReturnType<typeof loadReportConfig>,
): Promise<ReportModel> {
  const [contracts, cancellations, claims, pipelineAudits] = await Promise.all([
    repository.getContracts(),
    repository.getCancellations(),
    repository.getClaims(),
    repository.getLatestReconciliationAudits(),
  ]);
  console.log(
    `Extracted ${contracts.length} contract, ${cancellations.length} cancellation, and ${claims.length} claim documents.`,
  );
  if (pipelineAudits.length > 0) {
    console.log(
      `Pipeline Audits: ${pipelineAudits.map((a) => `${a.jobType}: ${a.reconciliation.status}`).join(' | ')}`,
    );
  }
  logMemory('2. After MongoDB Extraction (Raw Documents in RAM)');

  const model = new ReportTransformer(reportConfig).transform(
    contracts,
    cancellations,
    claims,
    pipelineAudits,
  );

  return model;
}

async function main(): Promise<void> {
  logMemory('1. Start');
  const reportConfig = loadReportConfig(process.argv.slice(2));
  const mongo = new MongoService(loadMongoUri());
  console.log(`Starting report for ${fileDate(reportConfig.asOfDate)}...`);

  try {
    await mongo.connect();
    const repository = new DataRepository(mongo, loadMongoSourceConfig());
    
    // Raw documents are isolated inside extractAndTransformModel and freed when it finishes
    const model = await extractAndTransformModel(repository, reportConfig);
    if (global.gc) {
      global.gc();
    }
    logMemory('3. After ReportTransformer (Raw Documents Released)');

    const workbook = await new ExcelService(reportConfig).generateFullReport(model);
    logMemory('4. After ExcelService.generateFullReport (Workbook in RAM)');

    await mkdir(reportConfig.outputDirectory, { recursive: true });
    const fileName = `${reportConfig.companyName.replace(/[^a-z0-9]+/gi, '_')}_Executive_Report_${fileDate(model.asOfDate)}.xlsx`;
    const outputPath = path.join(reportConfig.outputDirectory, fileName);
    await writeFile(outputPath, workbook);
    logMemory('5. After Writing Excel to Disk');

    console.log(`Report created: ${outputPath}`);
    console.log(`Data-quality issues: ${model.dataQualityIssues.length}`);
    await new EmailService(loadEmailConfig()).sendReport({
      companyName: reportConfig.companyName,
      asOfDate: model.asOfDate,
      fileName,
      workbook,
      dataQualityIssueCount: model.dataQualityIssues.length,
    });
  } finally {
    await mongo.close();
  }
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`Report generation failed: ${message}`);
  process.exitCode = 1;
});
