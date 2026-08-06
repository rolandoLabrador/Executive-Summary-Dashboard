import * as dotenv from 'dotenv';
import { mkdir, writeFile } from 'fs/promises';
import * as path from 'path';
import { loadEmailConfig, loadMongoSourceConfig, loadMongoUri, loadReportConfig } from './config';
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

async function main(): Promise<void> {
  const reportConfig = loadReportConfig(process.argv.slice(2));
  const mongo = new MongoService(loadMongoUri());
  console.log(`Starting report for ${fileDate(reportConfig.asOfDate)}...`);

  try {
    await mongo.connect();
    const repository = new DataRepository(mongo, loadMongoSourceConfig());
    const [contracts, cancellations, claims] = await Promise.all([
      repository.getContracts(),
      repository.getCancellations(),
      repository.getClaims(),
    ]);
    console.log(
      `Extracted ${contracts.length} contract, ${cancellations.length} cancellation, and ${claims.length} claim documents.`,
    );

    const model = new ReportTransformer(reportConfig).transform(contracts, cancellations, claims);
    const workbook = await new ExcelService(reportConfig).generateFullReport(model);
    await mkdir(reportConfig.outputDirectory, { recursive: true });
    const fileName = `${reportConfig.companyName.replace(/[^a-z0-9]+/gi, '_')}_Executive_Report_${fileDate(model.asOfDate)}.xlsx`;
    const outputPath = path.join(reportConfig.outputDirectory, fileName);
    await writeFile(outputPath, workbook);

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
