import { loadMongoSourceConfig, loadMongoUri, loadReportConfig } from '../../src/config';
import { DataRepository } from '../../src/services/data.repository';
import { MongoService } from '../../src/services/mongo.service';
import * as dotenv from 'dotenv';
import { ReportTransformer } from '../../src/utils/report.transformer';

dotenv.config();

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error('Usage: npx tsx test-pipeline.ts <DealerNumber>');
    process.exit(1);
  }
  const dealerNumber = args[0];

  console.log(`Testing pipeline for DealerNumber: ${dealerNumber}...`);

  const mongo = new MongoService(loadMongoUri());
  await mongo.connect();

  const repository = new DataRepository(mongo, loadMongoSourceConfig());
  const reportConfig = loadReportConfig([]);

  const contractsStream = repository.getContracts();
  const cancellationsStream = repository.getCancellations();
  const claimsStream = repository.getClaims();
  const pipelineAudits = await repository.getLatestReconciliationAudits();

  const model = await new ReportTransformer(reportConfig).transform(
    contractsStream,
    cancellationsStream,
    claimsStream,
    pipelineAudits,
  );

  const rolling = model.dealers.find((d) => d.name === dealerNumber);
  console.log(`\n--- ${dealerNumber} ROLLING 12 Stats ---`);
  console.dir(rolling, { depth: null });

  const itd = model.itdDealers.find((d) => d.name === dealerNumber);
  console.log(`\n--- ${dealerNumber} ITD Stats ---`);
  console.dir(itd, { depth: null });

  await mongo.close();
}

main().catch(console.error);
