import { loadMongoSourceConfig, loadMongoUri, loadReportConfig } from '../../src/config';
import { DataRepository } from '../../src/services/data.repository';
import { MongoService } from '../../src/services/mongo.service';
import * as dotenv from 'dotenv';
import { ReportTransformer } from '../../src/utils/report.transformer';

dotenv.config();

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error('Usage: npx tsx test-pipeline.ts <DealerNumber | AgentCodeOrName>');
    process.exit(1);
  }
  const target = args[0];
  const targetLower = target.toLowerCase();

  console.log(`Testing pipeline for Dealer or Agent: "${target}"...`);

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

  const isExwAlias = targetLower === 'exw' || targetLower === 'executive warranties';

  // Dealer lookup
  const rollingDealer = model.dealers.find(
    (d) => d.name.toLowerCase() === targetLower || d.displayName?.toLowerCase() === targetLower,
  );
  const itdDealer = model.itdDealers.find(
    (d) => d.name.toLowerCase() === targetLower || d.displayName?.toLowerCase() === targetLower,
  );

  // Agent lookup
  const rollingAgent = model.agents.find(
    (a) =>
      a.name.toLowerCase() === targetLower ||
      a.displayName?.toLowerCase() === targetLower ||
      (isExwAlias &&
        (a.name.toLowerCase() === 'exw' ||
          a.displayName?.toLowerCase() === 'executive warranties')),
  );
  const itdAgent = model.itdAgents.find(
    (a) =>
      a.name.toLowerCase() === targetLower ||
      a.displayName?.toLowerCase() === targetLower ||
      (isExwAlias &&
        (a.name.toLowerCase() === 'exw' ||
          a.displayName?.toLowerCase() === 'executive warranties')),
  );

  let found = false;

  if (rollingDealer || itdDealer) {
    found = true;
    console.log(
      `\n================ DEALER: ${rollingDealer?.displayName || rollingDealer?.name || target} ================`,
    );
    console.log(`--- ROLLING 12 Stats ---`);
    console.dir(rollingDealer, { depth: null });
    console.log(`--- ITD Stats ---`);
    console.dir(itdDealer, { depth: null });
  }

  if (rollingAgent || itdAgent) {
    found = true;
    console.log(
      `\n================ AGENT: ${rollingAgent?.displayName || rollingAgent?.name || target} ================`,
    );
    console.log(`--- ROLLING 12 Stats ---`);
    console.dir(rollingAgent, { depth: null });
    console.log(`--- ITD Stats ---`);
    console.dir(itdAgent, { depth: null });
  }

  if (!found) {
    console.log(`\nNo matching dealer or agent found for "${target}".`);
  }

  await mongo.close();
}

main().catch(console.error);
