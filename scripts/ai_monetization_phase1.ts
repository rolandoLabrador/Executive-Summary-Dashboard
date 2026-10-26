import * as dotenv from 'dotenv';
import { mkdir, writeFile } from 'fs/promises';
import * as path from 'path';
import { loadMongoSourceConfig, loadMongoUri, loadReportConfig } from '../src/config';
import { DataRepository } from '../src/services/data.repository';
import { MongoService } from '../src/services/mongo.service';
import { ReportTransformer } from '../src/utils/report.transformer';

dotenv.config();

function stdDev(array: number[], mean: number) {
  const n = array.length;
  if (n <= 1) return 0;
  const variance = array.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (n - 1);
  return Math.sqrt(variance);
}

function median(array: number[]): number {
  if (array.length === 0) return 0;
  array.sort((a, b) => a - b);
  const mid = Math.floor(array.length / 2);
  return array.length % 2 !== 0 ? (array[mid] ?? 0) : ((array[mid - 1] ?? 0) + (array[mid] ?? 0)) / 2;
}

async function main() {
  const reportConfig = loadReportConfig(process.argv.slice(2));
  const mongo = new MongoService(loadMongoUri());
  await mongo.connect();
  const repository = new DataRepository(mongo, loadMongoSourceConfig());

  const transformer = new ReportTransformer(reportConfig);
  
  async function* emptyStream() {}
  const claimsStream = repository.getClaims();

  console.log('Extracting claims...');
  const model = await transformer.transform(
    emptyStream(),
    emptyStream(),
    claimsStream,
    []
  );

  console.log(`Extracted ${model.claims.length} claims.`);
  
  const groups: Record<string, number[]> = {};
  for (const claim of model.claims) {
    if (!claim.lossCode || !claim.vehicleMake) continue;
    const key = `${claim.lossCode}|${claim.vehicleMake}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(claim.paidAmount);
  }

  const results = [];
  results.push("LossCode,VehicleMake,Count,Mean,Median,StdDev,Threshold_1_5_SD");

  let processedCount = 0;
  for (const [key, amounts] of Object.entries(groups)) {
    if (amounts.length < 5) continue; // skip low sample sizes
    const [lossCode, vehicleMake] = key.split('|');
    const sum = amounts.reduce((a, b) => a + b, 0);
    const mean = sum / amounts.length;
    const med = median([...amounts]);
    const sd = stdDev(amounts, mean);
    const threshold = med + (1.5 * sd);
    
    results.push(`"${lossCode}","${vehicleMake}",${amounts.length},${mean.toFixed(2)},${med.toFixed(2)},${sd.toFixed(2)},${threshold.toFixed(2)}`);
    processedCount++;
  }

  console.log(`Generated ${processedCount} valid benchmarking rules.`);

  const outputPath = path.join(__dirname, '../output/ai_claims_benchmarks.csv');
  await mkdir(path.dirname(outputPath), { recursive: true });
  await writeFile(outputPath, results.join('\n'));
  console.log(`Wrote benchmarks to ${outputPath}`);
  await mongo.close();
}

main().catch(console.error);
