import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
dotenv.config();

async function run() {
  const uri = process.env.MONGO_URI;
  if (!uri) throw new Error('MONGO_URI is missing');

  const client = new MongoClient(uri);
  await client.connect();

  const dbName = process.env.CONTRACT_DB || 'ContractDataDB';
  const collName = process.env.CONTRACT_COLLECTION || 'ContractData';
  const coll = client.db(dbName).collection(collName);

  console.log(`Connected to ${dbName}.${collName}`);

  // Check how many FL documents exist and examine markup fields
  const count = await coll.countDocuments({
    $or: [
      { 'metadata.DealerState/Province': /^FL/i },
      { 'metadata.DealerState/Province': 'Florida' },
    ],
  });
  console.log(`Total Florida contract documents: ${count}`);

  const sampleDocs = await coll
    .find({
      $or: [
        { 'metadata.DealerState/Province': /^FL/i },
        { 'metadata.DealerState/Province': 'Florida' },
      ],
    })
    .limit(100)
    .toArray();

  let maxFiMarkup = 0;
  let maxRetailMinusNet = 0;
  const sampleData: Record<string, unknown>[] = [];

  for (const d of sampleDocs) {
    const m = d.metadata || {};
    const fi = parseFloat(m['F&IMarkup'] || '0') || 0;
    const retail = parseFloat(m.RetailCost || '0') || 0;
    const net = parseFloat(m.NetCost || '0') || 0;
    const diff = retail - net;
    if (fi > maxFiMarkup) maxFiMarkup = fi;
    if (diff > maxRetailMinusNet) maxRetailMinusNet = diff;
    if (sampleData.length < 5) {
      sampleData.push({
        dealer: m.DealerName,
        dealerNumber: m.DealerNumber,
        state: m['DealerState/Province'],
        productType: m.ProductType,
        coverageCode: m.CoverageCode,
        coverageName: m.CoverageName,
        fiMarkup: m['F&IMarkup'],
        retailCost: m.RetailCost,
        netCost: m.NetCost,
        totalReserves: m.TotalReserves,
        submittedAmount: m.SubmittedAmount,
      });
    }
  }

  console.log('Sample Data:', JSON.stringify(sampleData, null, 2));
  console.log(`Max F&IMarkup in sample: ${maxFiMarkup}, Max Retail - Net: ${maxRetailMinusNet}`);

  await client.close();
}

run().catch(console.error);
