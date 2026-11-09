import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
dotenv.config();

async function run() {
  const uri = process.env.MONGO_URI!;
  const client = new MongoClient(uri);
  await client.connect();

  const coll = client
    .db(process.env.CONTRACT_DB || 'ContractDataDB')
    .collection(process.env.CONTRACT_COLLECTION || 'ContractData');

  const flMatch = {
    $or: [
      { 'metadata.DealerState/Province': /^FL/i },
      { 'metadata.DealerState/Province': 'Florida' },
    ],
  };

  // Check if any F&IMarkup > 0
  const countWithFiMarkup = await coll.countDocuments({
    ...flMatch,
    'metadata.F&IMarkup': { $nin: ['0', '0.0', '0.00', '0.0000', '', null] },
  });
  console.log(`FL contracts with non-zero F&IMarkup: ${countWithFiMarkup}`);

  // Let's check some examples of F&IMarkup if any
  if (countWithFiMarkup > 0) {
    const samples = await coll
      .find({
        ...flMatch,
        'metadata.F&IMarkup': { $nin: ['0', '0.0', '0.00', '0.0000', '', null] },
      })
      .limit(5)
      .project({
        'metadata.DealerName': 1,
        'metadata.DealerNumber': 1,
        'metadata.CoverageCode': 1,
        'metadata.CoverageName': 1,
        'metadata.F&IMarkup': 1,
        'metadata.RetailCost': 1,
        'metadata.NetCost': 1,
      })
      .toArray();
    console.log('Sample F&IMarkup docs:', JSON.stringify(samples, null, 2));
  }

  // Also check if any RetailCost != NetCost
  const diffSample = await coll
    .aggregate([
      { $match: flMatch },
      {
        $project: {
          dealer: '$metadata.DealerName',
          code: '$metadata.CoverageCode',
          fi: '$metadata.F&IMarkup',
          retail: { $toDouble: '$metadata.RetailCost' },
          net: { $toDouble: '$metadata.NetCost' },
        },
      },
      {
        $project: {
          dealer: 1,
          code: 1,
          fi: 1,
          retail: 1,
          net: 1,
          diff: { $subtract: ['$retail', '$net'] },
        },
      },
      { $match: { diff: { $gt: 0 } } },
      { $limit: 5 },
    ])
    .toArray();
  console.log('Sample Retail > Net docs:', diffSample);

  await client.close();
}

run().catch(console.error);
