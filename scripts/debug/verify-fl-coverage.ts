import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
dotenv.config();

async function check() {
  const uri = process.env.MONGO_URI!;
  const client = new MongoClient(uri);
  await client.connect();

  const coll = client
    .db(process.env.CONTRACT_DB || 'ContractDataDB')
    .collection(process.env.CONTRACT_COLLECTION || 'ContractData');

  // Check state variations in metadata['DealerState/Province']
  const stateValues = await coll.distinct('metadata.DealerState/Province');
  console.log('Distinct DealerState values in database:', stateValues);

  // Check if any zip starts with 32, 33, 34 where state is NOT FL
  const nonFlWithFlZip = await coll
    .aggregate([
      {
        $match: {
          'metadata.DealerZipCode': { $regex: /^(32|33|34)/ },
          'metadata.DealerState/Province': { $not: { $regex: /^fl/i } },
        },
      },
      {
        $group: {
          _id: {
            dealer: '$metadata.DealerName',
            state: '$metadata.DealerState/Province',
            zip: '$metadata.DealerZipCode',
          },
          count: { $sum: 1 },
        },
      },
    ])
    .toArray();
  console.log('Dealers with FL Zip (32xxx, 33xxx, 34xxx) but state not FL:', nonFlWithFlZip);

  // Check total dealers in DB vs FL dealers
  const allDealers = await coll.distinct('metadata.DealerName');
  const flDealers = await coll.distinct('metadata.DealerName', {
    $or: [
      { 'metadata.DealerState/Province': /^FL/i },
      { 'metadata.DealerState/Province': 'Florida' },
      { 'metadata.DealerZipCode': { $regex: /^(32|33|34)/ } },
    ],
  });
  console.log(
    `Total dealers in DB: ${allDealers.length}, Total Florida dealers: ${flDealers.length}`,
  );

  // Check total coverage codes in DB vs FL coverage codes
  const allCodes = await coll.distinct('metadata.CoverageCode');
  const flCodes = await coll.distinct('metadata.CoverageCode', {
    $or: [
      { 'metadata.DealerState/Province': /^FL/i },
      { 'metadata.DealerState/Province': 'Florida' },
      { 'metadata.DealerZipCode': { $regex: /^(32|33|34)/ } },
    ],
  });
  console.log(
    `Total coverage codes in DB: ${allCodes.length}, Coverage codes sold in FL: ${flCodes.length}`,
  );

  await client.close();
}

check().catch(console.error);
