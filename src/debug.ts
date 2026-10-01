import * as dotenv from 'dotenv';
import { MongoClient } from 'mongodb';
import { loadMongoUri } from './config';

dotenv.config();

async function run() {
  const uri = loadMongoUri();
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');
  const contracts = await db.collection('ContractData').find({ 'metadata.DealerName': 'Power Motors East' }).toArray();
  let total = 0;
  for (const c of contracts) {
    if (c.WrittenAmount && c.WrittenAmount.RESERVE && c.WrittenAmount.RESERVE.ADMIN) {
      total += Number(c.WrittenAmount.RESERVE.ADMIN);
    }
  }
  console.log('Total RESERVE.ADMIN in DB for Power Motors East:', total);
  await client.close();
}
run().catch(console.error);
