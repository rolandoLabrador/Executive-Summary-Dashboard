import { MongoClient } from 'mongodb';
import { env } from '../src/config/env';

async function run() {
  const client = new MongoClient(env.MONGODB_URI);
  await client.connect();
  const db = client.db('reporting');
  const contract = await db.collection('cancellations').findOne({ 'metadata.Contract#': 'DC11218481' });
  console.log('Contract:', contract?.metadata['Contract#']);
  console.log('WrittenAmount:', JSON.stringify(contract?.WrittenAmount, null, 2));
  await client.close();
}
run().catch(console.error);
