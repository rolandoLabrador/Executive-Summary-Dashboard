import * as fs from 'fs';
import * as path from 'path';
import { MongoClient } from 'mongodb';
import { config } from 'dotenv';
config();

async function run() {
  const client = new MongoClient(process.env.MONGODB_URI!);
  await client.connect();
  const db = client.db('reporting');
  
  const voided = await db.collection('contracts').find({
    'metadata.DealerName': 'Power Motors East',
    'metadata.ContractStatus': 'V'
  }).toArray();
  
  const badDates = await db.collection('contracts').find({
    'metadata.DealerName': 'Power Motors East',
    $or: [
      { 'metadata.ActivationDate': null },
      { 'metadata.ActivationDate': '' }
    ]
  }).toArray();

  const badCancels = await db.collection('cancellations').find({
    'metadata.DealerName': 'Power Motors East',
    $or: [
      { 'metadata.CancellationEffectiveDate': null },
      { 'metadata.CancellationEffectiveDate': '' }
    ]
  }).toArray();
  
  console.log('Voided Contracts for Power Motors East:');
  for (const doc of voided) {
    console.log(`- Contract ${doc.metadata['Contract#']}: ${JSON.stringify(doc.WrittenAmount?.ADMIN)}`);
  }
  
  console.log('\nMissing Date Contracts for Power Motors East:');
  for (const doc of badDates) {
     console.log(`- Contract ${doc.metadata['Contract#']} (Date: ${doc.metadata.ActivationDate}): ${JSON.stringify(doc.WrittenAmount?.ADMIN)}`);
  }

  console.log('\nMissing Date Cancellations for Power Motors East:');
  for (const doc of badCancels) {
     console.log(`- Contract ${doc.metadata['Contract#']} (Date: ${doc.metadata.CancellationEffectiveDate}): ${JSON.stringify(doc.WrittenAmount?.ADMIN)} / Cancelled: ${JSON.stringify(doc.CancelledAmount?.ADMIN)}`);
  }

  await client.close();
}
run().catch(console.error);
