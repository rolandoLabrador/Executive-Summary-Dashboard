import { MongoClient } from 'mongodb';
import { loadMongoUri } from '../../src/config';
import * as dotenv from 'dotenv';
dotenv.config();

function getAmount(container: unknown, category: string, sub: string): number {
  if (
    !container ||
    typeof container !== 'object' ||
    !(category in container)
  ) return 0;
  
  const catObj = (container as Record<string, unknown>)[category];
  if (!catObj || typeof catObj !== 'object' || !(sub in catObj)) return 0;
  
  let val = (catObj as Record<string, unknown>)[sub];
  if (typeof val === 'object' && val !== null) {
    // Check if it's a BSON type
    if ('value' in val && val.value !== undefined) return Number(val.value); // Some BSON Int32 parsers
    val = (val as { toString(): string }).toString();
  }
  return Number(val) || 0;
}

async function main() {
  const uri = loadMongoUri();
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');
  const cancelDb = client.db('CancelDataDB');
  
  const contracts = await db.collection('ContractData').find({
    $or: [{ 'metadata.DealerNumber': 'MS370' }, { 'metadata.DealerName': 'Power Motors East' }]
  }).toArray();
  
  const cancellations = await cancelDb.collection('CancelData').find({
    $or: [{ 'metadata.DealerNumber': 'MS370' }, { 'metadata.DealerName': 'Power Motors East' }]
  }).toArray();

  let manualAdmin = 0;
  let missingDateAdmin = 0;
  
  for (const doc of contracts) {
    const wa = doc.metadata.WrittenAmount;
    const adminWritten = 
      getAmount(wa, 'ADMIN', 'ADMIN') +
      getAmount(wa, 'ADMIN', 'BASEADMINMS') +
      getAmount(wa, 'ADMIN', 'BASEADMINOS') +
      getAmount(wa, 'ADMIN', 'MARKETINGFEE') +
      getAmount(wa, 'RESERVE', 'ADMIN');
      
    manualAdmin += adminWritten;
    
    const actDate = doc.metadata.ActivationDate;
    const hasDate = actDate && actDate.toString().length > 0;
    const isFuture = hasDate && new Date(actDate) > new Date('2026-08-31T23:59:59.999Z');
    
    if (doc.metadata.ContractStatus === 'V') {
      console.log(`[VOID] ${doc.metadata['Contract#']} : $${adminWritten}`);
    } else if (!hasDate) {
      console.log(`[MISSING_DATE] ${doc.metadata['Contract#']} : $${adminWritten}`);
      missingDateAdmin += adminWritten;
    } else if (isFuture) {
      console.log(`[FUTURE] ${doc.metadata['Contract#']} : $${adminWritten}`);
    }
  }

  for (const doc of cancellations) {
    const ca = doc.metadata.CancelledAmount;
    const adminCancelled = 
      getAmount(ca, 'ADMIN', 'ADMIN') +
      getAmount(ca, 'ADMIN', 'BASEADMINMS') +
      getAmount(ca, 'ADMIN', 'BASEADMINOS') +
      getAmount(ca, 'ADMIN', 'MARKETINGFEE') +
      getAmount(ca, 'RESERVE', 'ADMIN');
      
    manualAdmin -= adminCancelled;
  }

  console.log(`\nManual Net Admin: ${manualAdmin}`);
  console.log(`Total Admin dropped due to missing date: ${missingDateAdmin}`);
  
  await client.close();
}

main().catch(console.error);

