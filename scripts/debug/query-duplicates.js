const { MongoClient } = require('mongodb');
const { config } = require('dotenv');
config();

function businessId(value) {
  return (value === null || value === undefined ? '' : String(value).trim())
    .replace(/\s+/g, '')
    .toUpperCase();
}

async function main() {
  const uri = process.env.MONGO_URI.replace(/"/g, '').trim();
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db(process.env.CONTRACT_DB || 'ContractDataDB');
  const collection = db.collection(process.env.CONTRACT_COLLECTION || 'ContractData');

  const ms399 = await collection.find({ 'metadata.DealerNumber': 'MS399' }).limit(1).toArray();
  const contractNum = ms399[0].metadata['Contract#'];
  const bId = businessId(contractNum);

  console.log('Checking duplicates for:', bId);

  const docs = await collection.find({ 'metadata.Contract#': contractNum }).toArray();

  console.log('Found globally:', docs.length);
  const dealers = new Set(docs.map((d) => d.metadata.DealerNumber));
  console.log('Dealers containing this contract:', Array.from(dealers));

  await client.close();
}

main().catch(console.error);
