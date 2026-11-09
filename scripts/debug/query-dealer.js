const { MongoClient } = require('mongodb');
const { config } = require('dotenv');
config();

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error('Usage: node query-dealer.js <DealerNumber>');
    process.exit(1);
  }
  const dealerNumber = args[0];

  const uri = process.env.MONGO_URI.replace(/"/g, '').trim();
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db(process.env.CONTRACT_DB || 'ContractDataDB');
  const collection = db.collection(process.env.CONTRACT_COLLECTION || 'ContractData');

  console.log(`Querying MongoDB for DealerNumber: ${dealerNumber}...`);
  const docs = await collection.find({ 'metadata.DealerNumber': dealerNumber }).toArray();

  console.log(`\n--- ${dealerNumber} Summary ---`);
  console.log(`Total Documents: ${docs.length}`);

  const statusCounts = {};
  const monthlyCounts = {};

  for (const doc of docs) {
    const meta = doc.metadata || {};

    const status = meta.ContractStatus || 'Unknown';
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    const dt = meta.ActivationDate || doc.ActivationDate || 'Unknown';
    let yyyymm = String(dt).substring(0, 7);
    if (typeof dt === 'object') {
      yyyymm = dt.toISOString().substring(0, 7);
    }
    monthlyCounts[yyyymm] = (monthlyCounts[yyyymm] || 0) + 1;
  }

  console.log('\nContract Status Distribution:');
  console.table(statusCounts);

  console.log('\nActivation Date Distribution (YYYY-MM):');
  // Sort keys chronologically
  const sortedMonths = Object.keys(monthlyCounts)
    .sort()
    .reduce((obj, key) => {
      obj[key] = monthlyCounts[key];
      return obj;
    }, {});
  console.table(sortedMonths);

  await client.close();
}

main().catch(console.error);
