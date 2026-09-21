const { MongoClient } = require('mongodb');
require('dotenv').config();

async function run() {
  const uri = process.env.DB_URI || process.env.MONGO_URI;
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ClaimDataDB');
  
  const cursor = db.collection('ClaimData_Claim').find({});
  const counts = {};
  
  for await (const doc of cursor) {
    const amt = parseFloat(doc['Total Paid Amount']);
    if (amt < 0) {
      const desc = doc['Claim Description'] || 'N/A';
      counts[desc] = (counts[desc] || 0) + 1;
    }
  }
  
  console.log("Negative Claims by Description:");
  console.log(counts);
  await client.close();
}

run().catch(console.error);
