const { MongoClient } = require('mongodb');
require('dotenv').config();

async function run() {
  const client = new MongoClient(process.env.MONGO_URI);
  await client.connect();
  const db = client.db('ContractDataDB');
  const col = db.collection('ContractDatd');

  const twoYearsAgo = new Date();
  twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);

  console.log('Querying dates >=', twoYearsAgo);

  const count = await col.countDocuments({
    'metadata.ActivationDate': { $gte: twoYearsAgo },
  });

  console.log('Count from last 2 years:', count);

  const doc = await col.findOne({
    'metadata.ActivationDate': { $gte: twoYearsAgo },
  });

  if (doc) {
    console.log(
      'Type of ActivationDate:',
      typeof doc.metadata.ActivationDate,
      doc.metadata.ActivationDate,
    );
  } else {
    console.log('No documents found.');
  }

  await client.close();
}

run().catch(console.error);
