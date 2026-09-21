const { MongoClient } = require('mongodb');
require('dotenv').config();

async function run() {
  const client = new MongoClient(process.env.MONGO_URI);
  await client.connect();
  const db = client.db('ContractDataDB');
  const col = db.collection('ContractDatd');

  const total = await col.countDocuments({});
  console.log('Total documents in collection:', total);

  const doc = await col.findOne({});
  if (doc) {
    console.log('Doc _id:', doc._id);
    if (doc.metadata) {
      console.log(
        'ActivationDate:',
        doc.metadata.ActivationDate,
        'Type:',
        typeof doc.metadata.ActivationDate,
      );
      console.log('Is Date:', doc.metadata.ActivationDate instanceof Date);
    } else if (doc.ActivationDate) {
      console.log('ROOT ActivationDate:', doc.ActivationDate, 'Type:', typeof doc.ActivationDate);
      console.log('Is Date:', doc.ActivationDate instanceof Date);
    }
  }

  await client.close();
}

run().catch(console.error);
