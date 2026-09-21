const { MongoClient } = require('mongodb');
require('dotenv').config();

function newExcludedComponent(upper) {
  if (
    upper.includes('DEALER') ||
    upper.includes('DLR') ||
    upper.includes('COMMISSION') ||
    upper.includes('COMM') ||
    upper.includes('F&I') ||
    upper.includes('PACK')
  ) {
    return true;
  }
  
  if (
    ['CLIPFEE', 'PREMIUMTAX', 'CEDINGFEE', 'ADMIN'].includes(upper) ||
    upper.includes('PREMIUM TAX') ||
    upper.includes('CEEDING') ||
    upper.includes('CEDING') ||
    upper.includes('CLIP FEE')
  ) {
    return true;
  }
  return false;
}

async function check() {
  const uri = process.env.DB_URI || process.env.MONGO_URI;
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');
  const cancelDb = client.db('CancelDataDB');
  
  const cutoff = new Date('2026-09-01T00:00:00.000Z'); // Before Sep 1
  
  const cursor = db.collection('ContractData').find({
    "metadata.DealerNumber": "MS370",
    "metadata.ActivationDate": { $lt: cutoff }
  });
  
  let newWrittenSum = 0;
  
  for await (const doc of cursor) {
    if (doc.WrittenAmount && doc.WrittenAmount.RESERVE) {
      for (const [key, val] of Object.entries(doc.WrittenAmount.RESERVE)) {
        const upper = key.trim().toUpperCase();
        if (!newExcludedComponent(upper)) {
          newWrittenSum += Number(val);
        }
      }
    }
  }
  
  const cursor2 = cancelDb.collection('CancelData').find({
    "metadata.DealerNumber": "MS370",
    "metadata.CancelDate": { $lt: cutoff }
  });
  
  let newCancelSum = 0;
  
  for await (const doc of cursor2) {
    if (doc.CancelledAmount && doc.CancelledAmount.RESERVE) {
      for (const [key, val] of Object.entries(doc.CancelledAmount.RESERVE)) {
        const upper = key.trim().toUpperCase();
        if (!newExcludedComponent(upper)) {
          newCancelSum += Number(val);
        }
      }
    }
  }
  
  console.log(`New Code Net Reserve (MS370) cutoff 2026-08-31: ${newWrittenSum - newCancelSum}`);
  
  await client.close();
}

check().catch(console.error);

