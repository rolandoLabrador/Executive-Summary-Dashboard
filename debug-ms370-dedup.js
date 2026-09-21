const { MongoClient } = require('mongodb');
require('dotenv').config();

async function check() {
  const uri = process.env.DB_URI || process.env.MONGO_URI;
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');
  
  const cursor = db.collection('ContractData').find({
    "metadata.DealerNumber": "MS370"
  });
  
  const docs = await cursor.toArray();
  
  docs.sort((a, b) => {
    const da = a.metadata && a.metadata.ActivationDate ? new Date(a.metadata.ActivationDate).getTime() : 0;
    const db = b.metadata && b.metadata.ActivationDate ? new Date(b.metadata.ActivationDate).getTime() : 0;
    return da - db;
  });
  
  // Deduplicate by Contract Number (keep latest ActivationDate)
  const deduped = new Map();
  for (const doc of docs) {
    if (doc.metadata && doc.metadata['Contract#']) {
      deduped.set(doc.metadata['Contract#'], doc);
    }
  }
  
  let rawTotalReserves = 0;
  let rawBaseReserve = 0;
  let rawOverReserve = 0;
  let rawBaseReserveFTP = 0;
  let rawSiReserves = 0;
  let rawClipFee = 0;
  
  for (const doc of deduped.values()) {
    if (doc.metadata && doc.metadata.TotalReserves) {
      rawTotalReserves += Number(doc.metadata.TotalReserves);
    }
    if (doc.WrittenAmount && doc.WrittenAmount.RESERVE) {
      const res = doc.WrittenAmount.RESERVE;
      if (res.BASERESERVE) rawBaseReserve += Number(res.BASERESERVE);
      if (res.OVERRESERVE) rawOverReserve += Number(res.OVERRESERVE);
      if (res.BASERESERVEFTP) rawBaseReserveFTP += Number(res.BASERESERVEFTP);
      if (res.SIRESERVES) rawSiReserves += Number(res.SIRESERVES);
      if (res.CLIPFEE) rawClipFee += Number(res.CLIPFEE);
    }
  }
  
  console.log(`Deduped TotalReserves: ${rawTotalReserves}`);
  console.log(`Deduped BASERESERVE: ${rawBaseReserve}`);
  console.log(`Deduped OVERRESERVE: ${rawOverReserve}`);
  console.log(`Deduped BASERESERVEFTP: ${rawBaseReserveFTP}`);
  console.log(`Deduped SIRESERVES: ${rawSiReserves}`);
  console.log(`Sum 4 main: ${rawBaseReserve + rawOverReserve + rawBaseReserveFTP + rawSiReserves}`);
  console.log(`Deduped CLIPFEE: ${rawClipFee}`);
  
  await client.close();
}

check().catch(console.error);

