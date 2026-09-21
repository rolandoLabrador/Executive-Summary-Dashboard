const { MongoClient } = require('mongodb');
require('dotenv').config();

async function check() {
  const uri = process.env.DB_URI || process.env.MONGO_URI;
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');
  const cancelDb = client.db('CancelDataDB');
  
  const cursor = db.collection('ContractData').find({
    "metadata.DealerNumber": "MS370"
  });
  
  let rawTotalReserves = 0;
  let rawBaseReserve = 0;
  let rawOverReserve = 0;
  let rawBaseReserveFTP = 0;
  let rawSiReserves = 0;
  let rawClipFee = 0;
  let rawPremiumTax = 0;
  let rawCedingFee = 0;
  let rawAdmin = 0;
  
  for await (const doc of cursor) {
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
      if (res.PREMIUMTAX) rawPremiumTax += Number(res.PREMIUMTAX);
      if (res.CEDINGFEE) rawCedingFee += Number(res.CEDINGFEE);
      if (res.ADMIN) rawAdmin += Number(res.ADMIN);
    }
  }
  
  console.log(`Raw TotalReserves: ${rawTotalReserves}`);
  console.log(`Raw BASERESERVE: ${rawBaseReserve}`);
  console.log(`Raw OVERRESERVE: ${rawOverReserve}`);
  console.log(`Raw BASERESERVEFTP: ${rawBaseReserveFTP}`);
  console.log(`Raw SIRESERVES: ${rawSiReserves}`);
  console.log(`Sum 4 main: ${rawBaseReserve + rawOverReserve + rawBaseReserveFTP + rawSiReserves}`);
  console.log(`Raw CLIPFEE: ${rawClipFee}`);
  console.log(`Raw PREMIUMTAX: ${rawPremiumTax}`);
  console.log(`Raw CEDINGFEE: ${rawCedingFee}`);
  console.log(`Raw ADMIN: ${rawAdmin}`);
  
  await client.close();
}

check().catch(console.error);

