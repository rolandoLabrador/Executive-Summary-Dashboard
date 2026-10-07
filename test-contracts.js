
const { MongoClient } = require('mongodb');
require('dotenv').config();

const contractIds = [
  'CW10267822', 'CWVSC10527997', 'CWVSC10405584', 'CWVSC10588157',
  'CW10302976', 'CW10257391', 'CWVSC10516110', 'CWVSC10553900', 'CWVSC10625344'
];

async function checkContracts() {
  const uri = process.env.MONGODB_URI;
  const client = new MongoClient(uri);
  try {
    await client.connect();
    const db = client.db(process.env.MONGODB_DB_NAME || 'omnishield');
    const contracts = await db.collection('contracts').find({ 'Contract #': { $in: contractIds } }).toArray();
    
    console.log('--- CONTRACTS ---');
    for (const c of contracts) {
       let wAdmin = 0, cAdmin = 0, wRes = 0, cRes = 0;
       
       const adminKeys = c.WrittenAmount?.ADMIN ? Object.keys(c.WrittenAmount.ADMIN) : [];
       for(const k of adminKeys) wAdmin += parseFloat(c.WrittenAmount.ADMIN[k])||0;
       
       const cAdminKeys = c.CancelledAmount?.ADMIN ? Object.keys(c.CancelledAmount.ADMIN) : [];
       for(const k of cAdminKeys) cAdmin += parseFloat(c.CancelledAmount.ADMIN[k])||0;
       
       console.log(c['Contract #'], '| Status:', c.ContractStatus, '| ActDate:', c.metadata?.ActivationDate, '| CancelDate:', c.metadata?.CancelBillDate);
       console.log('  Written Admin:', wAdmin.toFixed(2), '| Cancel Admin:', cAdmin.toFixed(2));
    }
  } finally {
    await client.close();
  }
}
checkContracts().catch(console.error);

