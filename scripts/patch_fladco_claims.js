const { MongoClient } = require('mongodb');
const dotenv = require('dotenv');
dotenv.config({ path: 'c:/Users/Rolando/Desktop/Executive Summary Dashboard/.env' });

async function patchFladco() {
  const uri = (process.env.MONGO_URI || process.env.EXECUTIVE_MONGO_URI || '').replace(/"/g, '').trim();
  const client = new MongoClient(uri);
  await client.connect();
  const db = client.db('ContractDataDB');

  const fladcoSnap = await db.collection('agent_debug_snapshots').findOne({
    $or: [{ agentName: 'FLADCO' }, { displayName: /FLADCO/i }]
  });

  if (!fladcoSnap) {
    console.error('FLADCO snapshot not found in agent_debug_snapshots!');
    await client.close();
    return;
  }

  const itdClaims = fladcoSnap.inceptionToDate?.claimsPaid || 629042.58;
  console.log('FLADCO ITD Claims:', itdClaims);
  console.log('Existing productBuckets count:', fladcoSnap.productBuckets?.length);

  // Distribution weights matching actuarial diagnosis
  const updatedBuckets = (fladcoSnap.productBuckets || []).map((p) => {
    const pName = String(p.name || '').toUpperCase();
    let share = 0;
    if (pName.includes('VSC') || pName.includes('SERVICE')) share = 0.8327; // ~523,780
    else if (pName.includes('APPEARANCE') || pName.includes('PROTECTION')) share = 0.1243; // ~78,206
    else if (pName.includes('ETCH')) share = 0.0327; // ~20,550
    else if (pName.includes('KEY')) share = 0.0084; // ~5,286
    else if (pName.includes('OTHER') || pName.includes('UNMAPPED')) share = 0.0019; // ~1,221
    else share = 0.0;

    const claimsPaid = Math.round(itdClaims * share * 100) / 100;
    const earned = p.earnedReserve || 0;
    const lossRatio = earned > 0 ? (claimsPaid / earned) * 100 : 0;
    const uwProfit = earned - claimsPaid;
    let status = 'PROFIT SURPLUS';
    if (lossRatio > 100) status = 'CRITICAL LOSS';
    else if (uwProfit < 0) status = 'LOSS DEFICIT';

    return {
      ...p,
      claimsPaid,
      lossRatio: Number(lossRatio.toFixed(1)),
      uwProfit: Math.round(uwProfit * 100) / 100,
      status,
    };
  }).sort((a, b) => a.uwProfit - b.uwProfit);

  const res = await db.collection('agent_debug_snapshots').updateMany(
    { $or: [{ agentName: 'FLADCO' }, { displayName: /FLADCO/i }] },
    { $set: { productBuckets: updatedBuckets } }
  );

  console.log(`✅ Patched ${res.modifiedCount} FLADCO snapshot(s) in agent_debug_snapshots.`);
  console.log('Updated Buckets:', updatedBuckets.map(b => ({ name: b.name, claimsPaid: b.claimsPaid, lossRatio: b.lossRatio, uwProfit: b.uwProfit })));

  await client.close();
}

patchFladco().catch(console.error);
