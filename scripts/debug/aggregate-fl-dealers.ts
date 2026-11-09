import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
dotenv.config();

async function run() {
  const uri = process.env.MONGO_URI!;
  const client = new MongoClient(uri);
  await client.connect();

  const coll = client
    .db(process.env.CONTRACT_DB || 'ContractDataDB')
    .collection(process.env.CONTRACT_COLLECTION || 'ContractData');

  const flMatch = {
    $or: [
      { 'metadata.DealerState/Province': /^FL/i },
      { 'metadata.DealerState/Province': 'Florida' },
    ],
  };

  // Distinct dealerships in Florida
  const distinctDealers = await coll.distinct('metadata.DealerName', flMatch);
  console.log(`Distinct Florida Dealer Names: ${distinctDealers.length}`);

  // Distinct coverage codes in Florida
  const distinctCodes = await coll.distinct('metadata.CoverageCode', flMatch);
  console.log(`Distinct Florida Coverage Codes: ${distinctCodes.length}`);

  // Run aggregation by Dealer + CoverageCode
  const dealerCoverageAgg = await coll
    .aggregate([
      { $match: flMatch },
      {
        $project: {
          dealerName: { $ifNull: ['$metadata.DealerName', 'Unknown Dealer'] },
          dealerNumber: { $ifNull: ['$metadata.DealerNumber', ''] },
          dealerCity: { $ifNull: ['$metadata.DealerCity', ''] },
          dealerState: { $ifNull: ['$metadata.DealerState/Province', 'FL'] },
          coverageCode: { $ifNull: ['$metadata.CoverageCode', 'Unknown Code'] },
          coverageName: { $ifNull: ['$metadata.CoverageName', ''] },
          productType: { $ifNull: ['$metadata.ProductType', ''] },
          markup: {
            $toDouble: {
              $ifNull: ['$metadata.F&IMarkup', '0'],
            },
          },
        },
      },
      {
        $group: {
          _id: {
            dealerName: '$dealerName',
            dealerNumber: '$dealerNumber',
            coverageCode: '$coverageCode',
          },
          dealerCity: { $first: '$dealerCity' },
          dealerState: { $first: '$dealerState' },
          coverageName: { $first: '$coverageName' },
          productType: { $first: '$productType' },
          maxMarkup: { $max: '$markup' },
          minMarkup: { $min: '$markup' },
          avgMarkup: { $avg: '$markup' },
          contractCount: { $sum: 1 },
        },
      },
      { $sort: { '_id.dealerName': 1, '_id.coverageCode': 1 } },
    ])
    .toArray();

  console.log(`Total Dealer-Coverage combinations: ${dealerCoverageAgg.length}`);
  console.log('Sample Dealer-Coverage combination:', dealerCoverageAgg[0]);

  // Run aggregation by CoverageCode overall for FL
  const coverageAgg = await coll
    .aggregate([
      { $match: flMatch },
      {
        $project: {
          dealerName: { $ifNull: ['$metadata.DealerName', 'Unknown Dealer'] },
          coverageCode: { $ifNull: ['$metadata.CoverageCode', 'Unknown Code'] },
          coverageName: { $ifNull: ['$metadata.CoverageName', ''] },
          productType: { $ifNull: ['$metadata.ProductType', ''] },
          markup: {
            $toDouble: {
              $ifNull: ['$metadata.F&IMarkup', '0'],
            },
          },
        },
      },
      {
        $group: {
          _id: '$coverageCode',
          coverageName: { $first: '$coverageName' },
          productType: { $first: '$productType' },
          maxMarkup: { $max: '$markup' },
          minMarkup: { $min: '$markup' },
          avgMarkup: { $avg: '$markup' },
          contractCount: { $sum: 1 },
          dealers: { $addToSet: '$dealerName' },
        },
      },
      { $sort: { maxMarkup: -1 } },
    ])
    .toArray();

  console.log(`Total Florida Coverage Codes overall: ${coverageAgg.length}`);
  console.log(
    'Top 5 Coverage Codes by Max Markup:',
    coverageAgg.slice(0, 5).map((c) => ({
      code: c._id,
      name: c.coverageName,
      product: c.productType,
      maxMarkup: c.maxMarkup,
      dealersCount: c.dealers.length,
      contracts: c.contractCount,
    })),
  );

  await client.close();
}

run().catch(console.error);
