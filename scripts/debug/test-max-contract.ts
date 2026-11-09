import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
dotenv.config();

async function testMaxContract() {
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

  const sample = await coll
    .aggregate(
      [
        { $match: flMatch },
        {
          $project: {
            dealerName: '$metadata.DealerName',
            dealerNumber: '$metadata.DealerNumber',
            contractNumber: '$metadata.Contract#',
            coverageCode: '$metadata.CoverageCode',
            coverageName: '$metadata.CoverageName',
            productType: '$metadata.ProductType',
            saleDate: '$metadata.SaleDate',
            retailCost: { $toDouble: { $ifNull: ['$metadata.RetailCost', '0'] } },
            netCost: { $toDouble: { $ifNull: ['$metadata.NetCost', '0'] } },
            markup: { $toDouble: { $ifNull: ['$metadata.F&IMarkup', '0'] } },
          },
        },
        { $sort: { markup: -1 } },
        {
          $group: {
            _id: '$coverageCode',
            coverageName: { $first: '$coverageName' },
            productType: { $first: '$productType' },
            maxMarkup: { $first: '$markup' },
            maxMarkupContract: { $first: '$contractNumber' },
            maxMarkupDealer: { $first: '$dealerName' },
            maxMarkupRetail: { $first: '$retailCost' },
            maxMarkupNet: { $first: '$netCost' },
            maxMarkupSaleDate: { $first: '$saleDate' },
            minMarkup: { $min: '$markup' },
            avgMarkup: { $avg: '$markup' },
            contractCount: { $sum: 1 },
            dealers: { $addToSet: '$dealerName' },
          },
        },
        { $sort: { maxMarkup: -1 } },
        { $limit: 5 },
      ],
      { allowDiskUse: true },
    )
    .toArray();

  console.log('Top 5 Codes with Highest Markup Contract:', JSON.stringify(sample, null, 2));

  await client.close();
}

testMaxContract().catch(console.error);
