import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
dotenv.config();

async function testStream() {
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

  console.time('Stream and aggregate');

  interface CoverageSummary {
    coverageCode: string;
    coverageName: string;
    productType: string;
    maxMarkup: number;
    maxMarkupContract: string;
    maxMarkupDealer: string;
    maxMarkupRetail: number;
    maxMarkupNet: number;
    minMarkup: number;
    totalMarkup: number;
    contractCount: number;
    dealers: Set<string>;
  }

  interface DealerCoverageSummary {
    dealerName: string;
    dealerNumber: string;
    dealerCity: string;
    dealerState: string;
    coverageCode: string;
    coverageName: string;
    productType: string;
    maxMarkup: number;
    maxMarkupContract: string;
    maxMarkupRetail: number;
    maxMarkupNet: number;
    minMarkup: number;
    totalMarkup: number;
    contractCount: number;
  }

  const coverageMap = new Map<string, CoverageSummary>();
  const dealerCoverageMap = new Map<string, DealerCoverageSummary>();

  const cursor = coll.find(flMatch, {
    projection: {
      'metadata.DealerName': 1,
      'metadata.DealerNumber': 1,
      'metadata.DealerCity': 1,
      'metadata.DealerState/Province': 1,
      'metadata.Contract#': 1,
      'metadata.CoverageCode': 1,
      'metadata.CoverageName': 1,
      'metadata.ProductType': 1,
      'metadata.F&IMarkup': 1,
      'metadata.RetailCost': 1,
      'metadata.NetCost': 1,
    },
  });

  let count = 0;
  for await (const doc of cursor) {
    count++;
    const m = doc.metadata || {};
    const dealerName = m.DealerName || 'Unknown Dealer';
    const dealerNumber = m.DealerNumber || '';
    const dealerCity = m.DealerCity || '';
    const dealerState = m['DealerState/Province'] || 'FL';
    const contractNo = m['Contract#'] || '';
    const coverageCode = m.CoverageCode || 'Unknown Code';
    const coverageName = m.CoverageName || '';
    const productType = m.ProductType || '';
    const retail = parseFloat(m.RetailCost || '0') || 0;
    const net = parseFloat(m.NetCost || '0') || 0;
    const markup = parseFloat(m['F&IMarkup'] || '0') || 0;

    // 1. Coverage Map
    let cov = coverageMap.get(coverageCode);
    if (!cov) {
      cov = {
        coverageCode,
        coverageName,
        productType,
        maxMarkup: markup,
        maxMarkupContract: contractNo,
        maxMarkupDealer: dealerName,
        maxMarkupRetail: retail,
        maxMarkupNet: net,
        minMarkup: markup,
        totalMarkup: markup,
        contractCount: 1,
        dealers: new Set([dealerName]),
      };
      coverageMap.set(coverageCode, cov);
    } else {
      cov.contractCount++;
      cov.totalMarkup += markup;
      cov.dealers.add(dealerName);
      if (markup > cov.maxMarkup) {
        cov.maxMarkup = markup;
        cov.maxMarkupContract = contractNo;
        cov.maxMarkupDealer = dealerName;
        cov.maxMarkupRetail = retail;
        cov.maxMarkupNet = net;
      }
      if (markup < cov.minMarkup) {
        cov.minMarkup = markup;
      }
    }

    // 2. Dealer x Coverage Map
    const dcKey = `${dealerName}|${dealerNumber}|${coverageCode}`;
    let dc = dealerCoverageMap.get(dcKey);
    if (!dc) {
      dc = {
        dealerName,
        dealerNumber,
        dealerCity,
        dealerState,
        coverageCode,
        coverageName,
        productType,
        maxMarkup: markup,
        maxMarkupContract: contractNo,
        maxMarkupRetail: retail,
        maxMarkupNet: net,
        minMarkup: markup,
        totalMarkup: markup,
        contractCount: 1,
      };
      dealerCoverageMap.set(dcKey, dc);
    } else {
      dc.contractCount++;
      dc.totalMarkup += markup;
      if (markup > dc.maxMarkup) {
        dc.maxMarkup = markup;
        dc.maxMarkupContract = contractNo;
        dc.maxMarkupRetail = retail;
        dc.maxMarkupNet = net;
      }
      if (markup < dc.minMarkup) {
        dc.minMarkup = markup;
      }
    }
  }

  console.timeEnd('Stream and aggregate');
  console.log(`Processed ${count} contracts.`);
  console.log(`Unique Coverage Codes: ${coverageMap.size}`);
  console.log(`Unique Dealer-Coverage Pairs: ${dealerCoverageMap.size}`);

  // Top 5 coverage codes
  const sorted = Array.from(coverageMap.values()).sort((a, b) => b.maxMarkup - a.maxMarkup);
  console.log('Top 5 Codes with Highest Markup Contract:');
  sorted.slice(0, 5).forEach((c) => {
    console.log(
      `- ${c.coverageCode} (${c.coverageName}): Max Markup $${c.maxMarkup.toFixed(2)} on Contract #${c.maxMarkupContract} at ${c.maxMarkupDealer} (Retail: $${c.maxMarkupRetail}, Net: $${c.maxMarkupNet})`,
    );
  });

  await client.close();
}

testStream().catch(console.error);
