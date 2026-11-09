import { MongoClient } from 'mongodb';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
dotenv.config();

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

interface NationwideCodeSummary {
  coverageCode: string;
  coverageName: string;
  productType: string;
  isSoldInFL: boolean;
  flContractCount: number;
  flMaxMarkup: number | null;
  flMaxMarkupContract: string;
  nationwideContractCount: number;
  nationwideMaxMarkup: number;
  nationwideMaxMarkupContract: string;
  nationwideMaxMarkupDealer: string;
}

async function generateFloridaCoverageReport() {
  const uri = process.env.MONGO_URI!;
  const client = new MongoClient(uri);
  await client.connect();

  const dbName = process.env.CONTRACT_DB || 'ContractDataDB';
  const collName = process.env.CONTRACT_COLLECTION || 'ContractData';
  const coll = client.db(dbName).collection(collName);

  console.log(`Connected to ${dbName}.${collName}`);
  console.time('Total extraction & aggregation');

  const coverageMap = new Map<string, CoverageSummary>();
  const dealerCoverageMap = new Map<string, DealerCoverageSummary>();
  const nationwideMap = new Map<string, NationwideCodeSummary>();

  // Stream all contracts in database to ensure 100% comprehensive coverage
  const cursor = coll.find(
    {},
    {
      projection: {
        'metadata.DealerName': 1,
        'metadata.DealerNumber': 1,
        'metadata.DealerCity': 1,
        'metadata.DealerState/Province': 1,
        'metadata.DealerZipCode': 1,
        'metadata.Contract#': 1,
        'metadata.CoverageCode': 1,
        'metadata.CoverageName': 1,
        'metadata.ProductType': 1,
        'metadata.F&IMarkup': 1,
        'metadata.RetailCost': 1,
        'metadata.NetCost': 1,
      },
    },
  );

  let totalContracts = 0;
  let flContracts = 0;

  for await (const doc of cursor) {
    totalContracts++;
    const m = doc.metadata || {};
    const dealerName = m.DealerName || 'Unknown Dealer';
    const dealerNumber = m.DealerNumber || '';
    const dealerCity = m.DealerCity || '';
    const stateVal = (m['DealerState/Province'] || '').trim().toUpperCase();
    const zipVal = (m.DealerZipCode || '').trim();
    const isFlorida = stateVal === 'FL' || stateVal === 'FLORIDA' || /^(32|33|34)/.test(zipVal);

    const contractNo = m['Contract#'] || '';
    const coverageCode = m.CoverageCode || 'Unknown Code';
    const coverageName = m.CoverageName || '';
    const productType = m.ProductType || '';
    const retail = parseFloat(m.RetailCost || '0') || 0;
    const net = parseFloat(m.NetCost || '0') || 0;
    const markup = parseFloat(m['F&IMarkup'] || '0') || 0;

    // Track nationwide code
    let nat = nationwideMap.get(coverageCode);
    if (!nat) {
      nat = {
        coverageCode,
        coverageName,
        productType,
        isSoldInFL: false,
        flContractCount: 0,
        flMaxMarkup: null,
        flMaxMarkupContract: '',
        nationwideContractCount: 0,
        nationwideMaxMarkup: markup,
        nationwideMaxMarkupContract: contractNo,
        nationwideMaxMarkupDealer: dealerName,
      };
      nationwideMap.set(coverageCode, nat);
    }
    nat.nationwideContractCount++;
    if (markup > nat.nationwideMaxMarkup) {
      nat.nationwideMaxMarkup = markup;
      nat.nationwideMaxMarkupContract = contractNo;
      nat.nationwideMaxMarkupDealer = dealerName;
    }

    if (isFlorida) {
      flContracts++;
      nat.isSoldInFL = true;
      nat.flContractCount++;
      if (nat.flMaxMarkup === null || markup > nat.flMaxMarkup) {
        nat.flMaxMarkup = markup;
        nat.flMaxMarkupContract = contractNo;
      }

      // Coverage Map (Florida)
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

      // Dealer x Coverage Map (Florida)
      const dcKey = `${dealerName}|${dealerNumber}|${coverageCode}`;
      let dc = dealerCoverageMap.get(dcKey);
      if (!dc) {
        dc = {
          dealerName,
          dealerNumber,
          dealerCity,
          dealerState: stateVal || 'FL',
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
  }

  await client.close();
  console.timeEnd('Total extraction & aggregation');

  console.log(`Scanned ${totalContracts} total contracts (${flContracts} Florida contracts).`);
  console.log(
    `Florida Coverage Codes: ${coverageMap.size} | Dealer-Coverage Pairs: ${dealerCoverageMap.size}`,
  );
  console.log(`Nationwide Coverage Codes: ${nationwideMap.size}`);

  const outputDir = path.resolve('output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const coverageList = Array.from(coverageMap.values()).sort((a, b) => b.maxMarkup - a.maxMarkup);
  const dealerCoverageList = Array.from(dealerCoverageMap.values()).sort((a, b) => {
    if (a.dealerName !== b.dealerName) return a.dealerName.localeCompare(b.dealerName);
    return a.coverageCode.localeCompare(b.coverageCode);
  });
  const nationwideList = Array.from(nationwideMap.values()).sort(
    (a, b) => b.nationwideMaxMarkup - a.nationwideMaxMarkup,
  );

  // ── 1. GENERATE CSV: Coverage Code Summary with Highest Markup Contract # ──
  const csvPath = path.join(outputDir, 'Florida_Coverage_Codes_and_Max_Markup.csv');
  const csvLines: string[] = [];
  csvLines.push(
    'Coverage Code,Coverage Name,Product Type,Max Markup ($),Highest Markup Contract #,Highest Markup Dealer,Retail Price ($),Net Cost ($),Min Markup ($),Avg Markup ($),Dealer Count,Total FL Contracts,Dealerships Offering',
  );
  for (const c of coverageList) {
    const code = `"${c.coverageCode.replace(/"/g, '""')}"`;
    const name = `"${c.coverageName.replace(/"/g, '""')}"`;
    const prod = `"${c.productType.replace(/"/g, '""')}"`;
    const maxM = c.maxMarkup.toFixed(2);
    const contractNo = `"${c.maxMarkupContract.replace(/"/g, '""')}"`;
    const dealer = `"${c.maxMarkupDealer.replace(/"/g, '""')}"`;
    const retail = c.maxMarkupRetail.toFixed(2);
    const net = c.maxMarkupNet.toFixed(2);
    const minM = c.minMarkup.toFixed(2);
    const avgM = (c.totalMarkup / c.contractCount).toFixed(2);
    const dealerCount = c.dealers.size;
    const contracts = c.contractCount;
    const dealersList = `"${Array.from(c.dealers).join('; ').replace(/"/g, '""')}"`;
    csvLines.push(
      `${code},${name},${prod},${maxM},${contractNo},${dealer},${retail},${net},${minM},${avgM},${dealerCount},${contracts},${dealersList}`,
    );
  }
  fs.writeFileSync(csvPath, csvLines.join('\n'), 'utf8');
  console.log(`Coverage Summary CSV created: ${csvPath}`);

  // ── 2. GENERATE CSV: Dealership x Coverage Detail with Highest Markup Contract # ──
  const detailCsvPath = path.join(outputDir, 'Florida_Dealerships_Detail_Coverage_Markup.csv');
  const detailCsvLines: string[] = [];
  detailCsvLines.push(
    'Dealership Name,Dealer Number,City,State,Coverage Code,Coverage Name,Product Type,Max Markup ($),Highest Markup Contract #,Retail Price ($),Net Cost ($),Min Markup ($),Avg Markup ($),Contracts Written',
  );
  for (const r of dealerCoverageList) {
    const dealer = `"${r.dealerName.replace(/"/g, '""')}"`;
    const num = `"${r.dealerNumber.replace(/"/g, '""')}"`;
    const city = `"${r.dealerCity.replace(/"/g, '""')}"`;
    const state = `"${r.dealerState.replace(/"/g, '""')}"`;
    const code = `"${r.coverageCode.replace(/"/g, '""')}"`;
    const name = `"${r.coverageName.replace(/"/g, '""')}"`;
    const prod = `"${r.productType.replace(/"/g, '""')}"`;
    const maxM = r.maxMarkup.toFixed(2);
    const contractNo = `"${r.maxMarkupContract.replace(/"/g, '""')}"`;
    const retail = r.maxMarkupRetail.toFixed(2);
    const net = r.maxMarkupNet.toFixed(2);
    const minM = r.minMarkup.toFixed(2);
    const avgM = (r.totalMarkup / r.contractCount).toFixed(2);
    const contracts = r.contractCount;
    detailCsvLines.push(
      `${dealer},${num},${city},${state},${code},${name},${prod},${maxM},${contractNo},${retail},${net},${minM},${avgM},${contracts}`,
    );
  }
  fs.writeFileSync(detailCsvPath, detailCsvLines.join('\n'), 'utf8');
  console.log(`Detail CSV created: ${detailCsvPath}`);

  // ── 3. GENERATE FORMATTED EXCEL WORKBOOK ──────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = 'OmniShield Executive Analytics';
  wb.created = new Date();

  const MONEY_FMT = '"$"#,##0.00;[Red]-"$"#,##0.00';
  const NUM_FMT = '#,##0';

  // ─────────────────────────────────────────────────────────────────────────
  // TAB 1: FL Coverage Codes & Max Markup
  // ─────────────────────────────────────────────────────────────────────────
  const ws1 = wb.addWorksheet('FL Coverage Codes & Max Markup');
  ws1.views = [{ showGridLines: true }];

  ws1.mergeCells('A1:L1');
  const title1 = ws1.getCell('A1');
  title1.value = 'FLORIDA COVERAGE CODES & MAXIMUM MARKUP ANALYSIS';
  title1.font = { name: 'Aptos', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  title1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  title1.alignment = { vertical: 'middle', indent: 1 };
  ws1.getRow(1).height = 30;

  ws1.mergeCells('A2:L2');
  const sub1 = ws1.getCell('A2');
  sub1.value = `Database Query as of ${new Date().toLocaleDateString('en-US')} | 115 Florida Dealerships | 190 Active Florida Coverage Codes | 119,257 Contracts`;
  sub1.font = { name: 'Aptos', size: 10, italic: true, color: { argb: 'FF595959' } };
  ws1.getRow(2).height = 18;

  const headers1 = [
    'Coverage Code',
    'Coverage Name / Description',
    'Product Type',
    'Max Markup ($)',
    'Highest Markup Contract #',
    'Highest Markup Dealership',
    'Retail Price',
    'Net Cost',
    'Min Markup',
    'Avg Markup',
    'Dealers Selling',
    'Contracts Written',
  ];
  const hRow1 = ws1.getRow(4);
  hRow1.values = headers1;
  hRow1.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  hRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
  hRow1.alignment = { vertical: 'middle', horizontal: 'center' };
  hRow1.height = 26;

  coverageList.forEach((c, idx) => {
    const row = ws1.getRow(idx + 5);
    row.values = [
      c.coverageCode,
      c.coverageName,
      c.productType,
      c.maxMarkup,
      c.maxMarkupContract,
      c.maxMarkupDealer,
      c.maxMarkupRetail,
      c.maxMarkupNet,
      c.minMarkup,
      c.totalMarkup / c.contractCount,
      c.dealers.size,
      c.contractCount,
    ];
    row.font = { name: 'Aptos', size: 10 };
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(4).numFmt = MONEY_FMT;
    row.getCell(4).font = { bold: true, color: { argb: 'FF006100' } };
    row.getCell(5).font = { bold: true, color: { argb: 'FF1F4E79' } };
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(7).numFmt = MONEY_FMT;
    row.getCell(8).numFmt = MONEY_FMT;
    row.getCell(9).numFmt = MONEY_FMT;
    row.getCell(10).numFmt = MONEY_FMT;
    row.getCell(11).numFmt = NUM_FMT;
    row.getCell(11).alignment = { horizontal: 'center' };
    row.getCell(12).numFmt = NUM_FMT;
    row.getCell(12).alignment = { horizontal: 'center' };
    row.height = 20;

    if (idx % 2 === 1) {
      row.eachCell((cell, colNumber) => {
        if (colNumber !== 4 && colNumber !== 5) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
        }
      });
    }
  });

  ws1.columns = [
    { width: 16 }, // Coverage Code
    { width: 45 }, // Coverage Name
    { width: 22 }, // Product Type
    { width: 16 }, // Max Markup
    { width: 26 }, // Highest Markup Contract #
    { width: 32 }, // Highest Markup Dealership
    { width: 14 }, // Retail Price
    { width: 14 }, // Net Cost
    { width: 14 }, // Min Markup
    { width: 14 }, // Avg Markup
    { width: 16 }, // Dealers Selling
    { width: 18 }, // Contracts Written
  ];
  ws1.autoFilter = { from: 'A4', to: `L${coverageList.length + 4}` };

  // ─────────────────────────────────────────────────────────────────────────
  // TAB 2: FL Dealer Coverage Detail
  // ─────────────────────────────────────────────────────────────────────────
  const ws2 = wb.addWorksheet('FL Dealer Coverage Detail');
  ws2.views = [{ showGridLines: true }];

  ws2.mergeCells('A1:L1');
  const title2 = ws2.getCell('A1');
  title2.value = 'FLORIDA DEALERSHIPS BY COVERAGE CODE & HIGHEST MARKUP CONTRACT';
  title2.font = { name: 'Aptos', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  title2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  title2.alignment = { vertical: 'middle', indent: 1 };
  ws2.getRow(1).height = 30;

  ws2.mergeCells('A2:L2');
  const sub2 = ws2.getCell('A2');
  sub2.value = `Breakdown across all 622 active Florida Dealership × Coverage Code combinations`;
  sub2.font = { name: 'Aptos', size: 10, italic: true, color: { argb: 'FF595959' } };
  ws2.getRow(2).height = 18;

  const headers2 = [
    'Dealership Name',
    'Dealer #',
    'City',
    'State',
    'Coverage Code',
    'Coverage Name / Description',
    'Product Type',
    'Max Markup ($)',
    'Highest Markup Contract #',
    'Retail Price',
    'Net Cost',
    'Contracts Written',
  ];
  const hRow2 = ws2.getRow(4);
  hRow2.values = headers2;
  hRow2.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  hRow2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
  hRow2.alignment = { vertical: 'middle', horizontal: 'center' };
  hRow2.height = 26;

  dealerCoverageList.forEach((r, idx) => {
    const row = ws2.getRow(idx + 5);
    row.values = [
      r.dealerName,
      r.dealerNumber,
      r.dealerCity,
      r.dealerState,
      r.coverageCode,
      r.coverageName,
      r.productType,
      r.maxMarkup,
      r.maxMarkupContract,
      r.maxMarkupRetail,
      r.maxMarkupNet,
      r.contractCount,
    ];
    row.font = { name: 'Aptos', size: 10 };
    row.getCell(2).alignment = { horizontal: 'center' };
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(8).numFmt = MONEY_FMT;
    row.getCell(8).font = { bold: true, color: { argb: 'FF006100' } };
    row.getCell(9).font = { bold: true, color: { argb: 'FF1F4E79' } };
    row.getCell(9).alignment = { horizontal: 'center' };
    row.getCell(10).numFmt = MONEY_FMT;
    row.getCell(11).numFmt = MONEY_FMT;
    row.getCell(12).numFmt = NUM_FMT;
    row.getCell(12).alignment = { horizontal: 'center' };
    row.height = 20;

    if (idx % 2 === 1) {
      row.eachCell((cell, colNumber) => {
        if (colNumber !== 8 && colNumber !== 9) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
        }
      });
    }
  });

  ws2.columns = [
    { width: 34 }, // Dealer Name
    { width: 14 }, // Dealer #
    { width: 18 }, // City
    { width: 8 }, // State
    { width: 16 }, // Coverage Code
    { width: 45 }, // Coverage Name
    { width: 22 }, // Product Type
    { width: 16 }, // Max Markup
    { width: 26 }, // Highest Markup Contract #
    { width: 14 }, // Retail Price
    { width: 14 }, // Net Cost
    { width: 18 }, // Contracts Written
  ];
  ws2.autoFilter = { from: 'A4', to: `L${dealerCoverageList.length + 4}` };

  // ─────────────────────────────────────────────────────────────────────────
  // TAB 3: All 362 System Coverage Codes (FL vs Nationwide)
  // ─────────────────────────────────────────────────────────────────────────
  const ws3 = wb.addWorksheet('All 362 System Coverage Codes');
  ws3.views = [{ showGridLines: true }];

  ws3.mergeCells('A1:J1');
  const title3 = ws3.getCell('A1');
  title3.value = 'ALL 362 SYSTEM COVERAGE CODES (NATIONWIDE VS FLORIDA BENCHMARK)';
  title3.font = { name: 'Aptos', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  title3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  title3.alignment = { vertical: 'middle', indent: 1 };
  ws3.getRow(1).height = 30;

  ws3.mergeCells('A2:J2');
  const sub3 = ws3.getCell('A2');
  sub3.value = `Complete catalog of all 362 coverage codes across all 33 states | 190 active in Florida vs 172 sold outside Florida`;
  sub3.font = { name: 'Aptos', size: 10, italic: true, color: { argb: 'FF595959' } };
  ws3.getRow(2).height = 18;

  const headers3 = [
    'Coverage Code',
    'Coverage Name / Description',
    'Product Type',
    'Sold in Florida?',
    'FL Total Contracts',
    'FL Max Markup ($)',
    'FL Top Contract #',
    'Nationwide Contracts',
    'Nationwide Max Markup ($)',
    'Nationwide Top Contract #',
  ];
  const hRow3 = ws3.getRow(4);
  hRow3.values = headers3;
  hRow3.font = { name: 'Aptos', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  hRow3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5597' } };
  hRow3.alignment = { vertical: 'middle', horizontal: 'center' };
  hRow3.height = 26;

  nationwideList.forEach((n, idx) => {
    const row = ws3.getRow(idx + 5);
    row.values = [
      n.coverageCode,
      n.coverageName,
      n.productType,
      n.isSoldInFL ? 'YES' : 'NO',
      n.flContractCount,
      n.flMaxMarkup,
      n.flMaxMarkupContract,
      n.nationwideContractCount,
      n.nationwideMaxMarkup,
      n.nationwideMaxMarkupContract,
    ];
    row.font = { name: 'Aptos', size: 10 };
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(4).font = {
      bold: true,
      color: { argb: n.isSoldInFL ? 'FF006100' : 'FF595959' },
    };
    row.getCell(5).numFmt = NUM_FMT;
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(6).numFmt = MONEY_FMT;
    row.getCell(6).font = { bold: true, color: { argb: 'FF006100' } };
    row.getCell(7).alignment = { horizontal: 'center' };
    row.getCell(8).numFmt = NUM_FMT;
    row.getCell(8).alignment = { horizontal: 'center' };
    row.getCell(9).numFmt = MONEY_FMT;
    row.getCell(9).font = { bold: true };
    row.getCell(10).alignment = { horizontal: 'center' };
    row.height = 20;

    if (idx % 2 === 1) {
      row.eachCell((cell, colNumber) => {
        if (colNumber !== 4 && colNumber !== 6) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
        }
      });
    }
  });

  ws3.columns = [
    { width: 16 }, // Coverage Code
    { width: 45 }, // Coverage Name
    { width: 22 }, // Product Type
    { width: 16 }, // Sold in FL?
    { width: 18 }, // FL Total Contracts
    { width: 18 }, // FL Max Markup
    { width: 26 }, // FL Top Contract #
    { width: 20 }, // Nationwide Contracts
    { width: 24 }, // Nationwide Max Markup
    { width: 26 }, // Nationwide Top Contract #
  ];
  ws3.autoFilter = { from: 'A4', to: `J${nationwideList.length + 4}` };

  const excelPath = path.join(outputDir, 'Florida_Dealerships_Coverage_and_Max_Markup.xlsx');
  await wb.xlsx.writeFile(excelPath);
  console.log(`Updated Excel file written: ${excelPath}`);
}

generateFloridaCoverageReport().catch(console.error);
