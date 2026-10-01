import * as fs from 'fs';
import { parse } from 'csv-parse/sync';
import { MongoClient } from 'mongodb';
import { config } from 'dotenv';
config();

async function run() {
    const client = new MongoClient(process.env.MONGODB_URI!);
    await client.connect();
    const db = client.db('reporting');
    
    console.log(`Parsing output/ContractExport (79).csv...`);
    const content = fs.readFileSync('output/ContractExport (79).csv', 'utf-8');
    const records = parse(content, {
        columns: true,
        skip_empty_lines: true
    });
    
    let baseReserveWritten = 0;
    
    for (const record of records) {
        if (record.DealerName === 'Power Motors East' || !record.DealerName) {
            const raw = record['WrittenAmount.RESERVE.BASERESERVE'];
            if (raw) {
                baseReserveWritten += parseFloat(raw);
            }
        }
    }
    
    console.log(`Total BASERESERVE Written in CSV: ${baseReserveWritten}`);

    const badDates = await db.collection('contracts').find({
      'metadata.DealerName': 'Power Motors East',
      $or: [
        { 'metadata.ActivationDate': null },
        { 'metadata.ActivationDate': '' }
      ]
    }).toArray();
  
    const voided = await db.collection('contracts').find({
      'metadata.DealerName': 'Power Motors East',
      'metadata.ContractStatus': 'V'
    }).toArray();

    let dqBaseReserve = 0;
    for (const doc of badDates) {
        const val = doc.WrittenAmount?.RESERVE?.BASERESERVE || 0;
        dqBaseReserve += parseFloat(val);
        console.log(`- Contract ${doc.metadata['Contract#']} (Missing Date): $${val}`);
    }
    for (const doc of voided) {
        const val = doc.WrittenAmount?.RESERVE?.BASERESERVE || 0;
        dqBaseReserve += parseFloat(val);
        console.log(`- Contract ${doc.metadata['Contract#']} (Voided): $${val}`);
    }
    console.log(`Total BASERESERVE excluded by DQ engine: ${dqBaseReserve}`);
    
    await client.close();
}
run().catch(console.error);
