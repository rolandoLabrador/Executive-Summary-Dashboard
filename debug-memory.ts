import { MongoClient } from 'mongodb';
import * as dotenv from 'dotenv';
import { loadMongoSourceConfig, loadMongoUri } from './src/config';

dotenv.config();

function logMemory(checkpoint: string): void {
  const mem = process.memoryUsage();
  const heapUsedMB = Math.round(mem.heapUsed / 1024 / 1024);
  const heapTotalMB = Math.round(mem.heapTotal / 1024 / 1024);
  const rssMB = Math.round(mem.rss / 1024 / 1024);
  console.log(
    `⏱️ ${checkpoint.padEnd(30)} | Heap: ${heapUsedMB.toString().padStart(4)} MB / ${heapTotalMB.toString().padStart(4)} MB | RSS: ${rssMB.toString().padStart(4)} MB`,
  );
}

async function run() {
  const uri = loadMongoUri();
  const sources = loadMongoSourceConfig();
  console.log('Connecting to MongoDB...');
  const client = new MongoClient(uri);
  await client.connect();
  console.log('Connected to MongoDB.\n');

  const db = client.db(sources.contractDb);
  const collection = db.collection(sources.contractCollection);

  const total = await collection.countDocuments();
  console.log(`Found ${total.toLocaleString()} total contracts in the database.\n`);

  console.log('Testing raw download size (simulating .toArray())...');

  const cursor = collection.find({});
  const inMemoryArray = [];
  let count = 0;

  for await (const doc of cursor) {
    inMemoryArray.push(doc);
    count++;

    if (count % 50000 === 0) {
      logMemory(`Downloaded ${count.toLocaleString()} docs`);
    }
  }

  logMemory(`Final: ${count.toLocaleString()} docs`);
  await client.close();
}

run().catch(console.error);
