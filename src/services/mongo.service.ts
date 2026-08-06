import { MongoClient, type Db } from 'mongodb';

export class MongoService {
  private client: MongoClient;

  constructor(uri: string) {
    this.client = new MongoClient(uri, {
      appName: 'omnishield-executive-report',
      maxPoolSize: 10,
      serverSelectionTimeoutMS: 15_000,
    });
  }

  /**
   * Establishes a single connection pool to be reused across the application.
   */
  async connect(): Promise<void> {
    try {
      console.log('Connecting to MongoDB...');
      await this.client.connect();
      console.log('Connected to MongoDB.');
    } catch (error) {
      console.error(
        'MongoDB connection failed. Verify the URI, credentials, and Atlas network allowlist.',
      );
      throw error;
    }
  }

  /**
   * Returns a reference to a specific database.
   */
  getDb(dbName: string): Db {
    return this.client.db(dbName);
  }

  /**
   * Gracefully closes the connection pool.
   */
  async close(): Promise<void> {
    try {
      await this.client.close();
      console.log('MongoDB connection gracefully closed.');
    } catch (error) {
      console.error('Failed to close MongoDB connection:', error);
    }
  }
}
