import { type MongoService } from '../services/mongo.service';
import { type MongoSourceConfig, type UnknownDocument } from '../models/report.types';

export class DataRepository {
  constructor(
    private readonly mongoService: MongoService,
    private readonly sources: MongoSourceConfig,
  ) {}

  /**
   * Extracts raw contract data from ContractDataDB
   */
  async getContracts(): Promise<UnknownDocument[]> {
    return this.readCollection(this.sources.contractDb, this.sources.contractCollection);
  }

  /**
   * Extracts raw claim data from ClaimDataDB
   */
  async getClaims(): Promise<UnknownDocument[]> {
    return this.readCollection(this.sources.claimDb, this.sources.claimCollection);
  }

  /**
   * Extracts raw cancellation data from CancelDataDB
   */
  async getCancellations(): Promise<UnknownDocument[]> {
    return this.readCollection(this.sources.cancellationDb, this.sources.cancellationCollection);
  }

  private async readCollection(
    database: string,
    collectionName: string,
  ): Promise<UnknownDocument[]> {
    const collection = this.mongoService
      .getDb(database)
      .collection<UnknownDocument>(collectionName);
    // Project out known customer PII at the source. Field variants are intentionally included.
    const projection = {
      CustomerAddress1: 0,
      CustomerAddress2: 0,
      CustomerCity: 0,
      CustomerState: 0,
      CustomerZipCode: 0,
      CustomerCountry: 0,
      CustomerEmail: 0,
      CustomerFirstName: 0,
      CustomerLastName: 0,
      CustomerPhoneNumber: 0,
      CustomerSecondaryPhone: 0,
      'Customer Address 1': 0,
      'Customer Address 2': 0,
      'Customer City': 0,
      'Customer State': 0,
      'Customer Zip': 0,
      'Customer Country': 0,
      'Customer Email': 0,
      'Customer First Name': 0,
      'Customer Last name': 0,
      'Customer Phone': 0,
      VIN: 0,
    };
    return collection.find({}, { projection }).toArray();
  }
}
