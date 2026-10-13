import { BaseEntity } from '../models/ExecutionModels';

export interface DataReaderOptions {
  sheetNames?: string[] | string;
  maxRecords?: number;
}

export interface IDataReader<T extends BaseEntity> {
  read(sourcePath: string, options?: DataReaderOptions): Promise<T[]> | T[];
  inspectSheets?(sourcePath: string): Promise<string[]> | string[];
}
