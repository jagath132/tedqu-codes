import { BaseEntity, ExecutionResult } from '../models/ExecutionModels';

export interface ReportGenerationOptions {
  prefix?: string;
  autoOpen?: boolean;
}

export interface IReportGenerator<T extends BaseEntity = BaseEntity> {
  generate(results: ExecutionResult<T>[], options?: ReportGenerationOptions): Promise<string>;
}
