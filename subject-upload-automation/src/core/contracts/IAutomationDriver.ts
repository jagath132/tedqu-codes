import { BaseEntity, EntityStatus } from '../models/ExecutionModels';

export interface DriverExecutionResult {
  status: EntityStatus;
  message: string;
}

export interface IAutomationDriver<T extends BaseEntity> {
  initialize(): Promise<void>;
  processEntity(entity: T): Promise<DriverExecutionResult>;
  onFailure?(entity: T, error: Error): Promise<void>;
  teardown?(): Promise<void>;
}
