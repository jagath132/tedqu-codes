import { BaseEntity, ExecutionResult } from '../models/ExecutionModels';

export interface CheckpointPayload<T extends BaseEntity = BaseEntity> {
  sourcePath: string;
  totalRecords: number;
  completedRowNumbers: number[];
  results: ExecutionResult<T>[];
  lastUpdated: string;
}

export interface ICheckpointStore<T extends BaseEntity = BaseEntity> {
  saveCheckpoint(payload: CheckpointPayload<T>): Promise<void> | void;
  loadCheckpoint(): Promise<CheckpointPayload<T> | null> | CheckpointPayload<T> | null;
  clearCheckpoint(): Promise<void> | void;
}
