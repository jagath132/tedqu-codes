export type EntityStatus = 'PENDING' | 'RUNNING' | 'SUCCESS' | 'DUPLICATE' | 'FAILED' | 'SKIPPED';

export type AutomationRunMode = 'HEADLESS' | 'HEADED' | 'DRY_RUN';

export interface BaseEntity {
  rowNumber: number;
  [key: string]: any;
}

export interface ExecutionResult<T extends BaseEntity = BaseEntity> {
  entity: T;
  status: EntityStatus;
  message: string;
  executionTimeMs: number;
  timestamp: string;
  attempts?: number;
}

export interface RunSummary {
  total: number;
  processed: number;
  successCount: number;
  duplicateCount: number;
  failedCount: number;
  skippedCount: number;
  durationMs: number;
  status: 'IDLE' | 'RUNNING' | 'COMPLETED' | 'ABORTED' | 'FAILED';
  startedAt: string;
  completedAt?: string;
  results: ExecutionResult[];
}
