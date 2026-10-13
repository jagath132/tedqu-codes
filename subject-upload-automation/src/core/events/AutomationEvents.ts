import { BaseEntity, ExecutionResult, RunSummary } from '../models/ExecutionModels';

export enum AutomationEventType {
  RUN_STARTED = 'RUN_STARTED',
  RECORD_STARTED = 'RECORD_STARTED',
  RECORD_COMPLETED = 'RECORD_COMPLETED',
  PROGRESS_UPDATED = 'PROGRESS_UPDATED',
  LOG_EMITTED = 'LOG_EMITTED',
  RUN_COMPLETED = 'RUN_COMPLETED',
  RUN_FAILED = 'RUN_FAILED',
  RUN_ABORTED = 'RUN_ABORTED'
}

export interface RunStartedEvent {
  type: AutomationEventType.RUN_STARTED;
  totalRecords: number;
  sourcePath: string;
  mode: string;
  timestamp: string;
}

export interface RecordStartedEvent<T extends BaseEntity = BaseEntity> {
  type: AutomationEventType.RECORD_STARTED;
  entity: T;
  currentIndex: number;
  totalRecords: number;
  timestamp: string;
}

export interface RecordCompletedEvent<T extends BaseEntity = BaseEntity> {
  type: AutomationEventType.RECORD_COMPLETED;
  result: ExecutionResult<T>;
  currentIndex: number;
  totalRecords: number;
  timestamp: string;
}

export interface ProgressUpdatedEvent {
  type: AutomationEventType.PROGRESS_UPDATED;
  processed: number;
  total: number;
  percentage: number;
  successCount: number;
  duplicateCount: number;
  failedCount: number;
  skippedCount: number;
  currentItem?: string;
  timestamp: string;
}

export interface LogEmittedEvent {
  type: AutomationEventType.LOG_EMITTED;
  level: 'info' | 'warn' | 'error' | 'debug';
  message: string;
  timestamp: string;
}

export interface RunCompletedEvent {
  type: AutomationEventType.RUN_COMPLETED;
  summary: RunSummary;
  reportPath?: string;
  timestamp: string;
}

export interface RunFailedEvent {
  type: AutomationEventType.RUN_FAILED;
  error: string;
  summary?: RunSummary;
  timestamp: string;
}

export type AutomationEvent<T extends BaseEntity = BaseEntity> =
  | RunStartedEvent
  | RecordStartedEvent<T>
  | RecordCompletedEvent<T>
  | ProgressUpdatedEvent
  | LogEmittedEvent
  | RunCompletedEvent
  | RunFailedEvent;
