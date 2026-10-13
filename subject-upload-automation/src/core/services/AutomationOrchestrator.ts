import { BaseEntity, ExecutionResult, RunSummary, AutomationRunMode } from '../models/ExecutionModels';
import { IDataReader, DataReaderOptions } from '../contracts/IDataReader';
import { IValidator } from '../contracts/IValidator';
import { IAutomationDriver } from '../contracts/IAutomationDriver';
import { ICheckpointStore } from '../contracts/ICheckpointStore';
import { IReportGenerator } from '../contracts/IReportGenerator';
import { ITelemetryEmitter } from '../contracts/ITelemetryEmitter';
import { AutomationEventType } from '../events/AutomationEvents';

export interface OrchestratorOptions<T extends BaseEntity> {
  dataReader: IDataReader<T>;
  validator: IValidator<T>;
  driver?: IAutomationDriver<T>;
  checkpointStore?: ICheckpointStore<T>;
  reportGenerator?: IReportGenerator<T>;
  telemetryEmitter: ITelemetryEmitter;
  maxRetries?: number;
  retryDelayMs?: number;
}

export interface RunExecutionConfig {
  sourcePath: string;
  mode: AutomationRunMode;
  readerOptions?: DataReaderOptions;
  resume?: boolean;
  getItemDisplayName?: (entity: any) => string;
}

export class AutomationOrchestrator<T extends BaseEntity> {
  private dataReader: IDataReader<T>;
  private validator: IValidator<T>;
  private driver?: IAutomationDriver<T>;
  private checkpointStore?: ICheckpointStore<T>;
  private reportGenerator?: IReportGenerator<T>;
  private telemetryEmitter: ITelemetryEmitter;
  private maxRetries: number;
  private retryDelayMs: number;

  constructor(options: OrchestratorOptions<T>) {
    this.dataReader = options.dataReader;
    this.validator = options.validator;
    this.driver = options.driver;
    this.checkpointStore = options.checkpointStore;
    this.reportGenerator = options.reportGenerator;
    this.telemetryEmitter = options.telemetryEmitter;
    this.maxRetries = options.maxRetries ?? 2;
    this.retryDelayMs = options.retryDelayMs ?? 1000;
  }

  public async execute(config: RunExecutionConfig): Promise<RunSummary> {
    const startTime = Date.now();
    const startedAt = new Date().toISOString();

    // 1. Read input dataset
    const entities = await this.dataReader.read(config.sourcePath, config.readerOptions);
    const totalRecords = entities.length;

    this.telemetryEmitter.emit({
      type: AutomationEventType.RUN_STARTED,
      totalRecords,
      sourcePath: config.sourcePath,
      mode: config.mode,
      timestamp: startedAt
    });

    const results: ExecutionResult<T>[] = [];
    let completedRowNumbers = new Set<number>();

    // 2. Handle Checkpoint Recovery
    if (config.resume && this.checkpointStore) {
      const checkpoint = await this.checkpointStore.loadCheckpoint();
      if (checkpoint && checkpoint.results) {
        completedRowNumbers = new Set(checkpoint.completedRowNumbers);
        results.push(...checkpoint.results);
      }
    } else if (this.checkpointStore) {
      await this.checkpointStore.clearCheckpoint();
    }

    // 3. Dry Run execution branch
    if (config.mode === 'DRY_RUN') {
      return this.executeDryRun(entities, config, startTime, startedAt);
    }

    // 4. Browser-driven execution
    if (!this.driver) {
      throw new Error('Automation driver is required for non-dry-run modes.');
    }

    await this.driver.initialize();

    try {
      for (let i = 0; i < entities.length; i++) {
        const entity = entities[i];

        if (completedRowNumbers.has(entity.rowNumber)) {
          continue;
        }

        const itemDisplayName = config.getItemDisplayName ? config.getItemDisplayName(entity) : `Row #${entity.rowNumber}`;

        this.telemetryEmitter.emit({
          type: AutomationEventType.RECORD_STARTED,
          entity,
          currentIndex: i + 1,
          totalRecords,
          timestamp: new Date().toISOString()
        });

        const validation = this.validator.validate(entity);
        const recordStartTime = Date.now();
        let result: ExecutionResult<T>;

        if (!validation.isValid) {
          result = {
            entity,
            status: 'SKIPPED',
            message: `Validation failed: ${validation.errors.join('; ')}`,
            executionTimeMs: Date.now() - recordStartTime,
            timestamp: new Date().toISOString(),
            attempts: 0
          };
        } else {
          result = await this.processWithRetries(entity, recordStartTime);
        }

        results.push(result);

        // Update checkpoint
        if (this.checkpointStore) {
          await this.checkpointStore.saveCheckpoint({
            sourcePath: config.sourcePath,
            totalRecords,
            completedRowNumbers: results.map(r => r.entity.rowNumber),
            results,
            lastUpdated: new Date().toISOString()
          });
        }

        // Emit record completed & progress
        this.emitProgress(results, totalRecords, itemDisplayName);
      }
    } finally {
      if (this.driver.teardown) {
        await this.driver.teardown();
      }
    }

    // 5. Finalize run
    const summary = this.createSummary(results, totalRecords, startTime, startedAt, 'COMPLETED');

    let reportPath: string | undefined;
    if (this.reportGenerator) {
      reportPath = await this.reportGenerator.generate(results, { prefix: config.mode });
    }

    if (this.checkpointStore) {
      await this.checkpointStore.clearCheckpoint();
    }

    this.telemetryEmitter.emit({
      type: AutomationEventType.RUN_COMPLETED,
      summary,
      reportPath,
      timestamp: new Date().toISOString()
    });

    return summary;
  }

  private async processWithRetries(entity: T, recordStartTime: number): Promise<ExecutionResult<T>> {
    let attempts = 0;
    let lastError = '';

    while (attempts <= this.maxRetries) {
      attempts++;
      try {
        if (attempts > 1) {
          await this.sleep(this.retryDelayMs);
        }

        const driverResult = await this.driver!.processEntity(entity);
        if (driverResult.status === 'SUCCESS' || driverResult.status === 'DUPLICATE') {
          return {
            entity,
            status: driverResult.status,
            message: driverResult.message,
            executionTimeMs: Date.now() - recordStartTime,
            timestamp: new Date().toISOString(),
            attempts
          };
        }

        // If returned FAILED on non-terminal attempt, loop continues
      } catch (err) {
        lastError = (err as Error).message;
        if (this.driver!.onFailure) {
          await this.driver!.onFailure(entity, err as Error);
        }
      }
    }

    return {
      entity,
      status: 'FAILED',
      message: lastError ? `Error: ${lastError}` : 'Operation failed after retries',
      executionTimeMs: Date.now() - recordStartTime,
      timestamp: new Date().toISOString(),
      attempts
    };
  }

  private async executeDryRun(
    entities: T[],
    config: RunExecutionConfig,
    startTime: number,
    startedAt: string
  ): Promise<RunSummary> {
    const results: ExecutionResult<T>[] = [];
    const totalRecords = entities.length;

    for (let i = 0; i < entities.length; i++) {
      const entity = entities[i];
      const validation = this.validator.validate(entity);
      const itemDisplayName = config.getItemDisplayName ? config.getItemDisplayName(entity) : `Row #${entity.rowNumber}`;

      const result: ExecutionResult<T> = {
        entity,
        status: validation.isValid ? 'SUCCESS' : 'SKIPPED',
        message: validation.isValid ? 'Validation passed (Dry Run)' : `Validation failed: ${validation.errors.join('; ')}`,
        executionTimeMs: 1,
        timestamp: new Date().toISOString(),
        attempts: 1
      };

      results.push(result);
      this.emitProgress(results, totalRecords, itemDisplayName);
    }

    const summary = this.createSummary(results, totalRecords, startTime, startedAt, 'COMPLETED');

    let reportPath: string | undefined;
    if (this.reportGenerator) {
      reportPath = await this.reportGenerator.generate(results, { prefix: 'DRY_RUN' });
    }

    this.telemetryEmitter.emit({
      type: AutomationEventType.RUN_COMPLETED,
      summary,
      reportPath,
      timestamp: new Date().toISOString()
    });

    return summary;
  }

  private emitProgress(results: ExecutionResult<T>[], total: number, currentItem?: string): void {
    const processed = results.length;
    const successCount = results.filter(r => r.status === 'SUCCESS').length;
    const duplicateCount = results.filter(r => r.status === 'DUPLICATE').length;
    const failedCount = results.filter(r => r.status === 'FAILED').length;
    const skippedCount = results.filter(r => r.status === 'SKIPPED').length;
    const percentage = total > 0 ? Math.round((processed / total) * 100) : 0;

    const latestResult = results[results.length - 1];

    this.telemetryEmitter.emit({
      type: AutomationEventType.RECORD_COMPLETED,
      result: latestResult,
      currentIndex: processed,
      totalRecords: total,
      timestamp: new Date().toISOString()
    });

    this.telemetryEmitter.emit({
      type: AutomationEventType.PROGRESS_UPDATED,
      processed,
      total,
      percentage,
      successCount,
      duplicateCount,
      failedCount,
      skippedCount,
      currentItem,
      timestamp: new Date().toISOString()
    });
  }

  private createSummary(
    results: ExecutionResult<T>[],
    total: number,
    startTime: number,
    startedAt: string,
    status: RunSummary['status']
  ): RunSummary {
    return {
      total,
      processed: results.length,
      successCount: results.filter(r => r.status === 'SUCCESS').length,
      duplicateCount: results.filter(r => r.status === 'DUPLICATE').length,
      failedCount: results.filter(r => r.status === 'FAILED').length,
      skippedCount: results.filter(r => r.status === 'SKIPPED').length,
      durationMs: Date.now() - startTime,
      status,
      startedAt,
      completedAt: new Date().toISOString(),
      results
    };
  }

  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}
