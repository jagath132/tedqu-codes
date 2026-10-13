import { test } from '@playwright/test';
import { logger } from '../utils/logger';
import { ProgressBar } from '../utils/progressBar';
import { ExcelDataReader } from '../src/adapters/ExcelDataReader';
import { SubjectValidator } from '../src/adapters/SubjectValidator';
import { FileCheckpointStore } from '../src/adapters/FileCheckpointStore';
import { ExcelReportGenerator } from '../src/adapters/ExcelReportGenerator';
import { PlaywrightSubjectDriver } from '../src/adapters/PlaywrightSubjectDriver';
import { AutomationOrchestrator } from '../src/core/services/AutomationOrchestrator';
import { TelemetryEventBus } from '../src/core/events/TelemetryEventBus';
import { AutomationEventType, ProgressUpdatedEvent, RecordCompletedEvent } from '../src/core/events/AutomationEvents';
import { AutomationRunMode } from '../src/core/models/ExecutionModels';
import { Subject } from '../interfaces/Subject';

test.describe('Subject Creation Automation', () => {
  let excelPath: string;

  test.beforeAll(async () => {
    logger.info('Starting Enterprise Automation Suite...');
    if (!process.env.EXCEL_PATH) {
      throw new Error('EXCEL_PATH environment variable is required.');
    }
    excelPath = process.env.EXCEL_PATH;
  });

  test('Automate Subject Creation from Excel', async ({ page }) => {
    const isDryRun = process.env.DRY_RUN === 'true';
    const isResume = process.env.RESUME === 'true';
    const sheetNames = process.env.SHEET_NAMES ? process.env.SHEET_NAMES.split(',') : undefined;
    const runMode: AutomationRunMode = isDryRun ? 'DRY_RUN' : (process.env.HEADLESS === 'true' ? 'HEADLESS' : 'HEADED');

    const eventBus = TelemetryEventBus.getInstance();
    const dataReader = new ExcelDataReader();
    const validator = new SubjectValidator();
    const checkpointStore = new FileCheckpointStore();
    const reportGenerator = new ExcelReportGenerator();
    const driver = new PlaywrightSubjectDriver(page);

    // Wire terminal UI listeners via EventBus (Zero Coupling)
    eventBus.on<ProgressUpdatedEvent>(AutomationEventType.PROGRESS_UPDATED, (evt) => {
      ProgressBar.update(evt.processed, evt.total, evt.currentItem || '', evt.processed === evt.total ? 'DONE' : 'RUNNING');
    });

    eventBus.on<RecordCompletedEvent<Subject>>(AutomationEventType.RECORD_COMPLETED, (evt) => {
      const { result, currentIndex, totalRecords } = evt;
      const statusLabel = result.status;
      if (statusLabel === 'SUCCESS') {
        logger.info(`Row ${result.entity.rowNumber} (${currentIndex}/${totalRecords}): SUCCESS - ${result.entity.subjectName}`);
      } else if (statusLabel === 'DUPLICATE') {
        logger.warn(`Row ${result.entity.rowNumber} (${currentIndex}/${totalRecords}): DUPLICATE - ${result.entity.subjectName}`);
      } else if (statusLabel === 'SKIPPED') {
        logger.warn(`Row ${result.entity.rowNumber} (${currentIndex}/${totalRecords}): SKIPPED - ${result.message}`);
      } else {
        logger.error(`Row ${result.entity.rowNumber} (${currentIndex}/${totalRecords}): FAILED - ${result.message}`);
      }
    });

    // Instantiate clean orchestrator with Dependency Injection
    const orchestrator = new AutomationOrchestrator<Subject>({
      dataReader,
      validator,
      driver,
      checkpointStore,
      reportGenerator,
      telemetryEmitter: eventBus,
      maxRetries: 2,
      retryDelayMs: 1000
    });

    const summary = await orchestrator.execute({
      sourcePath: excelPath,
      mode: runMode,
      resume: isResume,
      readerOptions: { sheetNames },
      getItemDisplayName: (subject) => subject.subjectName
    });

    checkpointStore.markLiveStatusComplete('COMPLETED');
    logger.info(`Automation Run Finished with Status: ${summary.status}`);
  });
});
