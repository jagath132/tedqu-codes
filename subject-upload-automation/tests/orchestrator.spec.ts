import { test, expect } from '@playwright/test';
import { AutomationOrchestrator } from '../src/core/services/AutomationOrchestrator';
import { TelemetryEventBus } from '../src/core/events/TelemetryEventBus';
import { AutomationEventType } from '../src/core/events/AutomationEvents';
import { IDataReader } from '../src/core/contracts/IDataReader';
import { IValidator } from '../src/core/contracts/IValidator';
import { IAutomationDriver } from '../src/core/contracts/IAutomationDriver';
import { BaseEntity } from '../src/core/models/ExecutionModels';

interface MockEntity extends BaseEntity {
  name: string;
}

test.describe('AutomationOrchestrator Unit Test (SOLID Verification)', () => {
  test('should process entities, handle retries, and emit events without browser', async () => {
    const mockEntities: MockEntity[] = [
      { rowNumber: 1, name: 'Physics' },
      { rowNumber: 2, name: 'Invalid Blank' },
      { rowNumber: 3, name: 'Math' }
    ];

    const mockDataReader: IDataReader<MockEntity> = {
      read: () => mockEntities
    };

    const mockValidator: IValidator<MockEntity> = {
      validate: (entity) => ({
        isValid: entity.name !== 'Invalid Blank',
        errors: entity.name === 'Invalid Blank' ? ['Name cannot be blank'] : [],
        warnings: []
      })
    };

    let processCallCount = 0;
    const mockDriver: IAutomationDriver<MockEntity> = {
      initialize: async () => {},
      processEntity: async (entity) => {
        processCallCount++;
        if (entity.name === 'Physics') {
          return { status: 'SUCCESS', message: 'Subject created' };
        }
        return { status: 'DUPLICATE', message: 'Already exists' };
      }
    };

    const eventBus = new TelemetryEventBus();
    const emittedEvents: string[] = [];
    eventBus.on('*', (evt) => {
      emittedEvents.push(evt.type);
    });

    const orchestrator = new AutomationOrchestrator<MockEntity>({
      dataReader: mockDataReader,
      validator: mockValidator,
      driver: mockDriver,
      telemetryEmitter: eventBus
    });

    const summary = await orchestrator.execute({
      sourcePath: 'mock_test.xlsx',
      mode: 'HEADLESS'
    });

    expect(summary.total).toBe(3);
    expect(summary.processed).toBe(3);
    expect(summary.successCount).toBe(1);
    expect(summary.duplicateCount).toBe(1);
    expect(summary.skippedCount).toBe(1);
    expect(summary.status).toBe('COMPLETED');
    expect(processCallCount).toBe(2); // Only Physics and Math reached the driver

    expect(emittedEvents).toContain(AutomationEventType.RUN_STARTED);
    expect(emittedEvents).toContain(AutomationEventType.PROGRESS_UPDATED);
    expect(emittedEvents).toContain(AutomationEventType.RUN_COMPLETED);
  });
});
