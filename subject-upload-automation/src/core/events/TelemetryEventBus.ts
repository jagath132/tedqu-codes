import { EventEmitter } from 'events';
import { ITelemetryEmitter, EventListener } from '../contracts/ITelemetryEmitter';
import { AutomationEvent, AutomationEventType } from './AutomationEvents';

export class TelemetryEventBus implements ITelemetryEmitter {
  private static instance: TelemetryEventBus;
  private emitter: EventEmitter;

  constructor() {
    this.emitter = new EventEmitter();
    this.emitter.setMaxListeners(50);
  }

  public static getInstance(): TelemetryEventBus {
    if (!TelemetryEventBus.instance) {
      TelemetryEventBus.instance = new TelemetryEventBus();
    }
    return TelemetryEventBus.instance;
  }

  public emit<E extends AutomationEvent>(event: E): void {
    this.emitter.emit(event.type, event);
    this.emitter.emit('*', event);
  }

  public on<E extends AutomationEvent>(eventType: AutomationEventType | string, listener: EventListener<E>): void {
    this.emitter.on(eventType, listener as (...args: any[]) => void);
  }

  public off<E extends AutomationEvent>(eventType: AutomationEventType | string, listener: EventListener<E>): void {
    this.emitter.off(eventType, listener as (...args: any[]) => void);
  }

  public clearAllListeners(): void {
    this.emitter.removeAllListeners();
  }
}
