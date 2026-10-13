import { AutomationEvent, AutomationEventType } from '../events/AutomationEvents';

export type EventListener<E extends AutomationEvent = AutomationEvent> = (event: E) => void;

export interface ITelemetryEmitter {
  emit<E extends AutomationEvent>(event: E): void;
  on<E extends AutomationEvent>(eventType: AutomationEventType | string, listener: EventListener<E>): void;
  off<E extends AutomationEvent>(eventType: AutomationEventType | string, listener: EventListener<E>): void;
}
