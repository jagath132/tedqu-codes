import * as fs from 'fs';
import * as path from 'path';
import { ICheckpointStore, CheckpointPayload } from '../core/contracts/ICheckpointStore';
import { Subject } from '../../interfaces/Subject';

export class FileCheckpointStore implements ICheckpointStore<Subject> {
  private checkpointPath: string;
  private liveStatusPath: string;

  constructor(customDataDir?: string) {
    const dataDir = customDataDir || path.join(process.cwd(), 'data');
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }
    this.checkpointPath = path.join(dataDir, 'checkpoint.json');
    this.liveStatusPath = path.join(dataDir, 'live_status.json');
  }

  public saveCheckpoint(payload: CheckpointPayload<Subject>): void {
    try {
      fs.writeFileSync(this.checkpointPath, JSON.stringify(payload, null, 2), 'utf-8');

      // Mirror to live_status.json for legacy dashboard / external monitoring consumers
      const successCount = payload.results.filter(r => r.status === 'SUCCESS').length;
      const duplicateCount = payload.results.filter(r => r.status === 'DUPLICATE').length;
      const failedCount = payload.results.filter(r => r.status === 'FAILED').length;
      const skippedCount = payload.results.filter(r => r.status === 'SKIPPED').length;

      const legacyLiveStatus = {
        status: 'RUNNING',
        excelPath: payload.sourcePath,
        total: payload.totalRecords,
        processed: payload.results.length,
        successCount,
        duplicateCount,
        failedCount,
        skippedCount,
        lastUpdated: payload.lastUpdated,
        results: payload.results.map(r => ({
          ...r.entity,
          status: r.status,
          message: r.message,
          timestamp: r.timestamp,
          executionTime: r.executionTimeMs
        }))
      };

      fs.writeFileSync(this.liveStatusPath, JSON.stringify(legacyLiveStatus, null, 2), 'utf-8');
    } catch (err) {
      // Non-critical file write lock
    }
  }

  public loadCheckpoint(): CheckpointPayload<Subject> | null {
    if (!fs.existsSync(this.checkpointPath)) {
      return null;
    }
    try {
      const content = fs.readFileSync(this.checkpointPath, 'utf-8');
      return JSON.parse(content) as CheckpointPayload<Subject>;
    } catch {
      return null;
    }
  }

  public clearCheckpoint(): void {
    if (fs.existsSync(this.checkpointPath)) {
      try {
        fs.unlinkSync(this.checkpointPath);
      } catch {
        // Ignore
      }
    }
  }

  public markLiveStatusComplete(status: 'COMPLETED' | 'FAILED' | 'ABORTED' = 'COMPLETED'): void {
    try {
      if (fs.existsSync(this.liveStatusPath)) {
        const raw = JSON.parse(fs.readFileSync(this.liveStatusPath, 'utf-8'));
        raw.status = status;
        raw.completedAt = new Date().toISOString();
        fs.writeFileSync(this.liveStatusPath, JSON.stringify(raw, null, 2), 'utf-8');
      }
    } catch {
      // Ignore
    }
  }
}
