import * as fs from 'fs';
import * as path from 'path';
import { SubjectResult } from '../interfaces/Subject';
import { logger } from './logger';

export interface CheckpointData {
  excelPath: string;
  totalRecords: number;
  completedRows: number[];
  results: SubjectResult[];
  lastUpdated: string;
}

export class CheckpointManager {
  private static checkpointPath = path.join(process.cwd(), 'data', 'checkpoint.json');
  private static liveStatusPath = path.join(process.cwd(), 'data', 'live_status.json');

  public static initLiveStatus(total: number, excelPath: string, mode: string): void {
    const dataDir = path.join(process.cwd(), 'data');
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }

    const initialStatus = {
      status: 'RUNNING',
      mode,
      excelPath,
      total,
      processed: 0,
      successCount: 0,
      duplicateCount: 0,
      failedCount: 0,
      skippedCount: 0,
      startTime: new Date().toISOString(),
      lastUpdated: new Date().toISOString(),
      results: [] as SubjectResult[]
    };

    fs.writeFileSync(this.liveStatusPath, JSON.stringify(initialStatus, null, 2), 'utf-8');
  }

  public static updateLiveStatus(result: SubjectResult, total: number): void {
    try {
      let currentData: any = {};
      if (fs.existsSync(this.liveStatusPath)) {
        currentData = JSON.parse(fs.readFileSync(this.liveStatusPath, 'utf-8'));
      }

      const results: SubjectResult[] = currentData.results || [];
      const existingIdx = results.findIndex(r => r.rowNumber === result.rowNumber);
      if (existingIdx >= 0) {
        results[existingIdx] = result;
      } else {
        results.push(result);
      }

      const successCount = results.filter(r => r.status === 'SUCCESS').length;
      const duplicateCount = results.filter(r => r.status === 'DUPLICATE').length;
      const failedCount = results.filter(r => r.status === 'FAILED').length;
      const skippedCount = results.filter(r => r.status === 'SKIPPED').length;

      const updated = {
        ...currentData,
        processed: results.length,
        total,
        successCount,
        duplicateCount,
        failedCount,
        skippedCount,
        lastUpdated: new Date().toISOString(),
        results
      };

      fs.writeFileSync(this.liveStatusPath, JSON.stringify(updated, null, 2), 'utf-8');
    } catch (e) {
      // Non-critical, ignore write lock collisions
    }
  }

  public static finishLiveStatus(status: 'COMPLETED' | 'STOPPED' | 'FAILED' = 'COMPLETED'): void {
    try {
      if (fs.existsSync(this.liveStatusPath)) {
        const currentData = JSON.parse(fs.readFileSync(this.liveStatusPath, 'utf-8'));
        currentData.status = status;
        currentData.endTime = new Date().toISOString();
        fs.writeFileSync(this.liveStatusPath, JSON.stringify(currentData, null, 2), 'utf-8');
      }
    } catch (e) {
      // Ignore
    }
  }

  public static saveCheckpoint(excelPath: string, total: number, results: SubjectResult[]): void {
    try {
      const dataDir = path.join(process.cwd(), 'data');
      if (!fs.existsSync(dataDir)) {
        fs.mkdirSync(dataDir, { recursive: true });
      }

      const checkpoint: CheckpointData = {
        excelPath,
        totalRecords: total,
        completedRows: results.map(r => r.rowNumber),
        results,
        lastUpdated: new Date().toISOString()
      };

      fs.writeFileSync(this.checkpointPath, JSON.stringify(checkpoint, null, 2), 'utf-8');
    } catch (err) {
      logger.warn(`Could not save checkpoint: ${(err as Error).message}`);
    }
  }

  public static loadCheckpoint(): CheckpointData | null {
    if (!fs.existsSync(this.checkpointPath)) {
      return null;
    }
    try {
      const content = fs.readFileSync(this.checkpointPath, 'utf-8');
      return JSON.parse(content) as CheckpointData;
    } catch (err) {
      return null;
    }
  }

  public static clearCheckpoint(): void {
    if (fs.existsSync(this.checkpointPath)) {
      try {
        fs.unlinkSync(this.checkpointPath);
      } catch (err) {
        // Ignore
      }
    }
  }
}
