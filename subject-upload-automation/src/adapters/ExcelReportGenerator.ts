import * as ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { IReportGenerator, ReportGenerationOptions } from '../core/contracts/IReportGenerator';
import { ExecutionResult } from '../core/models/ExecutionModels';
import { Subject } from '../../interfaces/Subject';

export class ExcelReportGenerator implements IReportGenerator<Subject> {
  private reportsDir: string;

  constructor(customReportsDir?: string) {
    this.reportsDir = customReportsDir || path.join(process.cwd(), 'reports');
    if (!fs.existsSync(this.reportsDir)) {
      fs.mkdirSync(this.reportsDir, { recursive: true });
    }
  }

  public async generate(results: ExecutionResult<Subject>[], options?: ReportGenerationOptions): Promise<string> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Subject Creation Report');

    worksheet.columns = [
      { header: 'Excel Row', key: 'rowNumber', width: 12 },
      { header: 'Board', key: 'board', width: 20 },
      { header: 'Grade', key: 'grade', width: 15 },
      { header: 'Subject Name', key: 'subjectName', width: 30 },
      { header: 'Short Name', key: 'shortName', width: 20 },
      { header: 'Subject Code', key: 'subjectCode', width: 15 },
      { header: 'Status', key: 'status', width: 15 },
      { header: 'Message', key: 'message', width: 45 },
      { header: 'Timestamp', key: 'timestamp', width: 25 },
      { header: 'Execution Time (ms)', key: 'executionTime', width: 20 }
    ];

    // Style Headers
    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1F2937' }
    };

    // Add Rows
    for (const r of results) {
      const rowData = {
        rowNumber: r.entity.rowNumber,
        board: r.entity.board,
        grade: r.entity.grade,
        subjectName: r.entity.subjectName,
        shortName: r.entity.shortName || '',
        subjectCode: r.entity.subjectCode || '',
        status: r.status,
        message: r.message,
        timestamp: r.timestamp,
        executionTime: r.executionTimeMs
      };

      const row = worksheet.addRow(rowData);
      const statusCell = row.getCell('status');

      if (r.status === 'SUCCESS') {
        statusCell.font = { color: { argb: 'FF16A34A' }, bold: true };
      } else if (r.status === 'FAILED') {
        statusCell.font = { color: { argb: 'FFDC2626' }, bold: true };
      } else if (r.status === 'DUPLICATE') {
        statusCell.font = { color: { argb: 'FFD97706' }, bold: true };
      } else {
        statusCell.font = { color: { argb: 'FF6B7280' } };
      }
    }

    const now = new Date();
    const pad = (n: number) => n.toString().padStart(2, '0');
    const timestampStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;
    const prefix = options?.prefix || 'RUN';
    const fileName = `SubjectCreation_Report_${prefix}_${timestampStr}.xlsx`;
    const reportPath = path.join(this.reportsDir, fileName);
    const latestPath = path.join(this.reportsDir, 'SubjectCreation_Report_Latest.xlsx');

    try {
      await workbook.xlsx.writeFile(reportPath);
      await workbook.xlsx.writeFile(latestPath);

      if (options?.autoOpen && process.env.AUTO_OPEN_REPORT !== 'false') {
        this.openReportFile(reportPath);
      }

      return reportPath;
    } catch {
      return '';
    }
  }

  private openReportFile(filePath: string): void {
    try {
      if (process.platform === 'win32') {
        exec(`start "" "${filePath}"`);
      } else if (process.platform === 'darwin') {
        exec(`open "${filePath}"`);
      } else {
        exec(`xdg-open "${filePath}"`);
      }
    } catch {
      // Ignore
    }
  }
}
