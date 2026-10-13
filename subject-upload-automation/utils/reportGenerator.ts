import * as ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { SubjectResult } from '../interfaces/Subject';
import { logger } from './logger';

export class ReportGenerator {
  public static async generateReport(results: SubjectResult[], prefix: string = 'RUN'): Promise<string> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Subject Creation Report');

    // Add Headers
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
      fgColor: { argb: 'FF1F2937' } // Dark charcoal
    };

    // Add Data
    for (const result of results) {
      const row = worksheet.addRow(result);
      
      // Color code based on status
      const statusCell = row.getCell('status');
      if (result.status === 'SUCCESS') {
        statusCell.font = { color: { argb: 'FF16A34A' }, bold: true }; // Green
      } else if (result.status === 'FAILED') {
        statusCell.font = { color: { argb: 'FFDC2626' }, bold: true }; // Red
      } else if (result.status === 'DUPLICATE') {
        statusCell.font = { color: { argb: 'FFD97706' }, bold: true }; // Amber/Orange
      } else {
        statusCell.font = { color: { argb: 'FF6B7280' } }; // Gray
      }
    }

    const reportsDir = path.join(process.cwd(), 'reports');
    if (!fs.existsSync(reportsDir)) {
      fs.mkdirSync(reportsDir, { recursive: true });
    }

    // Format timestamp string YYYY-MM-DD_HH-mm-ss
    const now = new Date();
    const pad = (n: number) => n.toString().padStart(2, '0');
    const timestampStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;
    
    const fileName = `SubjectCreation_Report_${prefix}_${timestampStr}.xlsx`;
    const reportPath = path.join(reportsDir, fileName);
    const latestPath = path.join(reportsDir, 'SubjectCreation_Report_Latest.xlsx');

    try {
      await workbook.xlsx.writeFile(reportPath);
      await workbook.xlsx.writeFile(latestPath);
      logger.info(`Report generated successfully at: ${reportPath}`);
      this.printSummary(results);

      // Auto open report file if enabled (default true)
      if (process.env.AUTO_OPEN_REPORT !== 'false') {
        this.openReportFile(reportPath);
      }

      return reportPath;
    } catch (error) {
      logger.error(`Failed to generate report: ${(error as Error).message}`);
      return '';
    }
  }

  private static openReportFile(filePath: string): void {
    try {
      if (process.platform === 'win32') {
        exec(`start "" "${filePath}"`, (err) => {
          if (err) logger.warn(`Could not auto-open report: ${err.message}`);
        });
      } else if (process.platform === 'darwin') {
        exec(`open "${filePath}"`);
      } else {
        exec(`xdg-open "${filePath}"`);
      }
    } catch (e) {
      // Ignore open errors
    }
  }

  private static printSummary(results: SubjectResult[]): void {
    const total = results.length;
    const success = results.filter(r => r.status === 'SUCCESS').length;
    const duplicate = results.filter(r => r.status === 'DUPLICATE').length;
    const failed = results.filter(r => r.status === 'FAILED').length;
    const skipped = results.filter(r => r.status === 'SKIPPED').length;
    const totalTime = results.reduce((acc, curr) => acc + (curr.executionTime || 0), 0);

    console.log('\n======================================================');
    console.log('            SUBJECT CREATION EXECUTION REPORT         ');
    console.log('======================================================');
    console.log(`  Total Records   : ${total}`);
    console.log(`  ✓ Success       : ${success}`);
    console.log(`  ⚠ Duplicate     : ${duplicate}`);
    console.log(`  ✗ Failed        : ${failed}`);
    console.log(`  ⊘ Skipped       : ${skipped}`);
    console.log(`  ⏱ Total Time    : ${(totalTime / 1000).toFixed(2)}s`);
    console.log('======================================================\n');
  }
}
