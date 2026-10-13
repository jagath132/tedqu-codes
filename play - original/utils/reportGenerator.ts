import * as ExcelJS from 'exceljs';
import * as path from 'path';
import { SubjectResult } from '../interfaces/Subject';
import { logger } from './logger';

export class ReportGenerator {
  public static async generateReport(results: SubjectResult[]): Promise<void> {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Subject Creation Report');

    // Add Headers
    worksheet.columns = [
      { header: 'Excel Row', key: 'rowNumber', width: 10 },
      { header: 'Board', key: 'board', width: 20 },
      { header: 'Grade', key: 'grade', width: 15 },
      { header: 'Subject Name', key: 'subjectName', width: 30 },
      { header: 'Short Name', key: 'shortName', width: 20 },
      { header: 'Subject Code', key: 'subjectCode', width: 15 },
      { header: 'Status', key: 'status', width: 15 },
      { header: 'Message', key: 'message', width: 40 },
      { header: 'Timestamp', key: 'timestamp', width: 25 },
      { header: 'Execution Time (ms)', key: 'executionTime', width: 20 }
    ];

    // Style Headers
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' }
    };

    // Add Data
    for (const result of results) {
      const row = worksheet.addRow(result);
      
      // Color code based on status
      const statusCell = row.getCell('status');
      if (result.status === 'SUCCESS') {
        statusCell.font = { color: { argb: 'FF008000' } }; // Green
      } else if (result.status === 'FAILED') {
        statusCell.font = { color: { argb: 'FFFF0000' } }; // Red
      } else if (result.status === 'DUPLICATE') {
        statusCell.font = { color: { argb: 'FFFFA500' } }; // Orange
      } else {
        statusCell.font = { color: { argb: 'FF808080' } }; // Gray
      }
    }

    // Save File
    const reportPath = path.join(process.cwd(), 'reports', 'SubjectCreation_Report.xlsx');
    try {
      await workbook.xlsx.writeFile(reportPath);
      logger.info(`Report generated successfully at: ${reportPath}`);
      this.printSummary(results);
    } catch (error) {
      logger.error(`Failed to generate report: ${(error as Error).message}`);
    }
  }

  private static printSummary(results: SubjectResult[]): void {
    const total = results.length;
    const success = results.filter(r => r.status === 'SUCCESS').length;
    const duplicate = results.filter(r => r.status === 'DUPLICATE').length;
    const failed = results.filter(r => r.status === 'FAILED').length;
    const skipped = results.filter(r => r.status === 'SKIPPED').length;
    const totalTime = results.reduce((acc, curr) => acc + curr.executionTime, 0);

    console.log('\n=====================================');
    console.log('       SUBJECT CREATION REPORT       ');
    console.log('=====================================');
    console.log(`Total Records   : ${total}`);
    console.log(`Success         : ${success}`);
    console.log(`Duplicate       : ${duplicate}`);
    console.log(`Failed          : ${failed}`);
    console.log(`Skipped         : ${skipped}`);
    console.log(`Execution Time  : ${(totalTime / 1000).toFixed(2)}s`);
    console.log('=====================================\n');
  }
}
