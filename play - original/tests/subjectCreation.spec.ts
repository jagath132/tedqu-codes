import { test } from '@playwright/test';
import { ExcelReader } from '../utils/excelReader';
import { Validator } from '../utils/validator';
import { logger } from '../utils/logger';
import { ReportGenerator } from '../utils/reportGenerator';
import { ScreenshotUtil } from '../utils/screenshot';

import { LoginPage } from '../pages/LoginPage';
import { DashboardPage } from '../pages/DashboardPage';
import { SubjectPage } from '../pages/SubjectPage';
import { SubjectResult } from '../interfaces/Subject';
import { CONSTANTS } from '../config/constants';

test.describe('Subject Creation Automation', () => {
  let excelPath: string;

  test.beforeAll(async () => {
    logger.info('Starting Automation...');
    if (!process.env.EXCEL_PATH) {
      throw new Error('EXCEL_PATH environment variable is required.');
    }
    excelPath = process.env.EXCEL_PATH;
  });

  test('Automate Subject Creation from Excel', async ({ page }) => {
    // 1. Read Excel
    const subjects = ExcelReader.readSubjects(excelPath);
    const results: SubjectResult[] = [];

    // 2. Initialize Page Objects
    const loginPage = new LoginPage(page);
    const dashboardPage = new DashboardPage(page);
    const subjectPage = new SubjectPage(page);

    // 3. Login and Navigate
    await loginPage.login();
    await dashboardPage.navigateToSubjects();

    // 4. Iterate over subjects dynamically
    for (const subject of subjects) {
      const startTime = Date.now();
      
      let status: 'SUCCESS' | 'DUPLICATE' | 'FAILED' | 'SKIPPED' = 'SKIPPED';
      let message = '';

      if (!Validator.validateSubject(subject)) {
        status = 'SKIPPED';
        message = 'Validation failed (Missing mandatory fields)';
      } else {
        try {
          logger.info(`Processing Row ${subject.rowNumber}: ${subject.subjectName}`);
          
          status = await subjectPage.createSubject(subject);
          
          if (status === 'SUCCESS') {
            message = 'Subject created successfully';
            logger.info(`Row ${subject.rowNumber}: SUCCESS`);
          } else if (status === 'DUPLICATE') {
            message = 'Subject already exists';
            logger.warn(`Row ${subject.rowNumber}: DUPLICATE`);
          } else {
            message = 'Failed to create subject';
            logger.error(`Row ${subject.rowNumber}: FAILED`);
            await ScreenshotUtil.captureFailed(page, subject.rowNumber, subject.subjectName);
          }
        } catch (error) {
          status = 'FAILED';
          message = `Error: ${(error as Error).message}`;
          logger.error(`Row ${subject.rowNumber}: FAILED - ${message}`);
          await ScreenshotUtil.captureFailed(page, subject.rowNumber, subject.subjectName);
        }
      }

      const executionTime = Date.now() - startTime;

      // 5. Record Result
      results.push({
        ...subject,
        status,
        message,
        timestamp: new Date().toISOString(),
        executionTime
      });
    }

    // 6. Generate Report
    await ReportGenerator.generateReport(results);
  });
});
