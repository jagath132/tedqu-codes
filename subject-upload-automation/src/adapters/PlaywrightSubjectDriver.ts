import { Page } from '@playwright/test';
import { IAutomationDriver, DriverExecutionResult } from '../core/contracts/IAutomationDriver';
import { Subject } from '../../interfaces/Subject';
import { LoginPage } from '../../pages/LoginPage';
import { DashboardPage } from '../../pages/DashboardPage';
import { SubjectPage } from '../../pages/SubjectPage';
import { ScreenshotUtil } from '../../utils/screenshot';

export class PlaywrightSubjectDriver implements IAutomationDriver<Subject> {
  private page: Page;
  private loginPage: LoginPage;
  private dashboardPage: DashboardPage;
  private subjectPage: SubjectPage;

  constructor(page: Page) {
    this.page = page;
    this.loginPage = new LoginPage(page);
    this.dashboardPage = new DashboardPage(page);
    this.subjectPage = new SubjectPage(page);
  }

  public async initialize(): Promise<void> {
    await this.loginPage.login();
    await this.dashboardPage.navigateToSubjects();
  }

  public async processEntity(subject: Subject): Promise<DriverExecutionResult> {
    const outcome = await this.subjectPage.createSubject(subject);
    if (outcome === 'SUCCESS') {
      return { status: 'SUCCESS', message: 'Subject created successfully' };
    } else if (outcome === 'DUPLICATE') {
      return { status: 'DUPLICATE', message: 'Subject already exists' };
    } else {
      return { status: 'FAILED', message: 'Failed to create subject' };
    }
  }

  public async onFailure(subject: Subject, error: Error): Promise<void> {
    try {
      await ScreenshotUtil.captureFailed(this.page, subject.rowNumber, subject.subjectName);
    } catch {
      // Ignore screenshot errors
    }
  }
}
