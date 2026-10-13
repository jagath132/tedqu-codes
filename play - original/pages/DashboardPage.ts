import { Page } from '@playwright/test';
import { SELECTORS } from '../config/selectors';
import { logger } from '../utils/logger';

export class DashboardPage {
  private page: Page;

  constructor(page: Page) {
    this.page = page;
  }

  public async navigateToSubjects(): Promise<void> {
    try {
      logger.info('Navigating to Subjects Section...');
      
      // If already on the subject creation page (e.g. after direct login redirect)
      if (this.page.url().includes('/school_module/subject/')) {
        logger.info('Already on Subjects page.');
        return;
      }

      // Otherwise wait for the menu to be visible and click it
      if (await this.page.isVisible(SELECTORS.DASHBOARD.SUBJECTS_MENU)) {
        await this.page.click(SELECTORS.DASHBOARD.SUBJECTS_MENU);
      }
      
      if (await this.page.isVisible(SELECTORS.DASHBOARD.SUBJECTS_LINK)) {
        await this.page.click(SELECTORS.DASHBOARD.SUBJECTS_LINK);
      }

      await this.page.waitForLoadState('networkidle');
      logger.info('Successfully navigated to Subjects page.');
    } catch (error) {
      logger.error(`Failed to navigate to Subjects page: ${(error as Error).message}`);
      throw error;
    }
  }
}
