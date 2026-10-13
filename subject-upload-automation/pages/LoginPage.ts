import { Page } from '@playwright/test';
import { SELECTORS } from '../config/selectors';
import { logger } from '../utils/logger';

export class LoginPage {
  private page: Page;

  constructor(page: Page) {
    this.page = page;
  }

  public async login(): Promise<void> {
    const url = process.env.URL;
    const username = process.env.RUNTIME_USERNAME || process.env.LOGIN_USERNAME || process.env.USERNAME;
    const password = process.env.RUNTIME_PASSWORD || process.env.PASSWORD;

    if (!url || !username || !password) {
      throw new Error('Missing login credentials in .env file');
    }

    try {
      logger.info('Navigating to Login Page...');
      await this.page.goto(url);
      
      logger.info('Entering credentials...');
      await this.page.fill(SELECTORS.LOGIN.USERNAME_INPUT, username);
      await this.page.fill(SELECTORS.LOGIN.PASSWORD_INPUT, password);
      
      logger.info('Clicking Login button...');
      await this.page.click(SELECTORS.LOGIN.SUBMIT_BUTTON);
      
      await this.page.waitForLoadState('networkidle');
      logger.info('Login successful.');
    } catch (error) {
      logger.error(`Login failed: ${(error as Error).message}`);
      throw error;
    }
  }
}
