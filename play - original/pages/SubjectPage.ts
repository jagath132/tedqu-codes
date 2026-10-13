import { Page } from '@playwright/test';
import { SELECTORS } from '../config/selectors';
import { Subject } from '../interfaces/Subject';
import { logger } from '../utils/logger';

export class SubjectPage {
  private page: Page;

  constructor(page: Page) {
    this.page = page;
  }

  public async createSubject(subject: Subject): Promise<'SUCCESS' | 'DUPLICATE' | 'FAILED'> {
    try {
      await this.openSubjectForm();
      await this.selectBoard(subject.board);
      await this.selectGrade(subject.grade);
      await this.enterSubject(subject.subjectName);
      if (subject.shortName) await this.enterShortName(subject.shortName);
      if (subject.subjectCode) await this.enterCode(subject.subjectCode);
      if (subject.subjectDescription) await this.enterDescription(subject.subjectDescription);
      
      await this.save();

      const isDuplicate = await this.checkDuplicate();
      if (isDuplicate) {
        await this.closePopupIfAny();
        return 'DUPLICATE';
      }

      const isSuccess = await this.checkSuccess();
      if (isSuccess) {
        return 'SUCCESS';
      }

      return 'FAILED';
    } catch (error) {
      logger.error(`Error while creating subject '${subject.subjectName}': ${(error as Error).message}`);
      await this.closePopupIfAny();
      return 'FAILED';
    }
  }

  public async openSubjectForm(): Promise<void> {
    const isFormOpen = await this.page.isVisible(SELECTORS.SUBJECT.BOARD_DROPDOWN);
    if (!isFormOpen) {
      await this.page.waitForSelector(SELECTORS.SUBJECT.ADD_NEW_BUTTON, { state: 'visible' });
      await this.page.click(SELECTORS.SUBJECT.ADD_NEW_BUTTON);
      await this.page.waitForSelector(SELECTORS.SUBJECT.BOARD_DROPDOWN, { state: 'visible' });
    }
  }

  public async selectBoard(board: string): Promise<void> {
    await this.page.selectOption(SELECTORS.SUBJECT.BOARD_DROPDOWN, { label: board });
  }

  public async selectGrade(grade: string): Promise<void> {
    await this.page.selectOption(SELECTORS.SUBJECT.GRADE_DROPDOWN, { label: grade });
  }

  public async enterSubject(name: string): Promise<void> {
    await this.page.fill(SELECTORS.SUBJECT.NAME_INPUT, name);
  }

  public async enterShortName(shortName: string): Promise<void> {
    await this.page.fill(SELECTORS.SUBJECT.SHORT_NAME_INPUT, shortName);
  }

  public async enterCode(code: string): Promise<void> {
    await this.page.fill(SELECTORS.SUBJECT.CODE_INPUT, code);
  }

  public async enterDescription(description: string): Promise<void> {
    await this.page.fill(SELECTORS.SUBJECT.DESCRIPTION_INPUT, description);
  }

  public async save(): Promise<void> {
    await this.page.click(SELECTORS.SUBJECT.SAVE_AND_ADD_ANOTHER_BUTTON);
  }

  public async checkDuplicate(): Promise<boolean> {
    try {
      // Fast check for duplicate message
      const duplicateMsg = this.page.locator(SELECTORS.SUBJECT.DUPLICATE_MESSAGE);
      await duplicateMsg.waitFor({ state: 'visible', timeout: 3000 });
      return true;
    } catch (e) {
      return false;
    }
  }

  public async checkSuccess(): Promise<boolean> {
    try {
      // Wait for success toast or redirection
      const successMsg = this.page.locator(SELECTORS.SUBJECT.SUCCESS_MESSAGE);
      await successMsg.waitFor({ state: 'visible', timeout: 5000 });
      return true;
    } catch (e) {
      return false;
    }
  }

  private async closePopupIfAny(): Promise<void> {
    try {
      const closeBtn = this.page.locator(SELECTORS.SUBJECT.CLOSE_POPUP);
      if (await closeBtn.isVisible()) {
        await closeBtn.click();
      }
    } catch (e) {
      // Ignore if no popup
    }
  }
}
