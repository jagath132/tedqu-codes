import { Page } from '@playwright/test';
import * as path from 'path';

export class ScreenshotUtil {
  public static async captureFailed(page: Page, rowNumber: number, subjectName: string): Promise<void> {
    try {
      const sanitizedName = subjectName.replace(/[^a-z0-9]/gi, '_').toLowerCase();
      const filename = `Failed_Row_${rowNumber}_${sanitizedName}.png`;
      const filepath = path.join(process.cwd(), 'screenshots', filename);
      
      await page.screenshot({ path: filepath, fullPage: true });
    } catch (error) {
      console.error(`Failed to capture screenshot for row ${rowNumber}:`, error);
    }
  }
}
