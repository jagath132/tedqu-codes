import { defineConfig, devices } from '@playwright/test';
import * as dotenv from 'dotenv';

// Read from default ".env" file with override enabled.
dotenv.config(); // Do NOT use override:true — env vars from run.js must take priority over .env

export default defineConfig({
  testDir: './tests',
  /* Maximum time one test can run for. */
  timeout: 10 * 60 * 1000, // 10 minutes total for the entire automation flow
  expect: {
    timeout: 15000
  },
  /* Run tests in files in parallel */
  fullyParallel: false, // Subject creation needs to be sequential
  /* Fail the build on CI if you accidentally left test.only in the source code. */
  forbidOnly: !!process.env.CI,
  /* Retry on CI only */
  retries: process.env.CI ? 2 : 0,
  /* Opt out of parallel tests on CI. */
  workers: 1, // Ensure purely sequential execution
  /* Reporter to use. See https://playwright.dev/docs/test-reporters */
  reporter: 'html',
  /* Shared settings for all the projects below. See https://playwright.dev/docs/api/class-testoptions. */
  use: {
    /* Base URL to use in actions like `await page.goto('/')`. */
    baseURL: process.env.URL || 'http://localhost:3000',
    /* Collect trace when retrying the failed test. See https://playwright.dev/docs/trace-viewer */
    trace: 'on-first-retry',
    headless: process.env.HEADLESS === 'true',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    actionTimeout: 15000,
    navigationTimeout: 30000,
  },

  /* Configure projects for major browsers */
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
    // Adding firefox or webkit if necessary, but standard is chromium
  ],
});
