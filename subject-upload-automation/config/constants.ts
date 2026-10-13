export const CONSTANTS = {
  REQUIRED_COLUMNS: [
    'Education Board',
    'Grade',
    'Subject Name',
    'Subject Code',
    'Subject Description'
  ],
  STATUS: {
    SUCCESS: 'SUCCESS',
    DUPLICATE: 'DUPLICATE',
    FAILED: 'FAILED',
    SKIPPED: 'SKIPPED'
  } as const,
  TIMEOUTS: {
    SHORT: 5000,
    MEDIUM: 15000,
    LONG: 30000
  }
};
