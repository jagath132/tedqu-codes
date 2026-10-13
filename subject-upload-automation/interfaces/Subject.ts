export interface Subject {
  board: string;
  grade: string;
  subjectName: string;
  shortName?: string;
  subjectCode: string;
  subjectDescription: string;
  rowNumber: number;
}

export interface SubjectResult extends Subject {
  status: 'SUCCESS' | 'DUPLICATE' | 'FAILED' | 'SKIPPED';
  message: string;
  timestamp: string;
  executionTime: number;
}
