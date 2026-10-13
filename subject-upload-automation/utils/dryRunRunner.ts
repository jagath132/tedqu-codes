import { Subject, SubjectResult } from '../interfaces/Subject';
import { Validator } from './validator';
import { ProgressBar } from './progressBar';
import { ReportGenerator } from './reportGenerator';
import { CheckpointManager } from './checkpointManager';

export class DryRunRunner {
  public static async execute(subjects: Subject[], excelPath: string): Promise<void> {
    console.log('\n=============================================');
    console.log('   🔍 DRY RUN MODE (VALIDATION ONLY - NO BROWSER)   ');
    console.log('=============================================\n');

    CheckpointManager.initLiveStatus(subjects.length, excelPath, 'DRY_RUN');

    const results: SubjectResult[] = [];

    for (let i = 0; i < subjects.length; i++) {
      const subject = subjects[i];
      const startTime = Date.now();
      const isValid = Validator.validateSubject(subject);

      const status = isValid ? 'SUCCESS' : 'SKIPPED';
      const message = isValid 
        ? 'Data Validated: Ready for creation' 
        : 'Data Incomplete: Missing required fields (Board, Grade, or Subject Name)';

      const executionTime = Date.now() - startTime;
      const res: SubjectResult = {
        ...subject,
        status: status as any,
        message,
        timestamp: new Date().toISOString(),
        executionTime
      };

      results.push(res);
      ProgressBar.update(i + 1, subjects.length, subject.subjectName, status);
      CheckpointManager.updateLiveStatus(res, subjects.length);
    }

    CheckpointManager.finishLiveStatus('COMPLETED');
    console.log('\n');
    await ReportGenerator.generateReport(results, 'DRY_RUN');
  }
}
