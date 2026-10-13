import { Subject } from '../interfaces/Subject';
import { logger } from './logger';

export class Validator {
  public static validateSubject(subject: Subject): boolean {
    if (!subject.board) {
      logger.error(`Row ${subject.rowNumber}: Skipped - Board is empty`);
      return false;
    }
    
    if (!subject.grade) {
      logger.error(`Row ${subject.rowNumber}: Skipped - Grade is empty`);
      return false;
    }
    
    if (!subject.subjectName) {
      logger.error(`Row ${subject.rowNumber}: Skipped - Subject Name is empty`);
      return false;
    }

    if (!subject.subjectCode) {
      logger.warn(`Row ${subject.rowNumber}: Warning - Subject Code is empty`);
    }

    if (!subject.subjectDescription) {
      logger.warn(`Row ${subject.rowNumber}: Warning - Subject Description is empty`);
    }

    return true;
  }
}
