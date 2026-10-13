import { IValidator, ValidationResult } from '../core/contracts/IValidator';
import { Subject } from '../../interfaces/Subject';

export class SubjectValidator implements IValidator<Subject> {
  public validate(subject: Subject): ValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (!subject.board) {
      errors.push('Education Board is empty');
    }
    if (!subject.grade) {
      errors.push('Grade is empty');
    }
    if (!subject.subjectName) {
      errors.push('Subject Name is empty');
    }

    if (!subject.subjectCode) {
      warnings.push('Subject Code is missing');
    }
    if (!subject.subjectDescription) {
      warnings.push('Subject Description is missing');
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }
}
