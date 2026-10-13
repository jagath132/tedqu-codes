import { BaseEntity } from '../models/ExecutionModels';

export interface ValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

export interface IValidator<T extends BaseEntity> {
  validate(entity: T): ValidationResult;
}
