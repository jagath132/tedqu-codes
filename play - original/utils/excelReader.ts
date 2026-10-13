import * as fs from 'fs';
import * as xlsx from 'xlsx';
import { logger } from './logger';
import { Subject } from '../interfaces/Subject';

export class ExcelReader {
  public static readSubjects(filePath: string): Subject[] {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File does not exist: ${filePath}`);
    }
    
    if (!filePath.toLowerCase().endsWith('.xlsx')) {
      throw new Error(`Invalid file extension. Expected .xlsx, got: ${filePath}`);
    }

    try {
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      const rawData = xlsx.utils.sheet_to_json<any>(worksheet, { defval: '' });
      const subjects: Subject[] = [];

      // Row numbers in excel typically start from 2 (assuming row 1 is headers)
      let currentRow = 2; 

      for (const row of rawData) {
        // Ignore completely blank rows
        if (!row['Education Board'] && !row['Grade'] && !row['Subject Name'] && !row['Short Name'] && !row['Short name'] && !row['Subject Code'] && !row['Subject Description']) {
          currentRow++;
          continue;
        }

        subjects.push({
          board: String(row['Education Board'] || '').trim(),
          grade: String(row['Grade'] || '').trim(),
          subjectName: String(row['Subject Name'] || '').trim(),
          shortName: String(row['Short Name'] || row['Short name'] || row['short_name'] || row['ShortName'] || '').trim(),
          subjectCode: String(row['Subject Code'] || '').trim(),
          subjectDescription: String(row['Subject Description'] || '').trim(),
          rowNumber: currentRow
        });

        currentRow++;
      }

      logger.info('Excel Loaded Successfully');
      logger.info(`Total Records : ${subjects.length}`);
      
      return subjects;
    } catch (error) {
      logger.error(`Failed to read excel workbook: ${(error as Error).message}`);
      throw error;
    }
  }
}
