import * as fs from 'fs';
import * as xlsx from 'xlsx';
import { logger } from './logger';
import { Subject } from '../interfaces/Subject';

export class ExcelReader {
  public static getSheetNames(filePath: string): string[] {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File does not exist: ${filePath}`);
    }
    const workbook = xlsx.readFile(filePath);
    return workbook.SheetNames || [];
  }

  public static readSubjects(filePath: string, targetSheets?: string[] | string): Subject[] {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File does not exist: ${filePath}`);
    }
    
    if (!filePath.toLowerCase().endsWith('.xlsx')) {
      throw new Error(`Invalid file extension. Expected .xlsx, got: ${filePath}`);
    }

    try {
      const workbook = xlsx.readFile(filePath);
      let sheetsToRead: string[] = [];

      if (!targetSheets || targetSheets === 'ALL' || (Array.isArray(targetSheets) && targetSheets.length === 0)) {
        sheetsToRead = workbook.SheetNames;
      } else if (typeof targetSheets === 'string') {
        sheetsToRead = [targetSheets];
      } else if (Array.isArray(targetSheets)) {
        sheetsToRead = targetSheets;
      }

      const subjects: Subject[] = [];
      let globalRowCounter = 2;

      for (const sheetName of sheetsToRead) {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;

        const rawData = xlsx.utils.sheet_to_json<any>(worksheet, { defval: '' });

        for (const row of rawData) {
          // Ignore completely blank rows
          const board = String(row['Education Board'] || row['Board'] || '').trim();
          const grade = String(row['Grade'] || row['Class'] || '').trim();
          const subjectName = String(row['Subject Name'] || row['Subject'] || '').trim();
          const shortName = String(row['Short Name'] || row['Short name'] || row['short_name'] || row['ShortName'] || '').trim();
          const subjectCode = String(row['Subject Code'] || row['Code'] || '').trim();
          const subjectDescription = String(row['Subject Description'] || row['Description'] || '').trim();

          if (!board && !grade && !subjectName && !shortName && !subjectCode && !subjectDescription) {
            globalRowCounter++;
            continue;
          }

          subjects.push({
            board,
            grade,
            subjectName,
            shortName,
            subjectCode,
            subjectDescription,
            rowNumber: globalRowCounter
          });

          globalRowCounter++;
        }
      }

      logger.info(`Excel Loaded Successfully (${sheetsToRead.join(', ')})`);
      logger.info(`Total Records : ${subjects.length}`);
      
      return subjects;
    } catch (error) {
      logger.error(`Failed to read excel workbook: ${(error as Error).message}`);
      throw error;
    }
  }
}
