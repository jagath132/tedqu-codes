import * as fs from 'fs';
import * as xlsx from 'xlsx';
import { IDataReader, DataReaderOptions } from '../core/contracts/IDataReader';
import { Subject } from '../../interfaces/Subject';

export class ExcelDataReader implements IDataReader<Subject> {
  public inspectSheets(filePath: string): string[] {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File does not exist: ${filePath}`);
    }
    const workbook = xlsx.readFile(filePath);
    return workbook.SheetNames || [];
  }

  public read(filePath: string, options?: DataReaderOptions): Subject[] {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File does not exist: ${filePath}`);
    }

    if (!filePath.toLowerCase().endsWith('.xlsx')) {
      throw new Error(`Invalid file extension. Expected .xlsx, got: ${filePath}`);
    }

    const workbook = xlsx.readFile(filePath);
    let sheetsToRead: string[] = [];

    const targetSheets = options?.sheetNames;
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
        const board = String(row['Education Board'] || row['Board'] || '').trim();
        const grade = String(row['Grade'] || row['Class'] || '').trim();
        const subjectName = String(row['Subject Name'] || row['Subject'] || '').trim();
        const shortName = String(row['Short Name'] || row['Short name'] || row['short_name'] || row['ShortName'] || '').trim();
        const subjectCode = String(row['Subject Code'] || row['Code'] || '').trim();
        const subjectDescription = String(row['Subject Description'] || row['Description'] || '').trim();

        // Skip blank lines
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

    return subjects;
  }
}
