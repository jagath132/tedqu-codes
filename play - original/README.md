# Subject Creation Automation

Industry standard Playwright automation framework for creating Subjects in ERP Admin Panel.

## Installation

```bash
npm install
```

## Dependencies
- Playwright
- ExcelJS
- xlsx
- Winston Logger
- dotenv

## Folder Structure
```
project/
├── tests/             # Playwright test specs
├── pages/             # Page Object Model classes
├── utils/             # Helper utilities (logger, excelReader, etc)
├── interfaces/        # TypeScript interfaces and types
├── config/            # Selectors and constants
├── reports/           # Generated excel reports
├── screenshots/       # Error screenshots
├── data/              # Input data files (optional)
```

## Execution
Run the automation suite by executing:

```bash
npm run test
```
The terminal will prompt you to enter the absolute path to your Excel file.
Example: `C:\Users\Name\Desktop\Subjects.xlsx`

## Sample Excel Format
The automation expects an Excel file with the following columns:
1. Education Board
2. Grade
3. Subject Name
4. Subject Code
5. Subject Description

## Troubleshooting
- Check the `automation.log` for detailed logs.
- If a row fails, check the `screenshots/` directory for the captured state.
- Generated reports are available in the `reports/` directory with detailed execution status.
