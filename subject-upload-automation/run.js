const readline = require('readline');
const { spawn, exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

function ask(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(question, (ans) => { rl.close(); resolve(ans.trim()); }));
}

function askHidden(query) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    const stdin = process.openStdin();
    process.stdin.on('data', char => {
      char = char + '';
      switch (char) {
        case '\n':
        case '\r':
        case '\u0004':
          stdin.pause();
          break;
        default:
          break;
      }
    });
    rl.question(query, (value) => {
      rl.close();
      resolve(value.trim());
    });
  });
}

function startDashboardBackground() {
  const serverScript = path.join(__dirname, 'dashboard', 'server.js');
  const serverProcess = spawn('node', [serverScript], {
    detached: true,
    stdio: 'ignore'
  });
  serverProcess.unref();

  setTimeout(() => {
    if (process.platform === 'win32') {
      exec('start http://localhost:4000');
    } else if (process.platform === 'darwin') {
      exec('open http://localhost:4000');
    } else {
      exec('xdg-open http://localhost:4000');
    }
  }, 1000);
}

async function main() {
  console.log('\n======================================================');
  console.log('       ⚡ SUBJECT CREATION AUTOMATION SYSTEM ⚡      ');
  console.log('======================================================\n');

  // Step 1: Mode Selection
  console.log('Select Run Mode:');
  console.log('  1. Headless (background, no browser window)');
  console.log('  2. Headed   (browser window visible)');
  console.log('  3. Dry Run  (validate Excel only — no browser)');
  console.log('  4. Launch Live Web Dashboard Only (localhost:4000)\n');

  let modeAnswer;
  while (true) {
    modeAnswer = await ask('> Enter choice (1-4): ');
    if (['1', '2', '3', '4'].includes(modeAnswer)) break;
    console.log('  Invalid choice. Please enter 1, 2, 3, or 4.');
  }

  if (modeAnswer === '4') {
    startDashboardBackground();
    console.log('\n✓ Live Web Dashboard launched at: http://localhost:4000\n');
    return;
  }

  const isHeadless = modeAnswer === '1';
  const isDryRun = modeAnswer === '3';
  const modeLabel = isDryRun ? 'Dry Run (Validation Only)' : isHeadless ? 'Headless' : 'Headed';
  console.log(`\n  Selected Mode: ${modeLabel} ✓\n`);

  // Step 2: Excel File Path (with drag & drop sanitization)
  let excelPath = '';
  while (true) {
    const rawPath = await ask('Enter or Drag & Drop Excel File Path:\n> ');
    // Strip PowerShell invocation operator (& ), wrapping quotes (single/double), and whitespace
    excelPath = rawPath
      .trim()
      .replace(/^&\s*/, '') // Remove PowerShell '& ' prefix
      .replace(/^['"]+|['"]+$/g, '') // Remove wrapping single/double quotes
      .trim()
      .replace(/^['"]+|['"]+$/g, '') // Second pass for nested quotes
      .trim();

    if (!excelPath) {
      console.log('  File path cannot be empty.');
      continue;
    }

    if (!fs.existsSync(excelPath)) {
      console.log(`  ✗ File not found at: ${excelPath}. Please check path and try again.`);
      continue;
    }

    if (!excelPath.toLowerCase().endsWith('.xlsx')) {
      console.log('  ✗ Invalid file format. Expected a .xlsx Excel file.');
      continue;
    }

    break;
  }
  console.log(`  ✓ Loaded Excel: ${path.basename(excelPath)}\n`);

  // Step 3: Multi-Sheet Selection
  let targetSheets = '';
  try {
    const workbook = xlsx.readFile(excelPath);
    const sheetNames = workbook.SheetNames || [];
    if (sheetNames.length > 1) {
      console.log(`Found ${sheetNames.length} sheets in workbook:`);
      sheetNames.forEach((s, idx) => console.log(`  [${idx + 1}] ${s}`));
      console.log(`  [A] Process All Sheets\n`);

      const sheetChoice = await ask(`> Select sheet number (1-${sheetNames.length}) or 'A' for all [Default: A]: `);
      if (sheetChoice && sheetChoice.toUpperCase() !== 'A') {
        const selectedIdx = parseInt(sheetChoice, 10) - 1;
        if (selectedIdx >= 0 && selectedIdx < sheetNames.length) {
          targetSheets = sheetNames[selectedIdx];
          console.log(`  ✓ Processing sheet: ${targetSheets}\n`);
        }
      } else {
        targetSheets = 'ALL';
        console.log(`  ✓ Processing all sheets.\n`);
      }
    }
  } catch (err) {
    // Ignore sheet read warnings
  }

  // Step 4: Resume checkpoint detection
  let isResume = 'false';
  const checkpointPath = path.join(__dirname, 'data', 'checkpoint.json');
  if (!isDryRun && fs.existsSync(checkpointPath)) {
    try {
      const cp = JSON.parse(fs.readFileSync(checkpointPath, 'utf-8'));
      if (cp.completedRows && cp.completedRows.length > 0) {
        console.log(`\n⚠ Detected interrupted previous session: ${cp.completedRows.length}/${cp.totalRecords} records completed.`);
        const resumeAns = await ask('> Resume previous session? (y/n) [Default: y]: ');
        if (resumeAns.toLowerCase() !== 'n') {
          isResume = 'true';
          console.log('  ✓ Resuming previous run.');
        } else {
          try { fs.unlinkSync(checkpointPath); } catch(e) {}
          console.log('  ✓ Starting fresh run.');
        }
      }
    } catch (e) {}
  }

  // Step 5: Runtime credentials prompt
  let runtimeUser = '';
  let runtimePass = '';
  if (!isDryRun) {
    const credAns = await ask('\nUse saved credentials from .env? (y/n) [Default: y]: ');
    if (credAns.toLowerCase() === 'n') {
      runtimeUser = await ask('Enter Username: ');
      runtimePass = await ask('Enter Password: ');
      console.log('  ✓ Using custom session credentials.');
    }
  }

  // Step 6: Ask to open Live Dashboard
  const dashAns = await ask('\nOpen 24/7 Live Web Dashboard during run? (y/n) [Default: y]: ');
  if (dashAns.toLowerCase() !== 'n') {
    startDashboardBackground();
  }

  console.log('\n------------------------------------------------------');
  console.log('🚀 Starting Automation Process...');
  console.log('------------------------------------------------------\n');

  const extraArgs = process.argv.slice(2).filter(a => a !== '--headed');
  const command = `npx playwright test ${extraArgs.join(' ')}`;

  const envVars = {
    ...process.env,
    EXCEL_PATH: excelPath,
    HEADLESS: isHeadless ? 'true' : 'false',
    DRY_RUN: isDryRun ? 'true' : 'false',
    RESUME: isResume,
    SHEET_NAMES: targetSheets
  };

  if (runtimeUser && runtimePass) {
    envVars.RUNTIME_USERNAME = runtimeUser;
    envVars.RUNTIME_PASSWORD = runtimePass;
  }

  const playwrightProcess = spawn(command, {
    env: envVars,
    stdio: 'inherit',
    shell: true
  });

  playwrightProcess.on('exit', (code) => {
    process.exit(code);
  });
}

main();
