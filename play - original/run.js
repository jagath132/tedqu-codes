const readline = require('readline');
const { spawn } = require('child_process');

function ask(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => rl.question(question, (ans) => { rl.close(); resolve(ans.trim()); }));
}

async function main() {
  // Step 1: Choose run mode
  console.log('\nSelect Run Mode:');
  console.log('  1. Headless (background, no browser window)');
  console.log('  2. Headed   (browser window visible)\n');

  let modeAnswer;
  while (true) {
    modeAnswer = await ask('> Enter choice (1 or 2): ');
    if (modeAnswer === '1' || modeAnswer === '2') break;
    console.log('  Invalid choice. Please enter 1 or 2.');
  }

  const isHeadless = modeAnswer === '1';
  console.log(`  Mode: ${isHeadless ? 'Headless ✓' : 'Headed ✓'}\n`);

  // Step 2: Enter Excel file path
  const rawPath = await ask('Enter Excel File Path:\n> ');
  const excelPath = rawPath.replace(/^"|"$/g, '');

  if (!excelPath) {
    console.error('File path is required.');
    process.exit(1);
  }

  console.log(`\nStarting Playwright with EXCEL_PATH=${excelPath}`);

  const extraArgs = process.argv.slice(2).filter(a => a !== '--headed');
  const command = `npx playwright test ${extraArgs.join(' ')}`;

  const playwrightProcess = spawn(command, {
    env: {
      ...process.env,
      EXCEL_PATH: excelPath,
      HEADLESS: isHeadless ? 'true' : 'false'
    },
    stdio: 'inherit',
    shell: true
  });

  playwrightProcess.on('exit', (code) => {
    process.exit(code);
  });
}

main();

