const express = require('express');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const app = express();
const PORT = process.env.PORT || 4000;

// Security & Parsing Middlewares
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Security Headers Middleware
app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('X-XSS-Protection', '1; mode=block');
  next();
});

// Serve frontend static assets from public/
app.use(express.static(path.join(__dirname, 'public'), {
  etag: true,
  maxAge: '1h'
}));

const rootDir = path.join(__dirname, '..');
const dataDir = path.join(rootDir, 'data');
const liveStatusPath = path.join(dataDir, 'live_status.json');
const checkpointPath = path.join(dataDir, 'checkpoint.json');
const reportsDir = path.join(rootDir, 'reports');
const logPath = path.join(rootDir, 'automation.log');

if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}
if (!fs.existsSync(reportsDir)) {
  fs.mkdirSync(reportsDir, { recursive: true });
}

// Active Playwright child process reference
let activeRunProcess = null;

// Helper: Read live status safely
function getLiveStatus() {
  if (!fs.existsSync(liveStatusPath)) {
    return {
      status: 'IDLE',
      message: 'System ready. No active execution.',
      total: 0,
      processed: 0,
      successCount: 0,
      duplicateCount: 0,
      failedCount: 0,
      skippedCount: 0,
      results: []
    };
  }
  try {
    return JSON.parse(fs.readFileSync(liveStatusPath, 'utf-8'));
  } catch (err) {
    return { status: 'IDLE', results: [], error: 'Status file read error' };
  }
}

// 1. GET /api/status - Get current telemetry
app.get('/api/status', (req, res) => {
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.json(getLiveStatus());
});

// 2. GET /api/events - Server-Sent Events (SSE) stream for real-time reactive UI
app.get('/api/events', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  if (res.flushHeaders) res.flushHeaders();

  const sendStatus = () => {
    try {
      const data = getLiveStatus();
      res.write(`data: ${JSON.stringify(data)}\n\n`);
    } catch {
      // Ignore
    }
  };

  sendStatus();
  const interval = setInterval(sendStatus, 1000);

  req.on('close', () => {
    clearInterval(interval);
  });
});

// 3. GET /api/checkpoint - Check if an interrupted session exists
app.get('/api/checkpoint', (req, res) => {
  if (!fs.existsSync(checkpointPath)) {
    return res.json({ hasCheckpoint: false });
  }
  try {
    const cp = JSON.parse(fs.readFileSync(checkpointPath, 'utf-8'));
    res.json({
      hasCheckpoint: true,
      totalRecords: cp.totalRecords || 0,
      completedRows: cp.completedRows ? cp.completedRows.length : 0,
      excelPath: cp.excelPath,
      lastUpdated: cp.lastUpdated
    });
  } catch {
    res.json({ hasCheckpoint: false });
  }
});

// 4. POST /api/checkpoint/clear - Discard existing checkpoint
app.post('/api/checkpoint/clear', (req, res) => {
  try {
    if (fs.existsSync(checkpointPath)) {
      fs.unlinkSync(checkpointPath);
    }
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// 5. GET /api/sheets - Inspect sheets in an Excel workbook
app.get('/api/sheets', (req, res) => {
  const targetPath = req.query.path;
  if (!targetPath || typeof targetPath !== 'string') {
    return res.status(400).json({ error: 'Missing path query parameter' });
  }

  const resolved = path.resolve(targetPath);
  if (!fs.existsSync(resolved)) {
    return res.status(404).json({ error: 'File not found on system' });
  }

  try {
    const xlsx = require('xlsx');
    const workbook = xlsx.readFile(resolved);
    res.json({ sheets: workbook.SheetNames || [] });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read workbook: ' + err.message });
  }
});

// 6. POST /api/run/start - Launch automation run from Dashboard UI
app.post('/api/run/start', (req, res) => {
  if (activeRunProcess) {
    return res.status(400).json({ error: 'An automation run is already in progress.' });
  }

  const { excelPath, mode, resume, sheetNames, username, password } = req.body;

  if (!excelPath) {
    return res.status(400).json({ error: 'Excel file path is required.' });
  }

  const resolvedExcel = path.resolve(excelPath);
  if (!fs.existsSync(resolvedExcel)) {
    return res.status(404).json({ error: `File not found at: ${resolvedExcel}` });
  }

  const isHeadless = mode !== 'headed';
  const isDryRun = mode === 'dry_run';

  const envVars = {
    ...process.env,
    EXCEL_PATH: resolvedExcel,
    HEADLESS: isHeadless ? 'true' : 'false',
    DRY_RUN: isDryRun ? 'true' : 'false',
    RESUME: resume ? 'true' : 'false',
    SHEET_NAMES: sheetNames || 'ALL'
  };

  if (username && password) {
    envVars.RUNTIME_USERNAME = username;
    envVars.RUNTIME_PASSWORD = password;
  }

  const cmd = 'npx';
  const args = ['playwright', 'test'];

  activeRunProcess = spawn(cmd, args, {
    cwd: rootDir,
    env: envVars,
    shell: true
  });

  activeRunProcess.stdout.on('data', (data) => {
    // Append to automation.log
    try {
      fs.appendFileSync(logPath, data.toString());
    } catch {}
  });

  activeRunProcess.stderr.on('data', (data) => {
    try {
      fs.appendFileSync(logPath, data.toString());
    } catch {}
  });

  activeRunProcess.on('exit', (code) => {
    activeRunProcess = null;
    try {
      const current = getLiveStatus();
      current.status = code === 0 ? 'COMPLETED' : 'STOPPED';
      current.exitCode = code;
      fs.writeFileSync(liveStatusPath, JSON.stringify(current, null, 2), 'utf-8');
    } catch {}
  });

  res.json({
    success: true,
    message: 'Automation started successfully',
    mode: isDryRun ? 'Dry Run' : isHeadless ? 'Headless' : 'Headed'
  });
});

// 7. POST /api/run/stop - Abort active run
app.post('/api/run/stop', (req, res) => {
  if (!activeRunProcess) {
    return res.status(400).json({ error: 'No active automation run to abort.' });
  }

  try {
    if (process.platform === 'win32') {
      spawn('taskkill', ['/pid', activeRunProcess.pid.toString(), '/f', '/t']);
    } else {
      activeRunProcess.kill('SIGKILL');
    }
    activeRunProcess = null;

    const current = getLiveStatus();
    current.status = 'STOPPED';
    current.message = 'Run aborted by user from dashboard.';
    fs.writeFileSync(liveStatusPath, JSON.stringify(current, null, 2), 'utf-8');

    res.json({ success: true, message: 'Automation run stopped successfully.' });
  } catch (err) {
    res.status(500).json({ error: 'Failed to stop process: ' + err.message });
  }
});

// 8. GET /api/reports - List generated Excel reports with secure metadata
app.get('/api/reports', (req, res) => {
  if (!fs.existsSync(reportsDir)) {
    return res.json([]);
  }
  try {
    const files = fs.readdirSync(reportsDir)
      .filter(f => f.endsWith('.xlsx'))
      .map(name => {
        const fullPath = path.join(reportsDir, name);
        const stats = fs.statSync(fullPath);
        return {
          fileName: name,
          sizeBytes: stats.size,
          sizeFormatted: (stats.size / 1024).toFixed(1) + ' KB',
          createdAt: stats.mtime
        };
      })
      .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    res.json(files);
  } catch (e) {
    res.status(500).json({ error: 'Failed to read reports directory' });
  }
});

// 9. GET /api/reports/download/:fileName - Secure download with strict path sanitization
app.get('/api/reports/download/:fileName', (req, res) => {
  const fileName = path.basename(req.params.fileName);
  const filePath = path.resolve(reportsDir, fileName);

  // Security: Prevent path traversal outside reports directory
  if (!filePath.startsWith(path.resolve(reportsDir))) {
    return res.status(403).json({ error: 'Access denied: Invalid file path' });
  }

  if (fs.existsSync(filePath)) {
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.download(filePath, fileName);
  } else {
    res.status(404).json({ error: 'Report file not found' });
  }
});

// 10. GET /api/logs - Stream recent logs safely
app.get('/api/logs', (req, res) => {
  if (!fs.existsSync(logPath)) {
    return res.json({ logs: 'No logs generated yet.' });
  }
  try {
    const logs = fs.readFileSync(logPath, 'utf-8');
    const lines = logs.split('\n').slice(-250).join('\n');
    res.json({ logs: lines });
  } catch (err) {
    res.status(500).json({ error: 'Could not read log file' });
  }
});

// 11. POST /api/logs/clear - Clear log file
app.post('/api/logs/clear', (req, res) => {
  try {
    fs.writeFileSync(logPath, '', 'utf-8');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'Could not clear log file' });
  }
});

// Fallback to index.html for SPA client-side routing
app.get('*', (req, res) => {
  const indexPath = path.join(__dirname, 'public', 'index.html');
  if (fs.existsSync(indexPath)) {
    res.sendFile(indexPath);
  } else {
    res.status(404).send('Dashboard UI not found.');
  }
});

// Start Server
app.listen(PORT, () => {
  console.log(`\n==================================================================`);
  console.log(`  🛡️  ENTERPRISE AUTOMATION OPERATIONS HUB: http://localhost:${PORT}`);
  console.log(`==================================================================\n`);
});
