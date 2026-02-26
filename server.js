require('dotenv').config();
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const { verifyToken, signToken } = require('./middleware/auth');
const execData = require('./services/execDataService');

const app = express();
const PORT = process.env.PORT || 3000;

// ── Multer config for Excel uploads ───────────────────────────────
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
const upload = multer({
  dest: uploadDir,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.xlsx', '.xlsb', '.xls'].includes(ext)) cb(null, true);
    else cb(new Error('Only Excel files (.xlsx, .xlsb, .xls) are allowed'));
  },
});

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Data source state ─────────────────────────────────────────────────
let connectedSource = { type: null, url: null };

// ── Path to the Excel file ────────────────────────────────────────────
const EXCEL_PATH =
  process.env.EXCEL_PATH ||
  path.join(__dirname, '..', 'pov', '_Bandwidth Tracker.xlsx');

const CACHED_SHEET_PATH = path.join(__dirname, '_cached_sheet.xlsx');

// ══════════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════════

/** Extract Google Sheets spreadsheet ID from various URL formats */
function extractSheetId(url) {
  const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

/** Download a Google Sheet as XLSX and save locally */
async function downloadGoogleSheet(url) {
  const sheetId = extractSheetId(url);
  if (!sheetId) throw new Error('Invalid Google Sheets URL — could not extract sheet ID');

  const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  console.log(`Fetching Google Sheet: ${exportUrl}`);

  const resp = await fetch(exportUrl);
  if (!resp.ok) throw new Error(`Google Sheets download failed: ${resp.status} ${resp.statusText}`);

  const buffer = Buffer.from(await resp.arrayBuffer());
  fs.writeFileSync(CACHED_SHEET_PATH, buffer);
  console.log(`Google Sheet downloaded & cached (${(buffer.length / 1024).toFixed(1)} KB)`);
  return XLSX.read(buffer, { type: 'buffer' });
}

/** Load workbook — from Google Sheets if connected, else local file */
async function loadWorkbook() {
  // If a Google Sheet is connected, download fresh copy
  if (connectedSource.type === 'google' && connectedSource.url) {
    return await downloadGoogleSheet(connectedSource.url);
  }

  // Fall back to local Excel file
  if (!fs.existsSync(EXCEL_PATH)) {
    throw new Error(`Excel file not found at: ${EXCEL_PATH}`);
  }
  return XLSX.readFile(EXCEL_PATH);
}

/** Synchronous fallback for startup / simple cases */
function loadWorkbookSync() {
  if (fs.existsSync(CACHED_SHEET_PATH) && connectedSource.type === 'google') {
    return XLSX.readFile(CACHED_SHEET_PATH);
  }
  if (!fs.existsSync(EXCEL_PATH)) {
    throw new Error(`Excel file not found at: ${EXCEL_PATH}`);
  }
  return XLSX.readFile(EXCEL_PATH);
}

// Auto-connect Google Sheet from env var
if (process.env.GOOGLE_SHEET_URL) {
  connectedSource = { type: 'google', url: process.env.GOOGLE_SHEET_URL };
  console.log('Auto-connected Google Sheet from env:', process.env.GOOGLE_SHEET_URL);
}

function cleanString(val) {
  if (val == null) return '';
  return val.toString().replace(/\u200b/g, '').trim();
}

function formatDate(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const str = val.toString().trim();
  // YYYY-MM-DD
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  // DD/MM/YYYY
  const m2 = str.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
  return str;
}

/** Build lookup maps from the Drop Down sheet */
function buildDropdownLookup(wb) {
  const ws = wb.Sheets['Drop Down'];
  if (!ws) return { detailsByProject: {}, managerByProject: {} };

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  // Detect columns by header name
  const hdr = rows[0] || [];
  const colMap = {};
  hdr.forEach((h, i) => { if (h) colMap[cleanString(h).toLowerCase()] = i; });

  const projCol = colMap['project name'] ?? 1;
  const detailIdx = 2; // Column C has project details text
  const pmCol = colMap['project manager'] ?? 5;

  const detailsByProject = {};
  const managerByProject = {};

  for (let i = 1; i < rows.length; i++) {
    const projectName = cleanString(rows[i][projCol]);
    if (!projectName) continue;
    const detail = cleanString(rows[i][detailIdx]);
    const manager = cleanString(rows[i][pmCol]);
    if (detail) detailsByProject[projectName] = detail;
    if (manager) managerByProject[projectName] = manager;
  }
  return { detailsByProject, managerByProject };
}

/** Detect column indexes by header names (no hardcoded indexes) */
function detectColumns(headerRow) {
  const map = {};
  (headerRow || []).forEach((h, i) => {
    const key = cleanString(h).toLowerCase();
    if (key) map[key] = i;
  });
  return map;
}

// ══════════════════════════════════════════════════════════════════════
//  API: Bandwidth Tracker
// ══════════════════════════════════════════════════════════════════════

app.get('/api/bandwidth', async (req, res) => {
  try {
    const wb = await loadWorkbook();
    const { detailsByProject, managerByProject } = buildDropdownLookup(wb);
    const ws = wb.Sheets['Bandwidth Tracker'];
    if (!ws) return res.status(404).json({ error: 'Sheet "Bandwidth Tracker" not found' });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cols = detectColumns(rows[0]);

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const date = formatDate(r[cols['date'] ?? 0]);
      const project = cleanString(r[cols['project'] ?? 1]);
      if (!date && !project) continue;

      // Resolve Project Details & PM via VLOOKUP if empty
      let projectDetails = cleanString(r[cols['project details'] ?? 2]);
      let projectManager = cleanString(r[cols['project manager'] ?? 3]);
      if (!projectDetails || projectDetails.startsWith('='))
        projectDetails = detailsByProject[project] || '';
      if (!projectManager || projectManager.startsWith('='))
        projectManager = managerByProject[project] || '';

      data.push({
        date,
        project,
        projectDetails,
        projectManager,
        name: cleanString(r[cols['name'] ?? 4]),
        role: cleanString(r[cols['role'] ?? 5]),
        workItem: cleanString(r[cols['work item'] ?? 6]),
        description: cleanString(r[cols['description'] ?? 7]),
        time: r[cols['time'] ?? 8] || '',
        leaveStatus: cleanString(r[cols['leave status'] ?? 9]),
        freeBandwidth: cleanString(r[cols['free bandwidth'] ?? 10]),
      });
    }
    res.json(data);
  } catch (err) {
    console.error('Bandwidth API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ══════════════════════════════════════════════════════════════════════
//  API: QBR Date & Blockers
// ══════════════════════════════════════════════════════════════════════

app.get('/api/qbr', async (req, res) => {
  try {
    const wb = await loadWorkbook();
    const { detailsByProject, managerByProject } = buildDropdownLookup(wb);
    const ws = wb.Sheets['QBR Date & Blockers'];
    if (!ws) return res.status(404).json({ error: 'Sheet "QBR Date & Blockers" not found' });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cols = detectColumns(rows[0]);

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const date = formatDate(r[cols['date'] ?? 0]);
      const project = cleanString(r[cols['project'] ?? 1]);
      if (!date && !project) continue;

      let projectDetails = cleanString(r[cols['project details'] ?? 2]);
      let projectManager = cleanString(r[cols['project manager'] ?? 3]);
      if (!projectDetails || projectDetails.startsWith('='))
        projectDetails = detailsByProject[project] || '';
      if (!projectManager || projectManager.startsWith('='))
        projectManager = managerByProject[project] || '';

      data.push({
        date,
        project,
        projectDetails,
        projectManager,
        mbrQbrDate: formatDate(r[cols['mbr/qbr date'] ?? 4]),
        blockers: cleanString(r[cols['blockers / dependencies'] ?? 5]),
      });
    }
    res.json(data);
  } catch (err) {
    console.error('QBR API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ══════════════════════════════════════════════════════════════════════
//  API: Drop Down (lookup data)
// ══════════════════════════════════════════════════════════════════════

app.get('/api/dropdown', async (req, res) => {
  try {
    const wb = await loadWorkbook();
    const ws = wb.Sheets['Drop Down'];
    if (!ws) return res.status(404).json({ error: 'Sheet "Drop Down" not found' });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const cols = detectColumns(rows[0]);

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      data.push({
        member: cleanString(r[cols['member'] ?? 0]),
        projectName: cleanString(r[cols['project name'] ?? 1]),
        projectDetails: cleanString(r[2]), // col C
        workItem: cleanString(r[cols['work item'] ?? 3]),
        resourceType: cleanString(r[cols['resource type'] ?? 4]),
        projectManager: cleanString(r[cols['project manager'] ?? 5]),
      });
    }
    res.json(data);
  } catch (err) {
    console.error('Dropdown API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ══════════════════════════════════════════════════════════════════════
//  API: Connect external data source
// ══════════════════════════════════════════════════════════════════════

app.post('/api/connect-source', (req, res) => {
  try {
    const { type, url } = req.body;
    if (!type || !url) {
      return res.status(400).json({ success: false, error: 'Type and URL are required.' });
    }
    connectedSource = { type, url };
    console.log(`Data source connected: [${type}] ${url}`);
    res.json({ success: true, message: `${type} source connected successfully.` });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/source-status', (req, res) => {
  res.json({
    connected: connectedSource.type !== null,
    type: connectedSource.type,
    url: connectedSource.url,
  });
});

// ══════════════════════════════════════════════════════════════════
//  PHASE 2: AUTHENTICATION
// ══════════════════════════════════════════════════════════════════

app.post('/api/auth/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const validUser = process.env.MANAGER_USER || 'admin';
    const passHash = process.env.MANAGER_PASS_HASH || '';

    if (username !== validUser) {
      return res.status(401).json({ error: 'Invalid credentials.' });
    }

    const match = await bcrypt.compare(password, passHash);
    if (!match) {
      return res.status(401).json({ error: 'Invalid credentials.' });
    }

    const token = signToken({ username, role: 'manager' });
    res.json({ success: true, token, expiresIn: '30m' });
  } catch (err) {
    console.error('Login error:', err.message);
    res.status(500).json({ error: 'Login failed.' });
  }
});

app.get('/api/auth/verify', verifyToken, (req, res) => {
  res.json({ valid: true, user: req.user });
});

// ══════════════════════════════════════════════════════════════════
//  PHASE 2: EXECUTIVE APIs (all JWT-protected)
// ══════════════════════════════════════════════════════════════════

app.get('/api/exec/summary', verifyToken, (req, res) => {
  try {
    res.json(execData.computeSummary());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/team-capacity', verifyToken, (req, res) => {
  try {
    const team = execData.getData('team');
    const leave = execData.getData('leave');
    res.json({ team: team || {}, leave: leave || {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/kpi-scorecards', verifyToken, (req, res) => {
  try {
    const kpi = execData.getData('kpi');
    res.json(kpi || { kpis: [], summary: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/sow-financial', verifyToken, (req, res) => {
  try {
    const sow = execData.getData('sow');
    res.json(sow || { projects: [], summary: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/governance-risks', verifyToken, (req, res) => {
  try {
    const gov = execData.getData('governance');
    res.json(gov || { highlights: [], risks: [], audits: [], fteTrend: [], qbrSchedule: [], summary: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/leave-impact', verifyToken, (req, res) => {
  try {
    const leave = execData.getData('leave');
    res.json(leave || { currentMonth: {}, previousMonth: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/ftr-metrics', verifyToken, (req, res) => {
  try {
    const ftr = execData.getData('ftr');
    res.json(ftr || { accounts: [], metrics: [], summary: {} });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/project-health', verifyToken, (req, res) => {
  try {
    const health = execData.computeProjectHealth();
    res.json(health);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Upload Excel files for executive dashboard ────────────────────
const execUploadFields = upload.fields([
  { name: 'ftr', maxCount: 1 },
  { name: 'team', maxCount: 1 },
  { name: 'sow', maxCount: 1 },
  { name: 'governance', maxCount: 1 },
  { name: 'leave', maxCount: 1 },
  { name: 'kpi', maxCount: 1 },
]);

app.post('/api/exec/upload-sources', verifyToken, (req, res) => {
  execUploadFields(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    try {
      const uploaded = [];
      for (const [key, files] of Object.entries(req.files || {})) {
        if (files && files[0]) {
          const ext = path.extname(files[0].originalname).toLowerCase();
          const newPath = path.join(uploadDir, `${key}${ext}`);
          fs.renameSync(files[0].path, newPath);
          execData.setUploadedFile(key, newPath);
          uploaded.push(key);
        }
      }
      await execData.loadAll();
      res.json({ success: true, uploaded, message: `${uploaded.length} file(s) uploaded & data refreshed.` });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
});

// ── Connect SharePoint/cloud URLs ─────────────────────────────────
app.post('/api/exec/connect-sources', verifyToken, async (req, res) => {
  try {
    const { sources } = req.body; // { ftr: { url, type }, team: { url, type }, ... }
    if (!sources || typeof sources !== 'object') {
      return res.status(400).json({ error: 'Sources object is required.' });
    }
    const connected = [];
    for (const [key, cfg] of Object.entries(sources)) {
      if (cfg.url) {
        execData.setRemoteSource(key, cfg.url, cfg.type || 'sharepoint');
        connected.push(key);
      }
    }
    await execData.loadAll();
    res.json({ success: true, connected, message: `${connected.length} source(s) connected & data refreshed.` });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/exec/source-status', verifyToken, (req, res) => {
  res.json(execData.getSourceStatus());
});

// ── Executive data manual refresh ─────────────────────────────────
app.post('/api/exec/refresh', verifyToken, async (req, res) => {
  try {
    await execData.loadAll();
    res.json({ success: true, message: 'Data refreshed.', lastRefresh: new Date().toISOString() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Fallback → SPA ───────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, async () => {
  console.log(`Dashboard server running at http://localhost:${PORT}`);
  console.log(`Excel path: ${EXCEL_PATH}`);
  // Load executive data sources on startup
  try {
    await execData.loadAll();
    console.log('Executive data sources loaded on startup.');
  } catch (e) {
    console.warn('Could not load exec data on startup:', e.message);
  }
});
