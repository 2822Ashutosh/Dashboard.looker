const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const { parseFTR } = require('../parsers/ftrParser');
const { parseTeamDetails } = require('../parsers/teamParser');
const { parseSOW } = require('../parsers/sowParser');
const { parseGovernance } = require('../parsers/governanceParser');
const { parseLeaveTracker } = require('../parsers/leaveParser');
const { parseKPI } = require('../parsers/kpiParser');

// ══════════════════════════════════════════════════════════════════
//  DATA SOURCE CONFIGURATION
// ══════════════════════════════════════════════════════════════════

const DATA_SOURCES = {
    ftr: { path: process.env.EXEC_FTR_PATH || '', type: 'local', url: '' },
    team: { path: process.env.EXEC_TEAM_PATH || '', type: 'local', url: '' },
    sow: { path: process.env.EXEC_SOW_PATH || '', type: 'local', url: '' },
    governance: { path: process.env.EXEC_GOVERNANCE_PATH || '', type: 'local', url: '' },
    leave: { path: process.env.EXEC_LEAVE_PATH || '', type: 'local', url: '' },
    kpi: { path: process.env.EXEC_KPI_PATH || '', type: 'local', url: '' },
};

const FRIENDLY_NAMES = {
    ftr: 'FTR Tracker',
    team: 'Team Details',
    sow: 'SOW & PO Tracker',
    governance: 'Governance',
    leave: 'Leave Tracker',
    kpi: 'KPI',
};

// ══════════════════════════════════════════════════════════════════
//  IN-MEMORY CACHE
// ══════════════════════════════════════════════════════════════════

const cache = {
    ftr: null,
    team: null,
    sow: null,
    governance: null,
    leave: null,
    kpi: null,
    lastRefresh: null,
};

// ══════════════════════════════════════════════════════════════════
//  PARSING FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function safeParse(key, parserFn) {
    const src = DATA_SOURCES[key];
    const filePath = src.path;
    if (!filePath || !fs.existsSync(filePath)) {
        console.warn(`[ExecData] File not found for ${FRIENDLY_NAMES[key]}: ${filePath}`);
        return null;
    }
    try {
        const result = parserFn(filePath);
        console.log(`[ExecData] Parsed ${FRIENDLY_NAMES[key]} ✓`);
        return result;
    } catch (err) {
        console.error(`[ExecData] Error parsing ${FRIENDLY_NAMES[key]}:`, err.message);
        return null;
    }
}

/**
 * Download a SharePoint/OneDrive file and save locally.
 * For Google Sheets, export as xlsx.
 */
async function downloadRemoteFile(key) {
    const src = DATA_SOURCES[key];
    if (src.type !== 'sharepoint' && src.type !== 'google') return;
    if (!src.url) return;

    const uploadDir = path.join(__dirname, '..', 'uploads');
    if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });

    let downloadUrl = src.url;
    if (src.type === 'google') {
        const m = src.url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
        if (m) downloadUrl = `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`;
    }

    try {
        const resp = await fetch(downloadUrl);
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const buffer = Buffer.from(await resp.arrayBuffer());
        const ext = key === 'leave' ? '.xlsb' : '.xlsx';
        const localPath = path.join(uploadDir, `${key}${ext}`);
        fs.writeFileSync(localPath, buffer);
        src.path = localPath;
        console.log(`[ExecData] Downloaded ${FRIENDLY_NAMES[key]} from ${src.type} (${(buffer.length / 1024).toFixed(1)} KB)`);
    } catch (err) {
        console.error(`[ExecData] Download failed for ${FRIENDLY_NAMES[key]}:`, err.message);
    }
}

// ══════════════════════════════════════════════════════════════════
//  PUBLIC API
// ══════════════════════════════════════════════════════════════════

async function loadAll() {
    // Download remote sources first
    for (const key of Object.keys(DATA_SOURCES)) {
        if (DATA_SOURCES[key].type !== 'local') {
            await downloadRemoteFile(key);
        }
    }

    cache.ftr = safeParse('ftr', parseFTR);
    cache.team = safeParse('team', parseTeamDetails);
    cache.sow = safeParse('sow', parseSOW);
    cache.governance = safeParse('governance', parseGovernance);
    cache.leave = safeParse('leave', parseLeaveTracker);
    cache.kpi = safeParse('kpi', parseKPI);
    cache.lastRefresh = new Date().toISOString();

    console.log(`[ExecData] All sources loaded at ${cache.lastRefresh}`);
}

function getData(key) {
    return cache[key];
}

function getAllCached() {
    return cache;
}

function getSourceStatus() {
    const status = {};
    for (const [key, src] of Object.entries(DATA_SOURCES)) {
        status[key] = {
            name: FRIENDLY_NAMES[key],
            type: src.type,
            path: src.path,
            url: src.url || '',
            loaded: cache[key] !== null,
        };
    }
    status.lastRefresh = cache.lastRefresh;
    return status;
}

/**
 * Update a data source to use an uploaded file.
 */
function setUploadedFile(key, filePath) {
    if (DATA_SOURCES[key]) {
        DATA_SOURCES[key].type = 'local';
        DATA_SOURCES[key].path = filePath;
        DATA_SOURCES[key].url = '';
    }
}

/**
 * Update a data source to use a SharePoint/cloud URL.
 */
function setRemoteSource(key, url, type = 'sharepoint') {
    if (DATA_SOURCES[key]) {
        DATA_SOURCES[key].type = type;
        DATA_SOURCES[key].url = url;
    }
}

/**
 * Compute cross-file project health status.
 */
function computeProjectHealth() {
    const sowData = cache.sow;
    const govData = cache.governance;
    const kpiData = cache.kpi;

    if (!sowData) return [];

    return sowData.projects.map(p => {
        let health = 'Green';
        let reasons = [];

        // Check SOW/PO status
        if (p.sowStatus.toLowerCase() !== 'received') {
            health = 'Amber';
            reasons.push('SOW pending');
        }
        if (p.poStatus.toLowerCase() !== 'received' && p.poStatus.toLowerCase() !== 'ytr') {
            health = 'Amber';
            reasons.push('PO pending');
        }

        // Check risks
        if (govData) {
            const projectRisks = govData.risks.filter(r =>
                r.project.toLowerCase() === p.projectName.toLowerCase() &&
                r.status.toLowerCase() === 'ongoing'
            );
            if (projectRisks.length > 0) {
                health = projectRisks.some(r => r.impact.toLowerCase().includes('loss')) ? 'Red' : 'Amber';
                reasons.push(`${projectRisks.length} active risk(s)`);
            }
        }

        // Check project status
        if (p.projectStatus.toLowerCase() === 'yet to start') {
            health = 'Amber';
            reasons.push('Not yet started');
        }

        return {
            projectName: p.projectName,
            client: p.client,
            pm: p.pm,
            status: p.projectStatus,
            sowStatus: p.sowStatus,
            poStatus: p.poStatus,
            health,
            reasons: reasons.join('; '),
        };
    });
}

/**
 * Build the top-level executive summary.
 */
function computeSummary() {
    const team = cache.team;
    const sow = cache.sow;
    const kpi = cache.kpi;
    const ftr = cache.ftr;
    const leave = cache.leave;
    const gov = cache.governance;

    return {
        teamSize: team ? team.totalHeadcount : 0,
        activeProjects: sow ? sow.summary.activeProjects : 0,
        totalSOWValue: sow ? sow.summary.totalSOWValue : 0,
        kpiMetRate: kpi ? kpi.summary.metRate : 0,
        ftrAvgRating: ftr ? ftr.summary.avgRating : 0,
        onLeaveToday: leave && leave.currentMonth ? leave.currentMonth.onLeaveToday.length : 0,
        activeRisks: gov ? gov.summary.activeRisks : 0,
        lastRefresh: cache.lastRefresh,
    };
}

module.exports = {
    loadAll,
    getData,
    getAllCached,
    getSourceStatus,
    setUploadedFile,
    setRemoteSource,
    computeProjectHealth,
    computeSummary,
    DATA_SOURCES,
    FRIENDLY_NAMES,
};
