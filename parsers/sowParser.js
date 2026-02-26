const XLSX = require('xlsx');

/**
 * Parse SOW and PO Tracker 2026 Sample.xlsx
 * Sheets: "FTE Projects SOW & PO", "POC and Unitized Projects SOW", "Sheet1"
 */
function parseSOW(filePath) {
    const wb = XLSX.readFile(filePath);
    const projects = [];

    // ── Parse a SOW sheet ─────────────────────────────────
    function parseSOWSheet(sheetName, category) {
        const ws = wb.Sheets[sheetName];
        if (!ws) return;
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rows.forEach(r => {
            const projectName = (r['Project Name'] || '').toString().trim();
            if (!projectName) return;
            const totalSOW = parseFloat(r['Total SOW Value']) || 0;
            const deRevenue = parseFloat(r['DE Revenue Share']) || 0;
            projects.push({
                category,
                pm: (r['PM Name'] || '').toString().trim(),
                client: (r['Client'] || '').toString().trim(),
                subProject: (r['Sub Project'] || '').toString().trim(),
                projectId: (r['MySheets Project ID '] || '').toString().trim(),
                projectName,
                totalSOWValue: totalSOW,
                deRevenueShare: deRevenue,
                startDate: formatExcelDate(r['Project Start Date']),
                endDate: formatExcelDate(r['Project End Date']),
                sowStatus: (r['SOW Status'] || '').toString().trim(),
                poStatus: (r['PO Status'] || '').toString().trim(),
                projectStatus: (r['Project Status'] || '').toString().trim(),
                comments: (r['Comments'] || '').toString().trim(),
            });
        });
    }

    parseSOWSheet('FTE Projects SOW & PO', 'FTE');
    parseSOWSheet('POC and Unitized Projects SOW', 'POC');

    // ── Computed KPIs ─────────────────────────────────────
    const totalSOWValue = projects.reduce((s, p) => s + p.totalSOWValue, 0);
    const totalRevenueShare = projects.reduce((s, p) => s + p.deRevenueShare, 0);
    const sowReceived = projects.filter(p => p.sowStatus.toLowerCase() === 'received').length;
    const poReceived = projects.filter(p => p.poStatus.toLowerCase() === 'received').length;
    const activeProjects = projects.filter(p => p.projectStatus.toLowerCase() === 'in progress').length;
    const total = projects.length || 1;

    const projectsByStatus = {};
    projects.forEach(p => {
        const st = p.projectStatus || 'Unknown';
        projectsByStatus[st] = (projectsByStatus[st] || 0) + 1;
    });

    return {
        projects,
        summary: {
            totalProjects: projects.length,
            activeProjects,
            totalSOWValue,
            totalRevenueShare,
            sowReceivedPct: Math.round((sowReceived / total) * 100),
            poReceivedPct: Math.round((poReceived / total) * 100),
            projectsByStatus,
        },
    };
}

function formatExcelDate(val) {
    if (!val) return '';
    if (typeof val === 'number') {
        const d = XLSX.SSF.parse_date_code(val);
        return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }
    return val.toString().trim();
}

module.exports = { parseSOW };
