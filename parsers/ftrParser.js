const XLSX = require('xlsx');

/**
 * Parse FTR Tracker.xlsx
 * Sheets: "Account Details" (engagement metadata), "Bayer Quality Metrics" (QA data)
 */
function parseFTR(filePath) {
    const wb = XLSX.readFile(filePath);

    // ── Account Details ───────────────────────────────────
    const accounts = [];
    const accSheet = wb.Sheets['Account Details'];
    if (accSheet) {
        const rows = XLSX.utils.sheet_to_json(accSheet, { defval: '' });
        rows.forEach(r => {
            if (r['Engagement Name']) {
                accounts.push({
                    engagement: (r['Engagement Name'] || '').toString().trim(),
                    qaTracking: (r['QA Metrics Tracking'] || '').toString().trim(),
                    expectedProjects: (r['Expected projects per month'] || '').toString().trim(),
                    dataLayerType: (r['Datalayer/DOM scrapping'] || '').toString().trim(),
                    webTeam: (r['Web team- Internal/External'] || '').toString().trim(),
                });
            }
        });
    }

    // ── Bayer Quality Metrics ─────────────────────────────
    const metrics = [];
    const qmSheet = wb.Sheets['Bayer Quality Metrics '] || wb.Sheets['Bayer Quality Metrics'];
    if (qmSheet) {
        const rows = XLSX.utils.sheet_to_json(qmSheet, { defval: '' });
        rows.forEach(r => {
            const projectName = (r['Project Name'] || '').toString().trim();
            if (!projectName) return;
            const totalTags = Number(r['Total Tags']) || 0;
            const totalFailed = Number(r['Total Failed']) || 0;
            const rating = Number(r['Rating']) || 0;
            metrics.push({
                project: projectName,
                qaDate: formatExcelDate(r['QA Date']),
                global: Number(r['Global']) || 0,
                globalFailed: Number(r['Global Failed']) || 0,
                inpage: Number(r['Inpage']) || 0,
                inpageFailed: Number(r['Inpage Failed']) || 0,
                form: Number(r['Form']) || 0,
                formFailed: Number(r['Form Failed']) || 0,
                videos: Number(r['Videos']) || 0,
                videosFailed: Number(r['Videos Failed']) || 0,
                totalTags,
                totalFailed,
                rating,
            });
        });
    }

    // ── Computed KPIs ─────────────────────────────────────
    const totalProjects = metrics.length;
    const avgRating = totalProjects > 0
        ? metrics.reduce((s, m) => s + m.rating, 0) / totalProjects
        : 0;
    const passCount = metrics.filter(m => m.rating >= 0.95).length;
    const ftrPassRate = totalProjects > 0 ? passCount / totalProjects : 0;

    return {
        accounts,
        metrics,
        summary: {
            totalProjects,
            avgRating: Math.round(avgRating * 10000) / 100,
            ftrPassRate: Math.round(ftrPassRate * 10000) / 100,
            totalTagsChecked: metrics.reduce((s, m) => s + m.totalTags, 0),
            totalFailed: metrics.reduce((s, m) => s + m.totalFailed, 0),
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

module.exports = { parseFTR };
