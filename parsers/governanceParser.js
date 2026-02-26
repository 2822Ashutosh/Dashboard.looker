const XLSX = require('xlsx');

/**
 * Parse Governance Sample.xlsx
 * Sheets: "Highlights & Low lights", "QBR", "FTE", "Risks and Mitigation", "Audits"
 */
function parseGovernance(filePath) {
    const wb = XLSX.readFile(filePath);

    // ── Highlights & Lowlights ────────────────────────────
    const highlights = [];
    const hlSheet = wb.Sheets['Highlights & Low lights'];
    if (hlSheet) {
        const rows = XLSX.utils.sheet_to_json(hlSheet, { defval: '' });
        rows.forEach(r => {
            const pm = (r['Project Manager'] || '').toString().trim();
            if (!pm) return;
            highlights.push({
                pm,
                project: (r['Project'] || '').toString().trim(),
                month: (r['Month'] || '').toString().trim(),
                highlight: (r['Highlight'] || '').toString().trim(),
                lowlight: (r['Lowlight'] || '').toString().trim(),
            });
        });
    }

    // ── Risks and Mitigation ──────────────────────────────
    const risks = [];
    const riskSheet = wb.Sheets['Risks and Mitigation'];
    if (riskSheet) {
        const rows = XLSX.utils.sheet_to_json(riskSheet, { defval: '' });
        rows.forEach(r => {
            const project = (r['Project'] || '').toString().trim();
            if (!project) return;
            risks.push({
                month: (r['Month'] || '').toString().trim(),
                week: (r['Week'] || '').toString().trim(),
                pm: (r['Project Manager'] || '').toString().trim(),
                project,
                risk: (r['Risks'] || '').toString().trim(),
                owner: (r['Owner'] || '').toString().trim(),
                rootCause: (r['Root cause'] || '').toString().trim(),
                mitigation: (r['Mitigation Steps'] || '').toString().trim(),
                impact: (r['Impact'] || '').toString().trim(),
                status: (r['Status'] || '').toString().trim(),
            });
        });
    }

    // ── Audits ────────────────────────────────────────────
    const audits = [];
    const auditSheet = wb.Sheets['Audits'];
    if (auditSheet) {
        const rows = XLSX.utils.sheet_to_json(auditSheet, { defval: '' });
        rows.forEach(r => {
            const project = (r['Project'] || '').toString().trim();
            if (!project) return;
            audits.push({
                month: (r['Month'] || '').toString().trim(),
                pm: (r['Project Manager'] || '').toString().trim(),
                project,
                details: (r['Audit details'] || '').toString().trim(),
                closureDate: formatExcelDate(r['Audit Closure Date']),
            });
        });
    }

    // ── FTE Trend ─────────────────────────────────────────
    const fteTrend = [];
    const fteSheet = wb.Sheets['FTE'];
    if (fteSheet) {
        const rows = XLSX.utils.sheet_to_json(fteSheet, { header: 1, defval: '' });
        // Row 0 = header (months as serial numbers), Row 1 = labels, subsequent rows = data
        if (rows.length > 2) {
            const headerRow = rows[0];
            for (let i = 1; i < headerRow.length; i++) {
                const monthSerial = headerRow[i];
                if (typeof monthSerial === 'number') {
                    const d = XLSX.SSF.parse_date_code(monthSerial);
                    const monthLabel = `${d.y}-${String(d.m).padStart(2, '0')}`;
                    // Sum all non-empty numeric values in this column from data rows
                    let totalFTE = 0;
                    for (let r = 2; r < rows.length; r++) {
                        const val = Number(rows[r][i]);
                        if (!isNaN(val)) totalFTE += val;
                    }
                    fteTrend.push({ month: monthLabel, totalFTE });
                }
            }
        }
    }

    // ── QBR Schedule ──────────────────────────────────────
    const qbrSchedule = [];
    const qbrSheet = wb.Sheets['QBR'];
    if (qbrSheet) {
        const rows = XLSX.utils.sheet_to_json(qbrSheet, { defval: '' });
        rows.forEach(r => {
            const account = (r['Account'] || '').toString().trim();
            if (!account) return;
            qbrSchedule.push({
                account,
                pms: (r['PMs'] || '').toString().trim(),
                offshoreLeads: (r['Offshore Leads'] || '').toString().trim(),
                deliveryLeads: (r['Delivery Leads'] || '').toString().trim(),
                deliveryType: (r['QBR Delivery Type'] || '').toString().trim(),
                comments: (r['Comments'] || '').toString().trim(),
            });
        });
    }

    return {
        highlights,
        risks,
        audits,
        fteTrend,
        qbrSchedule,
        summary: {
            totalHighlights: highlights.length,
            totalRisks: risks.length,
            activeRisks: risks.filter(r => r.status.toLowerCase() === 'ongoing').length,
            totalAudits: audits.length,
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

module.exports = { parseGovernance };
