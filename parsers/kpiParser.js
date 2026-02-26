const XLSX = require('xlsx');

/**
 * Parse KPI Sample.xlsx
 * Sheet "KPI": 30 metrics with target, actual, status
 */
function parseKPI(filePath) {
    const wb = XLSX.readFile(filePath);

    const kpis = [];
    const ws = wb.Sheets['KPI'];
    if (ws) {
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        rows.forEach(r => {
            const metricName = (r['Metric Name (Standardized)'] || r['KPI / Metric Name ( as per Biz )'] || '').toString().trim();
            if (!metricName) return;

            const target = parseFloat(r['Target ']) || parseFloat(r['Target']) || 0;
            const actual = parseFloat(r['Actuals of Metrics Performance']) || 0;
            const status = (r['Status (Met/Not met/NA)'] || '').toString().trim();

            kpis.push({
                no: r['No.'] || '',
                buName: (r['BU Name'] || '').toString().trim(),
                department: (r['Department Name'] || '').toString().trim(),
                account: (r['Account Name (Standardized)'] || '').toString().trim(),
                metricName,
                metricDescription: (r['Metric Description'] || '').toString().trim(),
                formula: (r['Metrics Definition (Formula)'] || '').toString().trim(),
                frequency: (r['Frequency of Measurement'] || '').toString().trim(),
                unit: (r['Unit of Measurement'] || '').toString().trim(),
                dataSource: (r['Data Source'] || '').toString().trim(),
                partOfSOW: (r['Part of SOW ? ( Y/N)'] || '').toString().trim(),
                target,
                actual,
                status,
                htbOrLtb: (r['HTB or LTB'] || '').toString().trim(),
                poc: (r['POC from Ops team'] || '').toString().trim(),
            });
        });
    }

    // ── Computed KPIs ─────────────────────────────────────
    const scored = kpis.filter(k => k.status && k.status.toLowerCase() !== 'na');
    const metCount = scored.filter(k => k.status.toLowerCase().includes('met') && !k.status.toLowerCase().includes('not')).length;
    const notMetCount = scored.filter(k => k.status.toLowerCase().includes('not met')).length;
    const metRate = scored.length > 0 ? Math.round((metCount / scored.length) * 100) : 0;

    return {
        kpis,
        summary: {
            totalKPIs: kpis.length,
            scoredKPIs: scored.length,
            metCount,
            notMetCount,
            metRate,
        },
    };
}

module.exports = { parseKPI };
