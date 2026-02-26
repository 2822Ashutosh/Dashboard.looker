const XLSX = require('xlsx');

/**
 * Parse Team Details Sample.xlsx
 * Key sheets: "Team details 2026", "Exit resources"
 */
function parseTeamDetails(filePath) {
    const wb = XLSX.readFile(filePath);

    // ── Active Team Members ───────────────────────────────
    const activeMembers = [];
    const teamSheet = wb.Sheets['Team details 2026'];
    if (teamSheet) {
        const rows = XLSX.utils.sheet_to_json(teamSheet, { defval: '' });
        rows.forEach(r => {
            const name = (r['Name'] || '').toString().trim();
            if (!name) return;
            activeMembers.push({
                role: (r['Role'] || '').toString().trim(),
                name,
                empId: (r['Emp ID'] || '').toString().trim(),
                designation: (r['Designation'] || '').toString().trim(),
                email: (r['Email-ID'] || '').toString().trim(),
                doj: formatExcelDate(r['DOJ ']),
                location: (r['Current residing Location'] || '').toString().trim(),
            });
        });
    }

    // ── Exit Resources ────────────────────────────────────
    const exitedMembers = [];
    const exitSheet = wb.Sheets['Exit resources'];
    if (exitSheet) {
        const rows = XLSX.utils.sheet_to_json(exitSheet, { defval: '' });
        rows.forEach(r => {
            const name = (r['Name'] || '').toString().trim();
            if (!name) return;
            exitedMembers.push({
                name,
                empId: (r['Emp ID'] || '').toString().trim(),
                designation: (r['Designation'] || '').toString().trim(),
                lastWorkingDay: formatExcelDate(r['Last Working Day']),
            });
        });
    }

    // ── Role Distribution ─────────────────────────────────
    const roleDistribution = {};
    activeMembers.forEach(m => {
        const role = m.role || 'Unassigned';
        roleDistribution[role] = (roleDistribution[role] || 0) + 1;
    });

    return {
        activeMembers,
        exitedMembers,
        totalHeadcount: activeMembers.length,
        attritionCount: exitedMembers.length,
        roleDistribution,
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

module.exports = { parseTeamDetails };
