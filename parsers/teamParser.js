/**
 * Team Details Parser
 * Key sheets: "Team details 2026" (current team), "Exit resources", "Skillset Expertise 2025"
 * Team details 2026: headers at R0 (Role, Name, Emp ID, Designation, Email-ID, DOJ, Phone No, Birthday, Location)
 * Exit resources: headers at R0 (Name, Emp ID, Designation, Email-ID, DOJ, Phone No, Birthday, Last Working Day)
 */
const XLSX = require('xlsx');

function parseTeamDetails(filePath) {
    const wb = XLSX.readFile(filePath);
    const result = {
        activeMembers: [],
        exitResources: [],
        roleDistribution: {},
        totalHeadcount: 0,
        certifications: {},
    };

    // ── Primary: "Team details 2026" (most current) ──
    const teamSheet = findSheet(wb, ['Team details 2026', 'Team details 2025', 'Team details']);
    if (teamSheet) {
        const rows = XLSX.utils.sheet_to_json(teamSheet, { defval: '' });
        rows.forEach(r => {
            const name = r['Name'] || '';
            if (!name.trim()) return;
            const role = r['Role'] || r['Designation'] || '';
            const designation = r['Designation'] || role;
            result.activeMembers.push({
                name: name.trim(),
                role: designation.trim(),
                empId: r['Emp ID'] || '',
                email: r['Email-ID'] || '',
                doj: excelDate(r['DOJ']),
                phone: String(r['Phone No:'] || r['Phone No'] || ''),
                birthday: excelDate(r['Birthday (Month/date)'] || r['Birthday']),
                location: r['Current residing Location'] || '',
            });
            // Role distribution
            const roleKey = designation.trim() || 'Unspecified';
            result.roleDistribution[roleKey] = (result.roleDistribution[roleKey] || 0) + 1;
        });
    }

    // ── Fallback: older team sheets if no 2026 ──
    if (result.activeMembers.length === 0) {
        const fallback = findSheet(wb, ['Team details 2024', 'Team details 2023']);
        if (fallback) {
            const raw = XLSX.utils.sheet_to_json(fallback, { header: 1, defval: '' });
            // Headers at R1 for these sheets
            const headerRow = raw.findIndex(r => r.some(c => String(c).toLowerCase().includes('name') && String(c).toLowerCase() !== 'engagement name'));
            if (headerRow >= 0) {
                const headers = raw[headerRow].map(h => String(h).trim());
                for (let i = headerRow + 1; i < raw.length; i++) {
                    const row = raw[i];
                    const name = String(row[headers.indexOf('Name')] || '').trim();
                    if (!name) continue;
                    const desig = String(row[headers.indexOf('Designation')] || '').trim();
                    result.activeMembers.push({
                        name, role: desig,
                        empId: String(row[headers.indexOf('Emp ID')] || ''),
                        email: String(row[headers.indexOf('Email-ID')] || ''),
                        doj: excelDate(row[headers.indexOf('DOJ')]),
                        phone: String(row[headers.indexOf('Phone No:')] || ''),
                        birthday: '',
                        location: '',
                    });
                    result.roleDistribution[desig || 'Unspecified'] = (result.roleDistribution[desig || 'Unspecified'] || 0) + 1;
                }
            }
        }
    }

    // ── Exit Resources ──
    const exitSheet = findSheet(wb, ['Exit resources']);
    if (exitSheet) {
        const rows = XLSX.utils.sheet_to_json(exitSheet, { defval: '' });
        rows.forEach(r => {
            const name = r['Name'] || '';
            if (!name.trim()) return;
            result.exitResources.push({
                name: name.trim(),
                empId: r['Emp ID'] || '',
                designation: r['Designation'] || '',
                email: r['Email-ID'] || '',
                lastWorkingDay: excelDate(r['Last Working Day']),
            });
        });
    }

    // ── GA4 Certification ──
    const ga4Sheet = findSheet(wb, ['GA4 Certification']);
    if (ga4Sheet) {
        const rows = XLSX.utils.sheet_to_json(ga4Sheet, { defval: '' });
        let certified = 0, total = 0;
        rows.forEach(r => {
            const name = r['Name'] || '';
            if (!name.trim()) return;
            total++;
            if (String(r['GA4 Completion'] || '').toLowerCase() === 'yes') certified++;
        });
        result.certifications.ga4 = { certified, total, rate: total > 0 ? Math.round(certified / total * 100) : 0 };
    }

    result.totalHeadcount = result.activeMembers.length;
    return result;
}

function findSheet(wb, names) {
    for (const n of names) {
        const found = wb.SheetNames.find(s => s.toLowerCase().trim() === n.toLowerCase().trim());
        if (found) return wb.Sheets[found];
    }
    return null;
}

function excelDate(v) {
    if (!v) return '';
    if (typeof v === 'number') return new Date((v - 25569) * 86400000).toISOString().slice(0, 10);
    return String(v);
}

module.exports = { parseTeamDetails };
