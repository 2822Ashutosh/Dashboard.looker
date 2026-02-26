const XLSX = require('xlsx');

/**
 * Parse DDE Leave Tracker 2026.xlsb
 * Monthly sheets (Jan 26, Feb26, etc.) with:
 *   Row 0 = headers: Date, Day, Name1, Name2, ...
 *   Row 1+ = date rows with leave type codes (1=Planned, 2=Sick)
 */
function parseLeaveTracker(filePath) {
    const wb = XLSX.readFile(filePath);

    const now = new Date();
    const currentMonthSheets = findMonthSheets(wb.SheetNames, now);

    let currentMonth = { name: '', totalLeaves: 0, byPerson: {}, onLeaveToday: [] };
    let previousMonth = { name: '', totalLeaves: 0, byPerson: {} };

    if (currentMonthSheets.current) {
        currentMonth = parseMonthSheet(wb, currentMonthSheets.current, now);
        currentMonth.name = currentMonthSheets.current;
    }
    if (currentMonthSheets.previous) {
        previousMonth = parseMonthSheet(wb, currentMonthSheets.previous, null);
        previousMonth.name = currentMonthSheets.previous;
    }

    return { currentMonth, previousMonth };
}

/**
 * Find sheet names matching current & previous month.
 * Sheet names are like "Jan 26", "Feb26", "Dec 25", etc.
 */
function findMonthSheets(sheetNames, now) {
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const curMonth = months[now.getMonth()];
    const curYear = String(now.getFullYear()).slice(-2);
    const prevDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const prevMonth = months[prevDate.getMonth()];
    const prevYear = String(prevDate.getFullYear()).slice(-2);

    // Match variants: "Feb26", "Feb 26", "Feb-26"
    function findSheet(mon, yr) {
        return sheetNames.find(s => {
            const norm = s.replace(/[\s-]/g, '').toLowerCase();
            return norm === (mon + yr).toLowerCase() ||
                norm === (mon + ' ' + yr).replace(/\s/g, '').toLowerCase();
        });
    }

    return {
        current: findSheet(curMonth, curYear),
        previous: findSheet(prevMonth, prevYear),
    };
}

/**
 * Parse a single month sheet.
 * Header row has: Date, Day, Person1, Person2, ...
 * Data rows have: ExcelDate, DayName, leave-code or empty
 * Leave codes: 1 = Planned/Earned, 2 = Sick/Casual
 */
function parseMonthSheet(wb, sheetName, todayDate) {
    const ws = wb.Sheets[sheetName];
    if (!ws) return { totalLeaves: 0, byPerson: {}, onLeaveToday: [] };

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (rows.length < 2) return { totalLeaves: 0, byPerson: {}, onLeaveToday: [] };

    // Find the header row (contains "Date" in col 0)
    let headerIdx = 0;
    for (let i = 0; i < Math.min(rows.length, 5); i++) {
        const first = (rows[i][0] || '').toString().trim().toLowerCase();
        if (first === 'date') { headerIdx = i; break; }
    }

    const headers = rows[headerIdx];
    // Person names start from column 2 (col 0 = Date, col 1 = Day)
    const personNames = [];
    for (let c = 2; c < headers.length; c++) {
        const name = (headers[c] || '').toString().trim();
        if (name) personNames.push({ col: c, name });
    }

    const byPerson = {};
    personNames.forEach(p => { byPerson[p.name] = 0; });
    const onLeaveToday = [];
    let totalLeaves = 0;

    // Today's serial number for comparison
    let todaySerial = null;
    if (todayDate) {
        todaySerial = dateToExcelSerial(todayDate);
    }

    for (let r = headerIdx + 1; r < rows.length; r++) {
        const dateVal = rows[r][0];
        if (!dateVal) continue;

        const isToday = todaySerial && (typeof dateVal === 'number') && (dateVal === todaySerial);

        for (const { col, name } of personNames) {
            const cell = rows[r][col];
            if (cell === 1 || cell === 2 || cell === '1' || cell === '2') {
                byPerson[name] = (byPerson[name] || 0) + 1;
                totalLeaves++;
                if (isToday) {
                    onLeaveToday.push({ name, type: cell == 1 ? 'Planned' : 'Sick' });
                }
            }
        }
    }

    return { totalLeaves, byPerson, onLeaveToday };
}

function dateToExcelSerial(date) {
    // Excel serial: days since 1899-12-30
    const epoch = new Date(1899, 11, 30);
    const diff = date - epoch;
    return Math.floor(diff / (24 * 60 * 60 * 1000));
}

module.exports = { parseLeaveTracker };
