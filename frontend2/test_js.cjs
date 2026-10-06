const XLSX = require('xlsx');
const wb = XLSX.readFile('../Dry_System_Dummy_Test_Data.xlsx');
const ws = wb.Sheets['E-Log Book Dry Sys.'];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

const cell = (r, c) => {
    const row = aoa[r];
    if (!row) return null;
    const v = row[c];
    return (v === undefined || v === null || v === '') ? null : v;
};

const formatTime = (r, c) => {
    const v = cell(r, c)
    if (v == null || v === '') return null
    if (typeof v === 'number' && v >= 0 && v < 1) {
      const totalSeconds = Math.round(v * 24 * 3600)
      const h = Math.floor(totalSeconds / 3600)
      const m = Math.floor((totalSeconds % 3600) / 60)
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`
    }
    if (typeof v === 'string') {
      const match = v.match(/(\d{1,2})[:.](\d{2})/)
      if (match) return `${match[1].padStart(2, '0')}:${match[2]}`
      return v.trim() || null
    }
    return String(v).trim() || null
}

console.log('Cycle 3 for COMP-1 (Row idx 38):');
console.log('Raw start:', cell(38, 2), 'type:', typeof cell(38, 2));
console.log('Raw stop:', cell(38, 3), 'type:', typeof cell(38, 3));
console.log('Formatted start:', formatTime(38, 2));
console.log('Formatted stop:', formatTime(38, 3));
