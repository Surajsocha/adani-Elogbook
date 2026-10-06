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

const cellStr = (r, c) => { const v = cell(r, c); return v != null ? String(v).trim() : '' }

const EQ_TABLE1 = [
  { equipment: 'COMP-1', start: 2, stop: 3, avail: 5 },
  { equipment: 'COMP-2', start: 6, stop: 7, avail: 9 },
  { equipment: 'COMP-3', start: 10, stop: 11, avail: 13 },
  { equipment: 'COMP-4', start: 14, stop: 15, avail: 17 },
]

const equipment_status = []
const parseEqLegacy = (eqList, startRowIdx) => {
  for (let cycle = 1; cycle <= 3; cycle++) {
    const rowIdx = startRowIdx + (cycle - 1)
    for (const { equipment, start, stop, avail } of eqList) {
      const startTime = formatTime(rowIdx, start)
      const stopTime = formatTime(rowIdx, stop)
      const availSt = cellStr(rowIdx, avail)
      if (!startTime && !stopTime && !availSt) continue
      equipment_status.push({
        equipment, cycle,
        start_time: startTime || null,
        stop_time: stopTime || null,
        availability_status: availSt || null,
      })
    }
  }
}
parseEqLegacy(EQ_TABLE1, 35)

console.log(equipment_status.slice(8, 12));
