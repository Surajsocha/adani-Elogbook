export function mapWetSheets(sheets) {
  // ── Detect format ──
  const isLegacy = !!sheets['E-Log Book Wet Sys. ']
  const si = sheets['Shift Info']
  let formData = {}
  const esp_wet_deashing = []
  const esp_field_availability = []
  const pump_status = []

  if (isLegacy) {
    const aoa = sheets['E-Log Book Wet Sys. ']
    const cell = (r, c) => {
      const row = aoa[r]
      if (!row) return null
      const v = row[c]
      return (v === undefined || v === null || v === '') ? null : v
    }
    const cellStr = (r, c) => { const v = cell(r, c); return v != null ? String(v).trim() : '' }
    const num = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n }
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

    let date = cellStr(9, 3) || cellStr(9, 2)
    try {
      if (date.includes('-')) {
        const parts = date.split('-')
        if (parts[2].length === 4) date = `${parts[2]}-${parts[1]}-${parts[0]}`
      }
      date = new Date(date).toISOString().split('T')[0]
    } catch { date = new Date().toISOString().split('T')[0] }

    formData = {
      date,
      shift: cellStr(9, 17)?.[0] || 'A',
      unit1_load: num(cellStr(13, 3)),
      unit1_coal_flow: num(cellStr(14, 3)),
      unit2_load: num(cellStr(13, 23)),
      unit2_coal_flow: num(cellStr(14, 23)),
      u1_esp_fields_discharged: cellStr(12, 11),
      u1_esp_hopper_level_hi: cellStr(13, 11),
      u1_esp_cerm_avail: cellStr(14, 11),
      u1_esp_eerm_avail: cellStr(15, 11),
      u2_esp_fields_discharged: cellStr(12, 31),
      u2_esp_hopper_level_hi: cellStr(13, 31),
      u2_esp_cerm_avail: cellStr(14, 31),
      u2_esp_eerm_avail: cellStr(15, 31),
      u1_ba_start: formatTime(27, 3),
      u1_ba_stop: formatTime(27, 8),
      u2_ba_start: formatTime(28, 3),
      u2_ba_stop: formatTime(28, 8),
      incomer_osa_initial: num(cellStr(18, 5)),
      incomer_osa_final: num(cellStr(18, 8)),
      incomer_osc_initial: num(cellStr(18, 20)),
      incomer_osc_final: num(cellStr(18, 23)),
      ahp_lighting_initial: num(cellStr(21, 5)),
      ahp_lighting_final: num(cellStr(21, 8)),
      events_remarks: cellStr(79, 3) || '',
      follow_up: '',
      protection_bypassed: '',
      observations: '',
      document_no: 'ADTPS/AHP/OPN/F/02',
    }

    const ESP_PASS_COLS = [
      { pass: 'AB-1', start: 3, stop: 4 },
      { pass: 'CD-1', start: 6, stop: 7 },
      { pass: 'AB-2', start: 9, stop: 10 },
      { pass: 'CD-2', start: 12, stop: 13 },
      { pass: 'AB-3', start: 15, stop: 16 },
      { pass: 'CD-3', start: 18, stop: 19 },
      { pass: 'AB-4', start: 21, stop: 22 },
      { pass: 'CD-4', start: 24, stop: 25 },
      { pass: 'AB-5', start: 27, stop: 28 },
      { pass: 'CD-5', start: 30, stop: 31 },
      { pass: 'AB-6', start: 33, stop: 34 },
      { pass: 'CD-6', start: 36, stop: 37 },
    ]
    const parseEspLegacy = (startRowIdx, unit) => {
      for (let cycle = 1; cycle <= 3; cycle++) {
        const rowIdx = startRowIdx + (cycle - 1)
        for (const { pass, start, stop } of ESP_PASS_COLS) {
          const startTime = formatTime(rowIdx, start)
          const stopTime = formatTime(rowIdx, stop)
          if (!startTime && !stopTime) continue
          esp_wet_deashing.push({
            unit, component: pass, cycle,
            start_time: startTime || null,
            stop_time: stopTime || null,
            total_time: calcTimeDiff(startTime, stopTime),
          })
        }
      }
    }
    parseEspLegacy(33, '1')
    parseEspLegacy(40, '2')

    const PUMPS1 = [
      { pump: 'ADP-A', start: 3, stop: 4, avail: 6 },
      { pump: 'ADP-B', start: 7, stop: 8, avail: 10 },
      { pump: 'ADP-C', start: 11, stop: 12, avail: 14 },
      { pump: 'ADP-D', start: 15, stop: 16, avail: 18 },
      { pump: 'HPP-A', start: 19, stop: 20, avail: 22 },
      { pump: 'HPP-B', start: 23, stop: 24, avail: 26 },
      { pump: 'HPP-C', start: 27, stop: 28, avail: 30 },
      { pump: 'SWP-A', start: 31, stop: 32, avail: 34 },
      { pump: 'SWP-B', start: 35, stop: 36, avail: 38 },
    ]
    const parsePumpsLegacy = (pList, startRowIdx) => {
      for (let cycle = 1; cycle <= 4; cycle++) {
        const rowIdx = startRowIdx + (cycle - 1)
        for (const { pump, start, stop, avail } of pList) {
          const startTime = formatTime(rowIdx, start)
          const stopTime = formatTime(rowIdx, stop)
          const availSt = cellStr(rowIdx, avail)
          if (!startTime && !stopTime && !availSt) continue
          pump_status.push({
            pump, cycle,
            start_time: startTime || null,
            stop_time: stopTime || null,
            running_hrs: calcTimeDiff(startTime, stopTime),
            avail_status: availSt || null,
          })
        }
      }
    }
    parsePumpsLegacy(PUMPS1, 58)

  } else {
    // ── Template Format ──
    const hdrs = si?.[0] || []
    const dataR = si?.[1] || []
    const main = {}
    hdrs.forEach((h, i) => { main[safe(h)] = safe(dataR[i]) })
    const r = (k) => main[k] || ''
    const num = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n }
    const n = (val) => val === undefined || val === null || val === '' ? null : val

    let date = r('Date\n(YYYY-MM-DD)')
    try { date = new Date(date).toISOString().split('T')[0] }
    catch { date = new Date().toISOString().split('T')[0] }

    formData = {
      date,
      shift: r('Shift\n(A/B/C/G)')?.[0] || 'A',
      unit1_load: num(main['Unit-1 Load\n(MW)']),
      unit1_coal_flow: num(main['U1 Coal Flow\n(T/hr)']),
      unit2_load: num(main['Unit-2 Load\n(MW)']),
      unit2_coal_flow: num(main['U2 Coal Flow\n(T/hr)']),
      u1_esp_fields_discharged: r('U1 FTP Fields\nDischarged'),
      u1_esp_hopper_level_hi: r('U1 Hopper\nLevel Hi'),
      u1_esp_cerm_avail: r('U1 ESP CERM\nAvailability'),
      u1_esp_eerm_avail: r('U1 FHP EERM\nAvailability'),
      u2_esp_fields_discharged: r('U2 FTP Fields\nDischarged'),
      u2_esp_hopper_level_hi: r('U2 Hopper\nLevel Hi'),
      u2_esp_cerm_avail: r('U2 ESP CERM\nAvailability'),
      u2_esp_eerm_avail: r('U2 FHP EERM\nAvailability'),
      u1_ba_start: r('U1 BA Start\n(HH:MM)'),
      u1_ba_stop: r('U1 BA Stop\n(HH:MM)'),
      u2_ba_start: r('U2 BA Start\n(HH:MM)'),
      u2_ba_stop: r('U2 BA Stop\n(HH:MM)'),
      incomer_osa_initial: num(main['Incomer OSA\nInitial']),
      incomer_osa_final: num(main['Incomer OSA\nFinal']),
      incomer_osc_initial: num(main['Incomer OSC\nInitial']),
      incomer_osc_final: num(main['Incomer OSC\nFinal']),
      ahp_lighting_initial: num(main['AHP Lighting\nInitial']),
      ahp_lighting_final: num(main['AHP Lighting\nFinal']),
      events_remarks: r('Events /\nRemarks'),
      follow_up: r('Follow Up'),
      protection_bypassed: r('Protection\nBypassed'),
      observations: r('Observations'),
      document_no: 'ADTPS/AHP/OPN/F/02',
    }

    const parseWetEsp = (sheetAoa, unit, compKey = 'ESP Field') => {
      const rows = aoa2obj(sheetAoa, 1)
      for (const row of rows) {
        const component = safe(row[compKey]).trim()
        const cycle = parseInt(row['Cycle\n(1/2/3)'])
        if (!component || !cycle) continue
        esp_wet_deashing.push({
          unit, component, cycle,
          start_time: n(row['Start Time\n(HH:MM)']),
          stop_time: n(row['Stop Time\n(HH:MM)']),
          total_time: calcTimeDiff(n(row['Start Time\n(HH:MM)']), n(row['Stop Time\n(HH:MM)'])),
        })
      }
    }
    if (sheets['U1 ESP Wet Deashing']) parseWetEsp(sheets['U1 ESP Wet Deashing'], '1')
    if (sheets['U2 ESP Wet Deashing']) parseWetEsp(sheets['U2 ESP Wet Deashing'], '2')

    const ecoSheet = sheets['ECO APH Deashing']
    if (ecoSheet) {
      const rows = aoa2obj(ecoSheet, 1)
      for (const row of rows) {
        const unit = safe(row['Unit\n(1/2)']).trim()
        const component = safe(row['Component']).trim()
        const cycle = parseInt(row['Cycle\n(1/2/3)'])
        if (!unit || !component || !cycle) continue
        esp_wet_deashing.push({
          unit, component, cycle,
          start_time: n(row['Start Time\n(HH:MM)']),
          stop_time: n(row['Stop Time\n(HH:MM)']),
          total_time: calcTimeDiff(n(row['Start Time\n(HH:MM)']), n(row['Stop Time\n(HH:MM)'])),
        })
      }
    }

    const favSheet = sheets['Field Availability']
    if (favSheet) {
      const rows = aoa2obj(favSheet, 1)
      for (const row of rows) {
        const unit = safe(row['Unit\n(1/2)']).trim()
        const field = safe(row['Field / Component']).trim()
        if (!unit || !field || field.startsWith('===')) continue
        esp_field_availability.push({
          unit, field,
          avail_status: n(row['Availability\n(R/A/B/P)']),
        })
      }
    }

    const parsePumps = (sheetAoa) => {
      const rows = aoa2obj(sheetAoa, 1)
      for (const row of rows) {
        const pump = safe(row['Pump']).trim()
        const cycle = parseInt(row['Cycle\n(1/2/3/4)'])
        if (!pump || pump.startsWith('===') || !cycle) continue
        pump_status.push({
          pump, cycle,
          start_time: n(row['Start Time\n(HH:MM)']),
          stop_time: n(row['Stop Time\n(HH:MM)']),
          running_hrs: calcTimeDiff(n(row['Start Time\n(HH:MM)']), n(row['Stop Time\n(HH:MM)'])),
          avail_status: n(row['Status\n(R/A/B/P)']),
        })
      }
    }
    if (sheets['HT Pumps']) parsePumps(sheets['HT Pumps'])
    if (sheets['LT Pumps']) parsePumps(sheets['LT Pumps'])
  }

  return { formData, esp_wet_deashing, esp_field_availability, pump_status }
}
