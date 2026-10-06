const XLSX = require('xlsx');
const wb = XLSX.readFile('../Dry_System_Dummy_Test_Data.xlsx');
const ws = wb.Sheets['E-Log Book Dry Sys.'];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

for(let i=0; i<50; i++) {
  const r = aoa[i];
  if(r) {
    const nonEmp = Object.entries(r).filter(([k,v]) => v !== '').map(([k,v]) => `${k}:${v}`);
    if (nonEmp.length > 0) {
      console.log(`idx ${i}: `, nonEmp.join(', '));
    }
  }
}
