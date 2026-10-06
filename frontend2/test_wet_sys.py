import openpyxl

wb = openpyxl.load_workbook(r'c:\Users\surya\Desktop\adani new\Dry_System_Dummy_Test_Data.xlsx', data_only=True)
ws = wb['E-Log Book Wet Sys. '] # Space at end
aoa = [list(row) for row in ws.iter_rows(values_only=True)]

for i, row in enumerate(aoa[:80]):
    # skip mostly empty rows
    non_emp = [(c_idx+1, str(v).strip()[:20]) for c_idx, v in enumerate(row) if v is not None and str(v).strip() != '']
    if len(non_emp) > 0:
        print(f"idx {i}: {non_emp}")
