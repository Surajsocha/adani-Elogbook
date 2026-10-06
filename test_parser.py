import openpyxl

wb = openpyxl.load_workbook(r'c:\Users\surya\Desktop\adani new\Dry_System_Dummy_Test_Data.xlsx', data_only=True)
ws = wb['E-Log Book Dry Sys.']
aoa = [list(row) for row in ws.iter_rows(values_only=True)]

def cell(r, c):
    if r >= len(aoa): return None
    row = aoa[r]
    if c >= len(row): return None
    v = row[c]
    return None if (v is None or v == '') else v

EQ_TABLE1 = [
    {'equipment': 'COMP-1', 'start': 2, 'stop': 3, 'avail': 5},
    {'equipment': 'COMP-2', 'start': 6, 'stop': 7, 'avail': 9},
    {'equipment': 'COMP-3', 'start': 10, 'stop': 11, 'avail': 13},
    {'equipment': 'COMP-4', 'start': 14, 'stop': 15, 'avail': 17},
]
EQ_TABLE2 = [
    {'equipment': 'CT-FAN', 'start': 2, 'stop': 3, 'avail': 5},
    {'equipment': 'CT-PUMP-A', 'start': 6, 'stop': 7, 'avail': 9},
    {'equipment': 'CT-PUMP-B', 'start': 10, 'stop': 11, 'avail': 13},
    {'equipment': 'CLASS-A', 'start': 14, 'stop': 15, 'avail': 17},
    {'equipment': 'CLASS-B', 'start': 18, 'stop': 19, 'avail': 21},
]

items = []
def parseEq(eqList, startRowIdx):
    for cycle in range(1, 4):
        rowIdx = startRowIdx + (cycle - 1)
        for eq in eqList:
            start = cell(rowIdx, eq['start'])
            stop = cell(rowIdx, eq['stop'])
            avail = cell(rowIdx, eq['avail'])
            if start is None and stop is None and avail is None:
                continue
            items.append({
                'eq': eq['equipment'],
                'cycle': cycle,
                'start_raw': start,
                'stop_raw': stop,
                'avail': avail
            })

parseEq(EQ_TABLE1, 36)
parseEq(EQ_TABLE2, 43)

for i, item in enumerate(items):
    print(f"Item {i}: {item}")
