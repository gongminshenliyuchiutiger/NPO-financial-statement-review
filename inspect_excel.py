import sys
try:
    import openpyxl
    print("openpyxl available")
except ImportError:
    print("openpyxl missing")
    sys.exit(1)

try:
    wb = openpyxl.load_workbook('NPO_財務報表檢核範本_v2.xlsx', read_only=True)
    for sheet in wb.sheetnames:
        print(f"=== Sheet: {sheet} ===")
        ws = wb[sheet]
        rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
        for i, row in enumerate(rows):
            print(f"Row {i+1}: {row}")
except Exception as e:
    print(f"Error reading file: {e}")
