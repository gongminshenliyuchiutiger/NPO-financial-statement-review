import openpyxl
import sys

# Force utf-8 encoding for stdout
sys.stdout.reconfigure(encoding='utf-8')

try:
    wb = openpyxl.load_workbook('NPO_財務報表檢核範本_v2.xlsx')
    for sheet in wb.sheetnames:
        print(f"=== Sheet: {sheet} ===")
        ws = wb[sheet]
        row1 = [cell.value for cell in ws[1]]
        print(f"Headers: {row1}")
except Exception as e:
    print(f"Error: {e}")
