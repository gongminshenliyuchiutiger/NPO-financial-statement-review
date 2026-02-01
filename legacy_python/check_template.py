import openpyxl

try:
    wb = openpyxl.load_workbook('AI_財報自動檢核範本.xlsx')
    print("Sheets in AI_財報自動檢核範本.xlsx:")
    for sheet in wb.sheetnames:
        print(f"- {sheet}")
except Exception as e:
    print(f"Error reading excel: {e}")
