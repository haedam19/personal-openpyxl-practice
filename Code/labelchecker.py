from openpyxl import load_workbook
from pathlib import Path

# xlsx 파일은 Data 폴더에 넣을 것.
filePath = Path("Data") / input("Enter file name: ")
wb = load_workbook(filePath, data_only=True)
ws = wb.active  # active sheet 선택

ws.close()
