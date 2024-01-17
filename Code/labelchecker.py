from openpyxl import load_workbook
from pathlib import Path

# xlsx 파일은 Data 폴더에 넣을 것.
# Path 객체에서 상대 경로는 working directory를 기준으로 해석됨.
# 어디에서 코드를 실행할지 확실치 않으므로 Data 폴더 찾을 때 현재 코드 위치를 기준으로 찾도록 해야 함.
filePath = Path(__file__).parents[1] / "Data" / input("Enter file name: ")
wb = load_workbook(filePath, data_only=True)
ws = wb.active  # active sheet 선택

ws.close()
