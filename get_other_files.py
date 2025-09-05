from openpyxl import load_workbook
from typing import Any
import os
import sys

# excel_path = r"I:\Morgano\CVD Runsheets\20250708 Runsheet.xlsx"
# wb = load_workbook(excel_path)
wb_paths:list[str] = [] 
paths = [r"I:\Morgano\CVD Runsheets",
         r"I:\Curtis\CVD Run Sheet",
         r"I:\Gotera\CVD Data",
         r"I:\Zele\Backup\CVD Project\SierraRunsheets\2025\May 2025"
         ]



for path in paths:
    print(path)
    for file in os.listdir(path):
        print(file)
        if "runsheet" in file.lower():
            if os.path.join(path, file) not in wb_paths:
                print(os.path.join(path, file))
                wb_paths.append(os.path.join(path, file))