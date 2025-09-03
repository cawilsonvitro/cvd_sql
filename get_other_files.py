from openpyxl import load_workbook
from typing import Any
import os
import sys

excel_path = r"I:\Morgano\CVD Runsheets\20250708 Runsheet.xlsx"
wb = load_workbook(excel_path)

paths = [r"I:\Morgano\CVD Runsheets",
         r"I:\Curtis\CVD Run Sheet",
         ]