import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# 创建一个新的工作簿
wb = Workbook()

# 获取当前活动的工作表
ws = wb.active

xl=pd.ExcelFile('七孔測試.xlsx')
sheet_names = xl.sheet_names
new_wb = openpyxl.Workbook()
new_wb.remove(new_wb.active)