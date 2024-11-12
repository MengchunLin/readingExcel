import win32com.client as win32
import tkinter as tk
from tkinter import filedialog

def get_column_width(worksheet, col_letter):
    """獲取指定列的寬度"""
    return worksheet.Columns[col_letter].Width
def get_row_height(worksheet, row):
    """獲取指定行的高度"""
    return worksheet.Rows[row].Height
# 打開文件選擇對話框
root = tk.Tk()
root.withdraw()  # 隱藏主窗口
file_path = filedialog.askopenfilename(title="選擇 WMF 圖片文件", filetypes=[("WMF files", "*.wmf")])

# 啟動 Excel 應用程序
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

# 創建新工作簿和工作表
workbook = excel.Workbooks.Add()
worksheet = workbook.Worksheets(1)

# 定義矩形的列和行範圍
start_col = 'B'  # 矩形開始的列
start_row = 5    # 矩形開始的行
end_row = 10     # 矩形結束的行
# 獲取指定列的寬度和指定行的高度
col_width = get_column_width(worksheet, start_col)
row_height_start = get_row_height(worksheet, start_row)
row_height_end = get_row_height(worksheet, end_row)
# 計算矩形的左上角和右下角的坐標
left = worksheet.Columns[start_col].Left  # 矩形左邊界
top = worksheet.Rows[start_row].Top        # 矩形上邊界
height = row_height_end * (end_row - start_row + 1)  # 矩形的高度（從 B5 到 B10）
width = col_width 
# 插入矩形

shape = worksheet.Shapes.AddShape(1, left, top, width, height)

# 設定矩形填滿為圖片紋理
if file_path:
    # 設置填滿為圖片
    shape.Fill.UserPicture(file_path)

    # 設置填滿為圖片紋理
    shape.Fill.TextureTile = True

    # 調整刻度
    shape.Fill.TextureOffsetX = 0.05  # X 刻度: 10%
    shape.Fill.TextureOffsetY = 0.05  # Y 刻度: 10%
    shape.Fill.TextureHorizontalScale = 0.05  # X 刻度百分比: 10%
    shape.Fill.TextureVerticalScale = 0.05  # Y 刻度百分比: 10%

# 保存工作簿
workbook.SaveAs('example.xlsx')