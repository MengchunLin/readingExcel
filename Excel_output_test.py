import win32com.client as win32
import tkinter as tk
from tkinter import filedialog

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

# 插入矩形
shape = worksheet.Shapes.AddShape(1, 100, 100, 200, 100)  # 1 = msoShapeRectangle

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


