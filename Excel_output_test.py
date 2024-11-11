import win32com.client as win32
from tkinter import filedialog
import tkinter as tk

def insert_rectangle_with_picture(workbook_path, image_path, sheet_index=1, start_col='B', start_row=5, end_row=10):
    """
    在指定的 Excel 檔案內插入一個矩形，並用指定的圖片填滿。

    :param workbook_path: Excel 文件保存路徑
    :param image_path: 用於填滿矩形的圖片文件路徑
    :param sheet_index: 工作表索引（默認第一個工作表）
    :param start_col: 矩形起始列
    :param start_row: 矩形起始行
    :param end_row: 矩形結束行
    """
    # 啟動 Excel 應用程式
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # 隱藏 Excel 視窗

    # 創建新工作簿
    workbook = excel.Workbooks.Add()
    worksheet = workbook.Worksheets(sheet_index)

    # 獲取列的寬度和行的高度
    col_width = worksheet.Columns[start_col].Width
    row_height_start = worksheet.Rows[start_row].Height
    row_height_end = worksheet.Rows[end_row].Height

    # 計算矩形的坐標
    left = worksheet.Columns[start_col].Left  # 矩形左邊界
    top = worksheet.Rows[start_row].Top       # 矩形上邊界
    height = row_height_end * (end_row - start_row + 1)  # 矩形高度
    width = col_width

    # 插入矩形
    shape = worksheet.Shapes.AddShape(1, left, top, width, height)

    # 設定矩形填滿為圖片
    shape.Fill.UserPicture(image_path)

    # 保存工作簿
    workbook.SaveAs(workbook_path)
    workbook.Close()
    excel.Quit()
    print(f'矩形已插入並填滿圖片，保存到: {workbook_path}')

# 使用 tkinter 獲取圖片和保存路徑
root = tk.Tk()
root.withdraw()  # 隱藏主窗口

# 選擇圖片文件
image_file_path = filedialog.askopenfilename(title="選擇用於填滿的圖片", filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.wmf")])
if not image_file_path:
    print("未選擇圖片文件。")
    exit()

# 選擇保存路徑
save_file_path = filedialog.asksaveasfilename(title="保存 Excel 文件", defaultextension=".xlsx",
                                              filetypes=[("Excel Files", "*.xlsx")])
if not save_file_path:
    print("未選擇保存文件路徑。")
    exit()

# 插入矩形並填滿圖片
insert_rectangle_with_picture(save_file_path, image_file_path)
