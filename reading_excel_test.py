import numpy as np
import pandas as pd
import xlwings as xw
import tkinter as tk
from tkinter import filedialog

file_path = ""

# 連接到 Excel 應用程式
app = xw.App()

# 使用 Tkinter 檔案對話框選擇檔案
root = tk.Tk()
root.title('選擇檔案')
root.geometry('300x200')

def show():
    global file_path
    file_path = filedialog.askopenfilename()
    root.destroy()  # 選擇檔案後關閉 Tkinter 根視窗

btn = tk.Button(root,
                text='開啟檔案',
                font=('Arial', 20, 'bold'),
                command=show
            )
btn.pack()
root.mainloop()

print("檔案路徑:", file_path)

# 檢查檔案路徑是否不為空
if file_path:
    # 建立新的文字檔
    with open('data.txt', mode='w') as f:
        # 開啟工作簿
        wb = app.books.open(file_path)

        # 寫入文字檔
        f.write('0\n')
        f.write('0 0\n')
        f.write('1\n')

        # 鑽孔編號
        sheet = wb.sheets[1]
        f.write(str(sheet.range('C1').value) + '\n')

        # 鑽孔座標 (E / N)
        f.write(str(sheet.range('C6').value) + '\n')
        f.write(str(sheet.range('C7').value) + '\n')

        # 鑽孔孔頂高程 (EL+)
        f.write(str(sheet.range('C4').value) + '\n')

        # 鑽孔地下水位 (GL-)
        sheet = wb.sheets[4]
        f.write(str(sheet.range('C2').value) + '\n')

        # 地質圖元數 (分層數)
        df = pd.read_excel(file_path, sheet_name=8)
        row_count = df.shape[0]
        f.write(str(row_count) + '\n')

        # GL- 地質圖元代碼的 ASCII 內碼
        for index, row in df.iterrows():
            num = row['地質圖元代碼']
            if not pd.isna(num):
                sheet = wb.sheets[22]
                row_index = None
                for i, value in enumerate(sheet.range('A:A').value):
                    if value == num:
                        row_index = i + 1
                if row_index is not None:
                    value_in_row_a = sheet.range(f'A{row_index}').value
                    value_in_row_b = sheet.range(f'B{row_index}').value
                    f.write(f"{str(row['下限深度'])} {value_in_row_b}\n")
                    print(f"{num} 在 A 行的值: {value_in_row_a}")
                    print(f"對應的 B 行值: {value_in_row_b}")

        # 關閉工作簿
        wb.close()

    # 退出 Excel 應用程式
    app.quit()
else:
    print("未選擇檔案。")
