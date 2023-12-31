import numpy as np
import pandas as pd
import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog

file_path=""
# 連接到活動的Excel應用程序
app = xw.App()

#選取檔案
root = tk.Tk()
root.title('Choose a file')
root.geometry('300x200')

def show():
    global file_path
    file_path = filedialog.askopenfilename()   # 選擇檔案後回傳檔案路徑與名稱
    
# Button 設定 command 參數，點擊按鈕時執行 show 函式
btn = tk.Button(root,
                text='開啟檔案',
                font=('Arial',20,'bold'),
                command=show
            )

# def close_window():
#     root.destroy()

# btn.config(command=close_window)


btn.pack()

root.mainloop()
print("path name",file_path)

#新增一個txt檔
f=open('data.txt', mode='w')

# 打開工作簿

wb = app.books.open(file_path)

#不匯出點位資料
f.write('0''\n')
f.write('0'' ''0''\n')

#總孔數
f.write('1''\n')

# 鑽孔編號
sheet = wb.sheets[1]
f.write(sheet.range('C1').value+'\n')

#鑽孔 E / N 座標	
f.write(sheet.range('C6').value+'\n')
f.write(sheet.range('C7').value+'\n')

#鑽孔孔頂高程 EL+
f.write(sheet.range('C4').value+'\n')

#鑽孔地下水位 GL- (若地下水位在甚深處，則填 999，地下水位將不繪出)
sheet=wb.sheets[4]
f.write(sheet.range('C2').value+'\n')

#鑽孔地質圖元數(分層數)，不可為 0
df_dict = pd.read_excel(file_path, sheet_name=None)

# 获取sheet的名字列表
sheet_names = list(df_dict.keys())
print(sheet_names)

# 如果要读取第8个sheet，可以通过索引或者名字
sheet_index = 8  # 由于Python索引从0开始，所以这里是7
sheet_name = '岩石或土壤性質描述'  # 如果知道sheet的名字，可以直接使用名字

# 通过索引读取第8个sheet
df_sheet_8_by_index = df_dict[sheet_names[sheet_index]]

# 通过名字读取第8个sheet
df_sheet_8_by_name = df_dict[sheet_name]

#GL- 地質圖元代碼的 ASCII 內碼
f.write(str(len(df_sheet_8_by_name)) + '\n')
# 关闭工作簿
wb.close()

# 关闭Excel应用程序
app.quit()
