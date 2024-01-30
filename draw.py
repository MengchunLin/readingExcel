import pandas as pd
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
from pyautocad import Autocad
from pyautocad import Autocad, APoint

file_path = ""
Upper_depth = 5
Lower_depth = 0
distance=50

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

# 自動連線上 CAD
acad = Autocad(create_if_not_exists=True)
print(acad.doc.Name)

#單位設定
acad.ActiveDocument.SetVariable("INSUNITS", 6)  # 6 代表公尺


# 讀取 Excel 檔案中的所有工作表
xl = pd.ExcelFile(file_path)

# 透過 sheet_names 取得所有工作表的名稱
sheet_names = xl.sheet_names

# 印出每個工作表的資料
p0=APoint(0,0)
for index, sheet_name in enumerate(sheet_names):
    if index != 0:
        start_point=0
        end_point=start_point+10
        #鑽孔名稱
        df = pd.read_excel(xl, sheet_name)
        text_value = f"{sheet_name}"  # 使用 f-string 來格式化文字
        text = acad.model.AddText(text_value, APoint(0 + index * distance, 0), 2.5)

        #畫孔位-水平
        for layer_index, row in df.iterrows():
            Layer=row['Layer']
            if layer_index!=0:
                if not pd.isna(Layer):
            
                    print(Layer)


            
    
         

# 退出 Excel 應用程式
app.quit()