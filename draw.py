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
p1=(0,0)
p2=(0,0)

y_start=0
y_end=0
y_start_point= APoint(0, y_start)
y_end_point = APoint(0, y_end)
for index, sheet_name in enumerate(sheet_names):
    if index != 0:
        start_point = 0
        end_point = start_point + 10
        # 鑽孔名稱
        df = pd.read_excel(xl, sheet_name)
        text_value = f"{sheet_name}"  # 使用 f-string 來格式化文字
        text = acad.model.AddText(text_value, APoint(0 + index * distance, 5), 2.5)
        p1 = APoint(index * distance, 0)
        p2 = APoint(index * distance + 5, 0)
        acad.model.AddLine(p1, p2) 
        
        #畫孔位
        #讀取Layer列的數字
        previous_layer = 0
        for layer_index, row in df.iterrows():
            Layer = row['Layer']
            
            # Layer列數字迭代
            if layer_index != 0:
                if not pd.isna(Layer):
                    # 水平線點位
                    h_p1 = APoint(index * distance, -Layer)
                    h_p2 = APoint(index * distance + 5, -Layer)
                    # 水平線繪圖
                    acad.model.AddLine(h_p1, h_p2)
                #垂直線點位
                #x軸
                #左
                v_x1=index * distance
                v_x2=index * distance
                #右
                v_x3=index * distance+5
                v_x4=index * distance+5

                #y軸
                #左
                v_p1=APoint(v_x1,-previous_layer)
                v_p2=APoint(v_x2,-Layer)
                #右
                v_p3=APoint(v_x3,-previous_layer)
                v_p4=APoint(v_x4,-Layer)

                #垂直線繪圖
                acad.model.AddLine(v_p1, v_p2)
                acad.model.AddLine(v_p3, v_p4)
                previous_layer = Layer

acad.model.AddLine(y_start_point,y_end_point)

print('done')
# 退出 Excel 應用程式
app.quit()