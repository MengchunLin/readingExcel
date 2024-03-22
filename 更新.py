import pandas as pd
import numpy as np
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
from pyautocad import Autocad, APoint,aDouble
import array
import win32com.client
import math
import pythoncom

file_path = ""
Upper_depth = 5
Lower_depth = 0
distance=25
hole_width=2
#ruler
ruler_top=0
ruler_bottom=0

#放大倍數
scale_factor_h=0
scale_factor_w=0


# 連接到 Excel 應用程式
app = xw.App()
# 使用 Tkinter 檔案對話框選擇檔案
root = tk.Tk()
root.title('選擇檔案')
root.geometry('300x200')

#輸入框
label = tk.Label(root, text="長度比例:")
label.pack()
entry = tk.Entry(root)
entry.pack()

label = tk.Label(root, text="寬度比例:")
label.pack()
entry = tk.Entry(root)
entry.pack()

def show():
    global file_path,scale_factor_h, scale_factor_w
    file_path = filedialog.askopenfilename()
    scale_factor_h = float(entry.get())
    scale_factor_w = float(entry.get())
    root.destroy()  # 選擇檔案後關閉 Tkinter 根視窗

def read_excel_cell(file_path, sheet_name, row_index, col_name):
    # 讀取 Excel 檔案中特定儲存格的值
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    value = df.at[row_index, col_name]
    return value


btn = tk.Button(root,
                text='開啟檔案',
                font=('Arial', 20, 'bold'),
                command=show
            )
btn.pack()
root.mainloop()

# 自動連線上 CAD
wincad  = win32com.client.Dispatch("AutoCAD.Application")
# wincad  = Autocad(create_if_not_exists=True)

doc = wincad.ActiveDocument
msp = doc.ModelSpace

doc.SetVariable("INSUNITS", 6)

acad =Autocad().ActiveDocument.ModelSpace

# 讀取 Excel 檔案中的所有工作表
xl = pd.ExcelFile(file_path)

# 透過 sheet_names 取得所有工作表的名稱
sheet_names = xl.sheet_names

# 印出每個工作表的資料
p1=(0,0)
p2=(0,0)

y_start=0
y_end=0
y_start_point= APoint(0, y_start)
y_end_point = APoint(0, y_end)

def vtobj(obj):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

def vtfloat(lst):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)

for index, sheet_name in enumerate(sheet_names):
    #skip the first sheet
    N_1=0
    N_2=0
    E_1=0
    E_2=0
    if index != 0:
 #------------------------------------------------------------------------------------------------------------------------------------
        #孔頂高
        df = pd.read_excel(xl, sheet_name)
        row_index, col_index = np.where(df == 'Ground EL')
        col_index = col_index[0]
        row_index = row_index[0]
        next_col_index = col_index + 1
        Ground_EL = df.iloc[row_index, next_col_index]
        #print(Ground_EL)
#------------------------------------------------------------------------------------------------------------------------------------
        #位置distance
        row_index, col_index = np.where(df == 'N')
        col_index = col_index[0]
        row_index = row_index[0]
        next_col_index = col_index + 1
        N_2=df.iloc[row_index, next_col_index]
        row_index, col_index = np.where(df == 'E')
        col_index = col_index[0]
        row_index = row_index[0]
        next_col_index = col_index + 1
        E_2 = df.iloc[row_index, next_col_index]
        if index==1:
            distance=15
        else:
            distance_=pow(pow(E_2-E_1,2)+pow(N_2-N_1,2),0.5)/1000000*scale_factor_w
            E_1=E_2
            N_1=N_2
            distance=distance+distance_
        #print(distance)


 #------------------------------------------------------------------------------------------------------------------------------------
        # 鑽孔名稱 (hole_width/2*scale_factor))
        text_value = f"{sheet_name}"  # 使用 f-string 來格式化文字
        word_height=2
        #text_width = len(text_value) * 2.5 * scale_factor_w 
        insert_point=APoint((distance+hole_width/2)*scale_factor_w, (Ground_EL+5)*scale_factor_h,(Ground_EL+5)*scale_factor_h) 
        text = acad.AddText(text_value,insert_point, 1*scale_factor_w)
        text.Alignment=13
        text.TextAlignmentPoint = insert_point
        #EL
        text_value='EL  '+str(Ground_EL)
        insert_point=APoint((distance+hole_width/2)*scale_factor_w, (Ground_EL+3)*scale_factor_h,(Ground_EL+3)*scale_factor_h)
        text = acad.AddText(text_value,insert_point, 0.8*scale_factor_w)
        text.Alignment=13
        text.TextAlignmentPoint = insert_point
        #print(insert_point) 

 #------------------------------------------------------------------------------------------------------------------------------------
        #土層深度
        text_value='Depth(m)'
        insert_point=APoint((distance)*scale_factor_w, Ground_EL*scale_factor_h,Ground_EL*scale_factor_h) 
        text=acad.AddText(text_value,insert_point, 0.8*scale_factor_w)
        text.Alignment=14
        text.TextAlignmentPoint = insert_point
 #------------------------------------------------------------------------------------------------------------------------------------
        #spt
        text_value='SPT-N'
        #text=acad.AddText(text_value,APoint(index * scale_factor_w*distance+5*scale_factor_w, 2.5*scale_factor_h),2*scale_factor_w)
        insert_point=APoint((distance+hole_width)*scale_factor_w, Ground_EL*scale_factor_h,Ground_EL*scale_factor_h) 
        text=acad.AddText(text_value,insert_point, 0.8*scale_factor_w)
        text.Alignment=12
        text.TextAlignmentPoint = insert_point
        p1 = APoint(scale_factor_w*distance, Ground_EL)
        p2 = APoint(scale_factor_w*distance + 5, Ground_EL)
 #------------------------------------------------------------------------------------------------------------------------------------       
        #地下水位
        # row_index, col_index = np.where(df == 'G.W.L.')
        # col_index = col_index[0]
        # row_index = row_index[0]
        # next_col_index = col_index + 1
        # next_col_data = df.iloc[row_index, next_col_index]
        # GWL_point=APoint(distance*scale_factor_w-(10*scale_factor_w),-next_col_data*scale_factor_h)
        # #水位線
        # GWL_point_end=APoint(distance*scale_factor_w-(10*scale_factor_w)-(3*scale_factor_w),-next_col_data*scale_factor_h)
        # acad.AddLine(GWL_point,GWL_point_end)
        # #裝飾線
        # acad.AddLine(APoint(distance*scale_factor_w-(10*scale_factor_w)-(0.8*scale_factor_w),-next_col_data*scale_factor_h-0.2*scale_factor_h),
        #              APoint(distance*scale_factor_w-(10*scale_factor_w)-(2.2*scale_factor_w),-next_col_data*scale_factor_h-0.2*scale_factor_h))
        # acad.AddLine(APoint(distance*scale_factor_w-(10*scale_factor_w)-(1*scale_factor_w),-next_col_data*scale_factor_h-0.4*scale_factor_h),
        #              APoint(distance*scale_factor_w-(10*scale_factor_w)-(2*scale_factor_w),-next_col_data*scale_factor_h-0.4*scale_factor_h))
        # #箭頭
        # arrow_start=APoint((GWL_point.x+GWL_point_end.x)/2,-next_col_data*scale_factor_h)
        # acad.AddLine(arrow_start,APoint(arrow_start.x+(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)))
        # acad.AddLine(arrow_start,APoint(arrow_start.x-(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)))
        # acad.AddLine(APoint(arrow_start.x+(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)),
        #              APoint(arrow_start.x-(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)))
        # alignment = 1
        # text_position = APoint(arrow_start.x-5*scale_factor_w, arrow_start.y + 0.7*scale_factor_h)
        # next_col_data_str = str(next_col_data)
        # text_value='G.W.L.'+'  '+next_col_data_str
        # text = acad.AddText(text_value, text_position,1.5*scale_factor_w)
 #------------------------------------------------------------------------------------------------------------------------------------ 
        #畫孔位
        #讀取Layer列的數字
        previous_layer = 0
        y1=Ground_EL
        for layer_index, row in df.iterrows():
            Layer = row['Layer']
            hatch_num=row['LOG']
            depth=row['Depth']
            spt_n = row['N']
            #print(max(depth))
                
            # Layer列數字迭代
            if layer_index != 0:
                
                if not pd.isna(Layer):
                    y2=y1-Layer
                    p1=APoint(distance*scale_factor_w,y1*scale_factor_h)
                    p2=APoint((distance+hole_width)*scale_factor_w,y1*scale_factor_h)
                    p3=APoint(distance*scale_factor_w,y2*scale_factor_h)
                    p4=APoint((distance+hole_width)*scale_factor_w,y2*scale_factor_h)
                    y1=y2
                
                pnts=[p1.x,p1.y,
                      p2.x,p2.y,
                      p4.x,p4.y,
                      p3.x,p3.y,
                      p1.x,p1.y]
                #---------------------------------------------------------------------------------------------------------------------
                #填充線
                if not pd.isna(hatch_num):
                    pnts = vtfloat(pnts)
                    sq = msp.AddLightWeightPolyline(pnts)
                    sq.Closed = True
                    # Convert depth to a numeric type
                    depth = pd.to_numeric(depth, errors='coerce')
                    outerLoop = []
                    outerLoop.append(sq)
                    outerLoop = vtobj(outerLoop)
                    # 将 hatch_num 转换为整数类型
                    int_hatch_num = int(hatch_num)
                    hatchobj = msp.AddHatch(1, int_hatch_num, True)
                    hatchobj.PatternScale = 2*scale_factor_w  # 设置填充线比例为 2
                    hatchobj.AppendOuterLoop(outerLoop)
                    hatchobj.Evaluate()
                    if int_hatch_num in [1, 2, 5,6,7,11]:
                        rotation_angle_degrees = 45/180*3.1415926
                        hatchobj.PatternAngle = rotation_angle_degrees
                #---------------------------------------------------------------------------------------------------------------------
                #深度迭代
                if Ground_EL>ruler_top:
                    ruler_top=Ground_EL

                if (Ground_EL-depth)<ruler_bottom:
                    ruler_bottom=(Ground_EL-depth)

                # Check if depth is not NaN
                #分層深度
                #depth
                if not pd.isna(Layer):
                    Layer_text = f"{Layer:.1f}"
                    insert_point=APoint((distance-0.5)*scale_factor_w, y2*scale_factor_h,y2*scale_factor_h)
                    text = acad.AddText(Layer_text,insert_point, 0.5*scale_factor_w)
                    text.Alignment=11  
                    text.TextAlignmentPoint = insert_point
                    nan_encountered = False
                #spt
                if not pd.isna(depth):
                    if not pd.isna(spt_n):
                        text_value=spt_n
                        insert_point=APoint((distance+hole_width+0.5)*scale_factor_w, (Ground_EL-depth)*scale_factor_h,(Ground_EL-depth)*scale_factor_h)
                        text = acad.AddText(text_value,insert_point, 0.5*scale_factor_w)
                        text.Alignment=9
                        text.TextAlignmentPoint = insert_point
                else:
                    nan_encountered = True
                    break

#ruler


ruler_bottom=round(ruler_bottom+1)
ruler_top=round(ruler_top+1)
ruler_length = round((ruler_top-ruler_bottom))
insertion_point = APoint(0, ruler_top)
#insert_end=APoint(0,ruler_top-ruler_bottom)
acad.AddLine(insertion_point*scale_factor_h, APoint(0,(ruler_top-ruler_length)*scale_factor_h))
for i in range(ruler_top, ruler_bottom,-1):
    #if i==ruler_top:
        # text=ruler_top
        # acad.AddText(text,APoint(2 * scale_factor_w+4, i * scale_factor_h-(1.5 * scale_factor_w/2)),1 * scale_factor_w)
    if i % 10 == 0:
        # 画长刻度线
        acad.AddLine(APoint(0, i * scale_factor_h), APoint(2 * scale_factor_w, i * scale_factor_h))
        text=i/10*10
        acad.AddText(text,APoint(3 * scale_factor_w+4, i * scale_factor_h-(1.5 * scale_factor_w/2)),1 * scale_factor_w)
        text.Alignment=9
        #text.TextAlignmentPoint = insert_point
    elif i % 5 == 0:
        # 画中等长度的刻度线
        acad.AddLine(APoint(0, i * scale_factor_h), APoint(1 * scale_factor_w, i * scale_factor_h))
        text=i/5*5
        acad.AddText(text,APoint(3 * scale_factor_w+4,i * scale_factor_h-(1.5 * scale_factor_w/2)),1 * scale_factor_w)
        text.Alignment=9
        #text.TextAlignmentPoint = insert_point
    else:
        # 画短刻度线
        acad.AddLine(APoint(0, i * scale_factor_h), APoint(0.5 * scale_factor_w, i * scale_factor_h))
acad.AddLine(y_start_point,y_end_point)
# 退出 Excel 應用程式
app.quit()