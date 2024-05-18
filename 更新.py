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
from collections import Counter

file_path = ""
Upper_depth = 5
Lower_depth = 0
distance=25
hole_width=2
Ground_EL=0
num_lists = 0
lists = []
pre_num=0
example_list=[]
example_list_int=[]
#ruler
ruler_top=0
ruler_bottom=0

#放大倍數
scale_factor_h=0
scale_factor_w=0

#dictionary
# dictionary={'0.0':'回填土',
#             '1.0':'良好、不良級配卵礫石、砂礫石',
#             '2.0':'良好、不良級配卵礫石、砂礫石',
#             '3.0':'良好、不良級配卵礫石、砂礫石',
#             '4.0':'良好、不良級配砂土、礫砂土',
#             '5.0':'粉質砂土',
#             '6.0':'黏質砂土',
#             '7.0':'低至中等塑性無機黏土、低塑性粉質黏土',
#             '8.0':'黏質粉土、有機黏土',
#             '9.0':'高塑性無機黏土、中等至高塑性有機黏土',
#             '10.0':'無機質粉土',
#             '11.0':'砂質粉土',
#             '12.0':'安山岩',
#             '13.0':'卵礫石',
#             '14.0':'砂岩',
#             '15.0':'碎屑岩',
#             '16.0':'粉砂岩',
#             '17.0':'泥岩',
#             '18.0':'頁岩',
#             '19.0':'凝灰岩',
#             '20.0':'火山碎屑',
#             '21.0':'崩積層'
#             }
dictionary={'0':'回填土',
            '1':'良好、不良級配卵礫'+'\n'+'石、砂礫石',
            '2':'良好、不良級配卵石'+'\n'+'、砂礫石',
            '3':'良好、不良級配卵礫'+'\n'+'石、砂礫石',
            '4':'良好、不良級配砂土'+'\n'+'、礫砂土',
            '5':'粉質砂土',
            '6':'黏質砂土',
            '7':'低至中等塑性無機黏'+'\n'+'土、低塑性粉質黏土',
            '8':'黏質粉土、有機黏土',
            '9':'高塑性無機黏土、中等至高塑性有機黏土',
            '10':'無機質粉土',
            '11':'砂質粉土',
            '12':'安山岩',
            '13':'卵礫石',
            '14':'砂岩',
            '15':'碎屑岩',
            '16':'粉砂岩',
            '17':'泥岩',
            '18':'頁岩',
            '19':'凝灰岩',
            '20':'火山碎屑',
            '21':'崩積層'
            }

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
    num_lists+=1
    if index != 0:
 #------------------------------------------------------------------------------------------------------------------------------------
        #孔頂高
        df = pd.read_excel(xl, sheet_name)
        row_index, col_index = np.where(df == 'Ground EL')
        col_index = col_index[0]
        row_index = row_index[0]
        next_col_index = col_index + 1
        Ground_EL= df.iloc[row_index, next_col_index]
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
        row_index, col_index = np.where(df == 'G.W.L.')
        col_index = col_index[0]
        row_index = row_index[0]
        next_col_index = col_index + 1
        GWL = df.iloc[row_index, next_col_index]
        GWL_point=APoint(distance*scale_factor_w-(6*scale_factor_w),GWL*scale_factor_h)
        #水位線
        GWL_point_end=APoint(distance*scale_factor_w-(8*scale_factor_w),GWL*scale_factor_h)
        acad.AddLine(GWL_point,GWL_point_end)
        #裝飾線
        acad.AddLine(APoint(distance*scale_factor_w-(6.5*scale_factor_w),GWL*scale_factor_h-0.2*scale_factor_h),
                     APoint(distance*scale_factor_w-(7.5*scale_factor_w),GWL*scale_factor_h-0.2*scale_factor_h))
        acad.AddLine(APoint(distance*scale_factor_w-(6.6*scale_factor_w),GWL*scale_factor_h-0.4*scale_factor_h),
                     APoint(distance*scale_factor_w-(7.4*scale_factor_w),GWL*scale_factor_h-0.4*scale_factor_h))
        #箭頭
        arrow_start=APoint((GWL_point.x+GWL_point_end.x)/2,GWL*scale_factor_h)
        acad.AddLine(arrow_start,APoint(arrow_start.x+(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)))
        acad.AddLine(arrow_start,APoint(arrow_start.x-(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)))
        acad.AddLine(APoint(arrow_start.x+(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)),
                     APoint(arrow_start.x-(0.5*scale_factor_h/pow(3,0.5)),arrow_start.y+(0.5*scale_factor_h)))
        next_col_data_str = str(GWL)
        text='G.W.L.'+'  '+next_col_data_str
        insert_point=APoint((GWL_point.x+GWL_point_end.x)/2,arrow_start.y+(0.5*scale_factor_h))
        text=acad.AddText(text,insert_point, 0.6*scale_factor_w)
        text.Alignment=13
        text.TextAlignmentPoint = insert_point
 #------------------------------------------------------------------------------------------------------------------------------------ 
        #畫孔位
        #讀取Layer列的數字
        previous_layer = 0
        y1=Ground_EL
        depest=0
        num_element=0
        for layer_index, row in df.iterrows():
            Layer = row['Layer']
            hatch_num=row['LOG']
            depth=row['Depth']
            spt_n = row['N']

            
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
                    if int_hatch_num in [1, 2, 5,6,7,11,21]:
                        rotation_angle_degrees = 45/180*3.1415926
                        hatchobj.PatternAngle = rotation_angle_degrees
                    example_list.append(hatch_num)
                #---------------------------------------------------------------------------------------------------------------------
                #深度迭代
                if Ground_EL>ruler_top:
                    ruler_top=Ground_EL

                if depest<ruler_bottom:
                    ruler_bottom=depest

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
                    if Ground_EL-depth<depest:
                        depest=Ground_EL-depth
                    if not pd.isna(spt_n):
                        text_value=spt_n
                        insert_point=APoint((distance+hole_width+0.5)*scale_factor_w, (Ground_EL-depth)*scale_factor_h,(Ground_EL-depth)*scale_factor_h)
                        text = acad.AddText(text_value,insert_point, 0.5*scale_factor_w)
                        text.Alignment=9
                        text.TextAlignmentPoint = insert_point
                else:
                    nan_encountered = True
                    break

        acad.AddLine(APoint(distance*scale_factor_w,Ground_EL*scale_factor_h),APoint((distance+hole_width)*scale_factor_w,Ground_EL*scale_factor_h))
        acad.AddLine(APoint((distance+hole_width)*scale_factor_w,Ground_EL*scale_factor_h),APoint((distance+hole_width)*scale_factor_w,depest*scale_factor_h))
        acad.AddLine(APoint((distance+hole_width)*scale_factor_w,depest*scale_factor_h),APoint(distance*scale_factor_w,depest*scale_factor_h))
        acad.AddLine(APoint(distance*scale_factor_w,depest*scale_factor_h),APoint(distance*scale_factor_w,Ground_EL*scale_factor_h))
example_list=list(set(example_list))
for i in example_list:
    example_list_int.append(int(i))
print(example_list_int)
count=str(len(example_list_int))
#num=len(example_list_int)
#print(dictionary[example_list[1]])
# 將下面這段程式碼加入到你的程式中，用來印出字典中對應的值
n=0
i=0
square_h=1.5
text='圖例:'

text=acad.AddText(text,APoint(-30*scale_factor_w,0),1*scale_factor_w)
text.Alignment=6 #TopLeft
text.TextAlignmentPoint = APoint(-30*scale_factor_w,0)

#字
for i in example_list_int:
    # print(dictionary[str(i)])
    n+=2
    text=dictionary[str(i)]
    insert_point=APoint(-30*scale_factor_w,-n*scale_factor_w)
    print(insert_point)
    print(text)
    text=acad.AddText(text,insert_point,1*scale_factor_w)
    text.Alignment=6 #TopLeft
    text.TextAlignmentPoint = insert_point
    i+=1

print(example_list_int)
#圖例框
n=0
p1=APoint(-32*scale_factor_w,n*-2*scale_factor_w)
p2=APoint(-30.5*scale_factor_w,n*-2*scale_factor_w)
p3=APoint(-30.5*scale_factor_w,n*(-2-2)*scale_factor_w)
p4=APoint(-32*scale_factor_w,n*(-2-2)*scale_factor_w)
pnts=[p1.x,p1.y,
      p2.x,p2.y,
      p3.x,p3.y,
      p4.x,p4.y,
      p1.x,p1.y]
pnts = vtfloat(pnts)

for i in (example_list_int):
    square_h+=2
    n+=1
    sq = msp.AddLightWeightPolyline(pnts)
    sq.Closed = True
    depth = pd.to_numeric(depth, errors='coerce')
    outerLoop = []
    outerLoop.append(sq)
    outerLoop = vtobj(outerLoop)
    hatchobj = msp.AddHatch(1, i, True)
    hatchobj.PatternScale = 2*scale_factor_w  # 设置填充线比例为 2
    hatchobj.AppendOuterLoop(outerLoop)
    hatchobj.Evaluate()
    print(n)

    

#土層紀錄------------------------------------------------------------------------------------------------------------------------------------------------------
all_lists = []  # 創建一個空列表來存儲所有創建的列表
for index, sheet_name in enumerate(sheet_names):
    if index != 0:
        new_list = []  # 在每次迴圈開始時創建一個新列表
        for layer_index, row in df.iterrows():
            df = pd.read_excel(xl, sheet_name)
            LOG = row['LOG']
            if layer_index != 0:
                if not pd.isna(LOG):
                    new_list.append(LOG)  # 將每個 LOG 添加到新列表中
        all_lists.append(new_list)  # 將新列表添加到 all_lists 中

# 在迴圈外部檢視結果
for i, sheet_name in enumerate(sheet_names):
    if i != 0:
        #print(sheet_name + ": " + str(all_lists[i-1]))  # 輸出 sheet_name 和相應的列表
        continue
#---------------------------------------------------------------------------------------------------------------------------------------------------------------

ruler_bottom=round(ruler_bottom-1)
ruler_top=round(ruler_top+1)
ruler_length = round((ruler_top-ruler_bottom))
insertion_point = APoint(0, ruler_top)
#insert_end=APoint(0,ruler_top-ruler_bottom)
acad.AddLine(insertion_point*scale_factor_h, APoint(0,(ruler_top-ruler_length)*scale_factor_h))
for i in range(ruler_top, ruler_bottom,-1):

    if i % 10 == 0:
        # 画长刻度线
        acad.AddLine(APoint(0, i * scale_factor_h), APoint(2 * scale_factor_w, i * scale_factor_h))
        text=i/10*10
        insert_point=APoint(3 * scale_factor_w+4, i * scale_factor_h)
        text=acad.AddText(str(text),insert_point,0.7 * scale_factor_w)
        text.Alignment=9
        text.TextAlignmentPoint = insert_point
    elif i % 5 == 0:
        # 画中等长度的刻度线
        acad.AddLine(APoint(0, i * scale_factor_h), APoint(1 * scale_factor_w, i * scale_factor_h))
        text=i/5*5
        insert_point=APoint(3 * scale_factor_w+4,i * scale_factor_h)
        text=acad.AddText(str(text),insert_point,0.7 * scale_factor_w)
        text.Alignment=9
        text.TextAlignmentPoint = insert_point
    else:
        # 画短刻度线
        acad.AddLine(APoint(0, i * scale_factor_h), APoint(0.5 * scale_factor_w, i * scale_factor_h))
acad.AddLine(y_start_point,y_end_point)