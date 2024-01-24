import numpy as np
import pandas as pd
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
import pathlib

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

#print("檔案路徑:", file_path)

# 檢查檔案路徑是否不為空
if file_path:
    # 建立新的文字檔
    with open('data.txt', mode='w') as f:
        # 開啟工作簿
        wb = app.books.open(file_path)

        # 寫入文字檔
        #指標
        f.write('0\n')
        #;鑽孔座標表的左上角插入點座標(x,y)，若指標不是 1，本行亂填都可以
        f.write('0.00 0.00\n')

        #;總共要繪出的孔數，不可為 0 柱狀圖鑽孔寬度(m)
        folder_path = pathlib.Path(file_path).parent
        file_count = len(list(folder_path.glob('*')))
        
        f.write(str(file_count)+' ')
        f.write('1.0\n')

        #;圖面上的插入點座標(x,y)，為第一個輸入的鑽孔柱狀圖的左上基準點
        f.write('0.00'+' '+'0.00'+'\n')

        # 鑽孔編號
        sheet = wb.sheets[1]
        f.write(str(sheet.range('C1').value)+' '+'1'+'\n')

        # 鑽孔座標 (E / N)
        E_value=float(sheet.range('C6').value)
        N_value=float(sheet.range('C7').value)
        f.write(str('%.2f'%E_value)+'\n')
        f.write(str('%.2f'%N_value)+'\n')
        #f.write(str(sheet.range('C7').value) + '\n')

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
            df = pd.read_excel(file_path, sheet_name=8)
            num = row['地質圖元代碼']
            if not pd.isna(num):
                Upper_depth=row['上限深度']
                Lower_depth=row['下限深度']
                sheet = wb.sheets[22]
                row_index = None
                for i, value in enumerate(sheet.range('A:A').value):
                    if value == num:
                        row_index = i + 1
                        if row_index is not None:
                            value_in_row_B=sheet.range(f'B{row_index}').value
                            int_row_B=int(value_in_row_B)
                            f.write(str(Lower_depth) + ' ' + '%.0f' % value_in_row_B + '\n')




        #鑽孔 N 數(若無可填 0，以下便留白)
        df = pd.read_excel(file_path, sheet_name=7)
        row_count=df.shape[0]
        if row_count!=0:
            f.write(str(row_count) + '\n')
        else:
            f.write('0')

        #GL- SPT-N
        Upper_depth = df['上限深度']  #A
        Lower_depth = df['下限深度']  #B
        first_value = df['標準貫入N1值']  #C
        second_value = df['標準貫入N2值'] #D
        third_value = df['標準貫入N3值']  #E
        from decimal import Decimal, ROUND_HALF_UP

        def round_v3(num, decimal):
            str_deci = 1
            for _ in range(decimal):
                str_deci = str_deci / 10
            str_deci = str(str_deci)
            result = Decimal(str(num)).quantize(Decimal(str_deci), rounding=ROUND_HALF_UP)
            result = float(result)

            return result        
        for A, B, C, D, E in zip(Upper_depth, Lower_depth, first_value, second_value, third_value):
            avg1 = (A + B) / 2
            avg2 = C + D + E
            f.write(str(round_v3(avg1,2)) + ' ' + str(avg2) + '\n')
        #f.write('end SPT\n')
            
        #鑽孔 RQD 數(若無可填 0，以下便留白)
        df = pd.read_excel(file_path, sheet_name=9)
        row_count=df.shape[0]
        if row_count!=0:
            f.write(str(row_count) + '\n')
        else:
            f.write('0'+'\n')
        #f.write('end RQD count'+'\n')

        #GL- RQD
        df = pd.read_excel(file_path, sheet_name=9)
        Upper_depth1 = df['上限深度']  #A
        Lower_depth1 = df['下限深度']  #B
        RQD=df['岩石RQD值'] #C
        for A,B,C in zip(Upper_depth1,Lower_depth1,RQD):
            avg = (A + B) / 2
            float_avg=float(avg)
            f.write("{:.2f}".format(avg)+' '+str(C))
       
        # #鑽孔 USCS 數(若無可填 0，以下便留白)
        df = pd.read_excel(file_path, sheet_name=8)
        row_count=df.shape[0]
        if row_count!=0:
            f.write(str(row_count) + '\n')
        else:
            f.write('0'+'\n')

        #GL- USCS
        df=pd.read_excel(file_path,sheet_name=8)
        for index, row in df.iterrows():
            num = row['地質圖元代碼']
            depth_1=row['上限深度']
            depth_2=row['下限深度']
            avg=(depth_1+depth_2)/2
            float_avg=float(avg)
            if not pd.isna(num):
                sheet = wb.sheets[22]
                row_index = None
                for i, value in enumerate(sheet.range('A:A').value):
                    if value == num:
                        row_index = i + 2
                        if row_index is not None:
                            value_in_row_C=sheet.range(f'C{row_index}').value
                            #print(value_in_row_C)
                            #f.write({"{:.2f}".format(avg)} {value_in_row_C}+"\n")
                            f.write("{:.2f} {}".format(avg, value_in_row_C) + "\n")

        # #鑽孔 Smaple 數(若無可填 0，以下便留白)
        df = pd.read_excel(file_path, sheet_name=6)
        row_count=df.shape[0]
        if row_count!=0:
            f.write(str(row_count) + '\n')
        else:
            f.write('0'+'\n')
        #取樣上限 GL- 取樣下限 GL- 樣號
       
        for index, row in df.iterrows():
            Upper_depth = row['上限深度']  #A
            Lower_depth = row['下限深度']  #B
            sample=row['取樣編號']   #C
            if row_count!=0:
                # float_Upper_depth=float(Upper_depth)
                # float_Lower_depth=float(Lower_depth)
                #str_sample=str(sample)
                #f.write(str(round(Upper_depth,2)),str(round(Lower_depth,2))+' '+str(sample)+'\n')
                f.write(str('%.2f'%Upper_depth) + ' ' + str('%.2f'%Lower_depth) + ' ' + sample + '\n')

            else:
                print('no data')

        #鑽孔 W(LL/PI)數(若無可填 0，以下便留白)
        df = pd.read_excel(file_path, sheet_name=5)
        row_count=df.shape[0]
        if row_count!=0:
            f.write(str(row_count) + '\n')
        else:
            f.write('0'+'\n')       
        #GL- W LL PI
        for index, row in df.iterrows():
            Upper_depth=row['上限深度']
            Lower_depth=row['下限深度']
            LL=row['用水量']
            PL=row["迴水率"]
            f.write(str(Upper_depth)+' '+str(Lower_depth)+' '+str(LL)+' '+str(PL))

        # 關閉工作簿
        wb.close()

    # 退出 Excel 應用程式
    app.quit()
else:
    print("未選擇檔案。")