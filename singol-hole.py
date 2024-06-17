import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment

fn = 'singol_hole_output.xlsx'
wb = openpyxl.load_workbook(fn)

wb.active = 0
ws = wb.active

ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=25)
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=25)
ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=25)
img = Image('C:\\Users\\Janet\\Desktop\\readingExcel\\單孔柱狀圖\\萬頂圖樣.png')
target_cell = 'G1'
ws.add_image(img, target_cell)
img.width, img.height = 600, 60

# 选择特定的工作表
sheet = wb['工作表1']
sheet['A4'] = '地  質  鑽  探  及  土  壤  試  驗  一   覽  表'
sheet.alignment = Alignment(horizontal='center', vertical='center')
sheet['A5'] = 'SOIL EXPLORATION AND TESTING REPORT'

wb.save(fn)  # 保存文件，如果给不同文件名则相当于另存为
print('done')
