import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, Border, Side

fn = 'singol_hole_output.xlsx'
wb = openpyxl.load_workbook(fn)

wb.active = 0
ws = wb.active

# 合并单元格
ws.merge_cells(start_row=1, start_column=1, end_row=3, end_column=25)
ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=25)
ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=25)
ws.merge_cells(start_row=10, start_column=5, end_row=10, end_column=8)
ws.merge_cells(start_row=11, start_column=5, end_row=11, end_column=5)
ws.merge_cells(start_row=12, start_column=5, end_row=12, end_column=5)
ws.merge_cells(start_row=13, start_column=5, end_row=13, end_column=5)
ws.merge_cells(start_row=14, start_column=5, end_row=14, end_column=5)


# 插入图片
img = Image('C:\\Users\\Janet\\Desktop\\readingExcel\\單孔柱狀圖\\萬頂圖樣.png')
target_cell = 'G1'
ws.add_image(img, target_cell)
img.width, img.height = 600, 60

# 选择特定的工作表
sheet = wb['工作表1']

# 添加文本并设置加粗和对齐方式
cells_to_update = [
    ('A4', '地  質  鑽  探  及  土  壤  試  驗  一   覽  表'),
    ('A5', 'SOIL EXPLORATION AND TESTING REPORT'),
    ('A6', '工程名稱：'),
    ('A7', 'Project'),
    ('A8', '鑽孔編號：'),
    ('A9', 'Hole No.'),
    ('L6', '鑽探公司:'),
    ('L7', '座      標：'),
    ('L8', '鑽孔標高：'),
    ('L9', 'Surface Elev.'),
    ('Q8', '地下水位'),
    ('Q9', 'G. W. Depth'),
    ('T6', '地點：'),
    ('T7', 'Location'),
    ('T8', '頁次:'),
    ('T9', 'Page'),
    ('A11', '深度'),
    ('A13', 'Depth'),
    ('A14', '(M)'),
    ('B11', '柱'),
    ('B12', '狀'),
    ('B13', '圖'),
    ('C11','樣號'),
    ('C13','Sample'),
    ('C14','No.'),
    ('D11','擊數'),
    ('D12','No.of'),
    ('D13','Blows'),
    ('D14','Per ft.'),
    # ('E11','15cm'),
    # ('F11','15cm'),
    # ('G11','15cm'),

]

for cell, value in cells_to_update:
    sheet[cell] = value
    sheet[cell].font = Font(bold=True)
    sheet[cell].alignment = Alignment(horizontal='center', vertical='center')


border = Border(bottom=Side(style='medium'))

# 设置底部边框
for col in range(1, 26):  
    cell = sheet.cell(row=9, column=col)
    cell.border = border

wb.save(fn)  # 保存文件
print('done')
