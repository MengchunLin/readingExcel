import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

# 创建一个新的工作簿
wb = Workbook()

# 获取当前活动的工作表
ws = wb.active

# 设置工作表名称
ws.title = "工作表1"

def adjust_column_width(ws):
    ws.column_dimensions['A'].width = 6.11
    ws.column_dimensions['B'].width = 0.5

def setup_worksheet(start_row):
    # 合并单元格
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+2, end_column=24)
    ws.merge_cells(start_row=start_row+3, start_column=1, end_row=start_row+3, end_column=24)
    ws.merge_cells(start_row=start_row+4, start_column=1, end_row=start_row+4, end_column=24)
    ws.merge_cells(start_row=start_row+9, start_column=5, end_row=start_row+9, end_column=8)
    ws.merge_cells(start_row=start_row+10, start_column=5, end_row=start_row+10, end_column=8)
    ws.merge_cells(start_row=start_row+11, start_column=5, end_row=start_row+11, end_column=8)
    ws.merge_cells(start_row=start_row+12, start_column=5, end_row=start_row+12, end_column=8)
    ws.merge_cells(start_row=start_row+13, start_column=5, end_row=start_row+13, end_column=8)
    ws.merge_cells(start_row=start_row+14, start_column=5, end_row=start_row+15, end_column=5)
    ws.merge_cells(start_row=start_row+14, start_column=6, end_row=start_row+15, end_column=6)
    ws.merge_cells(start_row=start_row+14, start_column=7, end_row=start_row+15, end_column=7)
    ws.merge_cells(start_row=start_row+14, start_column=8, end_row=start_row+15, end_column=8)
    ws.merge_cells(start_row=start_row+9, start_column=11, end_row=start_row+10, end_column=14)
    ws.merge_cells(start_row=start_row+11, start_column=11, end_row=start_row+11, end_column=14)
    ws.merge_cells(start_row=start_row+9, start_column=22, end_row=start_row+9, end_column=23)
    ws.merge_cells(start_row=start_row+11, start_column=22, end_row=start_row+11, end_column=23)
    ws.merge_cells(start_row=start_row+12, start_column=22, end_row=start_row+12, end_column=23)

    # 插入图片
    img = Image('萬頂圖樣.png')
    target_cell = f'G{start_row}'
    ws.add_image(img, target_cell)
    img.width, img.height = 550, 60

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
        ('C11', '柱'),
        ('C12', '狀'),
        ('C13', '圖'),
        ('C14', 'Log.'),
        ('D11', '樣號'),
        ('D13', 'Sample'),
        ('D14', 'No.'),
        ('E11', '擊數'),
        ('E12', 'No.of'),
        ('E13', 'Blows'),
        ('E14', 'Per ft.'),
        ('E15', '15cm'),
        ('F15', '15cm'),
        ('G15', '15cm'),
        ('H15', 'N值'),
        ('I11', '地質說明'),
        ('I13', 'Soil Description'),
        ('J11', '分類'),
        ('J13', 'USCS'),
        ('J14', 'Classi-'),
        ('J15', 'fication'),
        ('K10','顆粒分析'),
        ('K12','Grain Size Analysis(%)'),
        ('K13','礫石'),
        ('K16','Gravel'),
        ('L13','砂'),
        ('L16','Sand'),
        ('M13','粉土'),
        ('M16','Silt'),
        ('N13','粘土'),
        ('N16','Clay'),
        ('O10','自然'),
        ('O11','含水量'),
        ('O13','Water'),
        ('O14','Content'),
        ('O16','W(%)'),
        ('P10','比重'),
        ('P13','Specific'),
        ('P14','Gravity'),
        ('P16','G'),
        ('Q10','當地'),
        ('Q11','密度'),
        ('Q13','Density'),
        ('Q14','γt'),
        ('Q16','T/M³'),
        ('R10','空隙比'),
        ('R13','Void'),
        ('R14','Ratio'),
        ('R16','e'),
        ('S10','塑性'),
        ('S11','指數'),
        ('S13','Liquid'),
        ('S14','Limit'),
        ('S16', 'WL(%)'),
        ('T10','塑性'),
        ('T11','指數'),
        ('T13','Plastic'),
        ('T14','Limit'),
        ('T16','IP(%)'),
        ('U10','單軸壓'),
        ('U11','縮強度'),
        ('U12','Uniaxial'),
        ('U13','Comp.'),
        ('U14','Strength'),
        ('U15','qu'),
        ('U16','(kgf/cm²)'),
        ('V10','強 度 參 數'),
        ('V12','Shear Strength'),
        ('V13','Parameter'),
        ('V15','ψ'),
        ('V16','(Degree)'),
        ('W15',"C'"),
        ('W16','(kgf/cm²)'),
        ('X10','岩石品'),
        ('X11','質指標'),
        ('X12','Rock'),
        ('X13','Quality'),
        ('X14','Design-'),
        ('X15','ation'),
        ('X16','R.Q.D.(%)'),
    ]

    def set_cell_style(cell, align='center'):
        cell.font = Font(name='新細明體', size=12, bold=True)
        cell.font = Font(name='Times New Roman', size=12, bold=True)
        cell.alignment = Alignment(horizontal=align, vertical='center')

    # 在 cells_to_update 循環中使用：
    for cell, value in cells_to_update:
        actual_cell = ws[cell[0] + str(int(cell[1:]) + start_row - 1)]
        actual_cell.value = value
        set_cell_style(actual_cell)
        border = Border(bottom=Side(style='medium'))


    # 设置底部边框
    for col in range(1, 25):
        cell = ws.cell(row=start_row + 8, column=col)
        cell.border = Border(bottom=Side(border_style='medium'))
        cell = ws.cell(row=start_row + 15, column=col)
        cell.border = Border(bottom=Side(border_style='thin'))
        cell = ws.cell(row=start_row + 45, column=col)
        cell.border = Border(bottom=Side(border_style='medium'))

    # 设置右侧边框
    for row in range(start_row+9, start_row + 46):
        cell = ws.cell(row=row, column=1)
        cell.border = Border(left=Side(border_style='medium'))
        cell = ws.cell(row=row, column=24)
        cell.border = Border(right=Side(border_style='medium'))

    #調整欄寬
    adjust_column_width(ws)

# 每46行执行一次
for i in range(0, 5):  # 可以调整范围以覆盖你需要的行数
    setup_worksheet(i * 46 + 1)

# 保存文件
fn = 'new_singol_hole_output.xlsx'
wb.save(fn)
print('done')
