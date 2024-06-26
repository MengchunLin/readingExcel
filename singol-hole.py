import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, Border, Side
import pandas as pd

# Load the existing Excel file
xl = pd.ExcelFile('七孔測試.xlsx')
sheet_names = xl.sheet_names
sheet_names.remove('工作表1')

# Create a new workbook and remove the default sheet
new_wb = Workbook()
new_wb.remove(new_wb.active)

# Define border style
border_style = Border(
    right=Side(border_style='thin'),
    bottom=Side(border_style='thin'),
    top=Side(border_style='medium'),
    left=Side(border_style='thin')
)

# Create corresponding sheets in the new workbook
for sheet_name in sheet_names:
    new_wb.create_sheet(title=sheet_name)

def adjust_column_width(ws):
    ws.column_dimensions['A'].width = 6.11
    ws.column_dimensions['B'].width = 1

def setup_worksheet(ws, start_row):
    # Merge cells
    merge_cells_instructions = [
        (start_row, 1, start_row+2, 24),
        (start_row+3, 1, start_row+3, 24),
        (start_row+4, 1, start_row+4, 24),
        (start_row+9, 5, start_row+9, 8),
        (start_row+10, 5, start_row+10, 8),
        (start_row+11, 5, start_row+11, 8),
        (start_row+12, 5, start_row+12, 8),
        (start_row+13, 5, start_row+13, 8),
        (start_row+14, 5, start_row+15, 5),
        (start_row+14, 6, start_row+15, 6),
        (start_row+14, 7, start_row+15, 7),
        (start_row+14, 8, start_row+15, 8),
        (start_row+9, 11, start_row+10, 14),
        (start_row+11, 11, start_row+11, 14),
        (start_row+9, 22, start_row+9, 23),
        (start_row+11, 22, start_row+11, 23),
        (start_row+12, 22, start_row+12, 23),
        (start_row+10, 22, start_row+10, 23),
        (start_row+13, 22, start_row+13, 23)
    ]
    for merge in merge_cells_instructions:
        ws.merge_cells(start_row=merge[0], start_column=merge[1], end_row=merge[2], end_column=merge[3])
    
    # Insert image
    img = Image('萬頂圖樣.png')
    img.width, img.height = 550, 60
    target_cell = f'G{start_row}'
    ws.add_image(img, target_cell)

    # Add text and set styles
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
        ('K10', '顆粒分析'),
        ('K12', 'Grain Size Analysis(%)'),
        ('K13', '礫石'),
        ('K16', 'Gravel'),
        ('L13', '砂'),
        ('L16', 'Sand'),
        ('M13', '粉土'),
        ('M16', 'Silt'),
        ('N13', '粘土'),
        ('N16', 'Clay'),
        ('O10', '自然'),
        ('O11', '含水量'),
        ('O13', 'Water'),
        ('O14', 'Content'),
        ('O16', 'W(%)'),
        ('P10', '比重'),
        ('P13', 'Specific'),
        ('P14', 'Gravity'),
        ('P16', 'G'),
        ('Q10', '當地'),
        ('Q11', '密度'),
        ('Q13', 'Density'),
        ('Q14', 'γt'),
        ('Q16', 'T/M³'),
        ('R10', '空隙比'),
        ('R13', 'Void'),
        ('R14', 'Ratio'),
        ('R16', 'e'),
        ('S10', '塑性'),
        ('S11', '指數'),
        ('S13', 'Liquid'),
        ('S14', 'Limit'),
        ('S16', 'WL(%)'),
        ('T10', '塑性'),
        ('T11', '指數'),
        ('T13', 'Plastic'),
        ('T14', 'Limit'),
        ('T16', 'IP(%)'),
        ('U10', '單軸壓'),
        ('U11', '縮強度'),
        ('U12', 'Uniaxial'),
        ('U13', 'Comp.'),
        ('U14', 'Strength'),
        ('U15', 'qu'),
        ('U16', '(kgf/cm²)'),
        ('V10', '強 度 參 數'),
        ('V12', 'Shear Strength'),
        ('V13', 'Parameter'),
        ('V15', 'ψ'),
        ('V16', '(Degree)'),
        ('W15', "C'"),
        ('W16', '(kgf/cm²)'),
        ('X10', '岩石品'),
        ('X11', '質指標'),
        ('X12', 'Rock'),
        ('X13', 'Quality'),
        ('X14', 'Design-'),
        ('X15', 'ation'),
        ('X16', 'R.Q.D.(%)'),
    ]

    def set_cell_style(cell, align='center'):
        cell.font = Font(name='Times New Roman', size=12, bold=True)
        cell.alignment = Alignment(horizontal=align, vertical='center')

    for cell, value in cells_to_update:
        actual_cell = ws[cell[0] + str(int(cell[1:]) + start_row - 1)]
        actual_cell.value = value
        set_cell_style(actual_cell)

    # Set borders
    for col in range(1, 25):
        ws.cell(row=start_row + 8, column=col).border = Border(bottom=Side(border_style='medium'))
        ws.cell(row=start_row + 46, column=col).border = Border(top=Side(border_style='medium'))

    for col in range(1, 24):
        for i in range(17, 47):
         ws.cell(row=i, column=col).border = Border(right=Side(border_style='thin'))

    # Set left and right borders
    for row in range(start_row + 16, start_row + 46):
        ws.cell(row=row, column=1).border = Border(left=Side(border_style='medium'))
        ws.cell(row=row, column=24).border = Border(right=Side(border_style='medium'))

    for row in ws['A10:A16']:
        for cell in row:
            cell.border = Border(left=Side(border_style='medium'),
                                 right=Side(border_style='thin'),
                                 top=Side(border_style='medium'),
                                 bottom=Side(border_style='thin'))


    # Adjust column widths
    adjust_column_width(ws)

# Process each worksheet
for sheet_name in sheet_names:
    df = pd.read_excel('七孔測試.xlsx', sheet_name=sheet_name)
    Layer = df.iloc[:, 20][5:].dropna().tolist()
    hatch_num = df.iloc[:, 21][5:].dropna().tolist()

    ws = new_wb[sheet_name]
    for i in range(len(sheet_names)):  # Adjust the range as needed
        setup_worksheet(ws, i * 46 + 1)

# Save the new workbook
fn = 'new_singol_hole_output.xlsx'
new_wb.save(fn)
print('done')
