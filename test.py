from openpyxl import Workbook

wb = Workbook()
ws = wb.active

cell = ws['S16']

# Unicode 下標字符
subscript_l = '\u2097'  # Unicode for subscript 'l'

cell.value = f'W{subscript_l}(%)'

wb.save('example.xlsx')