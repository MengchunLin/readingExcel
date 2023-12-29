import pandas as pd

# 指定Excel文件路径
excel_file_path = 'C:/Geo2010/TEMP/335065091060815B/05NT-LZ001.xls'

# 使用pandas的read_excel函数读取Excel文件
df = pd.read_excel(excel_file_path)

# 打印读取的数据
print(df)
