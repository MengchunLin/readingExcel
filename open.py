import chardet

# 打开文件并自动检测编码
with open('C:\Civilight_AutoLog.zip_1\Civilight_AutoLog\CVLTAL.fas', 'rb') as f:
    result = chardet.detect(f.read())
    encoding = result['encoding']

# 以指定编码读取文件内容
with open('C:\Civilight_AutoLog.zip_1\Civilight_AutoLog\CVLTAL.fas', encoding=encoding) as f:
    content = f.read()