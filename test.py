import pandas as pd

df = pd.read_excel('對照表.xlsx')
num = 112

# Check if the 'num' value exists in the 'code' column
if (df['code'] == num).any():
    # Generate a dictionary for the specific row with only the 'ASCII' column
    ASCII = {'ASCII': df.loc[df['code'] == num, 'ASCII'].values[0]}
    
    print(f'{ASCII["ASCII"]}')
else:
    print(f'{num} not found in the specified column.')
