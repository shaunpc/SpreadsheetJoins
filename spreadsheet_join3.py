import pandas as pd

# Notes: need to install from Settings/Interpreter packages
# Need  xlrd  and openpyxl
# --

# Load the input spreadsheets
df_left = pd.read_excel(r'ss_left.xlsx')
df_middle = pd.read_excel(r'ss_middle.xlsx')
df_right = pd.read_excel(r'ss_right.xlsx')
print(df_left)
print(df_middle)

# Both need to share the "Key" column
df_lm = pd.merge(df_left, df_middle, how='left', on='Key')
print(df_lm)

# Both need to share the "App" column
df_rm = pd.merge(df_lm, df_right, how='left', on='App')
print(df_rm)

df_rm.to_excel('ss_merged.xlsx', sheet_name='merged', index=False)

