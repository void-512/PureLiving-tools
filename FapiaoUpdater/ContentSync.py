import sys
import pandas as pd
import configparser
from openpyxl import load_workbook

cfg = 'FapiaoUpdaterConfig.cfg'

configObj = configparser.ConfigParser()
configObj.read(cfg)
pivot_src = configObj.get('src', 'pivot')
src = configObj.get('src', 'src')
src_sheet = configObj.get('src', 'sheet')
col_to_copy = configObj.get('src', 'cols_to_copy').split()

pivot_dest = configObj.get('dest', 'pivot')
dest = configObj.get('dest', 'dest')
dest_sheet = configObj.get('dest', 'sheet')

def write_df(df, pivot_val, col, msg):
    if col not in df.columns:
        df[col] = None
    df.loc[df[pivot_dest] == pivot_val, col] = msg

df_src = pd.read_excel(src, sheet_name=src_sheet)
df_dest = pd.read_excel(dest, sheet_name=dest_sheet)

for target_col in col_to_copy:
    if target_col not in df_src.columns:
        sys.exit(f'Column {target_col} does not exist')
    for index, row in df_src.iterrows():
        if not pd.isnull(row[target_col]):
            write_df(df_dest, row[pivot_src], target_col, row[target_col])

df_dest.to_excel('Content Updated.xlsx', index=False)