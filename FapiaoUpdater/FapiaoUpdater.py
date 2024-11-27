'''
    This program updates desired columns in the Fapiao table.
    For weekly task only, not compatible with other files.
'''

import sys
import pandas as pd
import configparser

cfg = 'FapiaoUpdaterConfig.cfg'

configObj = configparser.ConfigParser()
configObj.read(cfg)
pivot_src = configObj.get('src', 'pivot')
src = configObj.get('src', 'src')
src_sheet = configObj.get('src', 'sheet')
start_row_src = configObj.getint('src', 'skip_rows')
col_to_copy = configObj.get('src', 'cols_to_copy').split()

pivot_dest = configObj.get('dest', 'pivot')
dest = configObj.get('dest', 'dest')
dest_sheet = configObj.get('dest', 'sheet')
start_row_dest = configObj.getint('dest', 'skip_rows')

def write_df(df, pivot_val, col, msg):
    if col not in df.columns:
        df[col] = None
    df.loc[df[pivot_dest] == pivot_val, col] = msg

df_src = pd.read_excel(src, sheet_name=src_sheet, skiprows=start_row_src)
df_dest = pd.read_excel(dest, sheet_name=dest_sheet, skiprows=start_row_dest)

for target_col in col_to_copy:
    if target_col not in df_src.columns:
        sys.exit(f'Column {target_col} does not exist')
    for index, row in df_src.iterrows():
        if not pd.isnull(row[target_col]):
            write_df(df_dest, row[pivot_src], target_col, row[target_col])

df_dest.to_excel('NewExcel.xlsx', index=False)