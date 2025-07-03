import sys
import io
import tempfile
import os
import shutil
import pandas as pd

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

paths = []

if len(sys.argv) > 1:
    from glob import glob
    for arg in sys.argv[1:]:
        paths += glob(arg)
        if arg == '|':
            break
else:
    print('Usage: readxl <filename>.xlsx')
    exit(0)

xlFiles = [f for f in paths if str(f).endswith('.xlsx')]

def load_open_workbook(source_file_path):
    with tempfile.TemporaryDirectory() as temp_dir:
        # Copy a file into the temp dir
        dest_file_path = os.path.join(temp_dir, os.path.basename(source_file_path))
        shutil.copy2(source_file_path, dest_file_path)
        # Use the file at runtime
        wb_data = pd.read_excel(dest_file_path, sheet_name=None, dtype=str)
    return wb_data

for filepath in xlFiles:
    try:
        # os.
        df_dict = load_open_workbook(filepath)
    except:
        print(f'Error reading {filepath}')
        continue

    if len(paths) == 1:
        if '*' in paths[0]:
            prefix1 = ''
        else:
            prefix1 = f'{filepath}:'
    else:
        prefix1 = f'{filepath}:'
    
    df_str = ''
    for sheet, df in df_dict.items():
        if len(df_dict)==1:
            prefix2 = ''
        else:
            prefix2 = f'{sheet}:'

        df = df.fillna('')
        df.index = df.index+1
        for idx, row in df.iterrows():
            row_str = ','.join(row.tolist())
            row_str = f'{prefix1}{prefix2}{idx}:{row_str}'
            print(row_str)
