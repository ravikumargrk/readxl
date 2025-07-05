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

if not xlFiles:
    print('No XLSX files found with pattern: ' + ' '.join(sys.argv[1:]))
    exit(0)

def runTemp(source_file_path, func, *args, **kwargs):
    with tempfile.TemporaryDirectory() as temp_dir:
        dest_file_path = os.path.join(temp_dir, os.path.basename(source_file_path))
        shutil.copy2(source_file_path, dest_file_path)
        return func(dest_file_path, *args, **kwargs)

for filepath in xlFiles:
    try:
        # os.
        df_dict = runTemp(filepath, pd.read_excel, sheet_name=None, dtype=str)
    except:
        print(f'Error reading {filepath}')
        continue

    if (len(sys.argv) == 2) and (len(paths)==1):
        if sys.argv[1] == paths[0]:
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
            row_str = ','.join(row.tolist()).replace('\n', ' ')
            row_str = f'{prefix1}{prefix2}{idx}:{row_str}'
            print(row_str)
