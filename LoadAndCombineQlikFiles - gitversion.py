# James Caldwell
# 6/18/25

# Qlik files are several separate data files. This script loads them all into a single dataframe and saves it to an parquet file.
# I chose to use parquet files because it is much faster for reading/writing in python.
# Excel imports take ~2 mins, parquet imports take < 1 second.
import pandas as pd
import os

# pip install pyarrow

folder = r'folderpathhere'

df_list = []

print('loading files...')

for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('.xlsx'):
            file_path = os.path.join(root, file)
            print('loading: ', file_path)
            df = pd.read_excel(file_path)
            df_list.append(df) 
            print(f"Loaded: {file_path}")

# Combine into a single DataFrame
data = pd.concat(df_list, ignore_index=True)

data['CourseAndTerm'] = (
    data['Subject'].astype(str) + ' ' +
    data['Catalog Number'].astype(str) + ' - ' +
    data['Term Desc'].astype(str)
)
data['Subject and Catalog Number'] = data['Subject'] + ' ' + data['Catalog Number'].astype(str)
# print(data.head())

data.to_parquet('Course_Grade_Data.parquet')
print('Data saved to QlikData.parquet')
