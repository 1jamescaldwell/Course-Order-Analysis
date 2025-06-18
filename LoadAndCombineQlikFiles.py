# James Caldwell
# 6/18/25

# Qlik files are several separate data files. This script loads them all into a single dataframe and saves it to an parquet file.
# I chose to use parquet files because it is much faster for reading/writing in python.
# Excel imports take ~2 mins, parquet imports take < 1 second.
import pandas as pd
import os

folder = r'C:\Users\ywe4kw\OneDrive - University of Virginia\Documents\3Projects\Course Order Project\Qlik data files'

df_list = []

print('loading files...')

for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('.xlsx'):
            file_path = os.path.join(root, file)
            df = pd.read_excel(file_path)
            df_list.append(df)  # collect each one
            print(f"Loaded: {file_path}")

# Combine into a single DataFrame
data = pd.concat(df_list, ignore_index=True)

data.to_parquet('Course_Grade_Data.parquet')
print('Data saved to QlikData.parquet')
