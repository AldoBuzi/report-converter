import pandas as pd
import glob
import os
import re
from decimal import Decimal, ROUND_HALF_UP


#Find the file automatically
# We look for any CSV file that starts with 'productos_quickerp_'
file_pattern = 'productos_quickerp_*.xlsx'
found_files = glob.glob(file_pattern)

if not found_files:
    print("No file found matching the pattern 'productos_quickerp_*.csv'")
else:
    # Take the first matching file found
    file_path = found_files[0]
    print(f"Reading file: {file_path}")

    df = pd.read_excel(file_path, dtype={'CODPROD': str})
    print(df.head())
    # Formula: PRECIO_FINAL / 1.13 (Removing 13%)
    print(df['PRECIO_FINAL'])
    df['PRECIO_VENTA'] = df['PRECIO_FINAL'].astype(float) * 0.87
    print(df['PRECIO_VENTA'])
    df['PRECIO_VENTA'] = df['PRECIO_VENTA'].round().astype(int)
    df['CODPROD'] = df['CODPROD'].apply(lambda x: f'="{x}"')
    # We save to a new file to keep the original safe
    output_filename = 'productos_processed.csv'
    df.to_csv(output_filename, index=False)
    
    print(f"Successfully created '{output_filename}' with updated prices.")