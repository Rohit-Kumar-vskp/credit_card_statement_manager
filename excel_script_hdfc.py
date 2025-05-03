import os
import argparse
import camelot
import pandas as pd
from collections import defaultdict
# import aes
# import arc4
from cryptography.hazmat.primitives.ciphers.algorithms import AES, ARC4

def extract_df(path, password=None):
    # The default values from pdfminer are M = 2.0, W = 0.1 and L = 0.5
    laparams = {'char_margin': 2.0, 'word_margin': 0.2, 'line_margin': 1.0}

    # Extract all tables using the lattice algorithm
    lattice_tables = camelot.read_pdf(path, password=password, 
        pages='all', flavor='lattice', line_scale=50, layout_kwargs=laparams)

    # Extract bounding boxes
    regions = defaultdict(list)
    for table in lattice_tables:
        bbox = [table._bbox[i] for i in [0, 3, 2, 1]]
        regions[table.page].append(bbox)

    df = pd.DataFrame()

    # Extract tables using the stream algorithm
    for page, boxes in regions.items():
        areas = [','.join([str(int(x)) for x in box]) for box in boxes]
        stream_tables = camelot.read_pdf(path, password=password, pages=str(page),
            flavor='stream', table_areas=areas, row_tol=5, layout_kwargs=laparams)
        dataframes = [table.df for table in stream_tables]
        dataframes = pd.concat(dataframes)
        # df = df.append(dataframes)
        df = pd.concat([df, dataframes])
    
    return df

def filter_excel_by_regex_and_columns(file_path, regex_column_index, pattern, output_columns, output_column_names, sheet_name):
    # Read Excel file without headers
    df = pd.read_excel(file_path, header=None)

    # Filter rows based on regex applied to the specified column
    mask = df.iloc[:, regex_column_index].astype(str).str.contains(pattern, regex=True, na=False)
    filtered_df = df[mask]

    # Select specific columns
    filtered_df = filtered_df.iloc[:, output_columns]

    # Rename the columns
    filtered_df.columns = output_column_names

    # Remove rows where "Amount" contains "Cr"
    filtered_df = filtered_df[~filtered_df["Amount"].astype(str).str.contains("Cr", case=False, na=False)]
    # Convert Amount to numeric, keep NaNs (don’t drop any rows)
    filtered_df["Amount"] = (filtered_df["Amount"].astype(str).str.replace(",", "", regex=False).astype(float))

    # Compute sum only from valid numbers
    total_amount = filtered_df["Amount"].sum()

    # Add a total row
    total_row = pd.DataFrame({
        "Date": [""],
        "Description": ["Total"],
        "Amount": [total_amount]
    })

    # Append total row to the DataFrame
    final_df = pd.concat([filtered_df, total_row], ignore_index=True)
    # Overwrite the Excel file
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
        final_df.to_excel(writer, index=False, sheet_name=sheet_name)

def main(args):
    for file_name in os.listdir(args.in_dir):
        root, ext = os.path.splitext(file_name)
        if ext.lower() != '.pdf':
            continue
        pdf_path = os.path.join(args.in_dir, file_name)
        # print(f'Processing: {pdf_path}')
        df = extract_df(pdf_path, args.password)
        excel_name = 'HDFC_' + root + '.xlsx'
        excel_path = os.path.join(args.out_dir, excel_name)
        df.to_excel(excel_path)
        # print(f'Processed : {excel_path}')

        filter_excel_by_regex_and_columns(
            file_path=excel_path,
            regex_column_index=1,                             # Filter on 2nd column
            pattern=r"[0-9]{2}/[0-9]{2}/[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}",      # Regex to match
            output_columns=[1, 2, 4],                         # Keep 2nd, 3rd, and 5th columns
            output_column_names=["Date", "Description", "Amount"],  # Rename columns
            sheet_name='HDFC_'+root 
        )   

        print("HDFC: SUCCESS")

        
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--in-dir', type=str, required=True, help='directory to read statement PDFs from.')
    parser.add_argument('--out-dir', type=str, required=True, help='directory to store statement XLSX to.')
    parser.add_argument('--password', type=str, default=None, help='password for the statement PDF.')
    args = parser.parse_args()

    main(args)