import fitz  # PyMuPDF
import argparse
import sys
import pandas as pd
import re
import os

def extract_rohit_transactions(file_path, output_excel, sheet_name):
    with open(file_path, 'r') as f:
        lines = [line.strip() for line in f.readlines()]

    transactions = []
    capturing = False
    i = 0

    while i < len(lines):
        line = lines[i]
        if line == "TRANSACTIONS FOR ROHIT KUMAR":
            capturing = True
            i += 1
            continue

        if capturing:
            if i + 3 < len(lines):
                date = lines[i]
                description = lines[i + 1]
                amount = lines[i + 2]
                txn_type = lines[i + 3]

                if txn_type != "D":
                    break  # Stop when "D" is not found

                transactions.append([date, description, amount, txn_type])
                i += 4
            else:
                break
        else:
            i += 1

    # Convert to DataFrame
    df = pd.DataFrame(transactions, columns=["Date", "Description", "Amount", "Type"])

    df["Amount"] = (df["Amount"].astype(str).str.replace(",", "", regex=False).astype(float))

    # Compute sum only from valid numbers
    total_amount = df["Amount"].sum()

    # Add a total row
    total_row = pd.DataFrame({
        "Date": [""],
        "Description": ["Total"],
        "Amount": [total_amount],
        "Type": [""]
    })

    # Append total row to the DataFrame
    final_df = pd.concat([df, total_row], ignore_index=True)

    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)

    

def write_to_file(lines, file):

    with open(file, 'w') as f:
        for line in lines:
            f.write(line)


def extract_text_from_pdf(pdf_path, password=None):
    try:
        doc = fitz.open(pdf_path)
        
        if doc.needs_pass:
            if not password:
                print("🔒 PDF is encrypted and needs a password.")
                return ""
            if not doc.authenticate(password):
                print("❌ Incorrect password.")
                return ""
        
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    
    except Exception as e:
        print(f"⚠️ Error reading PDF: {e}")
        return ""

def main(args):
    
    for file_name in os.listdir(args.in_dir):
        root, ext = os.path.splitext(file_name)
        if ext.lower() != '.pdf':
            continue
        pdf_path = os.path.join(args.in_dir, file_name)
        # print(f'Processing: {pdf_path}')
        excel_name = 'SBI_' + root + '.xlsx'
        excel_path = os.path.join(args.out_dir, excel_name)

        text_name = 'SBI_' + root + '.txt'
        text_path = os.path.join(args.out_dir, text_name)

        text = extract_text_from_pdf(pdf_path, args.password)
        write_to_file(text, text_path)

        sheet_name = 'SBI_' + root
        extract_rohit_transactions(text_path, excel_path, sheet_name)

        print(f"SBI: SUCCESS")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--in-dir', type=str, required=True, help='directory to read statement PDFs from.')
    parser.add_argument('--out-dir', type=str, required=True, help='directory to store statement XLSX to.')
    parser.add_argument('--password', type=str, default=None, help='password for the statement PDF.')
    args = parser.parse_args()
    main(args)