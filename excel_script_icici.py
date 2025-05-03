import fitz  # PyMuPDF
import argparse
import sys
import pandas as pd
import re
import os

def extract_credit_card_transactions(txt_file_path, output_excel_path, sheet_name):
    def is_date(string):
        return re.match(r"\d{2}/\d{2}/\d{4}", string) is not None
    def is_valid_description(desc):
        return any(c.isalpha() for c in desc)  # contains at least one letter

    with open(txt_file_path, 'r', encoding='utf-8', errors='ignore') as file:
        lines = [line.strip() for line in file if line.strip()]

    transactions = []
    current_card = None
    i = 0

    while i < len(lines):
        line = lines[i]

        # Set card context
        if "4501XXXXXXXX1003" in line:
            current_card = "Sapphiro"
            i += 1
            continue
        elif "6528XXXXXXXX2001" in line:
            current_card = "Coral"
            i += 1
            continue

        # Start parsing transactions only if card is known and line is a valid date
        if current_card and is_date(line):
            try:
                date = lines[i]
                serial = lines[i + 1]
                desc1 = lines[i + 2]
                desc2 = lines[i + 3]

                if desc2.isdigit() and is_valid_description(desc1):
                    description = desc1
                    points = desc2
                    amount = lines[i + 4]
                    i += 5
                else:
                    combined_desc = f"{desc1} {desc2}"
                    if is_valid_description(combined_desc):
                        description = combined_desc
                        points = lines[i + 4]
                        amount = lines[i + 5]
                        i += 6
                    else:
                        i += 1
                        continue  # skip invalid

                transactions.append({
                    "Date": date,
                    "Description": description,
                    "Amount": amount.replace(",", "").strip(),
                    "Card": current_card
                })
            except IndexError:
                break
        else:
            i += 1

    # Convert to DataFrame
    df = pd.DataFrame(transactions)

    df = df[~df["Amount"].astype(str).str.contains("CR", case=False, na=False)]
    # Convert Amount to float
    df["Amount"] = pd.to_numeric(df["Amount"], errors='coerce')

    card1_total = df[df["Card"] == "Sapphiro"]["Amount"].sum()
    card2_total = df[df["Card"] == "Coral"]["Amount"].sum()
    total_total = df["Amount"].sum()

    # Add summary rows
    df = pd.concat([
        df,
        pd.DataFrame([
            {"Date": "", "Description": "Total - Sapphiro", "Amount": card1_total, "Card": ""},
            {"Date": "", "Description": "Total - Coral", "Amount": card2_total, "Card": ""},
            {"Date": "", "Description": "Net Total", "Amount": total_total, "Card": ""}
        ])
    ], ignore_index=True)

    # Save to Excel
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)    

def write_to_file(lines, file):

    with open(file, 'w', encoding='utf-8',errors="ignore") as f:
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
        # root = root.split('-')[0] + '-' + root.split('-')[1] + '-' + root.split('-')[2]
        if ext.lower() != '.pdf':
            continue
        pdf_path = os.path.join(args.in_dir, file_name)
        # print(f'Processing: {pdf_path}')
        excel_name = 'ICICI_' + root + '.xlsx'
        excel_path = os.path.join(args.out_dir, excel_name)

        text_name = 'ICICI_' + root + '.txt'
        text_path = os.path.join(args.out_dir, text_name)

        # text = extract_text_from_pdf(pdf_path, args.password)
        text = extract_text_from_pdf(pdf_path)

        write_to_file(text, text_path)
        sheet_name = 'ICICI_' + root

        extract_credit_card_transactions(text_path, excel_path, sheet_name)

        print(f"ICICI: SUCCESS")

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--in-dir', type=str, required=True, help='directory to read statement PDFs from.')
    parser.add_argument('--out-dir', type=str, required=True, help='directory to store statement XLSX to.')
    # parser.add_argument('--password', type=str, default=None, help='password for the statement PDF.')
    args = parser.parse_args()
    main(args)