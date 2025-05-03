import subprocess
import os
import glob
import pandas as pd
from datetime import datetime
import argparse

# === CONFIGURATION ===
script_extension = '.py'
script_dir = os.path.dirname(os.path.abspath(__file__))
final_dir = os.path.join(script_dir, "excel")
processed_log_path = os.path.join(final_dir, "last_processed_files.log")

# Parse command-line arguments for passwords and banks
parser = argparse.ArgumentParser(description="Process and merge credit card statements.")
parser.add_argument("--sbi-password", type=str, help="Password for SBI statements")
parser.add_argument("--hdfc-password", type=str, help="Password for HDFC statements")
parser.add_argument("--icici-password", type=str, help="Password for ICICI statements (if any)")
parser.add_argument("--banks", type=str, required=True, help="Comma-separated list of banks to process (e.g., 'SBI,HDFC')")
args = parser.parse_args()

# Read processed card+file entries into a set
if os.path.exists(processed_log_path):
    with open(processed_log_path, "r") as f:
        processed_files = set(line.strip() for line in f)
else:
    processed_files = set()

# Define configurations for each bank
card_configs = {
    "sbi": {
        "script": "excel_script_sbi.py",
        "dir": os.path.join(script_dir, "sbi"),
        "password": args.sbi_password,
    },
    "hdfc": {
        "script": "excel_script_hdfc.py",
        "dir": os.path.join(script_dir, "hdfc"),
        "password": args.hdfc_password,
    },
    "icici": {
        "script": "excel_script_icici.py",
        "dir": os.path.join(script_dir, "icici"),
        "password": args.icici_password,
    },
}

# Filter the banks to process based on user input
banks_to_process = [bank.lower() for bank in args.banks.split(",")]
card_configs = {key: value for key, value in card_configs.items() if key in banks_to_process}

if not card_configs:
    print("❌ No valid banks specified. Please provide valid bank names (e.g., 'SBI,HDFC').")
    exit(1)

# Process statements for the specified banks
start_year = 2023
end_year = datetime.now().year

for year in range(start_year, end_year + 1):
    for month in range(1, 13):
        mm = str(month).zfill(2)
        yyyy = str(year)
        base_filename = f"{mm}_{yyyy}"
        excel_basename = base_filename + ".xlsx"
        merged_excel_path = os.path.join(final_dir, excel_basename)

        temp_excel_paths = []
        newly_processed_files = []

        for card, config in card_configs.items():
            found_pdf_path = None
            for ext in [".pdf", ".PDF"]:
                candidate_path = os.path.join(config["dir"], base_filename + ext)
                if os.path.exists(candidate_path):
                    found_pdf_path = candidate_path
                    break

            if found_pdf_path:
                pdf_filename = os.path.basename(found_pdf_path)
                log_key = f"{card}:{pdf_filename}"

                if log_key in processed_files:
                    print(f"🔁 Already processed: {log_key}")
                    continue

                print(f"📄 Found {card.upper()} file: {pdf_filename}")
                cmd = ["python", os.path.join(script_dir, config["script"]),
                       "--in-dir", config["dir"],
                       "--out-dir", final_dir]
                if config["password"]:
                    cmd += ["--password", config["password"]]

                subprocess.run(cmd, check=True)
                print(f"✅ {card.upper()} processing complete")

                output_excel = os.path.join(final_dir, f"{card.upper()}_{base_filename}.xlsx")
                if os.path.exists(output_excel):
                    temp_excel_paths.append(output_excel)
                    newly_processed_files.append(log_key)
            else:
                print(f"⏩ No file for {card.upper()} {base_filename}.pdf")

        if temp_excel_paths:
            if os.path.exists(merged_excel_path):
                os.remove(merged_excel_path)
                print(f"🧹 Removed old merged Excel: {merged_excel_path}")

            with pd.ExcelWriter(merged_excel_path, engine="openpyxl") as writer:
                for excel_file in temp_excel_paths:
                    df = pd.read_excel(excel_file)
                    sheet_name = os.path.splitext(os.path.basename(excel_file))[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"📘 Merged Excel created: {merged_excel_path}")

            # Add newly processed items to the log
            with open(processed_log_path, "a") as f:
                for name in newly_processed_files:
                    f.write(name + "\n")

        else:
            print(f"❌ No new statements found for {mm}/{yyyy}, skipping merge.")

        # Cleanup temp Excel and txt files
        for bank in banks_to_process:
            pattern = os.path.join(final_dir, f"{bank}_??_????.xlsx")  # ?? = MM, ???? = YYYY
            for filepath in glob.glob(pattern):
                try:
                    os.remove(filepath)
                except Exception as e:
                    print(f"❗ Error deleting {filepath}: {e}")

        for txt_file in glob.glob(os.path.join(final_dir, "*.txt")):
            try:
                os.remove(txt_file)
            except Exception as e:
                print(f"❗ Error deleting file {txt_file}: {e}")