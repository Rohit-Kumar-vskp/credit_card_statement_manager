# Credit Card Statement Manager

This project automates the processing and merging of credit card statements from multiple banks (SBI, ICICI, HDFC). The script extracts data from PDF statements, converts them into Excel files, and merges them into a single Excel file for each month. (This is all done by final_merger.py)

There is an additional feature provided related to extracting the bank pdf statements from your e-mail. The process for this is explained at the last after the details for final_merger.py

---

## Folder Structure

```
credit_card_script_automation/
├── excel/                      # Final merged Excel files are stored here
├── hdfc/                       # Upload HDFC PDF statements here
├── icici/                      # Upload ICICI PDF statements here
├── sbi/                        # Upload SBI PDF statements here
├── excel_script_hdfc.py        # Script for processing HDFC statements
├── excel_script_icici.py       # Script for processing ICICI statements
├── excel_script_sbi.py         # Script for processing SBI statements
├── final_merger.py             # Main script to process and merge statements
├── README.md                   # Documentation
├── Code.gs                     # Google app scripts code
└── requirements.txt            # Python dependencies
```

---

## Features

- **PDF Processing**: Extracts transaction data from PDF statements for SBI, ICICI, and HDFC credit cards.
- **Excel Generation**: Converts the extracted data into Excel files for each bank.
- **Monthly Merging**: Merges individual bank Excel files into a single Excel file for each month.
- **Duplicate Handling**: Skips processing of already processed files to avoid duplication.
- **Custom Passwords**: Supports password-protected PDFs for secure processing.

---

## Prerequisites

1. **Python 3.8+**: Ensure Python is installed on your system.
2. **Dependencies**: Install the required Python libraries using the following command:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### 1. Prepare Input Files
- Place the PDF statements for each bank in their respective folders:
  - `hdfc/` for HDFC statements
  - `icici/` for ICICI statements
  - `sbi/` for SBI statements

### 2. Run the Script
Execute the `final_merger.py` script to process and merge the statements (if there is no password, use "None"):

#### Example 1: Process only SBI and HDFC
```bash
python final_merger.py --banks sbi,hdfc --sbi-password YOUR_SBI_PASSWORD --hdfc-password YOUR_HDFC_PASSWORD
```

#### Example 2: Process all banks
```bash
python final_merger.py --banks sbi,hdfc,icici --sbi-password YOUR_SBI_PASSWORD --hdfc-password YOUR_HDFC_PASSWORD --icici-password YOUR_ICICI_PASSWORD
```

### 3. Output
- Individual Excel files for each bank will be saved in the `excel/` folder.
- Merged Excel files for each month will also be saved in the `excel/` folder.

---

## Configuration

The script uses the following configuration for each bank:

- **SBI**:
  - Password: `Your password`
  - Script: `excel_script_sbi.py`
- **HDFC**:
  - Password: `Your password`
  - Script: `excel_script_hdfc.py`
- **ICICI**:
  - Password: `Your password`
  - Script: `excel_script_icici.py`

You can modify these settings in the `card_configs` dictionary in `final_merger.py`.

---

## Logs

- Processed files are logged in `excel/last_processed_files.log` to avoid reprocessing the same files.

---

## Cleanup

Temporary Excel and text files are automatically deleted after merging.

---

## Troubleshooting

- **Missing Dependencies**: Ensure all required libraries are installed using `pip install -r requirements.txt`.
- **Incorrect Password**: Verify the password for password-protected PDFs in the `card_configs` dictionary.
- **File Not Found**: Ensure the PDF files are placed in the correct folders.

---
## Additional Feature: Extract PDFs from Gmail Using Google Apps Script

This project includes a Google Apps Script (`Code.gs`) that automates the extraction of PDF attachments from Gmail emails with a specific subject format and stores them in Google Drive.

---

### How to Set Up the Google Apps Script

1. **Open Google Apps Script**  
   - Go to [Google Apps Script](https://script.google.com/).
   - Create a new project.

2. **Add the Script**  
   - Copy the content of `Code.gs` into the editor.

3. **Modify the Script**  
   - Update the script to filter emails based on your desired subject format. For example:
     ```javascript
     var subjectFilter = "Credit Card Statement"; // Replace with your desired subject format
     ```
   - Specify the Google Drive folder where the PDFs will be saved:
     ```javascript
     var folderName = "Credit Card Statements"; // Replace with your desired folder name
     ```
   - Adjust the code to handle your specific PDF naming formats for each bank, if necessary.

4. **Authorize the Script**  
   - Run the script for the first time and grant the necessary permissions to access Gmail and Google Drive.

5. **Set Up a Trigger**  
   - Automate the script to run periodically by setting up a trigger:
     1. In the Apps Script editor, go to **Triggers** (clock icon on the left panel).
     2. Click **Add Trigger**.
     3. Select the function to run (e.g., `extractPDFs`).
     4. Choose the event source as **Time-driven**.
     5. Set the frequency (e.g., daily, hourly) based on your needs.
     6. Save the trigger.

6. **Run the Script**  
   - Execute the script manually or let the trigger run it automatically to fetch PDF attachments from Gmail and save them to the specified Google Drive folder.

---

### How to Use the Extracted PDFs

1. **Download PDFs from Google Drive**  
   - Navigate to the specified folder in Google Drive and download the extracted PDFs.

2. **Place PDFs in the Project Folders**  
   - Move the downloaded PDFs to the respective folders in the project:
     - `hdfc/` for HDFC statements
     - `icici/` for ICICI statements
     - `sbi/` for SBI statements.

3. **Run the `final_merger.py` Script**  
   - Follow the instructions in the [Usage](#usage) section to process and merge the statements.

---

### Benefits of This Feature

- Automates the retrieval of credit card statements from Gmail.
- Saves time by eliminating the need to manually download and organize PDFs.
- Ensures all relevant statements are stored in a centralized location for processing.

---

### Troubleshooting

- **No Emails Found**: Ensure the subject filter in the script matches the email subject format.
- **Permission Issues**: Verify that the script has the necessary permissions to access Gmail and Google Drive.
- **Trigger Not Running**: Check the trigger settings in the Apps Script editor and ensure it is active.
- **Incorrect Folder**: Ensure the specified Google Drive folder exists and is correctly named in the script.