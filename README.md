# ASR Processing Tool

This tool automates the processing of ASR emails, organizes data into an Excel file, and creates ZIP archives for each FA. It is designed to streamline ASR data management and integrate seamlessly with Smartsheet.

---

## Features

- Extracts data from emails placed in the `_mail_asrs` folder.
- Generates an organized Excel file (`output.xlsx`) with the processed data.
- Automatically opens the Excel file upon completion.
- Creates ZIP archives for each FA with related ASRs in `_archive_pdf`.
- Easy integration with Smartsheet for managing ASR data.

---

## How to Use

### Prerequisites

- Ensure Python is installed (if running from source).
- For the `.exe` version, simply download the file and ensure the necessary folders (`_mail_asrs`, `_archive_pdf`, `_asr_pdfs`) exist.

### Steps to Run

1. **Clear the `_mail_asrs` Folder**  
   Ensure the `_mail_asrs` folder is empty. Remove any files leftover from previous runs.

2. **Add Email Files**  
   Drag and drop all email files (in `.msg` format) into the `_mail_asrs` folder.

3. **Run the Script**  
   Run the application:
   - If using the Python source, execute:
     ```bash
     python src/main.py
     ```
   - If using the `.exe` file, double-click `main.exe`.

4. **Wait for Processing**  
   The script will process the emails and automatically open an Excel file (`output.xlsx`) with the results.

5. **Transfer Data to Smartsheet**  
   Review the Excel file and transfer the data to your Smartsheet.

6. **Attach Archives (if necessary)**  
   Check the `_archive_pdf` folder for ZIP archives of ASRs grouped by FA. Attach these to Smartsheet if required.

7. **Clean Up**  
   Before starting a new batch, clear the folders:
   - `_mail_asrs`  
   - `_archive_pdf`  
   - `_asr_pdfs`  

---

## Notes

1. **Verify Results**  
   This tool is newly developed and may contain bugs. Always verify the processed data before using it in production.

2. **Key Folders**  
   - `_mail_asrs`: The main folder where emails are placed for processing. Always ensure this folder is empty before adding new files.
   - `_archive_pdf`: Contains ZIP files with ASRs grouped by FA. Clearing this folder is optional but recommended.
   - `_asr_pdfs`: Contains extracted and unsorted PDF files. Clearing this folder is optional.

3. **Excel File Handling**  
   - The Excel file (`output.xlsx`) opens automatically after processing. Ensure to **close it before running the script again**.
   - The file is overwritten with each new run, so save the previous data elsewhere if needed.

---

## Example Workflow

1. Place emails into `_mail_asrs`.
2. Run the tool by executing `main.exe`.
3. Wait for the tool to complete processing and open the Excel file.
4. Transfer the data to Smartsheet and attach relevant ZIP archives from `_archive_pdf`.
5. Clean the folders before the next batch.

---

## Troubleshooting

- If the script doesn't start, check that `output.xlsx` is closed.
- Ensure `_mail_asrs` is empty before adding new emails to avoid duplicates.
- If an archive is missing files, ensure the original PDFs are in `_asr_pdfs` and rerun the script.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
