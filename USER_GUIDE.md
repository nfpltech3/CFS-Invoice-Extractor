# CFS Invoice to Logisys User Guide

## Introduction
The **CFS Invoice Extractor** is an automated desktop application designed for accounts team. It scans physical CFS invoices (PDFs), intelligently extracts critical accounting data (Invoice Number, Date, Amount, Vendor Name), and links them against your daily Job Registry to build import narrations. Once verified, the tool generates a ready-to-upload Logisys CSV file.

## How to Use

### 1. Launching the App
1. Ensure the `CFS_Invoice_Extractor.exe` is in a folder.
2. In that **exact same folder**, ensure you have a file named `.env` containing your `GEMINI_API_KEY`.
3. If this is a new machine, ensure you have **Tesseract OCR** installed at `C:\Program Files\Tesseract-OCR`.
4. Double-click `CFS_Invoice_Extractor.exe`.

### 2. The Workflow (Step-by-Step)
1. **Load Invoice PDFs**: Click the first **Browse...** button and select one or multiple scanned vendor invoice `.pdf` files.
2. **Load Job Registry**: Click the second **Browse...** button and select your daily Job Registry (`.xlsx` or `.csv`).
3. **Process Invoices**: Click **PROCESS INVOICES**. The tool will analyze each PDF using AI and OCR.
   - *Note: This process may take 10-15 seconds per invoice depending on complexity.*
4. **Review Data**: The table will populate with the results. Look at the `Flag` column:
   - **`✓` (Green Check)**: Perfect match.
   - **`⚠` (Warning)**: The tool couldn't identify the vendor organization OR couldn't find a matching HBL/BOE in the Job Registry.
   - *Note: You can **Double-Click** any row with a warning to manually fix the Organization or Job Number!*
5. **Export to CSV**: Once everything looks good, click **Export CSV**.
   - *Result: Your formatted CSV will be saved into a new `CSV Output` folder next to your `.exe`.*

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Invoice PDFs Browse** | Selects the raw physical Scanned PDFs. | `.pdf` (Multiple selection allowed) |
| **Job Registry Browse** | Selects the Excel dump of your daily jobs. | `.xlsx` or `.csv` |
| **PROCESS INVOICES** | Begins the AI extraction. Cannot be clicked if no files are selected. | N/A |
| **Preview Table** | Displays extracted data (Org, Inv No, Amount, Job No, Flag). | N/A |
| **Export CSV** | Saves the table data directly into the Logisys CSV format. | N/A |

### Editor Popup Reference (Triggers on Double-Click)
| Field | Purpose |
| :--- | :--- |
| **Organization** | The Logisys-compatible name of the CFS (e.g., "ALLCARGO TERMINALS LIMITED"). |
| **Job No (Ref No)** | The Job Number to map against. |
| **Narration Name** | A shortened vendor name used specifically to piece together the final text narration. |

## Troubleshooting & Validations

If you see an error, check this table:

| Message / Visual Cue | What it means | Solution |
| :--- | :--- | :--- |
| **"Please select at least one PDF."** | You clicked Process before selecting invoice PDFs. | Select your PDF files using the Browse button. |
| **"Gemini API Client not initialized."** | The AI engine couldn't start because your Google API Key is missing. | Verify there is a `.env` file sitting right next to your `.exe` containing `GEMINI_API_KEY=your_key`. |
| **"All API keys exhausted for today."** | Your free-tier limits on Google Gemini have maxed out for the 24-hour cycle. | You must either wait for the quota reset, or add multiple comma-separated keys to your `.env` file. |
| **"Failed to load Job Registry."** | The Excel file you provided is corrupt or missing standard columns. | Ensure your Job Registry contains a header row, and it isn't completely empty. |
| **`⚠` Flag -> "UNKNOWN" Org** | The AI extracted the vendor name, but it doesn't match the internal lookup dictionary. | Double-click the row and manually type the correct Organization Name. |
| **`⚠` Flag -> "NOT FOUND" Job No** | The HBL or BOE from the invoice didn't match anything in your Job Registry. | Double-click the row and manually input the correct System Job Number. |
| **"Invalid amount extracted"** | The AI grabbed empty text or characters from the total amount box. | The invoice might be heavily blurred. Check the PDF manually. |
