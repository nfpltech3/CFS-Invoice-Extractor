# CFS Invoice Extractor

A desktop application designed to extract structured data from physical CFS invoices using a fast 2-step approach: pdfplumber (for text-based PDFs) and Gemini Vision (for direct processing of scanned image PDFs). Includes support for automated mapping against a Job Registry.

## Supported Vendors
Currently, the extraction system natively understands the layouts for the following 16 vendors:
1. Gateway Distriparks
2. Ameya Logistics
3. Allcargo Terminals
4. J M Baxi Ports & Logistics
5. JWR Logistics
6. JWC Logistics Park
7. Ashte Logistics
8. Seabird Marine Services
9. Navkar Corporation
10. Ekaiva Supply Chain
11. Central Warehousing Corporation (CWC)
12. Apollo Logisolutions
13. APM Terminals
14. Balmer Lawrie & Co
15. Continental Warehousing Corporation
16. EFC Logistics

## Tech Stack
- Python 3.14
- Frameworks/Libraries: Tkinter, PyMuPDF, pdfplumber, Pillow, Google GenAI SDK
- APIs: Google Gemini API (gemini-2.5-flash)

---

## Installation

### Clone
```bash
git clone https://github.com/username/project-name.git
cd project-name
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```bash
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. Run application
```bash
python CFS_Invoice_Extractor.py
```

---

### Build Executable (For Desktop Apps)

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Generate Spec file (First time only if missing):
```bash
pyinstaller --name="CFS_Invoice_Extractor" --onefile --windowed CFS_Invoice_Extractor.py
```

3. Build using the included Spec file (Ensure you do not run CFS_Invoice_Extractor.py directly):
```bash
pyinstaller CFS_Invoice_Extractor.spec
```

4. Locate Executable:
The application will be generated in the `dist/` folder.

---

## Environment Variables

Copy:
```bash
cp .env.example .env
```

Add required values:
- `GEMINI_API_KEY`: Provide your Google Gemini API key (supports comma-separated list of keys if needed).

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv or __pycache__.
- Run and test before pushing.
