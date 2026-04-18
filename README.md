# CFS Invoice Extractor

A desktop application designed to extract structured data from physical CFS invoices using a multi-tiered approach: pdfplumber (text), pytesseract (OCR), and Gemini Vision (AI), with support for automated mapping against a Job Registry.

## Tech Stack
- Python 3.14
- Frameworks/Libraries: Tkinter, PyMuPDF, pdfplumber, pytesseract, Pillow, Google GenAI SDK
- APIs: Google Gemini API (gemini-2.5-flash / gemini-2.5-flash-lite)

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
- `GEMINI_API_KEY`: Provide a single key or a comma-separated list of keys for automatic rotation to bypass free-tier rate limits.

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv or __pycache__.
- Run and test before pushing.
