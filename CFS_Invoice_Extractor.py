import os
import sys
import re
import csv
import json
import threading
import datetime
import time
from pathlib import Path
import io

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

import pdfplumber
import fitz  # PyMuPDF
import openpyxl
from dotenv import load_dotenv

from google import genai
from google.genai import types
from pydantic import BaseModel, Field

# Load environment variables
load_dotenv()


# ---------------------------------------------------------
# API KEY ROTATION
# ---------------------------------------------------------

# Parse comma-separated keys from .env
_raw_keys = os.environ.get("GEMINI_API_KEY", "")
API_KEY_LIST: list[str] = [k.strip() for k in _raw_keys.split(",") if k.strip()]
CURRENT_KEY_INDEX: int = 0

# Initialize client with the first key
try:
    gemini_client = genai.Client(api_key=API_KEY_LIST[0]) if API_KEY_LIST else None
    if API_KEY_LIST:
        print(f"Gemini client initialized with API Key #1 of {len(API_KEY_LIST)}")
except Exception as e:
    gemini_client = None
    print(f"Failed to initialize Gemini client: {e}")


def rotate_api_key():
    """Switch to the next API key in the pool. Raises Exception if none left."""
    global CURRENT_KEY_INDEX, gemini_client
    CURRENT_KEY_INDEX += 1
    if CURRENT_KEY_INDEX < len(API_KEY_LIST):
        gemini_client = genai.Client(api_key=API_KEY_LIST[CURRENT_KEY_INDEX])
        print(f"⚠ Rotated to API Key #{CURRENT_KEY_INDEX + 1} of {len(API_KEY_LIST)}")
    else:
        raise Exception("All API keys have been exhausted/failed. Check your Gemini API usage.")

# ---------------------------------------------------------
# CONSTANTS & MAPPINGS
# ---------------------------------------------------------

class InvoiceData(BaseModel):
    """Schema for structured extraction from CFS/port/terminal invoices.
    Only fields that come FROM the invoice are here.
    Fixed/hardcoded fields (Currency, GL code, Tax Type etc.) are applied in Python."""
    vendor_name: str | None = Field(description="Company that issued the invoice — from the letterhead/logo at the top")
    invoice_number: str | None = Field(description="Invoice number (NOT receipt number). Full alphanumeric string.")
    invoice_date: str | None = Field(description="Invoice date only (no time). Format: DD-MM-YYYY")
    hbl_number: str | None = Field(description="House Bill of Lading (HBL/HAWB). If only one BL exists on the invoice, put it here.")
    mbl_number: str | None = Field(description="Master Bill of Lading (MBL/MAWB). Only populated when BOTH MBL and HBL are present.")
    boe_number: str | None = Field(description="Bill of Entry number (labeled as BE No, BOE No, BOI No, B/E No, or BGE No on the invoice) — numeric portion only, no date")
    total_invoice_amount: float | None = Field(description="Final payable amount INCLUDING all taxes and round-off — the last grand total")


# ---------- Gemini Extraction Prompt ----------

EXTRACTION_PROMPT = """You are an expert invoice data extractor for Nagarkot Forwarders Pvt. Ltd., an Indian Customs House Agent (CHA).

CONTEXT:
You will receive a CFS invoice — either as extracted text or as a scanned image.

Known vendors: Gateway Distriparks, Ameya Logistics, Allcargo Terminals, J M Baxi Ports & Logistics, JWR Logistics, JWC Logistics Park, Ashte Logistics.

IMPORTANT: A PDF may contain BOTH a Tax Invoice page AND a Receipt page. Always PRIORITISE the Tax Invoice page for extracting invoice_number, invoice_date, and total_invoice_amount. Use the Receipt page only as a fallback if the Tax Invoice page is missing or illegible.

EXTRACT EXACTLY THESE 6 FIELDS:

1. vendor_name
   → The company that ISSUED the invoice. Found in the letterhead, logo, or header at the top.
   → This is the CFS/terminal operator — NOT "Nagarkot Forwarders", NOT the importer/consignee.

2. invoice_number
   → The INVOICE number (not the receipt number, not the acknowledgement number).
   → Look for labels: "Invoice No", "INV No", "Invoice No.", "Tax Invoice No", "Voice No".
   → Some formats show it in a table column labeled "INV No" or "Invoice No".
   → Extract the FULL alphanumeric string including prefixes/suffixes (e.g., GDLIH2627/001804, IMP0039456-25-26, IFI064546/25-26J, IMILG2026000930, EBI001925/25-26S).
   → JWC/JWR OCR WARNING: For JWC Logistics and JWR Logistics invoices, the invoice number typically starts with the uppercase letter 'I', NOT the number '1'. (e.g., 'I26000977', not '126000977'). Due to visual similarity, pay extremely close attention and NEVER extract a '1' if it should be an 'I'. Be aware that OCR might garble it (e.g., 'izeon1745' might actually be 'I26001745').

3. invoice_date
   → The date the INVOICE was issued (not receipt date, not validity date, not BOE date).
   → Look for labels: "Invoice Date", "Inv Date", "Date" (next to invoice number).
   → Output MUST be in DD-MM-YYYY format. Convert from any source format.
   → If the date includes a time component (e.g., "10-04-2026 13:24"), extract ONLY the date part.

4. hbl_number (House Bill of Lading)
   → This is the PRIMARY field used for matching invoices to jobs.
   → Look for labels: "HBL No", "HAWB No", "HBL", "HAWB".
   → CRITICAL RULE FOR SINGLE vs DUAL BL:
     • If the invoice shows ONLY ONE BL number (labeled "BL No", "B/L No", or similar):
       put it in hbl_number. Set mbl_number to null.
       Reason: when only one BL is present, it is almost always the House BL.
     • If the invoice shows TWO BL numbers (one labeled MBL/MAWB/Master and one labeled HBL/HAWB/House):
       put the House BL in hbl_number and the Master BL in mbl_number.
   → Extract ONLY the BL number string — not the date that may appear next to it.

5. mbl_number (Master Bill of Lading)
   → This field is ONLY for Master Bill of Lading numbers — numbers labeled EXPLICITLY with words like:
     "MBL No", "MAWB No", "Master BL", "Master Bill of Lading".
   → CRITICAL: mbl_number must NEVER contain:
     • A Bill of Entry number (BE No, BOE No) — that goes in boe_number
     • A cargo weight, container number, IGM number, or any other reference
     • The invoice number or any date
   → If only one BL exists on the invoice (no separate Master BL), set mbl_number to null.
   → When in doubt, set mbl_number to null rather than guessing.

6. boe_number (Bill of Entry number)
   → Look for labels: "BE No", "BOE No", "B/E No", "BE No/BE Date", "BOE", "BOI No", "BGE No", "BE No./BE Date".
   → FORMAT: 7 digits (e.g., 7936934, 8451015, 8260636).
     Sometimes it can be alphanumeric (e.g., 7936934__07/03/2026 where the numeric portion before the date is the BOE number).
   → If the field contains something like "BE No: 7936934__07/03/2026":
     extract ONLY "7936934" (the number BEFORE the date separator).
   → If the field contains something like "BOE No: TEMPTCNU32039 80000":
     extract the FULL identifier "TEMPTCNU32039" (the primary identifier, not just trailing digits).
     Do NOT extract just "80000" — that's incomplete.
   → Extract ONLY the BOE identifier — never include the BE Date or BGE Date.
   → COMMON MISTAKE: Don't confuse BE numbers with Master BL numbers. BE numbers are customs document numbers. If you see a 7-digit number labeled "BE No", it belongs in boe_number, NOT in mbl_number.

7. total_invoice_amount
   → The FINAL payable amount — this is the grand total AFTER adding all taxes (GST) and round-off.
   → This is always the LAST and LARGEST total on the invoice.
   → Look for labels: "Grand Total", "Total Amount After Tax", "Inv Amt", "Net Amount", "Total Invoice Amount", or simply the final "Total" row.
   → It MUST include GST. It MUST include round-off if present.
   → Output as a plain number with up to 2 decimal places. No commas, no currency symbols.

RULES:
- If a field is not visible, not legible, or not present: set it to null. NEVER guess or fabricate values.
- All dates must be DD-MM-YYYY format.
- All amounts must be plain numbers (no commas, no ₹ symbol).
- If the document has multiple pages, look at ALL pages before answering.
- Prefer Tax Invoice data over Receipt data when both exist in the same document.

COMMON MISTAKES TO AVOID:
- JWC/JWR INVOICES: Confusing the starting uppercase letter "I" with the number "1" in the invoice number (extracting '126000977' instead of the correct 'I26000977'). Watch out for garbled letters like 'izeon' which is actually 'I2600'.
- BOE numbers are typically 7 digits (e.g., 8451015). A 3-digit number like 992 is NOT a BOE — it's likely a weight or quantity field.
- BL numbers are typically 10+ characters. Read each digit ONE BY ONE — do not guess or interpolate.
- "Cargo Weight" is NOT a BOE number. Do not confuse numeric weight fields with BOE numbers.
- A number labeled "BE No" or "BOE No" is a Bill of Entry number → goes in boe_number, NEVER in mbl_number.
- An invoice with only ONE bill reference and a BE number has: hbl_number = the BL, mbl_number = null, boe_number = the BE.

CRITICAL LAYOUT NOTE — "Operational Details" section:
Many CFS invoices (especially Ameya, Allcargo, Gateway) have a TWO-COLUMN table in the "Operational Details" section.

EXAMPLE 1 — Ameya-style layout:
  LEFT COLUMN                    RIGHT COLUMN
  Shipping Line: YANG MING       Commodity Name: GENERAL CARGO
  IGM/Item No:   1184081/123     BL No:          1072298913
  CHA Name:      NAGARKOT...     Cargo Weight:   992
  BOE No:        8451015         BOE Date:       02-Apr-2026

  Correct extraction:
  - hbl_number = "1072298913" (single BL present)
  - mbl_number = null
  - boe_number = "8451015"

EXAMPLE 2 — Allcargo-style layout (two values in one row):
  BL No/BL Date :  BKK1265663__25/02/2026    BE No/BE Date : 7936934__07/03/2026
  IGM No:          1179376                   Vesse/Movement: X PRESS ANGLESEY

  Correct extraction:
  - hbl_number = "BKK1265663" (the BL number, before the date)
  - mbl_number = null (only one BL present — no Master BL mentioned anywhere)
  - boe_number = "7936934" (the BE number, before the date)
  
  WRONG extraction (do NOT do this):
  - mbl_number = "7936934" ❌ (This is a BE number, not an MBL!)

- "BOE No" and "BOE Date" are on the SAME ROW but DIFFERENT COLUMNS. They are SEPARATE fields.
- "BOE No" contains the Bill of Entry identifier → put this in boe_number.
- "BOE Date" contains a DATE → ignore this for boe_number.
- Do NOT skip BOE No just because BOE Date is next to it.
- When you see "BE No/BE Date: 7936934__07/03/2026" — the "__" is a visual separator. 7936934 is the BE number, 07/03/2026 is the date.
"""


# ---------- Organization Mapping ----------
# Maps vendor name keywords → (Logisys Name, Short Name for Narration)

ORG_MAPPING_RULES = {
    "gateway distriparks": ("GATEWAY DISTRIPARKS LTD.", "Gateway"),
    "gateway": ("GATEWAY DISTRIPARKS LTD.", "Gateway"),
    "ameya logistics": ("AMEYA LOGISTICS PVT. LTD.", "Ameya"),
    "ameya": ("AMEYA LOGISTICS PVT. LTD.", "Ameya"),
    "psa ameya": ("AMEYA LOGISTICS PVT. LTD.", "Ameya"),
    "allcargo terminals": ("ALLCARGO TERMINALS LIMITED", "Allcargo"),
    "allcargo": ("ALLCARGO TERMINALS LIMITED", "Allcargo"),
    "j m baxi": ("J M BAXI PORTS & LOGISTICS LTD.-V- (ICT INNFRA.PVT.LTD.)", "J M Baxi"),
    "jm baxi": ("J M BAXI PORTS & LOGISTICS LTD.-V- (ICT INNFRA.PVT.LTD.)", "J M Baxi"),
    "jwr logistics": ("JWR LOGISTICS PVT LTD", "JWR"),
    "jwc logistics": ("JWC LOGISTICS PARK PVT.LTD.", "JWC"),
    "ashte logistics": ("ASHTE LOGISTICS PVT LTD", "Ashte"),
}

def map_organization(vendor_name):
    """Map extracted vendor name to (Logisys org name, short name for narration, matched_keyword).
    Returns tuple: (logisys_name, short_name, matched_keyword)"""
    if not vendor_name:
        return "UNKNOWN VENDOR", "UNKNOWN", None
    v_lower = vendor_name.lower().strip()
    # Check longer keys first for specificity (e.g., "gateway distriparks" before "gateway")
    for key in sorted(ORG_MAPPING_RULES.keys(), key=len, reverse=True):
        if key in v_lower:
            return ORG_MAPPING_RULES[key][0], ORG_MAPPING_RULES[key][1], key
    return f"UNKNOWN - {vendor_name}", vendor_name, None


# ---------- Date Formatting ----------

MONTH_MAP = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
    "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
}

def format_date(date_str):
    """Convert various date formats to DD-MMM-YYYY (e.g., 09-Apr-2026)."""
    if not date_str:
        return ""
    # Try standard parsing first
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y", "%Y-%m-%d",
                "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %b %y",
                "%d-%B-%Y", "%d-%B-%y", "%d %B %Y"):
        try:
            dt = datetime.datetime.strptime(date_str.strip(), fmt)
            return dt.strftime("%d-%b-%Y")
        except ValueError:
            continue
    # Fallback: regex extraction
    try:
        m = re.search(r'(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})', date_str)
        if m:
            day, month, year = m.groups()
            month_str = MONTH_MAP.get(month.zfill(2), "")
            if not month_str:
                return date_str
            if len(year) == 2:
                year = "20" + year
            return f"{day.zfill(2)}-{month_str}-{year}"
    except Exception:
        pass
    return date_str


# ---------- Reference Number Normalization ----------

def normalize_ref_number(ref_num):
    """Strip all non-alphanumeric chars, uppercase — for fuzzy matching."""
    if not ref_num:
        return ""
    return re.sub(r'[^A-Z0-9]', '', str(ref_num).upper())


# ---------------------------------------------------------
# GLOBAL STATE
# ---------------------------------------------------------
selected_pdfs = []
selected_job_registry = ""
job_mapping_cache = {}  # { normalized_hbl: job_no, "BE_" + normalized_be: job_no }
processed_results = []  # list of row dicts for preview and export
batch_log_entries = []
batch_log_meta = {}


# ---------------------------------------------------------
# JOB REGISTRY
# ---------------------------------------------------------

def load_job_registry(filepath):
    """Load Job Registry Excel/CSV into memory for BL/BOE → Job No lookup."""
    global job_mapping_cache
    job_mapping_cache.clear()

    try:
        def process_headers_and_row(headers, row):
            row_dict = dict(zip(headers, row))
            hbl_key = next((k for k in headers if "hbl" in k or "hawb" in k), None)
            be_key = next((k for k in headers if "be no" in k or "be_no" in k or k == "be no"), None)
            job_key = next((k for k in headers if "job" in k and "no" in k), None)
            # Fallback: any column with "job" in name
            if not job_key:
                job_key = next((k for k in headers if "job" in k), None)

            hbl = normalize_ref_number(row_dict.get(hbl_key, "")) if hbl_key else ""
            be_no = normalize_ref_number(row_dict.get(be_key, "")) if be_key else ""
            job_no = str(row_dict.get(job_key, "")).strip() if job_key else ""

            if job_no and job_no != "None":
                if hbl:
                    job_mapping_cache[hbl] = job_no
                if be_no:
                    job_mapping_cache[f"BE_{be_no}"] = job_no

        ext = Path(filepath).suffix.lower()

        if ext == '.csv':
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                headers = []
                for row_idx, row in enumerate(reader):
                    if row_idx == 0:
                        headers = [str(cell).strip().lower() if cell else "" for cell in row]
                        continue
                    process_headers_and_row(headers, row)
        else:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheet = wb.active
            headers = []
            for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                if row_idx == 0:
                    headers = [str(cell).strip().lower() if cell else "" for cell in row]
                    continue
                process_headers_and_row(headers, row)

        return True
    except Exception as e:
        print(f"Error loading Job Registry: {e}")
        return False


def find_job_number(hbl_num, boe_num):
    """Look up Job No from registry.
    Primary: match HBL number against HAWB/HBL No column.
    Fallback: match BOE number against BE No column.
    Returns: (job_no, match_type, match_value)"""
    # Primary: match HBL (this covers both cases:
    #   - invoice had only one BL → it's stored as hbl_number → matches HAWB/HBL No in registry
    #   - invoice had both MBL+HBL → hbl_number has the House BL → matches HAWB/HBL No in registry)
    hbl_clean = normalize_ref_number(hbl_num)
    if hbl_clean and hbl_clean in job_mapping_cache:
        return job_mapping_cache[hbl_clean], "HBL", hbl_clean

    # Fallback: match BOE number
    boe_clean = normalize_ref_number(boe_num)
    if boe_clean and f"BE_{boe_clean}" in job_mapping_cache:
        return job_mapping_cache[f"BE_{boe_clean}"], "BOE", boe_clean

    return "NOT FOUND", "NONE", None


# ---------------------------------------------------------
# PDF EXTRACTION + GEMINI
# ---------------------------------------------------------

def extract_invoice_data(pdf_path):
    """
    Extract invoice data using a 2-step approach:
    Step A: Try pdfplumber to extract text. If >100 chars, send text to Gemini.
    Step B: If <100 chars (scanned), send PDF pages as IMAGES directly to Gemini.
    """
    print(f"  Processing: {Path(pdf_path).name}")
    print(f"{'='*60}")
    
    # Step A: Try pdfplumber (text-layer PDFs)
    extracted_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
    except Exception as e:
        print(f"pdfplumber failed: {e}")

    extracted_text = extracted_text.strip()
    print(f"  [Step A] pdfplumber: {len(extracted_text)} chars")
    
    if len(extracted_text) > 100:
        print(f"  [Step A] ✓ Sufficient text found → sending TEXT to Gemini")
        pdf_type = "Text-based"
        pdf_chars = len(extracted_text)
        data = call_gemini_extract(text_content=extracted_text)
        return data, pdf_type, pdf_chars

    # Step B: Gemini Vision (scanned PDFs)
    print(f"  [Step B] ⚠ Insufficient text → sending IMAGES directly to Gemini")
    pdf_type = "Scanned (Gemini Vision)"
    pdf_chars = len(extracted_text) if extracted_text else 0
    data = call_gemini_extract(pdf_path=pdf_path)
    return data, pdf_type, pdf_chars


def call_gemini_extract(text_content=None, pdf_path=None):
    """Call Gemini API for invoice data extraction with simple retry logic."""
    if not gemini_client:
        raise Exception("Gemini client not initialized. Check GEMINI_API_KEY in .env file.")

    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=InvoiceData,
        temperature=0.0,
    )

    # Build contents
    if text_content:
        contents = [
            EXTRACTION_PROMPT,
            f"--- INVOICE TEXT START ---\n{text_content}\n--- INVOICE TEXT END ---"
        ]
    elif pdf_path:
        doc = fitz.open(pdf_path)
        contents = [EXTRACTION_PROMPT]
        for i in range(min(3, len(doc))):
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='JPEG', quality=90)
            part = types.Part.from_bytes(
                data=img_byte_arr.getvalue(),
                mime_type='image/jpeg'
            )
            contents.append(part)
        doc.close()
    else:
        raise Exception("No content provided for extraction.")

    mode = "TEXT" if text_content else "IMAGE"
    print(f"  [API] Mode: {mode}")
    
    max_retries = 3
    attempt = 0
    
    while attempt < max_retries:
        print(f"  [API] Attempt {attempt + 1} (Key #{CURRENT_KEY_INDEX + 1})")
        try:
            response = gemini_client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
                config=config,
            )
            json_data = json.loads(response.text)
            print(f"  [API] ✓ Success on gemini-2.5-flash")
            return json_data

        except Exception as e:
            err_str = str(e)
            print(f"  [API Error]: {err_str[:120]}")

            # 1. API KEY EXHAUSTED OR QUOTA -> Rotate keys
            if "PerDay" in err_str or ("Quota" in err_str and "limit: 0" in err_str):
                print(f"  [API] ❌ Quota exhausted on Key #{CURRENT_KEY_INDEX + 1} → rotating...")
                rotate_api_key()
                attempt = 0 # reset attempts for the new key
                continue

            # 2. RPM THROTTLE -> Wait
            if "PerMinute" in err_str or "retryDelay" in err_str:
                wait_time = 15
                delay_match = re.search(r'retryDelay.*?(\d+)', err_str)
                if delay_match:
                    wait_time = int(delay_match.group(1))
                wait_time = min(wait_time, 60)
                print(f"  [API] ⏳ RPM throttle → waiting {wait_time}s")
                time.sleep(wait_time)
                attempt += 1
                continue

            # 3. SERVER ERROR -> Short wait
            if "503" in err_str or "UNAVAILABLE" in err_str:
                print(f"  [API] 🔄 Server busy → waiting 5s")
                time.sleep(5)
                attempt += 1
                continue

            # Unhandled errors
            print(f"  [API] ❌ Non-retryable error")
            raise

    raise Exception(f"Gemini extraction failed: exceeded max retries")


# ---------------------------------------------------------
# GUI
# ---------------------------------------------------------
BRAND_BLUE = "#1F3F6E"
BRAND_BG = "#F4F6F8"


class EditRowPopup(tk.Toplevel):
    """Popup to manually edit flagged rows (unknown vendor, missing job no)."""

    def __init__(self, parent, item_id, tree, row_data):
        super().__init__(parent)
        self.title("Edit Record")
        self.geometry("500x350")
        self.configure(bg=BRAND_BG)
        self.transient(parent)
        self.grab_set()

        self.item_id = item_id
        self.tree = tree
        self.row_data = row_data  # direct reference to the dict in processed_results

        ttk.Label(self, text="Edit Invoice Row", font=("Segoe UI", 12, "bold"),
                  background=BRAND_BG, foreground=BRAND_BLUE).pack(pady=10)

        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=20)

        # Organization
        ttk.Label(frame, text="Organization:").grid(row=0, column=0, sticky='w', pady=5)
        self.org_var = tk.StringVar(value=self.row_data.get('Organization', ''))
        ttk.Entry(frame, textvariable=self.org_var, width=50).grid(row=0, column=1, pady=5, padx=5)

        # Job Number
        ttk.Label(frame, text="Job No (Ref No):").grid(row=1, column=0, sticky='w', pady=5)
        self.job_var = tk.StringVar(value=self.row_data.get('Ref No', ''))
        ttk.Entry(frame, textvariable=self.job_var, width=50).grid(row=1, column=1, pady=5, padx=5)

        # Short Name (for narration)
        ttk.Label(frame, text="Narration Name:").grid(row=2, column=0, sticky='w', pady=5)
        self.short_name_var = tk.StringVar(value=self.row_data.get('_ShortName', ''))
        ttk.Entry(frame, textvariable=self.short_name_var, width=50).grid(row=2, column=1, pady=5, padx=5)

        # Info label
        ttk.Label(frame, text="(Narration: Being Entry posted for [Name] / CFS / [Job No])",
                  foreground="#6B7280").grid(row=3, column=0, columnspan=2, pady=5)

        ttk.Button(self, text="Save Changes", command=self._save).pack(pady=15)

    def _save(self):
        self.row_data['Organization'] = self.org_var.get()
        self.row_data['Ref No'] = self.job_var.get()
        self.row_data['_ShortName'] = self.short_name_var.get()

        # Rebuild narration
        short_name = self.row_data['_ShortName']
        job_no = self.row_data['Ref No']
        inv_no = self.row_data.get('Vendor Inv No', '')

        if job_no and job_no != "NOT FOUND":
            self.row_data['Narration'] = f"Being Entry posted for {short_name} / CFS / {job_no}"
        else:
            self.row_data['Narration'] = f"Being Entry posted for {short_name} / CFS / {inv_no}"

        # Update flag
        flag = "✓"
        if "UNKNOWN" in self.row_data['Organization']:
            flag = "⚠"
        if self.row_data['Ref No'] == "NOT FOUND":
            flag = "⚠"
        self.row_data['_Flag'] = flag

        # Refresh tree row
        self.tree.item(self.item_id, values=(
            self.row_data['Organization'],
            self.row_data['Vendor Inv No'],
            self.row_data['Amount'],
            self.row_data['Ref No'],
            flag
        ))

        self.destroy()


def resource_path(relative_path):
    """Get path for bundled resources (PyInstaller compatible)."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CFS Invoice → Logisys CSV Tool")
        self.state("zoomed")
        self.configure(bg=BRAND_BG)

        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background=BRAND_BG)
        style.configure('TLabel', background=BRAND_BG, foreground="#1E1E1E", font=("Segoe UI", 10))
        style.configure('Action.TButton', background=BRAND_BLUE, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=8)
        style.map('Action.TButton', background=[('active', "#2A528F")])
        style.configure('Secondary.TButton', background="white", foreground=BRAND_BLUE,
                        font=("Segoe UI", 10), padding=5)

        # Treeview brand styling
        style.configure('Treeview',
                        background="white",
                        foreground="#1E1E1E",
                        fieldbackground="white",
                        font=("Segoe UI", 10),
                        rowheight=28)
        style.configure('Treeview.Heading',
                        background=BRAND_BLUE,
                        foreground="white",
                        font=("Segoe UI", 10, "bold"),
                        relief="flat")
        style.map('Treeview.Heading',
                  background=[('active', "#2A528F")])
        style.map('Treeview',
                  background=[('selected', "#2A528F")],
                  foreground=[('selected', 'white')])

        # Progress bar brand styling
        style.configure('Brand.Horizontal.TProgressbar',
                        troughcolor='#E5E7EB',
                        background=BRAND_BLUE)

        # --- Footer (pack FIRST with side=BOTTOM so it anchors at the bottom) ---
        footer_frame = tk.Frame(self, bg=BRAND_BG)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=(0, 8))
        tk.Label(footer_frame, text="Nagarkot Forwarders Pvt. Ltd. \u00a9",
                 font=("Segoe UI", 8), bg=BRAND_BG, fg="#6B7280").pack(side=tk.LEFT)

        # --- Header ---
        header_bg = tk.Frame(self, bg="white")
        header_bg.pack(fill=tk.X)

        header_frame = tk.Frame(header_bg, bg="white")
        header_frame.pack(fill=tk.X, padx=20, pady=(16, 12))

        # Left-aligned logo
        logo_frame = tk.Frame(header_frame, bg="white")
        logo_frame.pack(side=tk.LEFT, padx=(10, 0))
        try:
            logo_path = resource_path("logo.png")
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                aspect = img.width / img.height
                img = img.resize((int(24 * aspect), 24), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(logo_frame, image=self.logo_img, bg="white").pack()
        except Exception:
            pass

        # Centered Title + Subtitle — true center relative to full window width
        center_container = tk.Frame(header_bg, bg="white")
        center_container.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        tk.Label(center_container, text="CFS Invoice to Logisys",
                 font=("Arial", 18, "bold"), bg="white", fg=BRAND_BLUE).pack(pady=(0, 2))
        tk.Label(center_container, text="Data Extractor & CSV Generator",
                 font=("Arial", 9), bg="white", fg="#6B7280").pack()

        # Separator line under header
        sep = tk.Frame(self, bg="#E5E7EB", height=1)
        sep.pack(fill=tk.X)

        # --- Body ---
        body_frame = ttk.Frame(self)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=10)

        # File selection
        sel_frame = ttk.Frame(body_frame)
        sel_frame.pack(fill=tk.X, pady=10)

        self.lbl_pdfs = ttk.Label(sel_frame, text="Invoice PDFs: 0 files selected")
        self.lbl_pdfs.grid(row=0, column=0, sticky="w", pady=5)
        ttk.Button(sel_frame, text="Browse...", style="Secondary.TButton",
                   command=self.browse_pdfs).grid(row=0, column=1, padx=10, pady=5)

        self.lbl_job = ttk.Label(sel_frame, text="Job Registry: None selected")
        self.lbl_job.grid(row=1, column=0, sticky="w", pady=5)
        ttk.Button(sel_frame, text="Browse...", style="Secondary.TButton",
                   command=self.browse_registry).grid(row=1, column=1, padx=10, pady=5)

        # Process button
        self.btn_process = ttk.Button(body_frame, text="PROCESS INVOICES",
                                      style="Action.TButton", command=self.run_process)
        self.btn_process.pack(pady=15)

        # Status + Progress
        self.lbl_status = ttk.Label(body_frame, text="Ready.")
        self.lbl_status.pack()
        self.progress = ttk.Progressbar(body_frame, orient=tk.HORIZONTAL,
                                        length=400, mode='determinate',
                                        style='Brand.Horizontal.TProgressbar')
        self.progress.pack(pady=5)

        # Preview table
        table_frame = ttk.Frame(body_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        columns = ("Org", "Inv No", "Amount", "Job No", "Flag")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.CENTER)
            if col == "Flag":
                self.tree.column(col, width=50, anchor=tk.CENTER)
            elif col == "Amount":
                self.tree.column(col, width=100, anchor=tk.CENTER)
            else:
                self.tree.column(col, width=200, anchor=tk.CENTER)

        # Alternating row colors
        self.tree.tag_configure('oddrow', background='#FFFFFF')
        self.tree.tag_configure('evenrow', background='#F0F4F8')

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # Bottom actions
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill=tk.X, padx=40, pady=10)

        self.btn_export = ttk.Button(bottom_frame, text="Export CSV", style="Action.TButton",
                                     command=self.export_csv, state=tk.DISABLED)
        self.btn_export.pack(side=tk.LEFT)

        # Error summary label
        self.lbl_summary = ttk.Label(bottom_frame, text="")
        self.lbl_summary.pack(side=tk.LEFT, padx=20)

    # --- File Selection ---

    def browse_pdfs(self):
        global selected_pdfs
        files = filedialog.askopenfilenames(title="Select Invoice PDFs",
                                            filetypes=[("PDF Files", "*.pdf")])
        if files:
            selected_pdfs = list(files)
            self.lbl_pdfs.config(text=f"Invoice PDFs: {len(selected_pdfs)} files selected")

    def browse_registry(self):
        global selected_job_registry
        file = filedialog.askopenfilename(title="Select Job Registry",
                                          filetypes=[("Spreadsheets", "*.xlsx *.csv")])
        if file:
            selected_job_registry = file
            self.lbl_job.config(text=f"Job Registry: {Path(file).name}")

    # --- Tree Interaction ---

    def on_tree_double_click(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return

        tree_vals = self.tree.item(item_id, 'values')
        target_id = None

        # Find the matching row in processed_results by _id
        # We store _id as the tree item's iid tag
        for r in processed_results:
            if r.get('Vendor Inv No') == tree_vals[1] and not r.get('_HasError'):
                target_id = r
                break

        if target_id:
            EditRowPopup(self, item_id, self.tree, target_id)

    # --- Processing ---

    def run_process(self):
        if not selected_pdfs:
            messagebox.showerror("Error", "Please select at least one PDF.")
            return
        if not selected_job_registry:
            messagebox.showerror("Error", "Please select the Job Registry Excel file.")
            return
        if not gemini_client:
            messagebox.showerror("Error",
                                 "Gemini API Client not initialized.\n"
                                 "Check GEMINI_API_KEY in your .env file.")
            return

        self.btn_process.config(state=tk.DISABLED)
        self.btn_export.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self.lbl_summary.config(text="")
        global processed_results
        processed_results = []

        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        try:
            self._update_status("Loading Job Registry...")
            if not load_job_registry(selected_job_registry):
                self.after(0, lambda: messagebox.showerror("Error", "Failed to load Job Registry."))
                self.after(0, lambda: self.btn_process.config(state=tk.NORMAL))
                return

            total = len(selected_pdfs)
            self.after(0, lambda: self.progress.configure(maximum=total, value=0))

            cur_date = datetime.datetime.now().strftime("%d-%b-%Y")
            error_count = 0
            warning_count = 0
            success_count = 0

            global batch_log_entries, batch_log_meta
            batch_log_entries.clear()
            batch_log_meta = {
                "date": datetime.datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
                "total": total,
                "registry": Path(selected_job_registry).name
            }

            for idx, pdf_path in enumerate(selected_pdfs):
                filename = Path(pdf_path).name
                self._update_status(f"Processing {idx + 1} of {total}: {filename}")
                self.after(0, lambda v=idx: self.progress.configure(value=v))

                pdf_type = "Unknown"
                pdf_chars = 0
                try:
                    # LAYER 1: Common Engine (OCR / Gemini extraction)
                    data, pdf_type, pdf_chars = extract_invoice_data(pdf_path)

                    vendor = data.get("vendor_name", "")
                    org_name, short_name, org_match_keyword = map_organization(vendor)

                    # LAYER 2: Vendor Detection
                    parser_type = "GENERIC"
                    if org_match_keyword == "ameya":
                        parser_type = "AMEYA"
                    elif org_match_keyword and "gateway" in org_match_keyword:
                        parser_type = "GATEWAY"
                    elif org_match_keyword and "allcargo" in org_match_keyword:
                        parser_type = "ALLCARGO"

                    # LAYER 3: Vendor-specific overrides (20% logic)
                    if parser_type == "AMEYA":
                        # Example: Override bad BOE or BL numbers specifically for Ameya if needed
                        pass
                    elif parser_type == "GATEWAY":
                        # Example: Gateway has specific formats we can strict-match
                        pass
                    
                    # Common Validation
                    inv_no = data.get("invoice_number", "")
                    raw_amount = data.get("total_invoice_amount")

                    if not inv_no:
                        raise Exception("Extracted invoice_number is empty.")

                    try:
                        amount = float(raw_amount)
                        if amount <= 0:
                            raise ValueError
                    except (TypeError, ValueError):
                        raise Exception(f"Invalid amount extracted: {raw_amount}")

                    # Map organization \u2192 (logisys_name, short_name)
                    # Look up job number from registry
                    job_no, job_match_type, job_match_value = find_job_number(data.get("hbl_number"), data.get("boe_number"))

                    # Build narration: "Being Entry posted for [Short Name] / CFS / [Job No]"
                    if job_no != "NOT FOUND":
                        narration = f"Being Entry posted for {short_name} / CFS / {job_no}"
                    else:
                        narration = f"Being Entry posted for {short_name} / CFS / {inv_no}"

                    # Build CSV row
                    # Fields marked FIXED never change. Fields marked BLANK are for future use.
                    
                    org_branch = "Mumbai"
                    if short_name in ["Allcargo", "J M Baxi", "Gateway", "JWC"]:
                        org_branch = "Navi Mumbai"
                    elif short_name == "Ashte":
                        org_branch = "Chembur"

                    row = {
                        # --- Extracted / Computed ---
                        "Entry Date": cur_date,                             # FIXED: current date
                        "Posting Date": cur_date,                           # FIXED: current date
                        "Organization": org_name,                           # MAPPED from vendor_name
                        "Organization Branch": org_branch,                  # DYNAMIC based on short_name
                        "Vendor Inv No": inv_no,                            # EXTRACTED
                        "Vendor Inv Date": format_date(data.get("invoice_date")),  # EXTRACTED
                        "Currency": "INR",                                  # FIXED
                        "ExchRate": "1",                                    # FIXED
                        "Narration": narration,                             # BUILT from short_name + job_no
                        "Due Date": "",                                     # BLANK
                        "Charge or GL": "Charge",                  # FIXED
                        "Charge or GL Name": "CFS CHARGES (1)",   # FIXED
                        "Charge or GL Amount": f"{round(amount)}",          # EXTRACTED: total with tax, rounded to nearest integer
                        "DR or CR": "DR",                                   # FIXED
                        "Cost Center": "CCL Import",                        # FIXED
                        "Branch": "HO",                                     # FIXED
                        "Charge Narration": "",                             # BLANK
                        "TaxGroup": "GSTIN",                                # FIXED
                        "Tax Type": "Pure Agent",                           # FIXED
                        "SAC or HSN": "999799",                             # FIXED
                        "Taxcode1": "",                                     # BLANK (future)
                        "Taxcode1 Amt": "",                                 # BLANK (future)
                        "Taxcode2": "",                                     # BLANK (future)
                        "Taxcode2 Amt": "",                                 # BLANK (future)
                        "Taxcode3": "",                                     # BLANK (future)
                        "Taxcode3 Amt": "",                                 # BLANK (future)
                        "Taxcode4": "",                                     # BLANK (future)
                        "Taxcode4 Amt": "",                                 # BLANK (future)
                        "Avail Tax Credit": "No",                           # FIXED
                        "LOB": "CCL IMP",                                   # FIXED
                        "Ref Type": "",                                     # BLANK
                        "Ref No": job_no,                                   # LOOKED UP from Job Registry

                        "Amount": f"{round(amount)}",                       # SAME as Charge or GL Amount
                        "Start Date": "",                                   # BLANK
                        "End Date": "",                                     # BLANK
                        "WH Tax Code": "",                                  # BLANK (future)
                        "WH Tax Percentage": "",                            # BLANK (future)
                        "WH Tax Taxable": "",                               # BLANK (future)
                        "WH Tax Amount": "",                                # BLANK (future)
                        "Round Off": "No",                                  # FIXED
                        "CC Code": "",                                      # BLANK
                        # --- Internal fields (excluded from CSV via extrasaction='ignore') ---
                        "_ShortName": short_name,
                        "_id": idx,
                    }

                    flag = "✓"
                    if "UNKNOWN" in org_name:
                        flag = "⚠"
                        warning_count += 1
                    elif job_no == "NOT FOUND":
                        flag = "⚠"
                        warning_count += 1
                    else:
                        success_count += 1

                    row["_Flag"] = flag
                    processed_results.append(row)

                    batch_log_entries.append({
                        "index": idx + 1,
                        "total": total,
                        "filename": filename,
                        "status": flag,
                        "pdf_type": pdf_type,
                        "pdf_chars": pdf_chars,
                        "gemini_raw": data,
                        "org_name": org_name,
                        "org_match_keyword": org_match_keyword,
                        "job_no": job_no,
                        "job_match_type": job_match_type,
                        "job_match_value": job_match_value,
                        "final_row": {
                            "Vendor Inv No": inv_no, 
                            "Vendor Inv Date": row["Vendor Inv Date"], 
                            "Amount": f"{round(amount)}", 
                            "Ref No": job_no, 
                            "Narration": narration
                        },
                        "error": None
                    })

                    self.after(0, self._add_to_tree,
                              org_name, inv_no, row["Amount"], job_no, flag)
                    print(f"  [Result] \u2713 Inv: {inv_no} | Amt: {round(amount)} | Job: {job_no} | Org: {short_name}")

                except Exception as ex:
                    error_count += 1
                    print(f"  [Result] \u274c FAILED: {filename} \u2192 {str(ex)[:100]}")
                    error_row = {
                        "Vendor Inv No": filename,
                        "Amount": "ERROR",
                        "Organization": "ERROR",
                        "Ref No": "FAILED",
                        "_Flag": "\u274c",
                        "_id": idx,
                        "_HasError": True,
                        "_ErrorDetail": str(ex),
                    }
                    processed_results.append(error_row)
                    
                    batch_log_entries.append({
                        "index": idx + 1,
                        "total": total,
                        "filename": filename,
                        "status": "\u274c",
                        "pdf_type": pdf_type,
                        "pdf_chars": pdf_chars,
                        "error": str(ex)
                    })

                    self.after(0, self._add_to_tree,
                              "ERROR", filename, "ERROR", "FAILED", "\u274c")

                # Small delay between invoices to respect RPM limits
                if idx < total - 1:
                    time.sleep(2)

            # Done
            self.after(0, lambda: self.progress.configure(value=total))
            self._update_status(f"Done. {success_count} OK, {warning_count} warnings, {error_count} errors.")
            print(f"\n{'='*60}")
            print(f"  BATCH COMPLETE: {success_count} \u2713 | {warning_count} \u26a0 | {error_count} \u274c")
            print(f"  API Key used: #{CURRENT_KEY_INDEX + 1} of {len(API_KEY_LIST)}")
            print(f"{'='*60}\n")
            self.after(0, lambda: self.lbl_summary.config(
                text=f"\u2713 {success_count}   \u26a0 {warning_count}   \u274c {error_count}"))
            self.after(0, lambda: self.btn_process.config(state=tk.NORMAL))
            self.after(0, lambda: self.btn_export.config(state=tk.NORMAL))

            # Show success popup if all invoices extracted cleanly
            if error_count == 0 and warning_count == 0 and success_count == total:
                self.after(0, lambda: messagebox.showinfo(
                    "Extraction Complete",
                    f"All {total} invoice(s) extracted successfully!\n\n"
                    f"You can now review the data and export to CSV."))

        except Exception as general_err:
            print(f"Processing Thread Error: {general_err}")
            self._update_status(f"Fatal error: {general_err}")
            self.after(0, lambda: self.btn_process.config(state=tk.NORMAL))

    def _update_status(self, text):
        self.after(0, lambda: self.lbl_status.config(text=text))

    def _add_to_tree(self, org, inv, amt, job, flag):
        row_count = len(self.tree.get_children())
        tag = 'evenrow' if row_count % 2 == 0 else 'oddrow'
        self.tree.insert("", tk.END, values=(org, inv, amt, job, flag), tags=(tag,))

    # --- CSV Export ---

    def export_csv(self):
        if not processed_results:
            return

        # Count exportable rows
        exportable = [r for r in processed_results if not r.get('_HasError')]
        if not exportable:
            messagebox.showwarning("No Data", "No successfully processed invoices to export.")
            return

        export_dir = os.path.join(os.path.abspath("."), "CSV Output")
        os.makedirs(export_dir, exist_ok=True)

        fpath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialdir=export_dir,
            initialfile=f"CFS_Export_{datetime.datetime.now().strftime('%d%b%y_%H%M%S')}.csv",
            title="Export CSV"
        )

        if not fpath:
            return

        columns = [
            "Entry Date", "Posting Date", "Organization", "Organization Branch", "Vendor Inv No",
            "Vendor Inv Date", "Currency", "ExchRate", "Narration", "Due Date",
            "Charge or GL", "Charge or GL Name", "Charge or GL Amount", "DR or CR", "Cost Center",
            "Branch", "Charge Narration", "TaxGroup", "Tax Type", "SAC or HSN",
            "Taxcode1", "Taxcode1 Amt", "Taxcode2", "Taxcode2 Amt", "Taxcode3", "Taxcode3 Amt",
            "Taxcode4", "Taxcode4 Amt", "Avail Tax Credit", "LOB", "Ref Type", "Ref No",
            "Amount", "Start Date", "End Date", "WH Tax Code", "WH Tax Percentage", "WH Tax Taxable",
            "WH Tax Amount", "Round Off", "CC Code"
        ]

        try:
            with open(fpath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=columns, extrasaction='ignore')
                writer.writeheader()
                for row in exportable:
                    writer.writerow(row)

            # --- Write Log File ---
            try:
                log_filename = f"CFS_Log_{datetime.datetime.now().strftime('%d%b%y_%H%M%S')}.txt"
                log_path = os.path.join(os.path.dirname(fpath), log_filename)
                write_batch_log(log_path, batch_log_meta, batch_log_entries)
            except Exception as le:
                print(f"Failed to write log file: {le}")

            messagebox.showinfo("Success", f"Exported {len(exportable)} rows to:\n{fpath}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))


def write_batch_log(log_path, meta, entries):
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write("─────────────────────────────────────────────────────────────\n")
        f.write("CFS Invoice → Logisys CSV Tool — Processing Log\n")
        f.write(f"Date: {meta.get('date')}\n")
        f.write(f"PDFs processed: {meta.get('total')}\n")
        f.write(f"Job Registry: {meta.get('registry')}\n")
        f.write("─────────────────────────────────────────────────────────────\n\n")

        for e in entries:
            f.write("================================================================================\n")
            f.write(f"[{e['index']}/{e['total']}] {e['filename']:<60} {e['status']}\n")
            f.write("================================================================================\n")
            f.write(f"PDF Type:   {e['pdf_type']} ({e['pdf_chars']} chars via pdfplumber)\n")
            if e.get("error"):
                f.write(f"Error:      {e['error']}\n\n")
                continue
            
            f.write("\n--- Gemini Extracted ---\n")
            raw = e.get("gemini_raw", {})
            f.write(f"vendor_name:          {raw.get('vendor_name') or 'null'}\n")
            f.write(f"invoice_number:       {raw.get('invoice_number') or 'null'}\n")
            f.write(f"invoice_date:         {raw.get('invoice_date') or 'null'}\n")
            f.write(f"hbl_number:           {raw.get('hbl_number') or 'null'}\n")
            f.write(f"mbl_number:           {raw.get('mbl_number') or 'null'}\n")
            f.write(f"boe_number:           {raw.get('boe_number') or 'null'}\n")
            f.write(f"total_invoice_amount: {raw.get('total_invoice_amount') or 'null'}\n")

            f.write("\n--- Mapping & Lookup ---\n")
            if e.get('org_match_keyword'):
                f.write(f"Organization:   {e['org_name']}  (matched keyword: \"{e['org_match_keyword']}\")\n")
            else:
                f.write(f"Organization:   {e['org_name']}  ⚠ (no keyword match found)\n")
            
            if e['job_match_type'] == "NONE":
                f.write(f"Job No:         NOT FOUND                    ⚠ (HBL {raw.get('hbl_number') or 'null'} not in registry, BOE {raw.get('boe_number') or 'null'} not in registry)\n")
            else:
                f.write(f"Job No:         {e['job_no']:<22} (matched {e['job_match_type']}: {e['job_match_value']})\n")

            f.write("\n--- Final CSV Values ---\n")
            f_row = e.get("final_row", {})
            f.write(f"Vendor Inv No:  {f_row.get('Vendor Inv No') or 'null'}\n")
            f.write(f"Vendor Inv Date: {f_row.get('Vendor Inv Date') or 'null'}\n")
            f.write(f"Amount:         {f_row.get('Amount') or 'null'}\n")
            f.write(f"Ref No:         {f_row.get('Ref No') or 'null'}\n")
            f.write(f"Narration:      {f_row.get('Narration') or 'null'}\n\n")

if __name__ == "__main__":
    app = App()
    app.mainloop()