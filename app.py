import streamlit as st
import json, re, os, tempfile, copy
from pathlib import Path
from datetime import datetime
from groq import Groq
import fitz
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from docx import Document as DocxDocument
import openpyxl
import pandas as pd

st.set_page_config(page_title="SeaBotics.ai", page_icon="S", layout="wide")

st.markdown("""
<style>
#MainMenu, footer, header { visibility: hidden; }
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
.field-label { font-size:0.68rem; color:#94a3b8; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:3px; }
.field-value { font-size:0.88rem; color:#1e293b; font-weight:500; background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; padding:6px 10px; min-height:32px; }
.field-empty { font-size:0.82rem; color:#cbd5e1; font-style:italic; background:#f8fafc; border:1px dashed #e2e8f0; border-radius:6px; padding:6px 10px; min-height:32px; }
.status-matched { background:#dcfce7; color:#166534; border-radius:12px; padding:2px 10px; font-size:0.7rem; font-weight:600; display:inline-block; }
.status-verify  { background:#fef9c3; color:#854d0e; border-radius:12px; padding:2px 10px; font-size:0.7rem; font-weight:600; display:inline-block; }
.status-manual  { background:#fee2e2; color:#991b1b; border-radius:12px; padding:2px 10px; font-size:0.7rem; font-weight:600; display:inline-block; }
.missing-tag { background:#fef2f2; color:#7f1d1d; border-radius:4px; padding:2px 7px; font-size:0.7rem; font-family:monospace; display:inline-block; margin:2px; }
.section-label { font-size:0.68rem; font-weight:700; letter-spacing:0.1em; text-transform:uppercase; color:#64748b; border-left:3px solid #1e3a5f; padding-left:8px; margin:1rem 0 0.5rem 0; }
.table-header { background:#f1f5f9; font-size:0.72rem; font-weight:600; color:#475569; padding:6px 8px; border-radius:4px 4px 0 0; }
div[data-testid="stDataFrame"] { border:1px solid #e2e8f0; border-radius:8px; overflow:hidden; }
</style>
""", unsafe_allow_html=True)

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
MODEL_NAME   = "llama-3.3-70b-versatile"

# ── SECTION COLUMN DEFINITIONS ────────────────────────────────
# Each multi-row section has defined columns
SECTION_COLUMNS = {
    "Medical": [
        "Document Type", "Document Name", "Certificate Number",
        "Date of Issue", "Date of Expiry"
    ],
    "Education": [
        "Institution Name", "Country", "Level of Education",
        "Field of Study", "Start Date", "Completion Date"
    ],
    "Courses": [
        "Document Name", "Certificate Number", "Type",
        "Place of Issue", "Date of Issue", "Date of Expiry"
    ],
    "Licenses and Certification": [
        "Document Name", "Certificate Number", "Type",
        "Issuing Authority", "Place of Issue", "Date of Issue", "Date of Expiry"
    ],
    "Travel Documents": [
        "Document Type", "Document Name", "Certificate Number",
        "Place of Issue", "Date of Issue", "Date of Expiry"
    ],
    "Sea Service Experience": [
        "Company", "Rank", "Vessel Name", "Vessel Type", "Voyage Type",
        "IMO Number", "Flag / Country", "GRT", "DWT", "BHP",
        "Main Engine Type", "Main Engine KW",
        "Sign On Port", "Sign Off Port", "Start Date", "End Date",
        "Reason for Sign Off"
    ],
    "Bank Details": [
        "Bank Name", "Branch Name", "Account Type",
        "Account Name", "Account Number", "IFSC / SWIFT Code"
    ],
}

# Personal info stays as single-record (key-value)
PERSONAL_FIELDS = [
    ("person_name","Full Name"), ("surname","Surname"), ("email","Email"),
    ("gender","Gender"), ("contact_number","Contact Number"),
    ("date_of_birth","Date of Birth"), ("nationality","Nationality"),
    ("blood_group","Blood Group"), ("rank","Current Rank"),
    ("identification_mark","Identification Mark"), ("availability_date","Availability Date"),
    ("address","Address"), ("country","Country"), ("state","State"),
    ("zip_code","ZIP / PIN Code"), ("domestic_airport","Domestic Airport"),
    ("international_airport","International Airport"),
    ("passport_number","Passport Number"), ("cdc_number","CDC / Seaman's Book Number"),
    ("height","Height"), ("weight","Weight"),
    ("boiler_suit_size","Boiler Suit Size"), ("safety_shoe_size","Safety Shoe Size"),
    ("shirt_size","Shirt Size"), ("trouser_size","Trouser Size"),
]

ALL_SECTIONS = ["Personal Information"] + list(SECTION_COLUMNS.keys())

# ── DOC TYPE → SECTION MAPPING ────────────────────────────────
DOC_TYPE_TO_SECTION = {
    "passport": "Travel Documents",
    "visa": "Travel Documents",
    "cdc": "Travel Documents",
    "continuous discharge": "Travel Documents",
    "seaman": "Travel Documents",
    "indos": "Travel Documents",
    "sid": "Travel Documents",
    "certificate of competency": "Licenses and Certification",
    "certificate of proficiency": "Licenses and Certification",
    "endorsement": "Licenses and Certification",
    "flag endorsement": "Licenses and Certification",
    "stcw": "Courses",
    "training": "Courses",
    "advanced fire": "Courses",
    "basic safety": "Courses",
    "proficiency": "Courses",
    "drug": "Medical",
    "alcohol": "Medical",
    "peme": "Medical",
    "medical fitness": "Medical",
    "lft": "Medical",
    "treadmill": "Medical",
    "yellow fever": "Medical",
    "covid": "Medical",
    "vaccine": "Medical",
    "vaccination": "Medical",
    "degree": "Education",
    "diploma": "Education",
    "bachelor": "Education",
    "master of": "Education",
    "pg": "Education",
    "bank": "Bank Details",
    "account": "Bank Details",
}

def guess_section(doc_type):
    if not doc_type: return None
    dt = doc_type.lower()
    for key, section in DOC_TYPE_TO_SECTION.items():
        if key in dt: return section
    return None

def init_state():
    if "profile" not in st.session_state:
        st.session_state["profile"] = {}
    if "sections" not in st.session_state:
        # Each section = list of row dicts
        st.session_state["sections"] = {s: [] for s in SECTION_COLUMNS}
    if "doc_tank" not in st.session_state:
        st.session_state["doc_tank"] = []
    if "popup" not in st.session_state:
        st.session_state["popup"] = None
    if "page" not in st.session_state:
        st.session_state["page"] = "My Profile"

init_state()

# ── SYSTEM PROMPT ─────────────────────────────────────────────
SYSTEM_PROMPT = """You are SeaBotics.ai — maritime crew ERP document extraction engine.

Extract ALL information. For resumes/CVs especially extract EVERY record from tables.

DOCUMENT TYPES: Passport, CDC, INDOS, SID, Visa, Certificate of Competency, Certificate of Proficiency, Endorsement, Flag Endorsement, Drug & Alcohol Test, PEME, LFT, Yellow Fever Vaccine, COVID Vaccine, Training Certificate, STCW, Degree, Bank Account, CV, Resume

DATE FORMAT: DD-MMM-YYYY (e.g. 15-Jul-2023). LIFETIME stays as LIFETIME.
CONFIDENCE: 1.0=exact 0.9=strong 0.75=inferred 0.6=weak 0.0=not found

CRITICAL FOR RESUMES: Extract EVERY row from sea service tables, EVERY document from document tables, EVERY course/certificate listed.

Return ONLY valid JSON, no markdown.

Schema:
{
  "document_type": "CV/Resume OR specific doc type",
  "is_resume": true or false,
  "personal": {
    "full_name": "", "surname": "", "email": "", "gender": "",
    "contact_number": "", "date_of_birth": "", "nationality": "",
    "blood_group": "", "current_rank": "", "identification_mark": "",
    "availability_date": "", "address": "", "country": "", "state": "",
    "zip_code": "", "domestic_airport": "", "international_airport": "",
    "passport_number": "", "cdc_number": "", "height": "", "weight": "",
    "boiler_suit_size": "", "safety_shoe_size": "", "shirt_size": "", "trouser_size": ""
  },
  "sea_service": [
    {
      "company": "", "rank": "", "vessel_name": "", "vessel_type": "",
      "voyage_type": "", "imo_number": "", "flag_country": "",
      "grt": "", "dwt": "", "bhp": "", "main_engine_type": "", "main_engine_kw": "",
      "sign_on_port": "", "sign_off_port": "",
      "start_date": "", "end_date": "", "reason_sign_off": ""
    }
  ],
  "travel_documents": [
    {
      "document_type": "", "document_name": "", "certificate_number": "",
      "place_of_issue": "", "date_of_issue": "", "date_of_expiry": ""
    }
  ],
  "medical": [
    {
      "document_type": "", "document_name": "", "certificate_number": "",
      "date_of_issue": "", "date_of_expiry": ""
    }
  ],
  "education": [
    {
      "institution_name": "", "country": "", "education_level": "",
      "field_of_study": "", "start_date": "", "completion_date": ""
    }
  ],
  "courses": [
    {
      "document_name": "", "certificate_number": "", "type": "",
      "place_of_issue": "", "date_of_issue": "", "date_of_expiry": ""
    }
  ],
  "licenses_certifications": [
    {
      "document_name": "", "certificate_number": "", "type": "",
      "issuing_authority": "", "place_of_issue": "",
      "date_of_issue": "", "date_of_expiry": ""
    }
  ],
  "bank_details": [
    {
      "bank_name": "", "branch_name": "", "account_type": "",
      "account_name": "", "account_number": "", "ifsc_swift": ""
    }
  ],
  "missing_fields": [],
  "validation_flags": [],
  "raw_text_summary": ""
}"""

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}

def preprocess(img):
    if img.mode != "RGB": img = img.convert("RGB")
    w, h = img.size
    if w < 1000: img = img.resize((1000, int(h*1000/w)), Image.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    img = ImageEnhance.Contrast(img).enhance(1.5)
    return img.convert("L")

def extract_text(path):
    ext = Path(path).suffix.lower()
    if ext in IMAGE_EXT:
        img = preprocess(Image.open(path))
        return max([pytesseract.image_to_string(img, config=c)
                    for c in ["--psm 3","--psm 6","--psm 11"]], key=len)
    if ext == ".pdf":
        doc = fitz.open(path); pages = []
        for p in doc:
            t = p.get_text("text").strip()
            if len(t) > 50: pages.append(t)
            else:
                pix = p.get_pixmap(dpi=250)
                img = preprocess(Image.frombytes("RGB",[pix.width,pix.height],pix.samples))
                pages.append(pytesseract.image_to_string(img))
        doc.close()
        return "\n--- PAGE BREAK ---\n".join(pages)
    if ext == ".docx":
        return "\n".join(p.text for p in DocxDocument(path).paragraphs if p.text.strip())
    if ext == ".txt":
        with open(path) as f: return f.read()
    if ext in [".xlsx",".xls"]:
        wb = openpyxl.load_workbook(path, data_only=True)
        return "\n".join("  |  ".join(str(c) for c in r if c)
                         for s in wb.worksheets for r in s.iter_rows(values_only=True))
    raise ValueError(f"Unsupported: {ext}")

def call_llm(text):
    client = Groq(api_key=GROQ_API_KEY)
    if len(text) > 12000: text = text[:12000]
    if not text.strip(): return None
    for attempt in range(3):
        try:
            r = client.chat.completions.create(
                model=MODEL_NAME, temperature=0.1, max_tokens=5000,
                messages=[
                    {"role":"system","content":SYSTEM_PROMPT},
                    {"role":"user","content":f"Extract all information from this document:\n\n{text}"}
                ])
            raw = r.choices[0].message.content.strip()
            if not raw: continue
            raw = re.sub(r"^```[a-z]*\n?","",raw)
            raw = re.sub(r"\n?```$","",raw)
            start = raw.find("{"); end = raw.rfind("}")+1
            if start == -1 or end == 0: continue
            return json.loads(raw[start:end])
        except json.JSONDecodeError: continue
        except Exception as e:
            st.error(f"API error: {e}"); return None
    st.error("Pipeline failed after 3 attempts. Please try again.")
    return None

def conf_color(c):
    if c >= 0.90: return "#22c55e"
    if c >= 0.60: return "#f59e0b"
    if c >  0.00: return "#ef4444"
    return "#e2e8f0"

# ── RESULT → SECTION ROWS ─────────────────────────────────────
def result_to_rows(result):
    """Convert LLM result into section rows ready to append."""
    rows = {s: [] for s in SECTION_COLUMNS}

    # Sea Service
    for entry in result.get("sea_service", []):
        if any(v for v in entry.values()):
            rows["Sea Service Experience"].append({
                "Company":          entry.get("company",""),
                "Rank":             entry.get("rank",""),
                "Vessel Name":      entry.get("vessel_name",""),
                "Vessel Type":      entry.get("vessel_type",""),
                "Voyage Type":      entry.get("voyage_type",""),
                "IMO Number":       entry.get("imo_number",""),
                "Flag / Country":   entry.get("flag_country",""),
                "GRT":              entry.get("grt",""),
                "DWT":              entry.get("dwt",""),
                "BHP":              entry.get("bhp",""),
                "Main Engine Type": entry.get("main_engine_type",""),
                "Main Engine KW":   entry.get("main_engine_kw",""),
                "Sign On Port":     entry.get("sign_on_port",""),
                "Sign Off Port":    entry.get("sign_off_port",""),
                "Start Date":       entry.get("start_date",""),
                "End Date":         entry.get("end_date",""),
                "Reason for Sign Off": entry.get("reason_sign_off",""),
            })

    # Travel Documents
    for entry in result.get("travel_documents", []):
        if any(v for v in entry.values()):
            rows["Travel Documents"].append({
                "Document Type":       entry.get("document_type",""),
                "Document Name":       entry.get("document_name",""),
                "Certificate Number":  entry.get("certificate_number",""),
                "Place of Issue":      entry.get("place_of_issue",""),
                "Date of Issue":       entry.get("date_of_issue",""),
                "Date of Expiry":      entry.get("date_of_expiry",""),
            })

    # Medical
    for entry in result.get("medical", []):
        if any(v for v in entry.values()):
            rows["Medical"].append({
                "Document Type":      entry.get("document_type",""),
                "Document Name":      entry.get("document_name",""),
                "Certificate Number": entry.get("certificate_number",""),
                "Date of Issue":      entry.get("date_of_issue",""),
                "Date of Expiry":     entry.get("date_of_expiry",""),
            })

    # Education
    for entry in result.get("education", []):
        if any(v for v in entry.values()):
            rows["Education"].append({
                "Institution Name":  entry.get("institution_name",""),
                "Country":           entry.get("country",""),
                "Level of Education": entry.get("education_level",""),
                "Field of Study":    entry.get("field_of_study",""),
                "Start Date":        entry.get("start_date",""),
                "Completion Date":   entry.get("completion_date",""),
            })

    # Courses
    for entry in result.get("courses", []):
        if any(v for v in entry.values()):
            rows["Courses"].append({
                "Document Name":      entry.get("document_name",""),
                "Certificate Number": entry.get("certificate_number",""),
                "Type":               entry.get("type",""),
                "Place of Issue":     entry.get("place_of_issue",""),
                "Date of Issue":      entry.get("date_of_issue",""),
                "Date of Expiry":     entry.get("date_of_expiry",""),
            })

    # Licenses and Certification
    for entry in result.get("licenses_certifications", []):
        if any(v for v in entry.values()):
            rows["Licenses and Certification"].append({
                "Document Name":      entry.get("document_name",""),
                "Certificate Number": entry.get("certificate_number",""),
                "Type":               entry.get("type",""),
                "Issuing Authority":  entry.get("issuing_authority",""),
                "Place of Issue":     entry.get("place_of_issue",""),
                "Date of Issue":      entry.get("date_of_issue",""),
                "Date of Expiry":     entry.get("date_of_expiry",""),
            })

    # Bank Details
    for entry in result.get("bank_details", []):
        if any(v for v in entry.values()):
            rows["Bank Details"].append({
                "Bank Name":        entry.get("bank_name",""),
                "Branch Name":      entry.get("branch_name",""),
                "Account Type":     entry.get("account_type",""),
                "Account Name":     entry.get("account_name",""),
                "Account Number":   entry.get("account_number",""),
                "IFSC / SWIFT Code": entry.get("ifsc_swift",""),
            })

    return rows

def personal_to_profile(personal_data):
    """Map personal dict from LLM to profile key-value store."""
    mapping = {
        "full_name": "Full Name", "surname": "Surname",
        "email": "Email Address", "gender": "Gender",
        "contact_number": "Phone Number", "date_of_birth": "Date of Birth",
        "nationality": "Nationality", "blood_group": "Blood Group",
        "current_rank": "Rank / Designation", "identification_mark": "Identification Mark",
        "availability_date": "Availability Date", "address": "Residential Address",
        "country": "Country", "state": "State", "zip_code": "ZIP / PIN Code",
        "domestic_airport": "Domestic Airport", "international_airport": "International Airport",
        "passport_number": "Passport Number", "cdc_number": "CDC / Seaman's Book Number",
        "height": "Height", "weight": "Weight",
        "boiler_suit_size": "Boiler Suit Size", "safety_shoe_size": "Safety Shoe Size",
        "shirt_size": "Shirt Size", "trouser_size": "Trouser Size",
    }
    result = {}
    for llm_key, profile_key in mapping.items():
        val = personal_data.get(llm_key,"")
        if val: result[profile_key] = val
    return result

# ── POPUP ─────────────────────────────────────────────────────
def render_popup():
    popup = st.session_state.get("popup")
    if not popup: return

    result        = popup.get("result", {})
    filename      = popup.get("filename", "document")
    source        = popup.get("source", "section")
    target_section = popup.get("target_section", None)
    is_resume     = result.get("is_resume", False)
    doc_type      = result.get("document_type", "—")
    personal      = result.get("personal", {})
    missing       = result.get("missing_fields", [])
    flags         = result.get("validation_flags", [])

    # Build a flat list of ALL extracted fields for the 5-column table
    # regardless of which section they belong to
    flat_fields = {}

    # Personal fields
    for k, v in personal.items():
        if v:
            flat_fields[f"personal__{k}"] = {
                "doc_label": k.replace("_", " ").title(),
                "value":     v,
                "confidence": 0.9,
                "section":   "Personal Information",
            }

    # All array sections
    section_map_llm = {
        "sea_service":             "Sea Service Experience",
        "travel_documents":        "Travel Documents",
        "medical":                 "Medical",
        "education":               "Education",
        "courses":                 "Courses",
        "licenses_certifications": "Licenses and Certification",
        "bank_details":            "Bank Details",
    }
    for llm_key, section_name in section_map_llm.items():
        for ridx, entry in enumerate(result.get(llm_key, [])):
            for field_key, val in entry.items():
                if val:
                    flat_fields[f"{llm_key}__{ridx}__{field_key}"] = {
                        "doc_label":  field_key.replace("_", " ").title(),
                        "value":      val,
                        "confidence": 0.85,
                        "section":    section_name,
                        "row_idx":    ridx,
                    }

    # All org field options for dropdown
    ALL_ORG_OPTIONS = sorted(set([
        "Full Name", "Surname", "Email Address", "Gender", "Phone Number",
        "Date of Birth", "Nationality", "Blood Group", "Rank / Designation",
        "Identification Mark", "Availability Date", "Residential Address",
        "Country", "State", "ZIP / PIN Code", "Domestic Airport",
        "International Airport", "Passport Number", "CDC / Seaman's Book Number",
        "Height", "Weight", "Boiler Suit Size", "Safety Shoe Size",
        "Shirt Size", "Trouser Size",
        "Document Type", "Document Name", "Document / Certificate Number",
        "Issuing Authority", "Place of Issue", "Date of Issue", "Date of Expiry",
        "Institution Name", "Country (Education)", "Level of Education",
        "Field of Study", "Start Date", "Completion Date",
        "Bank Name", "Branch Name", "Account Type", "Account Name",
        "Account Number", "IFSC / SWIFT Code",
        "Company Name", "IMO Number", "Vessel Name", "Vessel Type",
        "Flag State", "Voyage Type", "Sign On Date", "Sign On Port",
        "Sign Off Date", "Sign Off Port", "GRT", "DWT", "BHP",
        "Main Engine Type / Model", "Main Engine KW",
        "Reason for Sign Off", "Professional Summary",
        "Total Sea Experience", "Last Employer",
        "Certifications Held", "Skills",
        "Other / Unclassified"
    ]))

    # Org field pre-mapping based on doc_label
    PRE_MAP = {
        "full name": "Full Name", "surname": "Surname",
        "email": "Email Address", "gender": "Gender",
        "contact number": "Phone Number", "date of birth": "Date of Birth",
        "nationality": "Nationality", "blood group": "Blood Group",
        "current rank": "Rank / Designation", "rank": "Rank / Designation",
        "identification mark": "Identification Mark",
        "availability date": "Availability Date",
        "address": "Residential Address", "country": "Country",
        "state": "State", "zip code": "ZIP / PIN Code",
        "domestic airport": "Domestic Airport",
        "international airport": "International Airport",
        "passport number": "Passport Number",
        "cdc number": "CDC / Seaman's Book Number",
        "height": "Height", "weight": "Weight",
        "boiler suit size": "Boiler Suit Size",
        "safety shoe size": "Safety Shoe Size",
        "shirt size": "Shirt Size", "trouser size": "Trouser Size",
        "document type": "Document Type", "document name": "Document Name",
        "certificate number": "Document / Certificate Number",
        "issuing authority": "Issuing Authority",
        "place of issue": "Place of Issue",
        "date of issue": "Date of Issue", "date of expiry": "Date of Expiry",
        "institution name": "Institution Name",
        "education level": "Level of Education",
        "field of study": "Field of Study",
        "start date": "Start Date", "completion date": "Completion Date",
        "bank name": "Bank Name", "branch name": "Branch Name",
        "account type": "Account Type", "account name": "Account Name",
        "account number": "Account Number",
        "ifsc swift": "IFSC / SWIFT Code",
        "company": "Company Name", "imo number": "IMO Number",
        "vessel name": "Vessel Name", "vessel type": "Vessel Type",
        "flag country": "Flag State", "voyage type": "Voyage Type",
        "sign on port": "Sign On Port", "sign off port": "Sign Off Port",
        "start date": "Start Date", "end date": "Sign Off Date",
        "grt": "GRT", "dwt": "DWT", "bhp": "BHP",
        "main engine type": "Main Engine Type / Model",
        "main engine kw": "Main Engine KW",
        "reason sign off": "Reason for Sign Off",
    }

    title = "CV / Resume Review" if is_resume else "Document Review"

    with st.expander(f"Review & Approve — {title}: {filename}", expanded=True):

        # Summary metrics
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Document Type",   doc_type)
        m2.metric("CV / Resume",     "Yes" if is_resume else "No")
        m3.metric("Fields Extracted", len(flat_fields))
        high = len([f for f in flat_fields.values() if f["confidence"] >= 0.90])
        med  = len([f for f in flat_fields.values() if 0.6 <= f["confidence"] < 0.90])
        m4.metric("High Confidence", high)
        m5.metric("Verify",          med)

        if is_resume:
            section_counts = {}
            for f in flat_fields.values():
                s = f["section"]
                section_counts[s] = section_counts.get(s, 0) + 1
            if section_counts:
                st.markdown("**Sections detected in this document:**")
                ic = st.columns(min(len(section_counts), 4))
                for i, (sec, cnt) in enumerate(section_counts.items()):
                    with ic[i % 4]: st.info(f"**{sec}**: {cnt} field(s)")

        st.markdown("---")

        # 5-column table header
        h1, h2, h3, h4, h5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
        h1.markdown("**Document Field**")
        h2.markdown("**Extracted Value**")
        h3.markdown("**Confidence**")
        h4.markdown("**Status**")
        h5.markdown("**Org Database Field**")
        st.markdown("<hr style='margin:4px 0 8px 0;border-color:#e2e8f0;'>",
                    unsafe_allow_html=True)

        edited_values = {}
        edited_org    = {}

        # Group by section for visual separation
        current_section = None
        for fid, fdata in flat_fields.items():
            sec = fdata["section"]
            if sec != current_section:
                current_section = sec
                st.markdown(
                    f'<div class="section-label">{sec}</div>',
                    unsafe_allow_html=True
                )

            val   = fdata["value"]
            conf  = fdata["confidence"]
            label = fdata["doc_label"]
            color = conf_color(conf)

            if conf >= 0.90:
                status_txt, status_cls = "Matched",      "status-matched"
            elif conf >= 0.60:
                status_txt, status_cls = "Verify",       "status-verify"
            else:
                status_txt, status_cls = "Manual Entry", "status-manual"

            # Pre-map org field
            org_default = PRE_MAP.get(label.lower())
            options     = ["— Select org field —"] + ALL_ORG_OPTIONS
            default_idx = (ALL_ORG_OPTIONS.index(org_default) + 1) \
                          if org_default and org_default in ALL_ORG_OPTIONS else 0

            c1, c2, c3, c4, c5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
            with c1:
                st.markdown(
                    f'<div style="font-size:0.82rem;color:#475569;padding-top:8px;">'
                    f'{label}</div>',
                    unsafe_allow_html=True
                )
            with c2:
                edited_values[fid] = st.text_input(
                    f"v_{fid}", value=val,
                    key=f"v_popup_{fid}",
                    label_visibility="collapsed"
                )
            with c3:
                st.markdown(
                    f'<div style="padding-top:6px;">'
                    f'<div style="background:#f1f5f9;border-radius:4px;height:6px;width:70px;">'
                    f'<div style="width:{int(conf*100)}%;height:6px;'
                    f'background:{color};border-radius:4px;"></div></div>'
                    f'<div style="font-size:0.65rem;color:{color};margin-top:2px;">'
                    f'{int(conf*100)}%</div></div>',
                    unsafe_allow_html=True
                )
            with c4:
                st.markdown(
                    f'<div style="padding-top:8px;">'
                    f'<span class="{status_cls}">{status_txt}</span></div>',
                    unsafe_allow_html=True
                )
            with c5:
                sel = st.selectbox(
                    f"o_{fid}", options=options,
                    index=default_idx,
                    key=f"o_popup_{fid}",
                    label_visibility="collapsed"
                )
                edited_org[fid] = sel if sel != "— Select org field —" else ""

        if flags:
            st.markdown("---")
            for f in flags: st.warning(f)
        if missing:
            st.markdown(
                " ".join(f'<span class="missing-tag">{m}</span>' for m in missing),
                unsafe_allow_html=True
            )

        st.markdown("---")
        ba, bb, bc = st.columns([2, 2, 1])

        with ba:
            if st.button("Approve and Populate Profile", type="primary",
                         use_container_width=True, key="approve_popup"):

                # Rebuild result from edited values and org assignments
                # then push into sections the same way as before
                new_personal = dict(result.get("personal", {}))
                new_result   = copy.deepcopy(result)

                section_map_llm = {
                    "sea_service":             "Sea Service Experience",
                    "travel_documents":        "Travel Documents",
                    "medical":                 "Medical",
                    "education":               "Education",
                    "courses":                 "Courses",
                    "licenses_certifications": "Licenses and Certification",
                    "bank_details":            "Bank Details",
                }

                for fid, val in edited_values.items():
                    parts = fid.split("__")
                    if parts[0] == "personal":
                        new_personal[parts[1]] = val
                    elif len(parts) == 3:
                        llm_key, ridx, field_key = parts[0], int(parts[1]), parts[2]
                        if llm_key in new_result and ridx < len(new_result[llm_key]):
                            new_result[llm_key][ridx][field_key] = val

                # Populate personal info
                profile_data = personal_to_profile(new_personal)
                for k, v in profile_data.items():
                    st.session_state["profile"][k] = v

                # Cross-populate passport/CDC into personal
                for td in new_result.get("travel_documents", []):
                    dt = (td.get("document_type","") + td.get("document_name","")).lower()
                    if "passport" in dt and td.get("certificate_number"):
                        st.session_state["profile"]["Passport Number"] = td["certificate_number"]
                    if ("cdc" in dt or "seaman" in dt) and td.get("certificate_number"):
                        st.session_state["profile"]["CDC / Seaman's Book Number"] = td["certificate_number"]

                # Convert to rows and append to sections
                rows_to_add = result_to_rows(new_result)
                for section, rows in rows_to_add.items():
                    for row in rows:
                        if any(v for v in row.values()):
                            st.session_state["sections"][section].append(row)

                # Doc tank
                st.session_state["doc_tank"].append({
                    "timestamp":          datetime.now().strftime("%d-%b-%Y %H:%M"),
                    "filename":           filename,
                    "source":             source,
                    "document_type":      doc_type,
                    "is_resume":          is_resume,
                    "sections_populated": [s for s, r in rows_to_add.items() if r],
                    "personal_fields":    len(profile_data),
                    "total_rows":         sum(len(r) for r in rows_to_add.values()),
                    "raw_result":         result,
                })

                st.session_state["popup"] = None
                st.success("Profile populated successfully.")
                st.rerun()

        with bb:
            st.download_button(
                "Download Raw JSON",
                data=json.dumps(result, indent=2),
                file_name=f"{Path(filename).stem}_raw.json",
                mime="application/json",
                key="dl_popup",
                use_container_width=True
            )
        with bc:
            if st.button("Discard", use_container_width=True, key="discard_popup"):
                st.session_state["popup"] = None
                st.rerun()

# ── HEADER ────────────────────────────────────────────────────
col_brand, col_master = st.columns([3,1])
with col_brand:
    st.markdown("""
    <div style="background:linear-gradient(135deg,#0f172a,#1e3a5f);border-radius:12px;padding:16px 24px;margin-bottom:16px;">
        <div style="font-size:1.3rem;font-weight:700;color:#38bdf8;">SeaBotics.ai</div>
        <div style="font-size:0.72rem;color:#94a3b8;margin-top:2px;">Crew Profile Management — v4.0 with Resume Pipeline</div>
    </div>""", unsafe_allow_html=True)
with col_master:
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown("**Master Upload**")
    st.caption("Individual docs, merged PDFs, or CVs/Resumes")
    master_file = st.file_uploader("master", type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                   label_visibility="collapsed", key="master_uploader")
    if master_file and st.button("Run Pipeline", use_container_width=True, key="run_master"):
        with st.spinner("Extracting and classifying..."):
            ext = Path(master_file.name).suffix.lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(master_file.read()); tmp_path = tmp.name
            try:
                text = extract_text(tmp_path)
                result = call_llm(text)
                if result:
                    st.session_state["popup"] = {
                        "key":"master","result":result,
                        "filename":master_file.name,"source":"master","target_section":None
                    }
                    st.rerun()
            except Exception as e: st.error(f"Error: {e}")

# ── NAV ───────────────────────────────────────────────────────
n1,n2,n3 = st.columns([1,1,6])
with n1:
    if st.button("My Profile", use_container_width=True,
                 type="primary" if st.session_state["page"]=="My Profile" else "secondary"):
        st.session_state["page"]="My Profile"; st.rerun()
with n2:
    if st.button("Doc Tank", use_container_width=True,
                 type="primary" if st.session_state["page"]=="Doc Tank" else "secondary"):
        st.session_state["page"]="Doc Tank"; st.rerun()

# ── POPUP ─────────────────────────────────────────────────────
if st.session_state.get("popup"):
    render_popup()

# ════════════════════════════════════════════════════════════
#  MY PROFILE
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "My Profile":

    # ── PERSONAL INFORMATION (single record) ──────────────────
    with st.expander("Personal Information", expanded=True):
        uc1, uc2 = st.columns([4,1])
        with uc1:
            sec_file = st.file_uploader("Upload personal doc", type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                        key="up_personal", label_visibility="collapsed")
        with uc2:
            st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
            if sec_file and st.button("Extract", key="run_personal", use_container_width=True):
                with st.spinner("Extracting..."):
                    ext = Path(sec_file.name).suffix.lower()
                    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                        tmp.write(sec_file.read()); tmp_path = tmp.name
                    try:
                        text = extract_text(tmp_path)
                        result = call_llm(text)
                        if result:
                            st.session_state["popup"] = {"key":"sec_personal","result":result,
                                "filename":sec_file.name,"source":"section","target_section":"Personal Information"}
                            st.rerun()
                    except Exception as e: st.error(f"Error: {e}")

        st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)
        profile = st.session_state["profile"]
        cols_per = 3
        for i in range(0, len(PERSONAL_FIELDS), cols_per):
            row_fields = PERSONAL_FIELDS[i:i+cols_per]
            cols = st.columns(cols_per)
            for j,(fk,fl) in enumerate(row_fields):
                with cols[j]:
                    # Look up by org field name
                    from_map = {
                        "person_name":"Full Name","surname":"Surname","email":"Email Address",
                        "gender":"Gender","contact_number":"Phone Number","date_of_birth":"Date of Birth",
                        "nationality":"Nationality","blood_group":"Blood Group","rank":"Rank / Designation",
                        "identification_mark":"Identification Mark","availability_date":"Availability Date",
                        "address":"Residential Address","country":"Country","state":"State",
                        "zip_code":"ZIP / PIN Code","domestic_airport":"Domestic Airport",
                        "international_airport":"International Airport","passport_number":"Passport Number",
                        "cdc_number":"CDC / Seaman's Book Number","height":"Height","weight":"Weight",
                        "boiler_suit_size":"Boiler Suit Size","safety_shoe_size":"Safety Shoe Size",
                        "shirt_size":"Shirt Size","trouser_size":"Trouser Size",
                    }
                    org_key = from_map.get(fk, fl)
                    val = profile.get(org_key,"")
                    st.markdown(f'<div class="field-label">{fl}</div>',unsafe_allow_html=True)
                    if val: st.markdown(f'<div class="field-value">{val}</div>',unsafe_allow_html=True)
                    else:   st.markdown(f'<div class="field-empty">—</div>',unsafe_allow_html=True)

        if profile:
            if st.button("Clear Personal Info", key="clear_personal"):
                st.session_state["profile"] = {}; st.rerun()

    # ── MULTI-ROW SECTIONS ────────────────────────────────────
    for section_name, cols_def in SECTION_COLUMNS.items():
        rows = st.session_state["sections"].get(section_name, [])
        row_count = len(rows)

        with st.expander(f"{section_name}  ({row_count} record{'s' if row_count!=1 else ''})", expanded=False):

            # Section upload
            uc1,uc2 = st.columns([4,1])
            with uc1:
                sec_file = st.file_uploader(f"Upload for {section_name}",
                    type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                    key=f"up_{section_name}", label_visibility="collapsed")
            with uc2:
                st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
                if sec_file and st.button("Extract", key=f"run_{section_name}", use_container_width=True):
                    with st.spinner("Extracting..."):
                        ext = Path(sec_file.name).suffix.lower()
                        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                            tmp.write(sec_file.read()); tmp_path = tmp.name
                        try:
                            text = extract_text(tmp_path)
                            result = call_llm(text)
                            if result:
                                st.session_state["popup"] = {
                                    "key":f"sec_{section_name.replace(' ','_')}",
                                    "result":result,"filename":sec_file.name,
                                    "source":"section","target_section":section_name
                                }
                                st.rerun()
                        except Exception as e: st.error(f"Error: {e}")

            st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)

            if not rows:
                st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No records yet — upload a document above to populate.</div>', unsafe_allow_html=True)
            else:
                # Display as a dataframe table
                df = pd.DataFrame(rows, columns=cols_def)
                # Fill missing cols
                for c in cols_def:
                    if c not in df.columns: df[c] = ""
                df = df[cols_def].fillna("")
                st.dataframe(df, use_container_width=True, hide_index=True)

                # Edit individual rows
                with st.expander("Edit records", expanded=False):
                    for ridx, row in enumerate(rows):
                        st.markdown(f"**Record {ridx+1}**")
                        rcols = st.columns(min(len(cols_def),3))
                        for cidx, col_name in enumerate(cols_def):
                            with rcols[cidx%3]:
                                st.markdown(f'<div class="field-label">{col_name}</div>',unsafe_allow_html=True)
                                new_val = st.text_input(f"edit_{section_name}_{ridx}_{col_name}",
                                    value=row.get(col_name,""), key=f"edit_{section_name}_{ridx}_{col_name}",
                                    label_visibility="collapsed")
                                st.session_state["sections"][section_name][ridx][col_name] = new_val
                        st.markdown("---")

                # Add blank row
                if st.button(f"+ Add blank row", key=f"add_{section_name}"):
                    st.session_state["sections"][section_name].append({c:"" for c in cols_def})
                    st.rerun()

                if st.button(f"Clear all {section_name} records", key=f"clear_{section_name}"):
                    st.session_state["sections"][section_name] = []; st.rerun()

# ════════════════════════════════════════════════════════════
#  DOC TANK
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "Doc Tank":
    st.markdown("### Doc Tank — Upload Log and Extraction Archive")
    st.markdown("<div style='font-size:0.85rem;color:#64748b;margin-bottom:16px;'>All uploads, extracted data and section mappings for this session.</div>", unsafe_allow_html=True)

    tank = st.session_state.get("doc_tank",[])

    if not tank:
        st.info("No documents processed yet.")
    else:
        s1,s2,s3,s4 = st.columns(4)
        s1.metric("Total Uploads",   len(tank))
        s2.metric("Resumes / CVs",   len([d for d in tank if d.get("is_resume")]))
        s3.metric("Master Pipeline", len([d for d in tank if d["source"]=="master"]))
        s4.metric("Section Uploads", len([d for d in tank if d["source"]=="section"]))
        st.markdown("---")

        for i, doc in enumerate(reversed(tank)):
            idx = len(tank)-1-i
            label = f"[{doc['timestamp']}]  {doc['filename']}  —  {doc['document_type']}"
            with st.expander(label, expanded=(i==0)):
                d1,d2,d3,d4 = st.columns(4)
                d1.metric("Uploaded",        doc["timestamp"])
                d2.metric("Type",            doc["document_type"])
                d3.metric("Personal Fields", doc.get("personal_fields",0))
                d4.metric("Section Records", doc.get("total_rows",0))

                secs = doc.get("sections_populated",[])
                if secs:
                    st.markdown("**Sections populated:** " + "  |  ".join(secs))

                raw = doc.get("raw_result",{})
                st.download_button("Download Raw JSON", data=json.dumps(raw,indent=2),
                    file_name=f"{Path(doc['filename']).stem}_raw.json",
                    mime="application/json", key=f"dlr_{idx}", use_container_width=False)

        st.markdown("---")
        st.download_button("Export Full Doc Tank", data=json.dumps(tank,indent=2,default=str),
            file_name="seabotics_doc_tank.json", mime="application/json")
