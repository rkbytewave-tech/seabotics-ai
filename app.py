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
.field-label { font-size:0.65rem; color:#94a3b8; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:2px; }
.field-value { font-size:0.85rem; color:#1e293b; font-weight:500; background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; padding:5px 9px; min-height:28px; word-break:break-word; }
.field-empty { font-size:0.8rem; color:#cbd5e1; font-style:italic; background:#f8fafc; border:1px dashed #e2e8f0; border-radius:6px; padding:5px 9px; min-height:28px; }
.status-matched { background:#dcfce7; color:#166534; border-radius:12px; padding:2px 10px; font-size:0.68rem; font-weight:600; display:inline-block; }
.status-verify  { background:#fef9c3; color:#854d0e; border-radius:12px; padding:2px 10px; font-size:0.68rem; font-weight:600; display:inline-block; }
.status-manual  { background:#fee2e2; color:#991b1b; border-radius:12px; padding:2px 10px; font-size:0.68rem; font-weight:600; display:inline-block; }
.status-invalid { background:#fce7f3; color:#9d174d; border-radius:12px; padding:2px 10px; font-size:0.68rem; font-weight:600; display:inline-block; }
.missing-tag { background:#fef2f2; color:#7f1d1d; border-radius:4px; padding:2px 7px; font-size:0.68rem; font-family:monospace; display:inline-block; margin:2px; }
.section-label { font-size:0.65rem; font-weight:700; letter-spacing:0.1em; text-transform:uppercase; color:#64748b; border-left:3px solid #1e3a5f; padding-left:8px; margin:0.8rem 0 0.4rem 0; }
.grid-card { background:#ffffff; border:1px solid #e2e8f0; border-radius:8px; padding:12px; margin-bottom:8px; cursor:pointer; transition:box-shadow 0.15s; }
.grid-card:hover { box-shadow:0 2px 8px rgba(0,0,0,0.08); border-color:#93c5fd; }
.grid-card-title { font-size:0.82rem; font-weight:600; color:#1e293b; margin-bottom:6px; }
.grid-card-sub { font-size:0.72rem; color:#64748b; margin-bottom:2px; }
.invalid-flag { background:#fce7f3; border:1px solid #f9a8d4; border-radius:6px; padding:4px 8px; font-size:0.7rem; color:#9d174d; margin-top:3px; display:inline-block; }
.vessel-card { background:#f0f9ff; border:1px solid #bae6fd; border-radius:8px; padding:12px; margin-bottom:8px; }
</style>
""", unsafe_allow_html=True)

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
MODEL_NAME   = "llama-3.3-70b-versatile"

# ── DROPDOWN REFERENCE DATA ───────────────────────────────────

LICENSES_CERT_DROPDOWN = {
    "STCW - CERTIFICATE OF COMPETENCY": [
        "3RD OFFICER I/2, II/1","2ND ENGINEER MOTOR I/2, III/2",
        "OFFICER IN CHARGE - NAVIGATIONAL WATCH 1/2, II/1","2ND OFFICER I/2, II/1",
        "4TH ENGINEER MOTOR & STEAM 1/2, III/1","ELECTRICAL OFFICER I/2, III/6",
        "3RD ENGINEER MOTOR I/2, III/1","4TH ENGINEER STEAM I/2, III/1",
        "GMDSS GENERAL OPERATOR CERTIFICATE 1/2, IV/2",
        "OIC - ENGINEERING WATCH MOTOR & STEAM 1/2, III/1",
        "CHIEF ENGINEER MOTOR I/2, III/2","CHIEF OFFICER I/2, II/2",
        "OIC - ENGINEERING WATCH STEAM 1/2, III/1","2ND ENGINEER STEAM 1/2, III/2",
        "3RD ENGINEER MOTOR & STEAM 1/2, III/1","MASTER I/2, II/2",
        "Dangerous Cargo","CLASS 1","JUNIOR ENGINEER STEAM III/1",
    ],
    "STCW - CERTIFICATE OF PROFICIENCY": [
        "RATING ENGINEERING WATCH I/2, III/4","ABLE SEAMAN ENGINE 1/2, III/5",
    ],
    "STCW - ENDORSMENT": [
        "DANGEROUS CARGO (CHEM) ENDORSEMENT-OPERATIONS",
        "GMDSS CERTIFICATE OF ENDORSEMENT (NATIONAL)",
        "DANGEROUS CARGO (OIL) ENDORSEMENT - SUPPORT",
        "DANGEROUS CARGO (OIL) ENDORSEMENT - MANAGEMENT",
    ],
    "FLAG - FLAG ENDORSEMENT": [
        "DANGEROUS CARGO (OIL) ENDORSEMENT - OPERATI","SHIPS COOK MLC2006, 1/10",
        "DANGEROUS CARGO (OIL) ENDORSEMENT - SUPPOR","3RD OFFICER II/1, 1/10",
        "4TH ENGINEER MOTOR III/1, 1/10","ELECTRICAL OFFFICER III/6, 1/10",
        "ABLE SEAMAN DECK II/5, 1/10","RATING NAVIGATIONAL WATCH II/4, 1/10",
        "CHIEF ENGINEER MOTOR & STEAM III/2, 1/10",
        "DANGEROUS CARGO (OIL) ENDORSEMENT MANAGM",
        "OIC ENGINEERING WATCH MOTOR III/1, 1/10","CHIEF ENGINEER STEAM III/2, 1/10",
        "DANGEROUS CARGO (CHEM) ENDORSEMENT - OPERA",
        "RECEIPT OF APPLICATION (ROA) FOR COC/Certificate",
        "2ND ENGINEER MOTOR III/2, 1/10","2ND OFFICER II/1, 1/10",
        "3RD ENGINEER MOTOR III/1, 1/10",
        "DANGEROUS CARGO (CHEM) ENDORSEMENT - SUPPO",
        "RATING ENGINEERING WATCH III/4, 1/10",
        "RECEIPT OF APPLICATION (ROA) FOR GMDSS GOC FL.",
        "ELECTRICAL RATING III/7, 1/10","4TH ENGINEER MOTOR & STEAM III/1, 1/10",
        "ABLE SEAMAN ENGINE III/5, 1/10",
        "OIC ENGINEERING WATCH MOTOR & STEAM III/1, 1/10",
        "CHIEF ENGINEER MOTOR III/2, 1/10","2ND ENGINEER MOTOR & STEAM III/2, 1/10",
        "4TH ENGINEER STEAM III/1, 1/10","CHIEF OFFICER II/2, 1/10",
        "JUNIOR ENGINEER STEAM III/1, 1/10",
    ],
}

COURSES_DROPDOWN = {
    "STCW": [
        "ADVANCED CHEMICAL TANKER TRAINING",
        "ADVANCED TRAINING LIQUIFIED GAS TANKER CARGO OIL",
        "ADVANCED TRAINING CHEMICAL TANKER CARGO OPERATION",
        "ADVANCED OIL TANKER TRAINING","ADVANCED FIRE FIGHTING",
        "BASIC SAFETY TRAINING (FPFF, PSSR, PST, EFA)",
    ],
    "FLAG Endorsement": ["SHIP SECURITY OFFICER VI/5, I/10"],
    "Value Added": [
        "6G CERTIFICATE","ABB TURBO CHARGER COURSE",
        "AUTOMATIC IDENTIFICATION SYSTEM (AIS)","BASIC ARC WELDING",
        "BASIC LATHE MACHINE OPERATION","BEHAVIOUR BASED SAFETY","CATERING",
        "CRUDE OIL WASHING & INERT GAS SYSTEM (COW/IGS)","MARITIME ENGLISH",
        "MARITIME LABOUR CONVENTION (MLC) 2006","MARPOL",
        "PORT STATE CONTROL INSPECTION","RISK ASSESSMENT AND INCIDENT INVESTIGATION",
        "VERY LARGE VESSEL HANDLING COURSE","VESSEL HANDLING COURSE",
    ],
    "Other": [
        "ALFA LAVAL PURIFIER COURSE","HYDRAULICS AND PNEUMATICS",
        "NAVIGATIONAL AIDS (BRIDGE EQUIPMENT)",
        "REFRESHER - BRIDGE RESOURCE MANAGEMENT",
        "REFRESHER - ENGINE ROOM RESOURCE MANAGEMENT",
        "SHIP SAFETY OFFICER COURSE",
    ],
}

MEDICAL_DROPDOWN = {
    "Medical Test": [
        "D&A - Drug & Alcohol Test",
        "PEME - Pre-Employment Medical Examination",
        "LFT-Liver Function Test (Post Sign Off From Chemical Tanker)",
        "Treadmill Cardiogram",
    ],
    "Vaccination": [
        "Yellow Fever","COVID 19","VACCINE DOCUMENT","COVID 3rd Dose",
    ],
}

VOYAGE_TYPES   = ["Foreign Going","Near Coastal","Coastal","Home Trade"]
VESSEL_TYPES   = ["BULK CARRIER","CHEMICAL TANKER","CONTAINER SHIP","CRUISE SHIP",
                  "GAS CARRIER","GENERAL CARGO SHIP","LIVESTOCK CARRIER",
                  "OIL TANKER","RO-RO SHIP","TUGBOAT"]

VESSEL_CERT_CATEGORIES = {
    "Class": [
        "Annual Survey","Annual Survey (Fi-Fi)","Annual Survey (UMS)",
        "Annual Survey IG System","Auxiliary Boiler Survey",
        "Auxiliary Boiler Survey (Exhaust Gas)","Auxiliary Boiler Survey (Oil Fired) Port",
        "Auxiliary Boiler Survey (Oil Fired) Stbd","Cargo Gear Annual Survey",
        "Cargo Gear Renewal Survey","Continuous Survey (Fi-Fi)","Continuous Survey Hull",
        "Continuous Survey Machinery","Docking Survey","EIAPP Change of Flag Survey",
        "Exhaust Gas Steam Generators and Economisers Survey","General Examination",
        "In Water Survey","Intermediate Survey","Issuance of COC(SEEMP Part II)",
        "Issuance of COC(SEEMP Part III)","Main Boiler survey","Special Survey ( Fi-Fi)",
        "Special Survey (UMS)","Special Survey Hull","Special Survey IG system",
        "Special Survey Machiney","Tail Shaft Survey",
        "Tailshaft Condition Monitoring Annual Survey","Tailshaft Initial Survey",
        "Tailshaft Periodical Survey","Tailshaft Renewal Survey",
        "Thermal Oil Heating Systems Survey","Certificate of Class","Certificate of Fitness",
        "CERTIFICATE OF TEST AND THOROUGH EXAMINATION OF LIFTING APPLIANCES CG 2",
        "REGISTER OF LIFTING APPLIANCES AND CARGO HANDLING GEAR",
        "ENGINE INTERNATIONAL AIR POLLUTION PREVENTION CERTIFICATE",
    ],
    "Flag State": [
        "Safety Equipment Certificate","Safety Radio Certificate",
        "Safety Construction Certificate","International Loadline Certificate",
        "International Oil Pollution Prevention Certificate",
        "International Ship Security Certificate","Maritime Labour Certificate",
        "Minimum Safe Manning Certificate","ISM Safety Management Certificate",
        "Document of Compliance","USCG Certificate of Compliance",
        "Civil Liability Convention (CLC) 1992 Certificate:",
        "Civil Liability for Bunker Oil Pollution Damage Convention (CLBC) Certificate:",
        "Liability for the Removal of Wrecks Certificate",
        "U.S. Certificate of Financial Responsibility","Certificate of Registry",
        "International Sewage Pollution Prevention Certificate",
        "International Energy Efficiency Certificate",
        "International Air Pollution Prevention Certificate",
        "Ship Sanitation Control (SSCC)/Ship Sanitation Control Exemption",
        "International Ballast Water Management Certificate",
        "International Anti-Fouling System Certificate",
        "CONFIRMATION OF COMPLIANCE - SEEMP PART II",
        "CONFIRMATION OF COMPLIANCE - SEEMP PART III",
        "STATEMENT OF COMPLIANCE FOR INVENTORY OF HAZARDOUS MATERIALS",
        "DOCUMENT OF COMPLIANCE FOR THE CARRIAGE OF DANGEROUS GOODS",
        "CARRIAGE OF SOLID BULK CARGOES CERTIFICATE OF COMPLIANCE",
        "INTERNATIONAL TONNAGE CERTIFICATE",
        "INTERNATIONAL GARBAGE POLLUTION PREVENTION CERTIFICATE",
        "POLAR SHIP CERTIFICATE",
    ],
    "Statutory": [
        "AFS Change of Flag Survey","AFS Initial Survey","AFS Issuance Survey",
        "AFS Periodical Survey","IAPP Annual Survey","IAPP Change of Flag Survey",
        "IAPP General Examination Survey","IAPP Initial Survey","IAPP Intermediate Survey",
        "IAPP Renewal Survey","IBWMC Annual Survey","IBWMC Change of Flag Survey",
        "IBWMC General Examination Survey","IBWMC Renewal Survey",
        "ICOF ( CHEM ) Change of Flag Survey","ICOF ( CHEM ) General Examination Survey",
        "ICOF ( CHEM ) Initial Survey","ICOF ( CHEM ) Renewal Survey",
        "ICOF ( GAS ) Annual Survey","ICOF ( GAS ) Change of Flag Survey",
        "ICOF ( GAS ) General Examination Survey","ICOF ( GAS ) Initial Survey",
        "ICOF ( GAS ) Intermediate Survey","ICOF ( GAS ) Renewal Survey",
        "ICOF (CHEM ) Annual Survey","ICOF (CHEM ) Intermediate Survey",
        "IEEC Change of Flag Survey","IEEC Initial Survey","IEEC Issuance Survey",
        "IMDG Change of Flag Survey","IMDG Genaral Examination Survey",
        "IMDG Initial Survey","IMDG Renewal Survey","IMSBC Change of Flag Survey",
        "IMSBC Genaral Examination Survey","IMSBC Initial Survey","IMSBC Renewal Survey",
        "IGPP Change of Flag Survey","International Load Line Annual Survey",
        "International Load Line Change of Flag Survey",
        "International Load Line General Examination Survey",
        "International Load Line Initial Survey","International Load Line Renewal Survey",
        "International Tonnage Survey",
        "Inventory of Hazardous Material (HKC) Annual Survey",
        "Inventory of Hazardous Material (HKC) Change of Flag Survey",
        "Inventory of Hazardous Material (HKC) General Examination Survey",
        "Inventory of Hazardous Material (HKC) Renewal Survey",
        "Inventory of Hazardous Material(s)-EU Annual Survey",
        "Inventory of Hazardous Material(s)-EU Change of Flag Survey",
        "Inventory of Hazardous Material(s)-EU General Examination Survey",
        "Inventory of Hazardous Material(s)-EU Renewal Survey",
        "IOPP Annual Survey","IOPP Change of Flag Survey","IOPP General Examination Survey",
        "IOPP Initial Survey","IOPP Intermediate Survey","IOPP Renewal Survey",
        "ISM DOC Annual Survey","ISM DOC Initial Survey","ISM DOC interim Survey",
        "ISM DOC Renewal Survey","ISPP Change of Flag Survey",
        "ISPP General Examination Survey","ISPP Initial Survey","ISPP Renewal Survey",
        "ISSC Initial Survey","ISSC Interim Audit","ISSC Intermediate Survey",
        "ISSC Renewal Survey","MLC Interim Audit","PSC Intermediate Survey",
        "PSC Annual Survey","PSC Change of Flag Survey","PSC General Examination Survey",
        "PSC Initial Survey","PSC Renewal Survey","SAFCON Annual Survey",
        "SAFCON Change of Flag Survey","SAFCON General Examination Survey",
        "SAFCON Initial Survey","SAFCON Intermediate Survey","SAFCON Renewal Survey",
        "SEQ Annual Survey","SEQ Change of Flag Survey","SEQ General Examination Survey",
        "SEQ Initial Survey","SEQ Intermediate Survey","SEQ Periodical Survey",
        "SEQ Renewal Survey","SMC Initial Survey","SMC Interim Audit",
        "SMC Intermediate Survey","SMC Renewal Survey","SRT Annual Survey",
        "SRT Change of Flag Survey","SRT General Examination Survey","SRT Initial Survey",
        "SRT Periodical Survey","SRT Renewal Survey",
    ],
}

VESSEL_CERT_TYPES = ["Interim","Full Term","Short Term"]

ALL_LICENSE_NAMES  = [n for names in LICENSES_CERT_DROPDOWN.values() for n in names]
ALL_LICENSE_TYPES  = list(LICENSES_CERT_DROPDOWN.keys())
ALL_COURSE_NAMES   = [n for names in COURSES_DROPDOWN.values() for n in names]
ALL_COURSE_TYPES   = list(COURSES_DROPDOWN.keys())
ALL_MEDICAL_NAMES  = [n for names in MEDICAL_DROPDOWN.values() for n in names]
ALL_MEDICAL_TYPES  = list(MEDICAL_DROPDOWN.keys())
ALL_VESSEL_CATS    = list(VESSEL_CERT_CATEGORIES.keys())
ALL_VESSEL_NAMES   = [n for names in VESSEL_CERT_CATEGORIES.values() for n in names]

# ── SECTION DEFINITIONS ───────────────────────────────────────
PERSONAL_FIELDS = [
    ("person_name","Full Name"),("surname","Surname"),("email","Email"),
    ("gender","Gender"),("contact_number","Contact Number"),
    ("date_of_birth","Date of Birth"),("nationality","Nationality"),
    ("blood_group","Blood Group"),("rank","Current Rank"),
    ("identification_mark","Identification Mark"),("availability_date","Availability Date"),
    ("address","Address"),("country","Country"),("state","State"),
    ("zip_code","ZIP / PIN Code"),("domestic_airport","Domestic Airport"),
    ("international_airport","International Airport"),
    ("passport_number","Passport Number"),("cdc_number","CDC / Seaman's Book Number"),
    ("height","Height"),("weight","Weight"),
    ("boiler_suit_size","Boiler Suit Size"),("safety_shoe_size","Safety Shoe Size"),
    ("shirt_size","Shirt Size"),("trouser_size","Trouser Size"),
]

SECTION_COLUMNS = {
    "Medical": ["Document Type","Document Name","Certificate Number","Date of Issue","Date of Expiry"],
    "Education": ["Institution Name","Country","Level of Education","Field of Study","Start Date","Completion Date"],
    "Courses": ["Document Name","Certificate Number","Type","Place of Issue","Date of Issue","Date of Expiry"],
    "Licenses and Certification": ["Document Name","Certificate Number","Type","Issuing Authority","Place of Issue","Date of Issue","Date of Expiry"],
    "Travel Documents": ["Document Type","Document Name","Certificate Number","Place of Issue","Date of Issue","Date of Expiry"],
    "Sea Service Experience": ["Company","Rank","Vessel Name","Vessel Type","Voyage Type","IMO Number","Flag / Country","GRT","DWT","BHP","Main Engine Type","Main Engine KW","Sign On Port","Sign Off Port","Start Date","End Date","Reason for Sign Off"],
    "Bank Details": ["Bank Name","Branch Name","Account Type","Account Name","Account Number","IFSC / SWIFT Code"],
}

VESSEL_COLUMNS = ["Certificate Name","Certificate Category","Vessel","IMO Number","Certificate Number","Type","Issuing Authority","Classification Society","Date of Issue","Date of Expiry","Date of Survey"]

# Dropdown-validated fields per section
DROPDOWN_FIELDS = {
    "Medical": {
        "Document Type":  ALL_MEDICAL_TYPES,
        "Document Name":  ALL_MEDICAL_NAMES,
    },
    "Courses": {
        "Type":          ALL_COURSE_TYPES,
        "Document Name": ALL_COURSE_NAMES,
    },
    "Licenses and Certification": {
        "Type":          ALL_LICENSE_TYPES,
        "Document Name": ALL_LICENSE_NAMES,
    },
    "Sea Service Experience": {
        "Vessel Type":   VESSEL_TYPES,
        "Voyage Type":   VOYAGE_TYPES,
    },
    "Vessels": {
        "Certificate Category": ALL_VESSEL_CATS,
        "Certificate Name":     ALL_VESSEL_NAMES,
        "Type":                 VESSEL_CERT_TYPES,
    },
}

def fuzzy_match_dropdown(value, options):
    """Find closest matching dropdown option using simple fuzzy logic."""
    if not value or not options:
        return None
    val_upper = value.strip().upper()
    # Exact match first
    for opt in options:
        if opt.upper() == val_upper:
            return opt
    # Contains match
    for opt in options:
        if val_upper in opt.upper() or opt.upper() in val_upper:
            return opt
    # Word overlap
    val_words = set(val_upper.split())
    best_match = None
    best_score = 0
    for opt in options:
        opt_words = set(opt.upper().split())
        overlap = len(val_words & opt_words)
        if overlap > best_score:
            best_score = overlap
            best_match = opt
    return best_match if best_score >= 1 else None

def init_state():
    if "profile"  not in st.session_state: st.session_state["profile"]  = {}
    if "sections" not in st.session_state: st.session_state["sections"] = {s:[] for s in SECTION_COLUMNS}
    if "vessels"  not in st.session_state: st.session_state["vessels"]  = []
    if "doc_tank" not in st.session_state: st.session_state["doc_tank"] = []
    if "popup"    not in st.session_state: st.session_state["popup"]    = None
    if "subpopup" not in st.session_state: st.session_state["subpopup"] = None
    if "page"     not in st.session_state: st.session_state["page"]     = "My Profile"

init_state()

# ── SYSTEM PROMPT ─────────────────────────────────────────────
SYSTEM_PROMPT = """You are SeaBotics.ai — maritime crew ERP document extraction engine.

CRITICAL: First classify if this is a VESSEL document or a PERSONAL/CREW document.
- VESSEL document: contains ship certificates, surveys, class/flag state/statutory certificates with vessel name and IMO number but NO personal crew details like name, DOB, rank of a person
- PERSONAL document: passport, CDC, medical cert, crew certificate, CV, resume, training cert, bank details

DATE FORMAT: DD-MMM-YYYY (e.g. 15-Jul-2023). LIFETIME stays as LIFETIME.
CONFIDENCE: 1.0=exact 0.9=strong 0.75=inferred 0.6=weak 0.0=not found

For COURSES vs LICENSES AND CERTIFICATION:
- Licenses and Certification = Certificate of Competency, Certificate of Proficiency, Endorsements, Flag Endorsements (official qualifications issued by maritime authority)
- Courses = Training courses attended (STCW training, value-added courses, refresher courses)

Return ONLY valid JSON no markdown.

Schema:
{
  "document_type": "",
  "is_vessel_document": false,
  "is_resume": false,
  "personal": {
    "full_name":"","surname":"","email":"","gender":"","contact_number":"",
    "date_of_birth":"","nationality":"","blood_group":"","current_rank":"",
    "identification_mark":"","availability_date":"","address":"","country":"",
    "state":"","zip_code":"","domestic_airport":"","international_airport":"",
    "passport_number":"","cdc_number":"","height":"","weight":"",
    "boiler_suit_size":"","safety_shoe_size":"","shirt_size":"","trouser_size":""
  },
  "sea_service": [{"company":"","rank":"","vessel_name":"","vessel_type":"","voyage_type":"","imo_number":"","flag_country":"","grt":"","dwt":"","bhp":"","main_engine_type":"","main_engine_kw":"","sign_on_port":"","sign_off_port":"","start_date":"","end_date":"","reason_sign_off":""}],
  "travel_documents": [{"document_type":"","document_name":"","certificate_number":"","place_of_issue":"","date_of_issue":"","date_of_expiry":""}],
  "medical": [{"document_type":"","document_name":"","certificate_number":"","date_of_issue":"","date_of_expiry":""}],
  "education": [{"institution_name":"","country":"","education_level":"","field_of_study":"","start_date":"","completion_date":""}],
  "courses": [{"document_name":"","certificate_number":"","type":"","place_of_issue":"","date_of_issue":"","date_of_expiry":""}],
  "licenses_certifications": [{"document_name":"","certificate_number":"","type":"","issuing_authority":"","place_of_issue":"","date_of_issue":"","date_of_expiry":""}],
  "bank_details": [{"bank_name":"","branch_name":"","account_type":"","account_name":"","account_number":"","ifsc_swift":""}],
  "vessel_certificate": {
    "certificate_name":"","certificate_category":"","vessel_name":"","imo_number":"",
    "certificate_number":"","cert_type":"","issuing_authority":"",
    "classification_society":"","date_of_issue":"","date_of_expiry":"","date_of_survey":""
  },
  "missing_fields":[],
  "validation_flags":[],
  "raw_text_summary":""
}"""

IMAGE_EXT = {".png",".jpg",".jpeg",".tiff",".bmp",".webp"}

def preprocess(img):
    if img.mode != "RGB": img = img.convert("RGB")
    w,h = img.size
    if w < 1000: img = img.resize((1000,int(h*1000/w)),Image.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    img = ImageEnhance.Contrast(img).enhance(1.5)
    return img.convert("L")

def extract_text(path):
    ext = Path(path).suffix.lower()
    if ext in IMAGE_EXT:
        img = preprocess(Image.open(path))
        return max([pytesseract.image_to_string(img,config=c) for c in ["--psm 3","--psm 6","--psm 11"]],key=len)
    if ext == ".pdf":
        doc = fitz.open(path); pages=[]
        for p in doc:
            t = p.get_text("text").strip()
            if len(t)>50: pages.append(t)
            else:
                pix=p.get_pixmap(dpi=250)
                img=preprocess(Image.frombytes("RGB",[pix.width,pix.height],pix.samples))
                pages.append(pytesseract.image_to_string(img))
        doc.close(); return "\n--- PAGE BREAK ---\n".join(pages)
    if ext==".docx": return "\n".join(p.text for p in DocxDocument(path).paragraphs if p.text.strip())
    if ext==".txt":
        with open(path) as f: return f.read()
    if ext in [".xlsx",".xls"]:
        wb=openpyxl.load_workbook(path,data_only=True)
        return "\n".join("  |  ".join(str(c) for c in r if c) for s in wb.worksheets for r in s.iter_rows(values_only=True))
    raise ValueError(f"Unsupported: {ext}")

def call_llm(text):
    client=Groq(api_key=GROQ_API_KEY)
    if len(text)>12000: text=text[:12000]
    if not text.strip(): return None
    for attempt in range(3):
        try:
            r=client.chat.completions.create(
                model=MODEL_NAME,temperature=0.1,max_tokens=5000,
                messages=[{"role":"system","content":SYSTEM_PROMPT},
                          {"role":"user","content":f"Extract all information:\n\n{text}"}])
            raw=r.choices[0].message.content.strip()
            if not raw: continue
            raw=re.sub(r"^```[a-z]*\n?","",raw); raw=re.sub(r"\n?```$","",raw)
            start=raw.find("{"); end=raw.rfind("}")+1
            if start==-1 or end==0: continue
            return json.loads(raw[start:end])
        except json.JSONDecodeError: continue
        except Exception as e: st.error(f"API error: {e}"); return None
    st.error("Pipeline failed after 3 attempts."); return None

def conf_color(c):
    if c>=0.90: return "#22c55e"
    if c>=0.60: return "#f59e0b"
    if c>0.00:  return "#ef4444"
    return "#e2e8f0"

def check_dropdown_validity(section, field_name, value):
    """Returns True if value is valid for dropdown field, False if invalid"""
    drops = DROPDOWN_FIELDS.get(section,{})
    if field_name not in drops: return True  # not a dropdown field
    valid_options = drops[field_name]
    return any(value.strip().upper() == opt.upper() for opt in valid_options)

def get_dropdown_options(section, field_name):
    drops = DROPDOWN_FIELDS.get(section,{})
    return drops.get(field_name, [])

def personal_to_profile(p):
    m = {"full_name":"Full Name","surname":"Surname","email":"Email Address",
         "gender":"Gender","contact_number":"Phone Number","date_of_birth":"Date of Birth",
         "nationality":"Nationality","blood_group":"Blood Group","current_rank":"Rank / Designation",
         "identification_mark":"Identification Mark","availability_date":"Availability Date",
         "address":"Residential Address","country":"Country","state":"State",
         "zip_code":"ZIP / PIN Code","domestic_airport":"Domestic Airport",
         "international_airport":"International Airport","passport_number":"Passport Number",
         "cdc_number":"CDC / Seaman's Book Number","height":"Height","weight":"Weight",
         "boiler_suit_size":"Boiler Suit Size","safety_shoe_size":"Safety Shoe Size",
         "shirt_size":"Shirt Size","trouser_size":"Trouser Size"}
    return {v:p[k] for k,v in m.items() if p.get(k)}

def result_to_rows(result):
    rows={s:[] for s in SECTION_COLUMNS}
    for e in result.get("sea_service",[]):
        if any(v for v in e.values()):
            rows["Sea Service Experience"].append({
                "Company":e.get("company",""),"Rank":e.get("rank",""),
                "Vessel Name":e.get("vessel_name",""),"Vessel Type":e.get("vessel_type",""),
                "Voyage Type":e.get("voyage_type",""),"IMO Number":e.get("imo_number",""),
                "Flag / Country":e.get("flag_country",""),"GRT":e.get("grt",""),
                "DWT":e.get("dwt",""),"BHP":e.get("bhp",""),
                "Main Engine Type":e.get("main_engine_type",""),
                "Main Engine KW":e.get("main_engine_kw",""),
                "Sign On Port":e.get("sign_on_port",""),"Sign Off Port":e.get("sign_off_port",""),
                "Start Date":e.get("start_date",""),"End Date":e.get("end_date",""),
                "Reason for Sign Off":e.get("reason_sign_off",""),
            })
    for e in result.get("travel_documents",[]):
        if any(v for v in e.values()):
            rows["Travel Documents"].append({
                "Document Type":e.get("document_type",""),"Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Place of Issue":e.get("place_of_issue",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("medical",[]):
        if any(v for v in e.values()):
            rows["Medical"].append({
                "Document Type":e.get("document_type",""),"Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("education",[]):
        if any(v for v in e.values()):
            rows["Education"].append({
                "Institution Name":e.get("institution_name",""),"Country":e.get("country",""),
                "Level of Education":e.get("education_level",""),
                "Field of Study":e.get("field_of_study",""),
                "Start Date":e.get("start_date",""),"Completion Date":e.get("completion_date",""),
            })
    for e in result.get("courses",[]):
        if any(v for v in e.values()):
            rows["Courses"].append({
                "Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Type":e.get("type",""),"Place of Issue":e.get("place_of_issue",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("licenses_certifications",[]):
        if any(v for v in e.values()):
            rows["Licenses and Certification"].append({
                "Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Type":e.get("type",""),"Issuing Authority":e.get("issuing_authority",""),
                "Place of Issue":e.get("place_of_issue",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("bank_details",[]):
        if any(v for v in e.values()):
            rows["Bank Details"].append({
                "Bank Name":e.get("bank_name",""),"Branch Name":e.get("branch_name",""),
                "Account Type":e.get("account_type",""),"Account Name":e.get("account_name",""),
                "Account Number":e.get("account_number",""),"IFSC / SWIFT Code":e.get("ifsc_swift",""),
            })
    return rows

# Map cert name to category
CERT_NAME_TO_CATEGORY = {}
for cat, names in VESSEL_CERT_CATEGORIES.items():
    for name in names:
        CERT_NAME_TO_CATEGORY[name.upper()] = cat

def infer_cert_category(cert_name):
    if not cert_name: return ""
    matched = fuzzy_match_dropdown(cert_name, ALL_VESSEL_NAMES)
    if matched:
        return CERT_NAME_TO_CATEGORY.get(matched.upper(), "")
    return ""

def infer_cert_type(raw_type):
    if not raw_type: return "Full Term"
    rt = raw_type.upper()
    if "INTERIM" in rt: return "Interim"
    if "SHORT" in rt:   return "Short Term"
    return "Full Term"

def infer_issuing_authority(raw_text, cert_category, classification_society):
    """Issuing authority logic per category."""
    if cert_category == "Class":
        return classification_society or raw_text
    if cert_category == "Flag State":
        return raw_text  # flag state authority or delegated firm like DNV
    if cert_category == "Statutory":
        return raw_text  # flag state maritime authority or RO
    return raw_text

def vessel_result_to_row(result):
    vc = result.get("vessel_certificate",{})
    if not vc or not any(v for v in vc.values()): return None

    cert_name_raw  = vc.get("certificate_name","")
    cert_name_fuzzy = fuzzy_match_dropdown(cert_name_raw, ALL_VESSEL_NAMES) or cert_name_raw
    cert_category  = infer_cert_category(cert_name_raw) or vc.get("certificate_category","")
    cert_type      = infer_cert_type(vc.get("cert_type",""))
    class_society  = vc.get("classification_society","")
    issuing_auth   = infer_issuing_authority(vc.get("issuing_authority",""), cert_category, class_society)

    return {
        "Certificate Name":       cert_name_fuzzy,
        "Certificate Category":   cert_category,
        "Vessel":                 vc.get("vessel_name",""),
        "IMO Number":             vc.get("imo_number",""),
        "Certificate Number":     vc.get("certificate_number",""),
        "Type":                   cert_type,
        "Issuing Authority":      issuing_auth,
        "Classification Society": class_society,
        "Date of Issue":          vc.get("date_of_issue",""),
        "Date of Expiry":         vc.get("date_of_expiry",""),
        "Date of Survey":         vc.get("date_of_survey",""),
    }

# ── GRID CARD DISPLAY ─────────────────────────────────────────
def render_section_content(section_name, rows, key_prefix, view_mode="grid"):
    """Render section rows as grid cards or table, with full data visible."""
    cols_def = SECTION_COLUMNS.get(section_name, VESSEL_COLUMNS)
    drops = DROPDOWN_FIELDS.get(section_name, {})

    if not rows:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No records yet — upload a document above to populate.</div>', unsafe_allow_html=True)
        return

    if view_mode == "table":
        # Full dataframe view
        df = pd.DataFrame(rows, columns=cols_def)
        for c in cols_def:
            if c not in df.columns: df[c] = ""
        df = df[cols_def].fillna("")
        st.dataframe(df, use_container_width=True, hide_index=True)
        # Edit buttons per row below table
        st.markdown("**Edit individual records:**")
        for ridx, row in enumerate(rows):
            title = next((row.get(f,"") for f in ["Document Name","Certificate Name","Company","Institution Name","Bank Name","Vessel"] if row.get(f)), f"Record {ridx+1}")
            if st.button(f"Edit: {title}", key=f"{key_prefix}_tbl_edit_{ridx}"):
                st.session_state["subpopup"] = {
                    "section": section_name, "row_idx": ridx,
                    "row": copy.deepcopy(row), "key": f"{key_prefix}_{ridx}",
                    "source": "profile"
                }
                st.rerun()
    else:
        # Grid card view — show ALL fields in card
        for i in range(0, len(rows), 2):
            card_cols = st.columns(2)
            for j in range(2):
                idx = i + j
                if idx >= len(rows): break
                row = rows[idx]
                with card_cols[j]:
                    title_val = next((row.get(f,"") for f in ["Document Name","Certificate Name","Company","Institution Name","Bank Name","Vessel"] if row.get(f)), f"Record {idx+1}")
                    invalid_fields = [cn for cn,val in row.items() if val and cn in drops and not check_dropdown_validity(section_name,cn,val)]

                    # Build all field lines for card body
                    field_lines = "".join([
                        f'<div style="margin-bottom:3px;"><span style="font-size:0.65rem;color:#94a3b8;text-transform:uppercase;">{cn}:</span> <span style="font-size:0.78rem;color:#1e293b;font-weight:500;">{val}</span></div>'
                        for cn, val in row.items() if val
                    ])
                    invalid_html = "".join(f'<span class="invalid-flag">⚠ {f}</span> ' for f in invalid_fields)

                    st.markdown(f"""
                    <div class="grid-card">
                        <div class="grid-card-title">{title_val}</div>
                        {field_lines}
                        {invalid_html}
                    </div>
                    """, unsafe_allow_html=True)

                    if st.button("Edit", key=f"{key_prefix}_edit_{idx}", use_container_width=True):
                        st.session_state["subpopup"] = {
                            "section": section_name, "row_idx": idx,
                            "row": copy.deepcopy(row), "key": f"{key_prefix}_{idx}",
                            "source": "profile"
                        }
                        st.rerun()

# ── SUB-POPUP (edit individual record) ───────────────────────
def render_subpopup():
    sp = st.session_state.get("subpopup")
    if not sp: return

    section  = sp["section"]
    row_idx  = sp["row_idx"]
    row      = sp["row"]
    key      = sp["key"]
    cols_def = SECTION_COLUMNS.get(section, VESSEL_COLUMNS)
    drops    = DROPDOWN_FIELDS.get(section, {})

    ALL_ORG_OPTIONS = sorted(set([
        "Full Name","Surname","Email Address","Gender","Phone Number","Date of Birth",
        "Nationality","Blood Group","Rank / Designation","Identification Mark",
        "Availability Date","Residential Address","Country","State","ZIP / PIN Code",
        "Domestic Airport","International Airport","Passport Number",
        "CDC / Seaman's Book Number","Height","Weight","Boiler Suit Size",
        "Safety Shoe Size","Shirt Size","Trouser Size","Document Type","Document Name",
        "Document / Certificate Number","Issuing Authority","Place of Issue",
        "Date of Issue","Date of Expiry","Institution Name","Level of Education",
        "Field of Study","Start Date","Completion Date","Bank Name","Branch Name",
        "Account Type","Account Name","Account Number","IFSC / SWIFT Code",
        "Company Name","IMO Number","Vessel Name","Vessel Type","Flag State",
        "Voyage Type","Sign On Port","Sign Off Port","GRT","DWT","BHP",
        "Main Engine Type / Model","Main Engine KW","Reason for Sign Off",
        "Certificate Name","Certificate Category","Classification Society",
        "Date of Survey","Other / Unclassified"
    ]))

    ORG_PREMAP = {
        "Document Name": "Document Name",
        "Document Type": "Document Type",
        "Certificate Number": "Document / Certificate Number",
        "Date of Issue": "Date of Issue",
        "Date of Expiry": "Date of Expiry",
        "Place of Issue": "Place of Issue",
        "Issuing Authority": "Issuing Authority",
        "Type": "Document Type",
        "Company": "Company Name",
        "Rank": "Rank / Designation",
        "Vessel Name": "Vessel Name",
        "Vessel Type": "Vessel Type",
        "Voyage Type": "Voyage Type",
        "IMO Number": "IMO Number",
        "Flag / Country": "Flag State",
        "GRT": "GRT", "DWT": "DWT", "BHP": "BHP",
        "Main Engine Type": "Main Engine Type / Model",
        "Main Engine KW": "Main Engine KW",
        "Sign On Port": "Sign On Port",
        "Sign Off Port": "Sign Off Port",
        "Start Date": "Start Date",
        "End Date": "Sign Off Date",
        "Reason for Sign Off": "Reason for Sign Off",
        "Institution Name": "Institution Name",
        "Level of Education": "Level of Education",
        "Field of Study": "Field of Study",
        "Completion Date": "Completion Date",
        "Country": "Country",
        "Bank Name": "Bank Name",
        "Branch Name": "Branch Name",
        "Account Type": "Account Type",
        "Account Name": "Account Name",
        "Account Number": "Account Number",
        "IFSC / SWIFT Code": "IFSC / SWIFT Code",
        "Certificate Name": "Certificate Name",
        "Certificate Category": "Certificate Category",
        "Vessel": "Vessel Name",
        "Classification Society": "Classification Society",
        "Date of Survey": "Date of Survey",
    }

    with st.expander(f"Edit Record — {section}", expanded=True):
        st.markdown(f'<div class="section-label">Editing record {row_idx+1} in {section}</div>', unsafe_allow_html=True)

        h1,h2,h3,h4,h5 = st.columns([2,2.5,1.2,1.2,2.5])
        h1.markdown("**Field**")
        h2.markdown("**Value**")
        h3.markdown("**Confidence**")
        h4.markdown("**Status**")
        h5.markdown("**Org Database Field**")
        st.markdown("<hr style='margin:4px 0 8px 0;border-color:#e2e8f0;'>", unsafe_allow_html=True)

        edited_row = {}
        edited_org = {}

        for col_name in cols_def:
            val  = row.get(col_name,"") or ""
            conf = 0.9 if val else 0.0
            color = conf_color(conf)
            is_dropdown = col_name in drops
            dropdown_options = get_dropdown_options(section, col_name) if is_dropdown else []

            fuzzy_val = val
            if is_dropdown and val and dropdown_options:
                matched = fuzzy_match_dropdown(val, dropdown_options)
                if matched:
                    fuzzy_val = matched

            is_invalid = bool(val and is_dropdown and not check_dropdown_validity(section, col_name, fuzzy_val))

            if conf >= 0.90: status_txt,status_cls = "Matched","status-matched"
            elif conf >= 0.60: status_txt,status_cls = "Verify","status-verify"
            else: status_txt,status_cls = "Manual Entry","status-manual"
            if is_invalid: status_txt,status_cls = "Invalid","status-invalid"

            c1,c2,c3,c4,c5 = st.columns([2,2.5,1.2,1.2,2.5])

            with c1:
                st.markdown(
                    f'<div style="font-size:0.82rem;color:#1e293b;font-weight:500;padding-top:6px;">{col_name}</div>'
                    f'<div style="font-size:0.68rem;color:#94a3b8;font-style:italic;">Extracted: {val if val else "—"}</div>',
                    unsafe_allow_html=True
                )
                if is_invalid:
                    st.markdown('<span class="invalid-flag">Not in dropdown</span>', unsafe_allow_html=True)

            with c2:
                if is_dropdown and dropdown_options:
                    opts = ["— Select —"] + dropdown_options
                    cur_idx = next((i+1 for i,o in enumerate(dropdown_options) if o.upper()==fuzzy_val.upper()), 0)
                    edited_row[col_name] = st.selectbox(
                        f"v_{key}_{col_name}", options=opts, index=cur_idx,
                        key=f"v_sp_{key}_{col_name}", label_visibility="collapsed"
                    )
                    if edited_row[col_name] == "— Select —": edited_row[col_name] = ""
                else:
                    edited_row[col_name] = st.text_input(
                        f"v_{key}_{col_name}", value=fuzzy_val,
                        key=f"v_sp_{key}_{col_name}", label_visibility="collapsed"
                    )

            with c3:
                st.markdown(
                    f'<div style="padding-top:6px;">'
                    f'<div style="background:#f1f5f9;border-radius:4px;height:6px;width:70px;">'
                    f'<div style="width:{int(conf*100)}%;height:6px;background:{color};border-radius:4px;"></div></div>'
                    f'<div style="font-size:0.65rem;color:{color};margin-top:2px;">{int(conf*100)}%</div></div>',
                    unsafe_allow_html=True
                )

            with c4:
                st.markdown(
                    f'<div style="padding-top:8px;"><span class="{status_cls}">{status_txt}</span></div>',
                    unsafe_allow_html=True
                )

            with c5:
                org_default = ORG_PREMAP.get(col_name, "Other / Unclassified")
                org_opts = ["— Select org field —"] + ALL_ORG_OPTIONS
                default_idx = next((i+1 for i,o in enumerate(ALL_ORG_OPTIONS) if o==org_default), 0)
                edited_org[col_name] = st.selectbox(
                    f"o_{key}_{col_name}", options=org_opts, index=default_idx,
                    key=f"o_sp_{key}_{col_name}", label_visibility="collapsed"
                )

        st.markdown("---")
        sa,sb = st.columns(2)
        with sa:
            if st.button("Save Changes", type="primary", use_container_width=True, key=f"save_sp_{key}"):
                source = st.session_state["subpopup"].get("source","profile")
                if source == "profile":
                    if section == "Vessels":
                        if row_idx < len(st.session_state["vessels"]):
                            st.session_state["vessels"][row_idx] = edited_row
                    else:
                        if section in st.session_state["sections"] and row_idx < len(st.session_state["sections"][section]):
                            st.session_state["sections"][section][row_idx] = edited_row
                elif source == "popup":
                    popup = st.session_state.get("popup", {})
                    if "pending_rows" in popup and section in popup.get("pending_rows", {}):
                        if row_idx < len(popup["pending_rows"][section]):
                            popup["pending_rows"][section][row_idx] = edited_row
                            st.session_state["popup"] = popup
                    elif section == "Vessels":
                        st.session_state["popup"]["pending_vessel"] = edited_row
                st.session_state["subpopup"] = None
                st.rerun()
        with sb:
            if st.button("Cancel", use_container_width=True, key=f"cancel_sp_{key}"):
                st.session_state["subpopup"] = None
                st.rerun()

# ── MAIN POPUP (after upload) ─────────────────────────────────
def render_popup():
    popup = st.session_state.get("popup")
    if not popup: return

    result   = popup.get("result",{})
    filename = popup.get("filename","document")
    source   = popup.get("source","section")
    is_vessel = result.get("is_vessel_document", False)
    is_resume = result.get("is_resume", False)
    doc_type  = result.get("document_type","—")
    personal  = result.get("personal",{})
    missing   = result.get("missing_fields",[])
    flags     = result.get("validation_flags",[])

    # Build pending rows if not already built
    if "pending_rows" not in popup:
        popup["pending_rows"] = result_to_rows(result)
        popup["pending_vessel"] = vessel_result_to_row(result) if is_vessel else None
        st.session_state["popup"] = popup

    pending_rows   = popup["pending_rows"]
    pending_vessel = popup.get("pending_vessel")

    with st.expander(f"Review & Approve — {filename}", expanded=True):
        m1,m2,m3 = st.columns(3)
        m1.metric("Document Type", doc_type)
        m2.metric("Vessel Document", "Yes" if is_vessel else "No")
        m3.metric("CV / Resume", "Yes" if is_resume else "No")

        if is_vessel and pending_vessel:
            st.info(f"Vessel certificate detected — will populate **Vessels** section.")
        else:
            total_rows = sum(len(r) for r in pending_rows.values())
            pers_count = len([v for v in personal.values() if v])
            st.info(f"Personal fields: {pers_count}  |  Section records: {total_rows}")

        st.markdown("---")

        # Personal info (if not vessel doc)
        if not is_vessel and any(v for v in personal.values()):
            st.markdown('<div class="section-label">Personal Information</div>', unsafe_allow_html=True)
            p_items = [(k,v) for k,v in personal.items() if v]
            for i in range(0,len(p_items),3):
                row = p_items[i:i+3]
                cols = st.columns(3)
                for j,(k,v) in enumerate(row):
                    with cols[j]:
                        st.markdown(f'<div class="field-label">{k.replace("_"," ").title()}</div>', unsafe_allow_html=True)
                        personal[k] = st.text_input(f"pp_{k}", value=v, key=f"pp_popup_{k}", label_visibility="collapsed")

        # Vessel certificate preview
        if is_vessel and pending_vessel:
            st.markdown('<div class="section-label">Vessel Certificate</div>', unsafe_allow_html=True)
            cols_v = st.columns(3)
            for i,(col_name,val) in enumerate(pending_vessel.items()):
                with cols_v[i%3]:
                    st.markdown(f'<div class="field-label">{col_name}</div>', unsafe_allow_html=True)
                    if st.button("Edit", key=f"popup_vessel_edit_{i}", use_container_width=False):
                        st.session_state["subpopup"] = {
                            "section":"Vessels","row_idx":0,
                            "row":copy.deepcopy(pending_vessel),
                            "key":"popup_vessel","source":"popup"
                        }
                        st.rerun()
                    st.markdown(f'<div class="field-value">{val or "—"}</div>', unsafe_allow_html=True)

        # Section grids
        if not is_vessel:
            for section_name, rows in pending_rows.items():
                if not rows: continue
                st.markdown(f'<div class="section-label">{section_name} — {len(rows)} record(s)</div>', unsafe_allow_html=True)
                for ridx, row in enumerate(rows):
                    title_val = next((row[f] for f in ["Document Name","Company","Institution Name","Bank Name","Vessel Name"] if row.get(f)), f"Record {ridx+1}")
                    drops = DROPDOWN_FIELDS.get(section_name,{})
                    invalid_fields = [cn for cn,val in row.items() if val and cn in drops and not check_dropdown_validity(section_name,cn,val)]

                    cc1,cc2 = st.columns([4,1])
                    with cc1:
                        badge = f' <span class="invalid-flag">⚠ {len(invalid_fields)} invalid field(s)</span>' if invalid_fields else ""
                        st.markdown(f'<div class="grid-card"><div class="grid-card-title">{title_val}{badge}</div></div>', unsafe_allow_html=True)
                    with cc2:
                        if st.button("Edit", key=f"popup_edit_{section_name}_{ridx}", use_container_width=True):
                            st.session_state["subpopup"] = {
                                "section":section_name,"row_idx":ridx,
                                "row":copy.deepcopy(row),"key":f"popup_{section_name}_{ridx}",
                                "source":"popup"
                            }
                            st.rerun()

        if flags:
            st.markdown("---")
            for f in flags: st.warning(f)
        if missing:
            st.markdown(" ".join(f'<span class="missing-tag">{m}</span>' for m in missing), unsafe_allow_html=True)

        st.markdown("---")
        ba,bb,bc = st.columns([2,2,1])
        with ba:
            if st.button("Approve and Populate", type="primary", use_container_width=True, key="approve_popup"):
                if is_vessel:
                    if pending_vessel and any(v for v in pending_vessel.values()):
                        st.session_state["vessels"].append(pending_vessel)
                else:
                    profile_data = personal_to_profile(personal)
                    for k,v in profile_data.items():
                        st.session_state["profile"][k] = v
                    for td in pending_rows.get("Travel Documents",[]):
                        dt=(td.get("Document Type","")+td.get("Document Name","")).lower()
                        if "passport" in dt and td.get("Certificate Number"):
                            st.session_state["profile"]["Passport Number"]=td["Certificate Number"]
                        if ("cdc" in dt or "seaman" in dt) and td.get("Certificate Number"):
                            st.session_state["profile"]["CDC / Seaman's Book Number"]=td["Certificate Number"]
                    for section,rows in pending_rows.items():
                        for row in rows:
                            if any(v for v in row.values()):
                                st.session_state["sections"][section].append(row)

                st.session_state["doc_tank"].append({
                    "timestamp":  datetime.now().strftime("%d-%b-%Y %H:%M"),
                    "filename":   filename,"source":source,
                    "document_type": doc_type,"is_vessel": is_vessel,"is_resume": is_resume,
                    "raw_result": result,
                })
                st.session_state["popup"] = None
                st.success("Profile populated successfully.")
                st.rerun()

        with bb:
            st.download_button("Download Raw JSON", data=json.dumps(result,indent=2),
                file_name=f"{Path(filename).stem}_raw.json",
                mime="application/json", key="dl_popup", use_container_width=True)
        with bc:
            if st.button("Discard", use_container_width=True, key="discard_popup"):
                st.session_state["popup"]=None; st.rerun()

# ── HEADER ────────────────────────────────────────────────────
col_brand,col_master = st.columns([3,1])
with col_brand:
    st.markdown("""
    <div style="background:linear-gradient(135deg,#0f172a,#1e3a5f);border-radius:12px;padding:16px 24px;margin-bottom:16px;">
        <div style="font-size:1.3rem;font-weight:700;color:#38bdf8;">SeaBotics.ai</div>
        <div style="font-size:0.72rem;color:#94a3b8;margin-top:2px;">Crew & Vessel Profile Management — v5.0</div>
    </div>""", unsafe_allow_html=True)

with col_master:
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown("**Master Upload**")
    st.caption("Docs, merged PDFs, CVs, or Vessel Certs")
    master_file = st.file_uploader("master", type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                   label_visibility="collapsed", key="master_uploader")
    if master_file and st.button("Run Pipeline", use_container_width=True, key="run_master"):
        with st.spinner("Classifying and extracting..."):
            ext = Path(master_file.name).suffix.lower()
            with tempfile.NamedTemporaryFile(delete=False,suffix=ext) as tmp:
                tmp.write(master_file.read()); tmp_path=tmp.name
            try:
                text=extract_text(tmp_path)
                result=call_llm(text)
                if result:
                    st.session_state["popup"]={"key":"master","result":result,"filename":master_file.name,"source":"master"}
                    st.rerun()
            except Exception as e: st.error(f"Error: {e}")

# ── NAV ───────────────────────────────────────────────────────
n1,n2,n3,n4 = st.columns([1,1,1,5])
with n1:
    if st.button("My Profile", use_container_width=True,
                 type="primary" if st.session_state["page"]=="My Profile" else "secondary"):
        st.session_state["page"]="My Profile"; st.rerun()
with n2:
    if st.button("Vessels", use_container_width=True,
                 type="primary" if st.session_state["page"]=="Vessels" else "secondary"):
        st.session_state["page"]="Vessels"; st.rerun()
with n3:
    if st.button("Doc Tank", use_container_width=True,
                 type="primary" if st.session_state["page"]=="Doc Tank" else "secondary"):
        st.session_state["page"]="Doc Tank"; st.rerun()

# ── SUBPOPUP (always rendered on top if active) ───────────────
if st.session_state.get("subpopup"):
    render_subpopup()

# ── MAIN POPUP ────────────────────────────────────────────────
elif st.session_state.get("popup"):
    render_popup()

# ════════════════════════════════════════════════════════════
#  MY PROFILE
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "My Profile":

    # Personal Info
    with st.expander("Personal Information", expanded=True):
        uc1,uc2 = st.columns([4,1])
        with uc1:
            sec_file=st.file_uploader("Upload personal doc",type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                      key="up_personal",label_visibility="collapsed")
        with uc2:
            st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
            if sec_file and st.button("Extract",key="run_personal",use_container_width=True):
                with st.spinner("Extracting..."):
                    ext=Path(sec_file.name).suffix.lower()
                    with tempfile.NamedTemporaryFile(delete=False,suffix=ext) as tmp:
                        tmp.write(sec_file.read()); tmp_path=tmp.name
                    try:
                        text=extract_text(tmp_path)
                        result=call_llm(text)
                        if result:
                            st.session_state["popup"]={"key":"sec_personal","result":result,"filename":sec_file.name,"source":"section","target_section":"Personal Information"}
                            st.rerun()
                    except Exception as e: st.error(f"Error: {e}")

        st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)
        profile = st.session_state["profile"]
        ORG_MAP = {"person_name":"Full Name","surname":"Surname","email":"Email Address","gender":"Gender","contact_number":"Phone Number","date_of_birth":"Date of Birth","nationality":"Nationality","blood_group":"Blood Group","rank":"Rank / Designation","identification_mark":"Identification Mark","availability_date":"Availability Date","address":"Residential Address","country":"Country","state":"State","zip_code":"ZIP / PIN Code","domestic_airport":"Domestic Airport","international_airport":"International Airport","passport_number":"Passport Number","cdc_number":"CDC / Seaman's Book Number","height":"Height","weight":"Weight","boiler_suit_size":"Boiler Suit Size","safety_shoe_size":"Safety Shoe Size","shirt_size":"Shirt Size","trouser_size":"Trouser Size"}
        for i in range(0,len(PERSONAL_FIELDS),3):
            row_f=PERSONAL_FIELDS[i:i+3]; cols=st.columns(3)
            for j,(fk,fl) in enumerate(row_f):
                with cols[j]:
                    org_key=ORG_MAP.get(fk,fl); val=profile.get(org_key,"")
                    st.markdown(f'<div class="field-label">{fl}</div>',unsafe_allow_html=True)
                    if val: st.markdown(f'<div class="field-value">{val}</div>',unsafe_allow_html=True)
                    else:   st.markdown(f'<div class="field-empty">—</div>',unsafe_allow_html=True)
        if profile:
            if st.button("Clear Personal Info",key="clear_personal"):
                st.session_state["profile"]={};st.rerun()

    # Multi-row sections
    for section_name,cols_def in SECTION_COLUMNS.items():
        rows=st.session_state["sections"].get(section_name,[])
        with st.expander(f"{section_name}  ({len(rows)} record{'s' if len(rows)!=1 else ''})",expanded=False):
            uc1,uc2=st.columns([4,1])
            with uc1:
                sec_file=st.file_uploader(f"Upload for {section_name}",
                    type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                    key=f"up_{section_name}",label_visibility="collapsed")
            with uc2:
                st.markdown("<div style='height:4px'></div>",unsafe_allow_html=True)
                if sec_file and st.button("Extract",key=f"run_{section_name}",use_container_width=True):
                    with st.spinner("Extracting..."):
                        ext=Path(sec_file.name).suffix.lower()
                        with tempfile.NamedTemporaryFile(delete=False,suffix=ext) as tmp:
                            tmp.write(sec_file.read()); tmp_path=tmp.name
                        try:
                            text=extract_text(tmp_path)
                            result=call_llm(text)
                            if result:
                                st.session_state["popup"]={"key":f"sec_{section_name.replace(' ','_')}","result":result,"filename":sec_file.name,"source":"section","target_section":section_name}
                                st.rerun()
                        except Exception as e: st.error(f"Error: {e}")

            st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>",unsafe_allow_html=True)

# View toggle
            view_key = f"view_{section_name}"
            if view_key not in st.session_state:
                st.session_state[view_key] = "grid"
            vt1, vt2, _ = st.columns([1,1,6])
            with vt1:
                if st.button("Grid", key=f"btn_grid_{section_name}",
                             type="primary" if st.session_state[view_key]=="grid" else "secondary"):
                    st.session_state[view_key] = "grid"; st.rerun()
            with vt2:
                if st.button("Table", key=f"btn_table_{section_name}",
                             type="primary" if st.session_state[view_key]=="table" else "secondary"):
                    st.session_state[view_key] = "table"; st.rerun()

            render_section_content(section_name, rows, key_prefix=f"profile_{section_name}", view_mode=st.session_state[view_key])
            
            ba,bb=st.columns(2)
            with ba:
                if st.button(f"+ Add blank row",key=f"add_{section_name}"):
                    st.session_state["sections"][section_name].append({c:"" for c in cols_def})
                    st.rerun()
            with bb:
                if rows and st.button(f"Clear all",key=f"clear_{section_name}"):
                    st.session_state["sections"][section_name]=[]; st.rerun()

# ════════════════════════════════════════════════════════════
#  VESSELS
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "Vessels":
    st.markdown("### Vessels — Certificate Registry")
    st.markdown("<div style='font-size:0.85rem;color:#64748b;margin-bottom:16px;'>Upload vessel certificates to populate. Each certificate is stored as one row.</div>", unsafe_allow_html=True)

    vc1,vc2=st.columns([4,1])
    with vc1:
        v_file=st.file_uploader("Upload vessel certificate",type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                key="up_vessel",label_visibility="collapsed")
    with vc2:
        st.markdown("<div style='height:4px'></div>",unsafe_allow_html=True)
        if v_file and st.button("Extract",key="run_vessel",use_container_width=True):
            with st.spinner("Extracting vessel certificate..."):
                ext=Path(v_file.name).suffix.lower()
                with tempfile.NamedTemporaryFile(delete=False,suffix=ext) as tmp:
                    tmp.write(v_file.read()); tmp_path=tmp.name
                try:
                    text=extract_text(tmp_path)
                    result=call_llm(text)
                    if result:
                        st.session_state["popup"]={"key":"vessel_cert","result":result,"filename":v_file.name,"source":"vessel"}
                        st.rerun()
                except Exception as e: st.error(f"Error: {e}")

    st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>",unsafe_allow_html=True)

    vessels = st.session_state.get("vessels",[])
    if not vessels:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No vessel certificates yet.</div>', unsafe_allow_html=True)
    else:
        view_key = "view_vessels"
        if view_key not in st.session_state: st.session_state[view_key] = "grid"
        vt1,vt2,_ = st.columns([1,1,6])
        with vt1:
            if st.button("Grid", key="btn_grid_vessels",
                         type="primary" if st.session_state[view_key]=="grid" else "secondary"):
                st.session_state[view_key]="grid"; st.rerun()
        with vt2:
            if st.button("Table", key="btn_table_vessels",
                         type="primary" if st.session_state[view_key]=="table" else "secondary"):
                st.session_state[view_key]="table"; st.rerun()

        render_section_content("Vessels", vessels, key_prefix="vessel_grid", view_mode=st.session_state[view_key])

# ════════════════════════════════════════════════════════════
#  DOC TANK
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "Doc Tank":
    st.markdown("### Doc Tank — Upload Log")
    tank=st.session_state.get("doc_tank",[])
    if not tank:
        st.info("No documents processed yet.")
    else:
        s1,s2,s3,s4=st.columns(4)
        s1.metric("Total",      len(tank))
        s2.metric("CV/Resume",  len([d for d in tank if d.get("is_resume")]))
        s3.metric("Vessel",     len([d for d in tank if d.get("is_vessel")]))
        s4.metric("Crew Docs",  len([d for d in tank if not d.get("is_vessel") and not d.get("is_resume")]))
        st.markdown("---")
        for i,doc in enumerate(reversed(tank)):
            idx=len(tank)-1-i
            with st.expander(f"[{doc['timestamp']}]  {doc['filename']}  —  {doc['document_type']}",expanded=(i==0)):
                d1,d2,d3=st.columns(3)
                d1.metric("Uploaded",      doc["timestamp"])
                d2.metric("Document Type", doc["document_type"])
                d3.metric("Vessel Doc",    "Yes" if doc.get("is_vessel") else "No")
                st.download_button("Download Raw JSON",
                    data=json.dumps(doc.get("raw_result",{}),indent=2),
                    file_name=f"{Path(doc['filename']).stem}_raw.json",
                    mime="application/json",key=f"dlr_{idx}")
        st.markdown("---")
        st.download_button("Export Full Doc Tank",
            data=json.dumps(tank,indent=2,default=str),
            file_name="seabotics_doc_tank.json",mime="application/json")
