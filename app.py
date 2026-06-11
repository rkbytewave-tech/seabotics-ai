import streamlit as st
import json, re, os, tempfile, copy
from difflib import SequenceMatcher
from pathlib import Path
from datetime import datetime
from groq import Groq
import fitz
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from docx import Document as DocxDocument
import openpyxl
import pandas as pd

st.set_page_config(page_title="SeaBotics.ai", page_icon="S", layout="wide", initial_sidebar_state="expanded")

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

ALL_GENDER_OPTIONS    = ["Male", "Female", "Prefer Not to mention", "Other"]
ALL_RANK_OPTIONS      = [
    "Master", "Chief Officer", "Second Officer", "Third Officer",
    "Chief Engineer", "Second Engineer", "Third Engineer", "Fourth Engineer",
    "Electrical Officer", "Bosun", "Able Seaman", "Ordinary Seaman",
    "Deck Cadet", "Engine Cadet", "Fitter", "Oiler", "Wiper",
    "Chief Steward", "Steward", "Cook", "Messman",
]
ALL_SHIRT_SIZES       = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
ALL_TROUSER_SIZES     = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
ALL_BOILER_SUIT_SIZES = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]

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
    "Personal Information": {
        "Gender":           ALL_GENDER_OPTIONS,
        "Current Rank":     ALL_RANK_OPTIONS,
        "Shirt Size":       ALL_SHIRT_SIZES,
        "Trouser Size":     ALL_TROUSER_SIZES,
        "Boiler Suit Size": ALL_BOILER_SUIT_SIZES,
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

def reclassify_courses_and_licenses(rows_dict):
    """
    Bottom-up reclassification for Courses and Licenses and Certification.
    Fuzzy-matches extracted document names against reference lists.
    Moves misclassified rows to the correct section and sets the correct type.
    """
    # Build reverse lookup: name -> type category
    COURSE_NAME_TO_TYPE = {}
    for cat, names in COURSES_DROPDOWN.items():
        for name in names:
            COURSE_NAME_TO_TYPE[name.upper()] = cat

    LICENSE_NAME_TO_TYPE = {}
    for cat, names in LICENSES_CERT_DROPDOWN.items():
        for name in names:
            LICENSE_NAME_TO_TYPE[name.upper()] = cat

    final_courses  = []
    final_licenses = []

    # Pool all rows from both sections for reclassification
    pool = []
    for row in rows_dict.get("Courses", []):
        pool.append(("Courses", row))
    for row in rows_dict.get("Licenses and Certification", []):
        pool.append(("Licenses and Certification", row))

    for origin, row in pool:
        doc_name = row.get("Document Name", "")

        # Try matching against course names first
        course_match = fuzzy_match_dropdown(doc_name, ALL_COURSE_NAMES)
        license_match = fuzzy_match_dropdown(doc_name, ALL_LICENSE_NAMES)

        # Score: exact/strong match wins
        course_score  = _match_score(doc_name, course_match)
        license_score = _match_score(doc_name, license_match)

        if license_score >= course_score and license_match:
            # Goes to Licenses and Certification
            new_row = dict(row)
            new_row["Document Name"] = license_match
            new_row["Type"] = LICENSE_NAME_TO_TYPE.get(license_match.upper(), row.get("Type",""))
            # Ensure license fields exist
            if "Issuing Authority" not in new_row: new_row["Issuing Authority"] = ""
            if "Place of Issue"    not in new_row: new_row["Place of Issue"]    = ""
            final_licenses.append(new_row)

        elif course_match:
            # Goes to Courses
            new_row = dict(row)
            new_row["Document Name"] = course_match
            new_row["Type"] = COURSE_NAME_TO_TYPE.get(course_match.upper(), row.get("Type",""))
            final_courses.append(new_row)

        else:
            # No match found - keep in original section, flag it
            if origin == "Courses":
                final_courses.append(row)
            else:
                final_licenses.append(row)

    rows_dict["Courses"] = final_courses
    rows_dict["Licenses and Certification"] = final_licenses
    return rows_dict


def _match_score(original, matched):
    """Score how well a fuzzy match landed. Higher = better."""
    if not original or not matched:
        return 0
    orig_upper    = original.strip().upper()
    matched_upper = matched.strip().upper()
    if orig_upper == matched_upper:
        return 3  # exact
    if orig_upper in matched_upper or matched_upper in orig_upper:
        return 2  # contains
    orig_words    = set(orig_upper.split())
    matched_words = set(matched_upper.split())
    overlap       = len(orig_words & matched_words)
    return overlap  # word overlap score

def init_state():
    if "profile"  not in st.session_state: st.session_state["profile"]  = {}
    if "sections" not in st.session_state: st.session_state["sections"] = {s:[] for s in SECTION_COLUMNS}
    if "vessels"  not in st.session_state: st.session_state["vessels"]  = []
    if "doc_tank" not in st.session_state: st.session_state["doc_tank"] = []
    if "popup"    not in st.session_state: st.session_state["popup"]    = None
    if "subpopup" not in st.session_state: st.session_state["subpopup"] = None
    if "page"     not in st.session_state: st.session_state["page"]     = "Crew"
    if "dt_popup" not in st.session_state: st.session_state["dt_popup"] = None

def save_to_doc_tank(result, filename, source, doc_type, is_vessel, is_resume, section_name=None):
    """Save any uploaded doc to doc tank immediately as draft."""
    # Avoid duplicate entries for same filename+timestamp
    ts = datetime.now().strftime("%d-%b-%Y %H:%M")
    st.session_state["doc_tank"].append({
        "timestamp":     ts,
        "filename":      filename,
        "source":        source,
        "section_name":  section_name or "",
        "document_type": doc_type,
        "is_vessel":     is_vessel,
        "is_resume":     is_resume,
        "raw_result":    result,
        "status":        "Pending",
        "tank_id":       f"{filename}_{ts}",
    })

init_state()

# ── CLASSIFIER PROMPT (Master Upload - call 1) ────────────────
CLASSIFIER_PROMPT = """You are a maritime document classifier. Read the text and return ONLY valid JSON, no markdown.

Classify the document and return:
{
  "is_vessel": false,
  "is_resume": false,
  "target_section": "",
  "doc_type": ""
}

target_section must be one of:
"Personal Information", "Medical", "Education", "Courses",
"Licenses and Certification", "Travel Documents",
"Sea Service Experience", "Bank Details", "Vessel"

Rules:
- is_vessel = true if document contains ship/vessel certificates, surveys, class/flag/statutory certs with IMO number but NO personal crew details
- is_resume = true if document is a CV, resume, or sailing experience record covering multiple sections
- If is_resume = true, set target_section = "Resume"
- If is_vessel = true, set target_section = "Vessel"
- Travel Documents: passport, CDC, seaman's book, visa, yellow fever card
- Medical: PEME, D&A test, vaccination, medical fitness
- Licenses and Certification: Certificate of Competency, Certificate of Proficiency, Endorsements, Flag Endorsements
- Courses: STCW training, value-added courses, refresher courses, basic safety training
- Education: degree, diploma, school certificate
- Sea Service Experience: discharge book, sea service letter, vessel employment record
- Bank Details: bank statement, account details, SWIFT/IFSC letter
- Personal Information: any ID document not covered above
"""

# ── VESSEL PROMPT ─────────────────────────────────────────────
VESSEL_PROMPT = """You are a maritime vessel certificate extraction engine.

CLASSIFICATION SOCIETIES (these go in classification_society ONLY):
ABS, American Bureau of Shipping, DNV, Det Norske Veritas, Lloyd's Register, LR,
ClassNK, Nippon Kaiji Kyokai, Bureau Veritas, BV, China Classification Society, CCS,
RINA, Registro Italiano Navale, Korean Register, KR, Indian Register of Shipping, IRS,
Polish Register of Shipping, PRS, Croatian Register of Shipping, CRS

ISSUING AUTHORITY = the flag state government, maritime administration, or delegated authority
that legally issued the certificate (e.g. "Republic of Panama - Maritime Authority",
"Bahamas Maritime Authority", "Marshall Islands Registry", "UK Maritime and Coastguard Agency").
A classification society can be a delegated issuing authority for statutory certs - if the cert
was physically issued by DNV on behalf of a flag state, issuing_authority = flag state,
classification_society = DNV.

DATE FORMAT: DD-MMM-YYYY. LIFETIME stays as LIFETIME.
Return ONLY valid JSON, no markdown.

{
  "is_vessel_document": true,
  "document_type": "",
  "vessel_certificate": {
    "certificate_name": "",
    "certificate_category": "",
    "vessel_name": "",
    "imo_number": "",
    "certificate_number": "",
    "cert_type": "",
    "issuing_authority": "",
    "classification_society": "",
    "date_of_issue": "",
    "date_of_expiry": "",
    "date_of_survey": ""
  },
  "missing_fields": [],
  "validation_flags": [],
  "raw_text_summary": ""
}

certificate_category must be one of: "Class", "Flag State", "Statutory"
cert_type must be one of: "Interim", "Full Term", "Short Term"
"""

# ── RESUME PROMPT ─────────────────────────────────────────────
RESUME_PROMPT = """You are SeaBotics.ai - maritime crew ERP document extraction engine for CVs and resumes.

CRITICAL DISTINCTION - read carefully:
- licenses_certifications = Official legal documents issued by a maritime authority granting rank/qualification to work.
  Examples: Certificate of Competency (CoC), Certificate of Proficiency (CoP), GMDSS General Operator Certificate,
  Flag Endorsements, STCW Endorsements. These have an issuing authority (MARINA, MCA, flag state).
- courses = Training classes attended at a training centre. These prove attendance, not legal qualification.
  Examples: Basic Safety Training, ECDIS course, Advanced Fire Fighting, ARPA, crowd management.
  These are issued by training institutions, not maritime authorities.
  Rule: if it sounds like a class you attended at a school/centre, it is a course.
  Rule: if it is a government/authority issued document granting you the right to work at a rank, it is a license.

DATE FORMAT: DD-MMM-YYYY. LIFETIME stays as LIFETIME.
Return ONLY valid JSON, no markdown.

{
  "document_type": "",
  "is_vessel_document": false,
  "is_resume": true,
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
  "missing_fields":[],
  "validation_flags":[],
  "raw_text_summary":""
}"""

# ── SECTION-SPECIFIC PROMPTS ──────────────────────────────────
SECTION_PROMPTS = {

"Medical": """You are a maritime medical document extractor.
Extract medical certificate / vaccination / test data. Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","date_of_birth":"","nationality":"","passport_number":"","cdc_number":""},
  "medical": [{"document_type":"","document_name":"","certificate_number":"","date_of_issue":"","date_of_expiry":""}],
  "missing_fields":[]
}
document_type must be one of: "Medical Test","Vaccination"
document_name examples: "PEME - Pre-Employment Medical Examination","D&A - Drug & Alcohol Test","Yellow Fever","COVID 19"
""",

"Travel Documents": """You are a maritime travel document extractor.
Extract passport, CDC, seaman's book, visa, and similar documents. Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","surname":"","date_of_birth":"","nationality":"","gender":"","blood_group":"","current_rank":"","address":"","country":"","state":"","zip_code":"","domestic_airport":"","international_airport":""},
  "travel_documents": [{"document_type":"","document_name":"","certificate_number":"","place_of_issue":"","date_of_issue":"","date_of_expiry":""}],
  "missing_fields":[]
}
document_type examples: "Passport","CDC","Seaman's Book","Visa","Yellow Fever Card"
""",

"Licenses and Certification": """You are a maritime licensing document extractor.
Extract Certificate of Competency, Certificate of Proficiency, Endorsements, Flag Endorsements.
Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","date_of_birth":"","nationality":"","cdc_number":"","current_rank":""},
  "licenses_certifications": [{"document_name":"","certificate_number":"","type":"","issuing_authority":"","place_of_issue":"","date_of_issue":"","date_of_expiry":""}],
  "missing_fields":[]
}
type must be one of: "STCW - CERTIFICATE OF COMPETENCY","STCW - CERTIFICATE OF PROFICIENCY","STCW - ENDORSMENT","FLAG - FLAG ENDORSEMENT"
""",

"Courses": """You are a maritime training course document extractor.
Extract STCW training, value-added courses, refresher courses, basic safety training.
Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","date_of_birth":"","nationality":""},
  "courses": [{"document_name":"","certificate_number":"","type":"","place_of_issue":"","date_of_issue":"","date_of_expiry":""}],
  "missing_fields":[]
}
type must be one of: "STCW","FLAG Endorsement","Value Added","Other"
""",

"Education": """You are an education document extractor for maritime crew profiles.
Extract degrees, diplomas, school/college certificates. Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","date_of_birth":"","nationality":""},
  "education": [{"institution_name":"","country":"","education_level":"","field_of_study":"","start_date":"","completion_date":""}],
  "missing_fields":[]
}
""",

"Sea Service Experience": """You are a maritime sea service record extractor.
Extract employment/discharge records, sea service letters, vessel history.
Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","cdc_number":"","current_rank":""},
  "sea_service": [{"company":"","rank":"","vessel_name":"","vessel_type":"","voyage_type":"","imo_number":"","flag_country":"","grt":"","dwt":"","bhp":"","main_engine_type":"","main_engine_kw":"","sign_on_port":"","sign_off_port":"","start_date":"","end_date":"","reason_sign_off":""}],
  "missing_fields":[]
}
vessel_type must be one of: "BULK CARRIER","CHEMICAL TANKER","CONTAINER SHIP","CRUISE SHIP","GAS CARRIER","GENERAL CARGO SHIP","LIVESTOCK CARRIER","OIL TANKER","RO-RO SHIP","TUGBOAT"
voyage_type must be one of: "Foreign Going","Near Coastal","Coastal","Home Trade"
""",

"Bank Details": """You are a bank details extractor for maritime crew profiles.
Extract bank account information. Also extract any personal info found.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {"full_name":"","nationality":""},
  "bank_details": [{"bank_name":"","branch_name":"","account_type":"","account_name":"","account_number":"","ifsc_swift":""}],
  "missing_fields":[]
}
""",

"Personal Information": """You are a personal information extractor for maritime crew profiles.
Extract all personal/identity details from any ID document.
DATE FORMAT: DD-MMM-YYYY. Return ONLY valid JSON, no markdown.
{
  "document_type": "",
  "personal": {
    "full_name":"","surname":"","email":"","gender":"","contact_number":"",
    "date_of_birth":"","nationality":"","blood_group":"","current_rank":"",
    "identification_mark":"","availability_date":"","address":"","country":"",
    "state":"","zip_code":"","domestic_airport":"","international_airport":"",
    "passport_number":"","cdc_number":"","height":"","weight":"",
    "boiler_suit_size":"","safety_shoe_size":"","shirt_size":"","trouser_size":""
  },
  "missing_fields":[]
}
""",
}

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

def _llm_call(system_prompt, user_text, max_tokens=4000, text_limit=12000):
    """Core LLM caller - shared by all pipelines."""
    client = Groq(api_key=GROQ_API_KEY)
    if len(user_text) > text_limit:
        user_text = user_text[:text_limit]
    if not user_text.strip():
        return None
    for attempt in range(3):
        try:
            r = client.chat.completions.create(
                model=MODEL_NAME, temperature=0.1, max_tokens=max_tokens,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user",   "content": f"Extract all information:\n\n{user_text}"}
                ])
            raw = r.choices[0].message.content.strip()
            if not raw: continue
            raw = re.sub(r"^```[a-z]*\n?", "", raw)
            raw = re.sub(r"\n?```$", "", raw)
            start = raw.find("{"); end = raw.rfind("}") + 1
            if start == -1 or end == 0: continue
            return json.loads(raw[start:end])
        except json.JSONDecodeError:
            continue
        except Exception as e:
            st.error(f"API error: {e}"); return None
    st.error("Pipeline failed after 3 attempts.")
    return None


def classify_document(text):
    """Step 1 for Master Upload - lightweight classifier, ~200 tokens."""
    # Only feed first 3000 chars to classifier - enough to identify doc type
    # saves TPM and avoids rate limit on the classify call
    return _llm_call(CLASSIFIER_PROMPT, text, max_tokens=200, text_limit=3000)


def call_llm(text, target_section=None, is_vessel=False, is_resume=False):
    """
    Unified entry point.
    - target_section set  -> use section-specific prompt (1 call)
    - is_vessel = True    -> use vessel prompt (1 call)
    - is_resume = True    -> use resume prompt (1 call)
    - all None/False      -> full resume prompt as fallback
    """
    if is_vessel:
        return _llm_call(VESSEL_PROMPT, text, max_tokens=1500)
    if is_resume:
        # Resume gets more text but still capped to avoid TPM breach on Groq free tier
        # 8000 chars ~ 2000 tokens input, well within 6000 TPM limit
        return _llm_call(RESUME_PROMPT, text, max_tokens=5000, text_limit=8000)
    if target_section and target_section in SECTION_PROMPTS:
        return _llm_call(SECTION_PROMPTS[target_section], text, max_tokens=2000)
    # fallback - treat as resume/full extract
    return _llm_call(RESUME_PROMPT, text, max_tokens=5000)

def conf_color(c):
    if c>=0.90: return "#22c55e"
    if c>=0.60: return "#f59e0b"
    if c>0.00:  return "#ef4444"
    return "#e2e8f0"

def field_coherence(field_label, value):
    """Return 0.0-1.0 coherence score for a free-text value given its field label."""
    if not value:
        return 0.0
    v    = value.strip()
    fl   = field_label.lower()
    _len = max(len(v), 1)
    # Name fields: only letters, spaces, hyphens, apostrophes, dots
    if fl in ("full name", "surname"):
        valid = len(re.findall(r"[a-zA-Z \-\.\']", v))
        return round(valid / _len, 2)
    # Numeric fields
    if fl in ("height", "weight"):
        return 1.0 if re.match(r"^\d+(\.\d+)?\s*(cm|kg|lbs?|m)?$", v, re.IGNORECASE) else 0.3
    # Email
    if fl == "email":
        return 1.0 if re.match(r"^[\w.+\-]+@[\w\-]+\.[\w.]+$", v) else 0.2
    # Phone
    if fl == "contact number":
        valid = len(re.findall(r"[\d\+\-\(\) ]", v))
        return round(min(1.0, valid / _len), 2)
    # Document ID numbers
    if fl in ("passport number", "cdc / seaman's book number"):
        valid = len(re.findall(r"[a-zA-Z0-9\-\/]", v))
        return round(min(1.0, valid / _len), 2)
    # Mostly-alpha fields
    if fl in ("nationality", "country", "state", "address"):
        valid = len(re.findall(r"[a-zA-Z0-9 \-\/\.,]", v))
        return round(min(1.0, valid / _len), 2)
    # Date fields - defer to validate_field_value flag
    if "date" in fl:
        _, flag = validate_field_value(field_label, v)
        return 0.7 if flag else 1.0
    # General: penalise obviously junk characters ($#@!%^&*)
    bad = len(re.findall(r"[$#@!%^&*=|<>~`]", v))
    if bad == 0:
        return 1.0
    return max(0.1, round(1.0 - (bad / _len) * 3, 2))

def render_field_alerts(issues):
    """Render a compact inline badge strip summarising field issues.
    issues: list of (field_label, status) or (field_label, status, note)
    status in ("Invalid", "Verify", "Missing")
    """
    if not issues:
        return
    n_inv = sum(1 for x in issues if x[1] == "Invalid")
    n_ver = sum(1 for x in issues if x[1] == "Verify")
    n_mis = sum(1 for x in issues if x[1] == "Missing")
    badges = []
    if n_inv:
        badges.append(
            f"<span style=\"background:#fef2f2;color:#ef4444;padding:2px 7px;border-radius:4px;font-size:0.7rem;font-weight:600;\">{n_inv} Invalid</span>"
        )
    if n_ver:
        badges.append(
            f"<span style=\"background:#fffbeb;color:#d97706;padding:2px 7px;border-radius:4px;font-size:0.7rem;font-weight:600;\">{n_ver} to Verify</span>"
        )
    if n_mis:
        badges.append(
            f"<span style=\"background:#f8fafc;color:#94a3b8;padding:2px 7px;border-radius:4px;font-size:0.7rem;font-weight:600;\">{n_mis} Missing</span>"
        )
    st.markdown(
        "<div style=\"display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin:2px 0 8px 0;\">"
        + "".join(badges)
        + "</div>",
        unsafe_allow_html=True
    )

DATE_PATTERN = re.compile(
    r"^\d{1,2}-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4}$",
    re.IGNORECASE
)

VALID_GENDERS = {"male", "female", "m", "f", "other"}

def validate_field_value(col_name, value):
    """
    Returns (confidence_penalty, flag_message) for illogical or malformed values.
    confidence_penalty is subtracted from base confidence (0.0 = no issue).
    """
    if not value or not value.strip():
        return 0.0, None

    v = value.strip()
    col_lower = col_name.lower()

    # Date fields
    if "date" in col_lower:
        if v.upper() == "LIFETIME":
            return 0.0, None
        if not DATE_PATTERN.match(v):
            return 0.3, f"{col_name}: '{v}' is not in DD-MMM-YYYY format"

    # Gender field
    if col_lower == "gender":
        if v.lower() not in VALID_GENDERS:
            return 0.3, f"Gender: '{v}' is not a recognised value (expected Male/Female)"

    # Numeric fields that should not contain letters
    if col_lower in {"grt","dwt","bhp","main engine kw","height","weight"}:
        clean = v.replace(",","").replace(".","").replace(" ","")
        if not clean.isdigit():
            return 0.2, f"{col_name}: '{v}' expected a numeric value"

    # Certificate/document numbers that are suspiciously short
    if "certificate number" in col_lower or "account number" in col_lower:
        if len(v) < 3:
            return 0.2, f"{col_name}: '{v}' seems too short to be valid"

    return 0.0, None

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
    rows = reclassify_courses_and_licenses(rows)
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
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No records yet - upload a document above to populate.</div>', unsafe_allow_html=True)
        return

    if view_mode == "table":
        cols_def = SECTION_COLUMNS.get(section_name, VESSEL_COLUMNS)
        # Build clean dataframe
        df = pd.DataFrame(rows)
        for c in cols_def:
            if c not in df.columns:
                df[c] = ""
        df = df[cols_def].fillna("")

        # Read-only table view
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        # Grid card view - show ALL fields in card
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

# ── SUB-POPUP (edit individual record) ───────────────────────
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

def render_subpopup():
    sp = st.session_state.get("subpopup")
    if not sp: return

    section  = sp["section"]
    row_idx  = sp["row_idx"]
    row      = sp["row"]
    key      = sp["key"]
    if section == "Personal Information":
        cols_def = [fl for fk, fl in PERSONAL_FIELDS]
    else:
        cols_def = SECTION_COLUMNS.get(section, VESSEL_COLUMNS)
    drops    = DROPDOWN_FIELDS.get(section, {})

    with st.expander(f"Edit Record - {section}", expanded=True):
        st.markdown(f'<div class="section-label">Editing record {row_idx+1} in {section}</div>', unsafe_allow_html=True)

        # Field alert strip for this row
        _sp_issues = []
        for _sp_col in cols_def:
            _sp_v = row.get(_sp_col, "") or ""
            if not _sp_v:
                _sp_issues.append((_sp_col, "Missing"))
                continue
            if _sp_col in drops and not check_dropdown_validity(section, _sp_col, _sp_v):
                _sp_issues.append((_sp_col, "Invalid"))
            else:
                _, _sp_vf = validate_field_value(_sp_col, _sp_v)
                if _sp_vf or field_coherence(_sp_col, _sp_v) < 0.60:
                    _sp_issues.append((_sp_col, "Verify"))
        h1,h2,h3,h4,h5 = st.columns([2,2.5,1.2,1.2,2.5])
        h1.markdown("**Field**")
        h2.markdown("**Value**")
        h3.markdown("**Confidence**")
        h4.markdown("**Status**")
        h5.markdown("**System Fields**")
        st.markdown("<hr style='margin:4px 0 8px 0;border-color:#e2e8f0;'>", unsafe_allow_html=True)

        edited_row = {}
        edited_org = {}

        _ABBR_MAP = {
            "M": "Male", "F": "Female",
            "MALE": "Male", "FEMALE": "Female", "OTHER": "Other",
            "NA": "N/A", "N/A": "N/A", "NONE": "",
        }
        for col_name in cols_def:
            val  = row.get(col_name,"") or ""
            _, val_flag = validate_field_value(col_name, val)
            is_dropdown = col_name in drops
            dropdown_options = get_dropdown_options(section, col_name) if is_dropdown else []

            fuzzy_val = val
            val_pre   = val
            if is_dropdown and val and dropdown_options:
                val_pre = _ABBR_MAP.get(val.strip().upper(), val)
                matched = fuzzy_match_dropdown(val_pre, dropdown_options)
                if matched:
                    fuzzy_val = matched

            # Live confidence: read current widget value from session state
            _sp_wkey = f"v_sp_{key}_{col_name}"
            _sp_cur  = st.session_state.get(_sp_wkey, fuzzy_val)
            if _sp_cur == "- Select -":
                _sp_cur = ""
            _, val_flag  = validate_field_value(col_name, _sp_cur if _sp_cur else val)
            is_invalid = bool(_sp_cur and is_dropdown and not check_dropdown_validity(section, col_name, _sp_cur))

            if not _sp_cur:
                conf = 0.0
            elif is_invalid:
                conf = 0.0
            elif is_dropdown:
                conf = round(SequenceMatcher(None, val_pre.lower(), _sp_cur.lower()).ratio(), 2)
            else:
                conf = field_coherence(col_name, _sp_cur)

            color = conf_color(conf)

            if is_invalid:
                status_txt,status_cls = "Invalid","status-invalid"
            elif val_flag:
                status_txt,status_cls = "Verify","status-verify"
            elif conf >= 0.90:
                status_txt,status_cls = "Matched","status-matched"
            elif conf >= 0.60:
                status_txt,status_cls = "Verify","status-verify"
            else:
                status_txt,status_cls = "Manual Entry","status-manual"

            c1,c2,c3,c4,c5 = st.columns([2,2.5,1.2,1.2,2.5])

            with c1:
                st.markdown(
                    f'<div style="font-size:0.82rem;color:#1e293b;font-weight:500;padding-top:6px;">{col_name}</div>'
                    f'<div style="font-size:0.68rem;color:#94a3b8;font-style:italic;">Extracted: {val if val else "-"}</div>',
                    unsafe_allow_html=True
                )
                if is_invalid:
                    st.markdown('<span class="invalid-flag">Not in dropdown</span>', unsafe_allow_html=True)
                if val_flag:
                    st.markdown(f'<span class="invalid-flag">Format issue</span>', unsafe_allow_html=True)

            with c2:
                if is_dropdown and dropdown_options:
                    opts = ["- Select -"] + dropdown_options
                    cur_idx = next((i+1 for i,o in enumerate(dropdown_options) if o.upper()==fuzzy_val.upper()), 0)
                    edited_row[col_name] = st.selectbox(
                        f"v_{key}_{col_name}", options=opts, index=cur_idx,
                        key=f"v_sp_{key}_{col_name}", label_visibility="collapsed"
                    )
                    if edited_row[col_name] == "- Select -": edited_row[col_name] = ""
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
                org_opts = ["- Select org field -"] + ALL_ORG_OPTIONS
                default_idx = next((i+1 for i,o in enumerate(ALL_ORG_OPTIONS) if o==org_default), 0)
                edited_org[col_name] = st.selectbox(
                    f"o_{key}_{col_name}", options=org_opts, index=default_idx,
                    key=f"o_sp_{key}_{col_name}", label_visibility="collapsed"
                )

        st.markdown("---")
        sa,sb = st.columns(2)
        with sa:
            if st.button("Save Changes", type="primary", use_container_width=True, key=f"save_sp_{key}"):
                source = st.session_state["subpopup"].get("source", "profile")
                if source == "personal":
                    # Write back to profile using PERSONAL_FIELDS label as key
                    for fk, fl in PERSONAL_FIELDS:
                        if fl in edited_row:
                            org_key = {fl2: fk for fk2, fl2 in PERSONAL_FIELDS}.get(fl, fl)
                            ORG_MAP_INV = {
                                "Full Name":"Full Name","Surname":"Surname",
                                "Email":"Email Address","Gender":"Gender",
                                "Contact Number":"Phone Number","Date of Birth":"Date of Birth",
                                "Nationality":"Nationality","Blood Group":"Blood Group",
                                "Rank":"Rank / Designation","Identification Mark":"Identification Mark",
                                "Availability Date":"Availability Date","Address":"Residential Address",
                                "Country":"Country","State":"State","ZIP / PIN Code":"ZIP / PIN Code",
                                "Domestic Airport":"Domestic Airport",
                                "International Airport":"International Airport",
                                "Passport Number":"Passport Number",
                                "CDC / Seaman's Book Number":"CDC / Seaman's Book Number",
                                "Height":"Height","Weight":"Weight",
                                "Boiler Suit Size":"Boiler Suit Size",
                                "Safety Shoe Size":"Safety Shoe Size",
                                "Shirt Size":"Shirt Size","Trouser Size":"Trouser Size",
                            }
                            profile_key = ORG_MAP_INV.get(fl, fl)
                            if edited_row[fl]:
                                st.session_state["profile"][profile_key] = edited_row[fl]
                elif source == "profile":
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
def approve_and_populate(result, is_vessel, pending_rows, pending_vessel, personal, selected_sections=None):
    """
    Shared approval logic. Fills empty fields only - skips already populated ones.
    selected_sections: set of section names to populate. None = all sections.
    """
    if is_vessel:
        if selected_sections is None or "Vessels" in selected_sections:
            if pending_vessel and any(v for v in pending_vessel.values()):
                st.session_state["vessels"].append(pending_vessel)
    else:
        if selected_sections is None or "Personal Information" in selected_sections:
            profile_data = personal_to_profile(personal)
            for k, v in profile_data.items():
                # Only fill if field is currently empty
                if v and not st.session_state["profile"].get(k):
                    st.session_state["profile"][k] = v

            for td in pending_rows.get("Travel Documents", []):
                dt = (td.get("Document Type","") + td.get("Document Name","")).lower()
                if "passport" in dt and td.get("Certificate Number"):
                    if not st.session_state["profile"].get("Passport Number"):
                        st.session_state["profile"]["Passport Number"] = td["Certificate Number"]
                if ("cdc" in dt or "seaman" in dt) and td.get("Certificate Number"):
                    if not st.session_state["profile"].get("CDC / Seaman's Book Number"):
                        st.session_state["profile"]["CDC / Seaman's Book Number"] = td["Certificate Number"]

        for section, rows in pending_rows.items():
            if selected_sections is not None and section not in selected_sections:
                continue
            for row in rows:
                if any(v for v in row.values()):
                    st.session_state["sections"][section].append(row)

def render_popup():
    popup = st.session_state.get("popup")
    if not popup: return

    result    = popup.get("result", {})
    filename  = popup.get("filename", "document")
    is_vessel = result.get("is_vessel_document", False)
    is_resume = result.get("is_resume", False)
    doc_type  = result.get("document_type", "-")
    personal  = result.get("personal", {})
    missing   = result.get("missing_fields", [])
    flags     = result.get("validation_flags", [])

    # Build pending rows once; reset section nav on each new popup
    if "pending_rows" not in popup:
        _draft_state = popup.get("draft_state")
        if _draft_state:
            popup["pending_rows"]   = _draft_state.get("pending_rows", result_to_rows(result))
            popup["pending_vessel"] = _draft_state.get("pending_vessel", vessel_result_to_row(result) if is_vessel else None)
            if _draft_state.get("personal"):
                result["personal"] = _draft_state["personal"]
        else:
            popup["pending_rows"]   = result_to_rows(result)
            popup["pending_vessel"] = vessel_result_to_row(result) if is_vessel else None
        st.session_state["popup"] = popup
        st.session_state["popup_active_section"] = None
        st.session_state["popup_discard_confirm"] = False
        # Flush widget keys from any previous popup to prevent stale confidence scores
        for _wk in list(st.session_state.keys()):
            if _wk.startswith(("pp_popup_", "vc_popup_", "v_sp_")):
                del st.session_state[_wk]
        # Initialise section selections - all True by default
        _init_sel = {}
        if is_vessel:
            _init_sel["Vessels"] = True
        else:
            if any(v for v in personal.values()):
                _init_sel["Personal Information"] = True
            for _sn, _srows in popup["pending_rows"].items():
                if _srows:
                    _init_sel[_sn] = True
        st.session_state["popup_section_selections"] = _init_sel

    pending_rows   = popup["pending_rows"]
    pending_vessel = popup.get("pending_vessel")
    _dt_approved   = (popup.get("source") == "doc_tank" and popup.get("tank_status", "") in ("Approved", "Partial"))

    # Count low-confidence fields in one section row
    def _row_low_conf(section_name, row):
        drops = DROPDOWN_FIELDS.get(section_name, {})
        count = 0
        for col_name, val in row.items():
            if not val:
                continue
            penalty, _ = validate_field_value(col_name, val)
            conf = max(0.0, 0.9 - penalty)
            if col_name in drops and not check_dropdown_validity(section_name, col_name, val):
                conf = 0.0
            if conf < 0.90:
                count += 1
        return count

    # Count personal fields that fail format validation
    def _personal_low_conf():
        count = 0
        for k, v in personal.items():
            if not v:
                continue
            penalty, _ = validate_field_value(k, v)
            if penalty > 0:
                count += 1
        return count

    with st.expander(f"Review & Approve - {filename}", expanded=True):

        # Doc tank status banner
        if popup.get("source") == "doc_tank":
            tank_status = popup.get("tank_status", "Pending")
            if tank_status == "Approved":
                st.info("This document was already approved. To update, re-upload the document.")
            else:
                st.warning("This document is pending approval. Review and approve below.")

        m1, m2, m3 = st.columns(3)
        m1.metric("Document Type", doc_type)
        m2.metric("Vessel Document", "Yes" if is_vessel else "No")
        m3.metric("CV / Resume", "Yes" if is_resume else "No")

        st.markdown("---")

        # ── VESSEL FLOW: inline editable confidence table ─────
        if is_vessel and pending_vessel:
            if not _dt_approved:
                _vsl_chk = st.checkbox(
                    "Include in Approve and Populate",
                    value=st.session_state.get("popup_section_selections", {}).get("Vessels", True),
                    key="ov_chk_vessel"
                )
                st.session_state.setdefault("popup_section_selections", {})["Vessels"] = _vsl_chk
            st.info("Vessel certificate detected - will populate the **Vessels** section.")
            st.markdown('<div class="section-label">Vessel Certificate</div>', unsafe_allow_html=True)
            _vsl_drops = DROPDOWN_FIELDS.get("Vessels", {})
            vh1, vh2, vh3, vh4, vh5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
            vh1.markdown("**Field**")
            vh2.markdown("**Value**")
            vh3.markdown("**Confidence**")
            vh4.markdown("**Status**")
            vh5.markdown("**System Fields**")
            st.markdown("<hr style='margin:4px 0 8px 0;border-color:#e2e8f0;'>", unsafe_allow_html=True)
            for _vc_col in VESSEL_COLUMNS:
                _vc_val = pending_vessel.get(_vc_col, "") or ""
                _, _vc_flag = validate_field_value(_vc_col, _vc_val)
                _vc_drop_opts = _vsl_drops.get(_vc_col, [])
                _vc_is_drop   = bool(_vc_drop_opts)
                _vc_matched   = _vc_val
                if _vc_is_drop and _vc_val:
                    _vc_pm = fuzzy_match_dropdown(_vc_val, _vc_drop_opts)
                    if _vc_pm:
                        _vc_matched = _vc_pm
                # Live confidence: read current widget value from session state
                _vc_wkey = f"vc_popup_{_vc_col}"
                _vc_cur  = st.session_state.get(_vc_wkey, _vc_matched)
                if _vc_cur == "- Select -":
                    _vc_cur = ""
                _, _vc_flag = validate_field_value(_vc_col, _vc_cur if _vc_cur else _vc_val)
                _vc_invalid = bool(_vc_cur and _vc_is_drop and not check_dropdown_validity("Vessels", _vc_col, _vc_cur))
                if not _vc_cur:
                    _vc_conf = 0.0
                elif _vc_invalid:
                    _vc_conf = 0.0
                elif _vc_is_drop:
                    _vc_conf = round(SequenceMatcher(None, _vc_val.lower(), _vc_cur.lower()).ratio(), 2)
                else:
                    _vc_conf = field_coherence(_vc_col, _vc_cur)
                _vc_color = conf_color(_vc_conf)
                if _vc_invalid:
                    _vc_status_txt, _vc_status_cls = "Invalid",      "status-invalid"
                elif _vc_flag:
                    _vc_status_txt, _vc_status_cls = "Verify",       "status-verify"
                elif _vc_conf >= 0.90:
                    _vc_status_txt, _vc_status_cls = "Matched",      "status-matched"
                elif _vc_conf > 0.0:
                    _vc_status_txt, _vc_status_cls = "Verify",       "status-verify"
                else:
                    _vc_status_txt, _vc_status_cls = "Manual Entry", "status-manual"
                vc1, vc2, vc3, vc4, vc5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
                with vc1:
                    st.markdown(
                        f'<div style="font-size:0.82rem;color:#1e293b;font-weight:500;padding-top:6px;">{_vc_col}</div>'
                        f'<div style="font-size:0.68rem;color:#94a3b8;font-style:italic;">Extracted: {_vc_val if _vc_val else "-"}</div>',
                        unsafe_allow_html=True
                    )
                    if _vc_invalid:
                        st.markdown('<span class="invalid-flag">Not in dropdown</span>', unsafe_allow_html=True)
                    if _vc_flag:
                        st.markdown('<span class="invalid-flag">Format issue</span>', unsafe_allow_html=True)
                with vc2:
                    if _vc_is_drop and not _dt_approved:
                        _vc_opts = ["- Select -"] + _vc_drop_opts
                        _vc_idx  = next((i+1 for i, o in enumerate(_vc_drop_opts) if o.casefold() == _vc_matched.casefold()), 0)
                        _vc_sel  = st.selectbox(f"vc_{_vc_col}", options=_vc_opts, index=_vc_idx,
                                                key=f"vc_popup_{_vc_col}", label_visibility="collapsed")
                        pending_vessel[_vc_col] = "" if _vc_sel == "- Select -" else _vc_sel
                    elif not _dt_approved:
                        pending_vessel[_vc_col] = st.text_input(
                            f"vc_{_vc_col}", value=_vc_matched,
                            key=f"vc_popup_{_vc_col}", label_visibility="collapsed"
                        )
                    else:
                        st.markdown(
                            f'<div style="padding-top:6px;font-size:0.85rem;color:#1e293b;">{_vc_val or "-"}</div>',
                            unsafe_allow_html=True
                        )
                with vc3:
                    st.markdown(
                        f'<div style="padding-top:6px;">'
                        f'<div style="background:#f1f5f9;border-radius:4px;height:6px;width:70px;">'
                        f'<div style="width:{int(_vc_conf*100)}%;height:6px;background:{_vc_color};border-radius:4px;"></div></div>'
                        f'<div style="font-size:0.65rem;color:{_vc_color};margin-top:2px;">{int(_vc_conf*100)}%</div></div>',
                        unsafe_allow_html=True
                    )
                with vc4:
                    st.markdown(
                        f'<div style="padding-top:8px;"><span class="{_vc_status_cls}">{_vc_status_txt}</span></div>',
                        unsafe_allow_html=True
                    )
                with vc5:
                    _vc_org_default = ORG_PREMAP.get(_vc_col, "Other / Unclassified")
                    _vc_org_opts = ["- Select org field -"] + ALL_ORG_OPTIONS
                    _vc_org_idx  = next((i+1 for i,o in enumerate(ALL_ORG_OPTIONS) if o == _vc_org_default), 0)
                    st.selectbox(f"o_vc_{_vc_col}", options=_vc_org_opts, index=_vc_org_idx,
                                 key=f"o_vc_popup_{_vc_col}", label_visibility="collapsed")

        # ── CREW FLOW: section card overview + drill-in ───────
        if not is_vessel:
            active_sec = st.session_state.get("popup_active_section")

            # ── OVERVIEW: one card per section that has data ──
            if active_sec is None:
                _sel_state = st.session_state.get("popup_section_selections", {})
                show_personal = any(v for v in personal.values())
                if show_personal:
                    p_low   = _personal_low_conf()
                    p_total = len([v for v in personal.values() if v])
                    p_badge = (
                        f'<span class="status-verify">{p_low} to verify</span>'
                        if p_low else
                        '<span class="status-matched">All matched</span>'
                    )
                    ov_chk, ov1, ov2 = st.columns([0.5, 3.5, 1])
                    with ov_chk:
                        if not _dt_approved:
                            st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
                            _pi_chk = st.checkbox("", value=_sel_state.get("Personal Information", True),
                                                  key="ov_chk_personal", label_visibility="collapsed")
                            st.session_state.setdefault("popup_section_selections", {})["Personal Information"] = _pi_chk
                    with ov1:
                        st.markdown(
                            f'<div class="grid-card">'
                            f'<div class="grid-card-title">Personal Information</div>'
                            f'<div class="grid-card-sub">{p_total} field(s) extracted &nbsp;|&nbsp; {p_badge}</div>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
                    with ov2:
                        if st.button("Review", key="popup_open_personal", use_container_width=True):
                            st.session_state["popup_active_section"] = "Personal Information"
                            st.rerun()

                for section_name, rows in pending_rows.items():
                    if not rows:
                        continue
                    total_low = sum(_row_low_conf(section_name, row) for row in rows)
                    s_badge = (
                        f'<span class="status-verify">{total_low} to verify</span>'
                        if total_low else
                        '<span class="status-matched">All matched</span>'
                    )
                    ov_chk, ov1, ov2 = st.columns([0.5, 3.5, 1])
                    with ov_chk:
                        if not _dt_approved:
                            st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
                            _sec_chk = st.checkbox("", value=_sel_state.get(section_name, True),
                                                   key=f"ov_chk_{section_name}", label_visibility="collapsed")
                            st.session_state.setdefault("popup_section_selections", {})[section_name] = _sec_chk
                    with ov1:
                        st.markdown(
                            f'<div class="grid-card">'
                            f'<div class="grid-card-title">{section_name}</div>'
                            f'<div class="grid-card-sub">{len(rows)} record(s) &nbsp;|&nbsp; {s_badge}</div>'
                            f'</div>',
                            unsafe_allow_html=True
                        )
                    with ov2:
                        if st.button("Review", key=f"popup_open_{section_name}", use_container_width=True):
                            st.session_state["popup_active_section"] = section_name
                            st.rerun()

            # ── PERSONAL INFO DRILL-IN ────────────────────────
            elif active_sec == "Personal Information":
                if st.button("< Back", key="popup_back_personal"):
                    st.session_state["popup_active_section"] = None
                    st.rerun()
                st.markdown('<div class="section-label">Personal Information</div>', unsafe_allow_html=True)
                if popup.get("source") == "section":
                    st.info("Personal info found in this document - review below before approving.")
                if _dt_approved:
                    # Read-only display for approved documents
                    _PK_RO = {"person_name": "full_name", "rank": "current_rank"}
                    _ro_items = [(fl, personal.get(_PK_RO.get(fk, fk), "") or "") for fk, fl in PERSONAL_FIELDS]
                    _filled = [(fl, v) for fl, v in _ro_items if v]
                    st.caption(f"{len(_filled)} of {len(PERSONAL_FIELDS)} fields extracted")
                    for i in range(0, len(_ro_items), 3):
                        grp = _ro_items[i:i+3]
                        ro_cols = st.columns(3)
                        for j, (fl, v) in enumerate(grp):
                            with ro_cols[j]:
                                st.markdown(f'<div class="field-label">{fl}</div>', unsafe_allow_html=True)
                                if v:
                                    st.markdown(f'<div class="field-value">{v}</div>', unsafe_allow_html=True)
                                else:
                                    st.markdown(f'<div class="field-empty">-</div>', unsafe_allow_html=True)
                else:
                    # Confidence table header
                    ph1, ph2, ph3, ph4, ph5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
                    ph1.markdown("**Field**")
                    ph2.markdown("**Value**")
                    ph3.markdown("**Confidence**")
                    ph4.markdown("**Status**")
                    ph5.markdown("**System Fields**")
                    st.markdown("<hr style='margin:4px 0 8px 0;border-color:#e2e8f0;'>", unsafe_allow_html=True)
                    # LLM returns these under different keys than PERSONAL_FIELDS fk
                    _P_KEY   = {"person_name": "full_name", "rank": "current_rank"}
                    _p_drops = DROPDOWN_FIELDS.get("Personal Information", {})
                    _P_ABBR_MAP = {
                        "M": "Male", "F": "Female",
                        "MALE": "Male", "FEMALE": "Female", "OTHER": "Other",
                        "NA": "N/A", "N/A": "N/A", "NONE": "",
                    }
                    for fk, fl in PERSONAL_FIELDS:
                        _pk  = _P_KEY.get(fk, fk)
                        val  = personal.get(_pk, "") or ""
                        _p_drop_opts = _p_drops.get(fl, [])
                        _is_p_drop   = bool(_p_drop_opts)
                        _p_val_pre   = val
                        _p_matched   = val
                        if _is_p_drop and val:
                            _p_val_pre = _P_ABBR_MAP.get(val.strip().upper(), val)
                            _pm = fuzzy_match_dropdown(_p_val_pre, _p_drop_opts)
                            if _pm:
                                _p_matched = _pm
                        # Live confidence: read current widget value from session state
                        _pp_wkey = f"pp_popup_{_pk}"
                        _pp_cur  = st.session_state.get(_pp_wkey, _p_matched)
                        if _pp_cur == "- Select -":
                            _pp_cur = ""
                        _, val_flag = validate_field_value(fl, _pp_cur if _pp_cur else val)
                        _is_p_invalid = bool(_pp_cur and _is_p_drop and not check_dropdown_validity("Personal Information", fl, _pp_cur))

                        if not _pp_cur:
                            conf = 0.0
                        elif _is_p_invalid:
                            conf = 0.0
                        elif _is_p_drop:
                            conf = round(SequenceMatcher(None, _p_val_pre.lower(), _pp_cur.lower()).ratio(), 2)
                        else:
                            conf = field_coherence(fl, _pp_cur)

                        color = conf_color(conf)

                        if _is_p_invalid:
                            status_txt, status_cls = "Invalid",      "status-invalid"
                        elif val_flag:
                            status_txt, status_cls = "Verify",       "status-verify"
                        elif conf >= 0.90:
                            status_txt, status_cls = "Matched",      "status-matched"
                        elif conf > 0.0:
                            status_txt, status_cls = "Verify",       "status-verify"
                        else:
                            status_txt, status_cls = "Manual Entry", "status-manual"

                        pc1, pc2, pc3, pc4, pc5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
                        with pc1:
                            st.markdown(
                                f'<div style="font-size:0.82rem;color:#1e293b;font-weight:500;padding-top:6px;">{fl}</div>'
                                f'<div style="font-size:0.68rem;color:#94a3b8;font-style:italic;">Extracted: {val if val else "-"}</div>',
                                unsafe_allow_html=True
                            )
                            if _is_p_invalid:
                                st.markdown('<span class="invalid-flag">Not in dropdown</span>', unsafe_allow_html=True)
                            if val_flag:
                                st.markdown('<span class="invalid-flag">Format issue</span>', unsafe_allow_html=True)
                        with pc2:
                            if _is_p_drop:
                                _p_opts = ["- Select -"] + _p_drop_opts
                                _p_idx  = next((i+1 for i, o in enumerate(_p_drop_opts) if o.casefold() == _p_matched.casefold()), 0)
                                _sel = st.selectbox(f"pp_{_pk}", options=_p_opts, index=_p_idx,
                                                    key=f"pp_popup_{_pk}", label_visibility="collapsed")
                                personal[_pk] = "" if _sel == "- Select -" else _sel
                            else:
                                personal[_pk] = st.text_input(
                                    f"pp_{_pk}", value=val, key=f"pp_popup_{_pk}", label_visibility="collapsed"
                                )
                        with pc3:
                            st.markdown(
                                f'<div style="padding-top:6px;">'
                                f'<div style="background:#f1f5f9;border-radius:4px;height:6px;width:70px;">'
                                f'<div style="width:{int(conf*100)}%;height:6px;background:{color};border-radius:4px;"></div></div>'
                                f'<div style="font-size:0.65rem;color:{color};margin-top:2px;">{int(conf*100)}%</div></div>',
                                unsafe_allow_html=True
                            )
                        with pc4:
                            st.markdown(
                                f'<div style="padding-top:8px;"><span class="{status_cls}">{status_txt}</span></div>',
                                unsafe_allow_html=True
                            )
                        with pc5:
                            _pp_org_default = ORG_PREMAP.get(fl, "Other / Unclassified")
                            _pp_org_opts = ["- Select org field -"] + ALL_ORG_OPTIONS
                            _pp_org_idx  = next((i+1 for i,o in enumerate(ALL_ORG_OPTIONS) if o == _pp_org_default), 0)
                            st.selectbox(f"o_pp_{_pk}", options=_pp_org_opts, index=_pp_org_idx,
                                         key=f"o_pp_popup_{_pk}", label_visibility="collapsed")

            # ── SECTION RECORDS DRILL-IN (table view) ────────────
            else:
                rows     = pending_rows.get(active_sec, [])
                cols_def = SECTION_COLUMNS.get(active_sec, VESSEL_COLUMNS)
                if st.button("< Back", key=f"popup_back_{active_sec}"):
                    st.session_state["popup_active_section"] = None
                    st.rerun()
                st.markdown(f'<div class="section-label">{active_sec} - {len(rows)} record(s)</div>', unsafe_allow_html=True)
                if rows:
                    df = pd.DataFrame(rows)
                    for c in cols_def:
                        if c not in df.columns:
                            df[c] = ""
                    df = df[cols_def].fillna("")
                    if _dt_approved:
                        st.dataframe(df, use_container_width=True, hide_index=True)
                    else:
                        t_col, e_col = st.columns([6, 1])
                        with t_col:
                            st.dataframe(df, use_container_width=True, hide_index=True)
                        with e_col:
                            st.markdown("<div style='height:36px'></div>", unsafe_allow_html=True)
                            for ridx, row in enumerate(rows):
                                if st.button("Edit", key=f"popup_edit_{active_sec}_{ridx}", use_container_width=True):
                                    st.session_state["subpopup"] = {
                                        "section":  active_sec,
                                        "row_idx":  ridx,
                                        "row":      copy.deepcopy(row),
                                        "key":      f"popup_{active_sec}_{ridx}",
                                        "source":   "popup"
                                    }
                                    st.rerun()
                                st.markdown("<div style='height:21px'></div>", unsafe_allow_html=True)

        if flags:
            st.markdown("---")
            for f in flags:
                st.warning(f)
        if missing:
            st.markdown(
                " ".join(f'<span class="missing-tag">{m}</span>' for m in missing),
                unsafe_allow_html=True
            )

        st.markdown("---")
        ba, bb, bc, bd = st.columns([2, 1, 1, 1])
        with ba:
            tank_status       = popup.get("tank_status", "Pending")
            is_dt_source      = popup.get("source") == "doc_tank"
            _already_approved = is_dt_source and tank_status in ("Approved", "Partial")
            _confirming       = st.session_state.get("popup_confirm_approve", False)
            _in_drill         = not is_vessel and st.session_state.get("popup_active_section") is not None
            if _in_drill:
                if st.button("Save Changes", use_container_width=True, key="save_changes_popup"):
                    st.session_state["popup_active_section"] = None
                    st.session_state["popup_confirm_approve"] = False
                    st.rerun()
            elif not _confirming:
                if st.button("Approve and Populate", type="primary", use_container_width=True,
                             key="approve_popup", disabled=_already_approved):
                    st.session_state["popup_confirm_approve"] = True
                    st.rerun()
            else:
                st.warning("Populate profile with selected sections?")
                _conf1, _conf2 = st.columns(2)
                with _conf1:
                    if st.button("Yes, Approve", type="primary", use_container_width=True, key="approve_confirm_yes"):
                        _selections   = st.session_state.get("popup_section_selections", {})
                        _selected_set = set(k for k, v in _selections.items() if v)

                        # Resume wipe clears both profile and all crew sections
                        if result.get("is_resume", False) and not is_vessel:
                            st.session_state["profile"] = {}
                            for _wipe_sec in SECTION_COLUMNS:
                                st.session_state["sections"][_wipe_sec] = []

                        approve_and_populate(result, is_vessel, pending_rows, pending_vessel, personal,
                                             selected_sections=_selected_set if _selected_set else None)

                        # Build approved fields list (only from selected sections)
                        approved_fields = []
                        if not is_vessel:
                            if _selections.get("Personal Information", True):
                                approved_fields += [v for v in personal.values() if v]
                            for sec_name, sec_rows in pending_rows.items():
                                if _selections.get(sec_name, True):
                                    for row in sec_rows:
                                        for k, v in row.items():
                                            if v:
                                                approved_fields.append(f"{k}: {v[:30]}")
                        if is_vessel and pending_vessel:
                            if _selections.get("Vessels", True):
                                approved_fields += [f"{k}: {v[:30]}" for k, v in pending_vessel.items() if v]

                        # Determine Approved vs Partial
                        _all_available = set()
                        if is_vessel:
                            _all_available.add("Vessels")
                        else:
                            if any(v for v in personal.values()):
                                _all_available.add("Personal Information")
                            for _sn, _srows in pending_rows.items():
                                if _srows:
                                    _all_available.add(_sn)
                        _final_status = "Approved" if _selected_set >= _all_available else "Partial"

                        # Mark the doc tank entry
                        tank_id = popup.get("tank_id")
                        for dt_entry in st.session_state["doc_tank"]:
                            if dt_entry.get("tank_id") == tank_id:
                                dt_entry["status"] = _final_status
                                dt_entry["approved_fields"] = approved_fields
                                break

                        st.session_state["popup_confirm_approve"] = False
                        st.session_state["popup_discard_confirm"] = False
                        st.session_state["popup"] = None
                        st.session_state["popup_active_section"] = None
                        st.success("Profile populated successfully.")
                        st.rerun()
                with _conf2:
                    if st.button("Cancel", use_container_width=True, key="approve_confirm_cancel"):
                        st.session_state["popup_confirm_approve"] = False
                        st.rerun()

        with bb:
            st.download_button(
                "Download Raw JSON", data=json.dumps(result, indent=2),
                file_name=f"{Path(filename).stem}_raw.json",
                mime="application/json", key="dl_popup", use_container_width=True
            )
        with bc:
            if not _already_approved:
                if st.button("Save as Draft", use_container_width=True, key="save_draft_popup"):
                    _draft_personal = dict(result.get("personal", {}))
                    _draft_state = {
                        "pending_rows":   copy.deepcopy(pending_rows),
                        "pending_vessel": copy.deepcopy(pending_vessel) if pending_vessel else None,
                        "personal":       _draft_personal,
                    }
                    tank_id = popup.get("tank_id")
                    for dt_entry in st.session_state["doc_tank"]:
                        if dt_entry.get("tank_id") == tank_id:
                            dt_entry["status"]      = "Draft"
                            dt_entry["draft_state"] = _draft_state
                            break
                    st.session_state["popup"] = None
                    st.session_state["popup_active_section"] = None
                    st.session_state["popup_confirm_approve"] = False
                    st.session_state["popup_discard_confirm"] = False
                    st.rerun()
        with bd:
            discard_label = "Close" if popup.get("source") == "doc_tank" else "Discard"
            if _already_approved:
                if st.button(discard_label, use_container_width=True, key="discard_popup"):
                    st.session_state["popup"] = None
                    st.session_state["popup_active_section"] = None
                    st.session_state["popup_confirm_approve"] = False
                    st.session_state["popup_discard_confirm"] = False
                    st.rerun()
            elif st.session_state.get("popup_discard_confirm"):
                st.markdown(
                    '<div style="font-size:0.72rem;color:#f59e0b;font-weight:600;margin-bottom:4px;text-align:center;">Save as Draft?</div>',
                    unsafe_allow_html=True
                )
                _dc1, _dc2 = st.columns(2)
                with _dc1:
                    if st.button("Save Draft", use_container_width=True, key="discard_save_draft"):
                        _d_draft_state = {
                            "pending_rows":   copy.deepcopy(pending_rows),
                            "pending_vessel": copy.deepcopy(pending_vessel) if pending_vessel else None,
                            "personal":       dict(result.get("personal", {})),
                        }
                        _d_tank_id = popup.get("tank_id")
                        for _d_entry in st.session_state["doc_tank"]:
                            if _d_entry.get("tank_id") == _d_tank_id:
                                _d_entry["status"]      = "Draft"
                                _d_entry["draft_state"] = _d_draft_state
                                break
                        st.session_state["popup"] = None
                        st.session_state["popup_active_section"] = None
                        st.session_state["popup_confirm_approve"] = False
                        st.session_state["popup_discard_confirm"] = False
                        st.rerun()
                with _dc2:
                    if st.button("Discard", use_container_width=True, key="discard_confirm_close"):
                        st.session_state["popup"] = None
                        st.session_state["popup_active_section"] = None
                        st.session_state["popup_confirm_approve"] = False
                        st.session_state["popup_discard_confirm"] = False
                        st.rerun()
            else:
                if st.button(discard_label, use_container_width=True, key="discard_popup"):
                    st.session_state["popup_discard_confirm"] = True
                    st.rerun()

# ── HEADER ────────────────────────────────────────────────────
st.markdown("""
<div style="background:linear-gradient(135deg,#0f172a,#1e3a5f);border-radius:12px;padding:16px 24px;margin-bottom:16px;">
    <div style="font-size:1.3rem;font-weight:700;color:#38bdf8;">SeaBotics.ai</div>
    <div style="font-size:0.72rem;color:#94a3b8;margin-top:2px;">Crew & Vessel Profile Management - v5.0</div>
</div>""", unsafe_allow_html=True)

# ── NAV ───────────────────────────────────────────────────────
_n1, _n2, _n3, _ = st.columns([1, 1, 1, 5])
for _ncol, _label, _key in [(_n1, "Crew", "Crew"), (_n2, "Vessel", "Vessel"), (_n3, "Doc Tank", "Doc Tank")]:
    with _ncol:
        _btn_type = "primary" if st.session_state["page"] == _key else "secondary"
        if st.button(_label, use_container_width=True, type=_btn_type, key=f"nav_{_key}"):
            st.session_state["page"] = _key
            st.rerun()

# ── SUBPOPUP (always rendered on top if active) ───────────────
if st.session_state.get("subpopup"):
    render_subpopup()

# ── MAIN POPUP ────────────────────────────────────────────────
elif st.session_state.get("popup"):
    render_popup()

# ════════════════════════════════════════════════════════════
#  MY PROFILE
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "Crew":

    # All Crew tabs
    _crew_tabs = st.tabs([
        "Personal Info", "Experience", "Licenses and Certifications",
        "Medical", "Education", "Courses", "Bank Details",
        "Travel Documents", "Doc Check-List", "Travel Details",
    ])

    # Personal Info tab
    with _crew_tabs[0]:
        st.markdown(
            '<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">'
            '<span style="font-size:1.1rem;">✨</span>'
            '<span style="font-size:0.9rem;font-weight:700;color:#0f172a;">AI-Assisted Profile Fill</span>'
            '<span style="font-size:0.75rem;color:#64748b;">'
            '&nbsp;Upload your CV or Resume - AI reads it and fills all sections for your review'
            '</span></div>',
            unsafe_allow_html=True
        )
        _rcu1, _rcu2 = st.columns([5, 1])
        with _rcu1:
            resume_file = st.file_uploader("Upload CV or Resume",
                type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                key="up_resume", label_visibility="collapsed")
        with _rcu2:
            st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
            if resume_file and st.button("Extract", key="run_resume", use_container_width=True):
                with st.spinner("Running resume pipeline..."):
                    ext = Path(resume_file.name).suffix.lower()
                    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                        tmp.write(resume_file.read()); tmp_path = tmp.name
                    try:
                        text = extract_text(tmp_path)
                        result = call_llm(text, is_resume=True)
                        if result:
                            result["is_vessel_document"] = False
                            result["is_resume"] = True
                            save_to_doc_tank(result, resume_file.name, "master",
                                             result.get("document_type", "CV / Resume"),
                                             False, True)
                            st.session_state["popup"] = {
                                "key": "resume_personal",
                                "result": result,
                                "filename": resume_file.name,
                                "source": "master",
                                "target_section": "Resume",
                                "tank_id": st.session_state["doc_tank"][-1]["tank_id"]
                            }
                            st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

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
                    else:   st.markdown(f'<div class="field-empty">-</div>',unsafe_allow_html=True)
        if profile:
            if st.button("Clear Personal Info", key="clear_personal"):
                st.session_state["profile"] = {}; st.rerun()

    # Multi-row sections - ordered display list (internal key, display label)
    _CREW_ORDER = [
        ("Sea Service Experience",     "Experience"),
        ("Licenses and Certification", "Licenses and Certifications"),
        ("Medical",                    "Medical"),
        ("Education",                  "Education"),
        ("Courses",                    "Courses"),
        ("Bank Details",               "Bank Details"),
        ("Travel Documents",           "Travel Documents"),
    ]
    for (section_name, display_name), _tab_sec in zip(_CREW_ORDER, _crew_tabs[1:8]):
        cols_def = SECTION_COLUMNS[section_name]
        rows=st.session_state["sections"].get(section_name,[])
        with _tab_sec:
            st.markdown(
                f'<div style="font-size:0.8rem;color:#64748b;margin-bottom:8px;">'
                f'{len(rows)} record{"s" if len(rows)!=1 else ""}</div>',
                unsafe_allow_html=True
            )
            # Top action buttons
            _btn_add, _btn_up = st.columns(2)
            with _btn_add:
                if st.button(f"+ Add {display_name}", key=f"add_{section_name}", use_container_width=True):
                    st.session_state["sections"][section_name].append({c:"" for c in cols_def})
                    st.rerun()
            with _btn_up:
                if st.button("✨ AI Autofill", key=f"toggle_up_{section_name}", use_container_width=True):
                    _tk = f"show_up_{section_name}"
                    st.session_state[_tk] = not st.session_state.get(_tk, False)
                    st.rerun()
                st.caption("Upload a document - AI extracts and fills for you")

            if st.session_state.get(f"show_up_{section_name}", False):
                uc1, uc2 = st.columns([4, 1])
                with uc1:
                    sec_file = st.file_uploader(f"Upload for {section_name}",
                        type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                        key=f"up_{section_name}", label_visibility="collapsed")
                with uc2:
                    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
                    if sec_file and st.button("Extract", key=f"run_{section_name}", use_container_width=True):
                        with st.spinner(f"Extracting {section_name}..."):
                            ext = Path(sec_file.name).suffix.lower()
                            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                                tmp.write(sec_file.read()); tmp_path = tmp.name
                            try:
                                text = extract_text(tmp_path)
                                # Direct section-specific call - no classify step needed
                                result = call_llm(text, target_section=section_name)
                                if result:
                                    result["is_vessel_document"] = False
                                    result["is_resume"] = False
                                    save_to_doc_tank(result, sec_file.name, "section",
                                                     result.get("document_type", section_name),
                                                     False, False, section_name=section_name)
                                    st.session_state["popup"] = {
                                        "key": f"sec_{section_name.replace(' ','_')}",
                                        "result": result, "filename": sec_file.name,
                                        "source": "section", "target_section": section_name,
                                        "tank_id": st.session_state["doc_tank"][-1]["tank_id"]
                                    }
                                    st.rerun()
                            except Exception as e:
                                st.error(f"Error: {e}")

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

            if rows and st.button(f"Clear all", key=f"clear_{section_name}"):
                st.session_state["sections"][section_name]=[]; st.rerun()

    # Placeholder tabs - no functionality yet
    with _crew_tabs[8]:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">Coming soon.</div>', unsafe_allow_html=True)
    with _crew_tabs[9]:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">Coming soon.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
#  VESSELS
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "Vessel":
    st.markdown("### Vessels - Certificate Registry")
    st.markdown(
        '<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">'
        '<span style="font-size:1.1rem;">✨</span>'
        '<span style="font-size:0.9rem;font-weight:700;color:#0f172a;">AI Certificate Import</span>'
        '<span style="font-size:0.75rem;color:#64748b;">'
        '&nbsp;Upload a vessel certificate - AI extracts the details for your review'
        '</span></div>',
        unsafe_allow_html=True
    )

    vc1,vc2=st.columns([4,1])
    with vc1:
        v_file=st.file_uploader("Upload vessel certificate",type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                key="up_vessel",label_visibility="collapsed")
    with vc2:
        st.markdown("<div style='height:4px'></div>",unsafe_allow_html=True)
        if v_file and st.button("Extract", key="run_vessel", use_container_width=True):
            with st.spinner("Extracting vessel certificate..."):
                ext = Path(v_file.name).suffix.lower()
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(v_file.read()); tmp_path = tmp.name
                try:
                    text = extract_text(tmp_path)
                    # Direct vessel call - no classify step
                    result = call_llm(text, is_vessel=True)
                    if result:
                        result["is_vessel_document"] = True
                        result["is_resume"] = False
                        save_to_doc_tank(result, v_file.name, "vessel",
                                         result.get("document_type","Vessel Certificate"),
                                         True, False)
                        st.session_state["popup"] = {
                            "key": "vessel_cert", "result": result,
                            "filename": v_file.name, "source": "vessel",
                            "target_section": "Vessel",
                            "tank_id": st.session_state["doc_tank"][-1]["tank_id"]
                        }
                        st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

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
    st.markdown("### Doc Tank - Upload Log")

    # ── MASTER UPLOAD ─────────────────────────────────────────
    mu1, mu2 = st.columns([4, 1])
    with mu1:
        st.markdown(
            '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">'
            '<span style="font-size:1.1rem;">✨</span>'
            '<span style="font-size:0.9rem;font-weight:700;color:#0f172a;">AI Smart Import</span>'
            '<span style="font-size:0.75rem;color:#64748b;">'
            '&nbsp;Upload any doc - CV, certificate, or merged PDF - AI classifies and fills automatically'
            '</span></div>',
            unsafe_allow_html=True
        )
        master_file = st.file_uploader("master", type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                       label_visibility="collapsed", key="master_uploader")
    with mu2:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if master_file and st.button("Run Pipeline", use_container_width=True, key="run_master"):
            with st.spinner("Classifying document..."):
                ext = Path(master_file.name).suffix.lower()
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(master_file.read()); tmp_path = tmp.name
                try:
                    text = extract_text(tmp_path)

                    # Call 1 - lightweight classify
                    clf = classify_document(text)
                    if not clf:
                        st.error("Classification failed."); st.stop()

                    is_vessel  = clf.get("is_vessel", False)
                    is_resume  = clf.get("is_resume", False)
                    target_sec = clf.get("target_section", "")
                    doc_type   = clf.get("doc_type", "")

                    # Call 2 - targeted extraction
                    with st.spinner(f"Extracting: {doc_type or target_sec}..."):
                        result = call_llm(text,
                                          target_section=target_sec,
                                          is_vessel=is_vessel,
                                          is_resume=is_resume)

                    if result:
                        # Stamp doc_type from classifier if LLM didn't fill it
                        if not result.get("document_type") and doc_type:
                            result["document_type"] = doc_type
                        result["is_vessel_document"] = is_vessel
                        result["is_resume"] = is_resume
                        save_to_doc_tank(result, master_file.name, "master",
                                         result.get("document_type", doc_type),
                                         is_vessel, is_resume)
                        st.session_state["popup"] = {
                            "key": "master", "result": result,
                            "filename": master_file.name, "source": "master",
                            "target_section": target_sec,
                            "tank_id": st.session_state["doc_tank"][-1]["tank_id"]
                        }
                        st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)

    tank = st.session_state.get("doc_tank", [])

    if "dt_view" not in st.session_state:
        st.session_state["dt_view"] = "crew"
    _dtb1, _dtb2, _ = st.columns([2, 2, 8])
    with _dtb1:
        if st.button("Crew", key="dt_btn_crew", use_container_width=True,
                     type="primary" if st.session_state["dt_view"] == "crew" else "secondary"):
            st.session_state["dt_view"] = "crew"; st.rerun()
    with _dtb2:
        if st.button("Vessel", key="dt_btn_vessel", use_container_width=True,
                     type="primary" if st.session_state["dt_view"] == "vessel" else "secondary"):
            st.session_state["dt_view"] = "vessel"; st.rerun()
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    _show_vessel = (st.session_state["dt_view"] == "vessel")

    if not tank:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No documents processed yet.</div>', unsafe_allow_html=True)
    else:
        s1,s2,s3,s4,s5,s6 = st.columns(6)
        s1.metric("Total",    len(tank))
        s2.metric("Pending",  len([d for d in tank if d.get("status")=="Pending"]))
        s3.metric("Draft",    len([d for d in tank if d.get("status")=="Draft"]))
        s4.metric("Approved", len([d for d in tank if d.get("status") in ("Approved","Partial")]))
        s5.metric("Vessel",   len([d for d in tank if d.get("is_vessel")]))
        s6.metric("CV/Resume",len([d for d in tank if d.get("is_resume")]))
        st.markdown("---")
        _filtered = [
            (len(tank)-1-i, doc)
            for i, doc in enumerate(reversed(tank))
            if bool(doc.get("is_vessel")) == _show_vessel
        ]
        if not _filtered:
            st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No documents in this category.</div>', unsafe_allow_html=True)
        else:
            for idx, doc in _filtered:
                status     = doc.get("status", "Pending")
                if status == "Approved":
                    status_cls = "status-matched"
                elif status in ("Partial", "Draft"):
                    status_cls = "status-verify"
                else:
                    status_cls = "status-manual"
                approved_fields = doc.get("approved_fields", [])
                col_label, col_status, col_open, col_del = st.columns([4, 1, 1, 1])
                with col_label:
                    _src      = doc.get("source", "")
                    _sec      = doc.get("section_name", "")
                    if _src == "vessel":
                        _src_label = "Vessel Registry"
                        _src_color = "#0ea5e9"
                    elif _src == "section" and _sec:
                        _src_label = f"Crew / {_sec}"
                        _src_color = "#8b5cf6"
                    elif _src == "section":
                        _src_label = "Crew / Section Upload"
                        _src_color = "#8b5cf6"
                    elif _src == "master" and doc.get("is_resume"):
                        _src_label = "Crew / Resume Upload"
                        _src_color = "#8b5cf6"
                    else:
                        _src_label = "Doc Tank / Smart Import"
                        _src_color = "#64748b"
                    st.markdown(
                        f'<div style="padding:8px 0;">'
                        f'<div style="font-size:0.85rem;font-weight:600;color:#1e293b;">{doc["filename"]}</div>'
                        f'<div style="font-size:0.72rem;color:#64748b;">{doc["timestamp"]} &nbsp;|&nbsp; {doc["document_type"]}</div>'
                        f'<div style="font-size:0.68rem;color:{_src_color};margin-top:2px;">&#x2022; {_src_label}</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                with col_status:
                    st.markdown(
                        f'<div style="padding-top:12px;">'
                        f'<span class="{status_cls}">{status}</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                with col_open:
                    if st.button("Open", key=f"dt_open_{idx}", use_container_width=True):
                        raw = doc.get("raw_result", {})
                        st.session_state["popup"] = {
                            "key":            f"dt_{idx}",
                            "result":         copy.deepcopy(raw),
                            "filename":       doc["filename"],
                            "source":         "doc_tank",
                            "target_section": doc.get("target_section", ""),
                            "tank_id":        doc.get("tank_id", ""),
                            "tank_idx":       idx,
                            "tank_status":    status,
                            "approved_fields": approved_fields,
                            "draft_state":    doc.get("draft_state"),
                        }
                        st.session_state["popup_confirm_approve"] = False
                        st.rerun()
                with col_del:
                    if st.button("Delete", key=f"dt_del_{idx}", use_container_width=True):
                        st.session_state["doc_tank"].pop(idx)
                        st.rerun()
                if approved_fields:
                    with st.expander(f"Approved fields ({len(approved_fields)})", expanded=False):
                        af_cols = st.columns(3)
                        for afi, af in enumerate(approved_fields):
                            with af_cols[afi % 3]:
                                st.markdown(f'<span class="status-matched" style="margin:2px;display:inline-block;">{af}</span>', unsafe_allow_html=True)
                st.markdown("<hr style='margin:4px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)
