import re

from data import (
    ALL_COURSE_NAMES, ALL_LICENSE_NAMES,
    COURSES_DROPDOWN, LICENSES_CERT_DROPDOWN,
    ALL_VESSEL_NAMES, CERT_NAME_TO_CATEGORY,
    SECTION_COLUMNS, VESSEL_COLUMNS, DROPDOWN_FIELDS,
)

# ── DATE / GENDER VALIDATION CONSTANTS ───────────────────────

DATE_PATTERN = re.compile(
    r"^\d{1,2}-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4}$",
    re.IGNORECASE
)

VALID_GENDERS = {"male", "female", "m", "f", "other"}

# ── FUZZY MATCHING ────────────────────────────────────────────

def fuzzy_match_dropdown(value, options):
    """Find closest matching dropdown option using simple fuzzy logic."""
    if not value or not options:
        return None
    val_upper = value.strip().upper()
    for opt in options:
        if opt.upper() == val_upper:
            return opt
    for opt in options:
        if val_upper in opt.upper() or opt.upper() in val_upper:
            return opt
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


def _match_score(original, matched):
    """Score how well a fuzzy match landed. Higher = better."""
    if not original or not matched:
        return 0
    orig_upper    = original.strip().upper()
    matched_upper = matched.strip().upper()
    if orig_upper == matched_upper:
        return 3
    if orig_upper in matched_upper or matched_upper in orig_upper:
        return 2
    orig_words    = set(orig_upper.split())
    matched_words = set(matched_upper.split())
    overlap       = len(orig_words & matched_words)
    return overlap


# ── COURSE / LICENSE RECLASSIFICATION ────────────────────────

def reclassify_courses_and_licenses(rows_dict):
    """
    Bottom-up reclassification for Courses and Licenses and Certification.
    Fuzzy-matches extracted document names against reference lists.
    Moves misclassified rows to the correct section and sets the correct type.
    """
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

    pool = []
    for row in rows_dict.get("Courses", []):
        pool.append(("Courses", row))
    for row in rows_dict.get("Licenses and Certification", []):
        pool.append(("Licenses and Certification", row))

    for origin, row in pool:
        doc_name = row.get("Document Name", "")

        course_match  = fuzzy_match_dropdown(doc_name, ALL_COURSE_NAMES)
        license_match = fuzzy_match_dropdown(doc_name, ALL_LICENSE_NAMES)

        course_score  = _match_score(doc_name, course_match)
        license_score = _match_score(doc_name, license_match)

        if license_score >= course_score and license_match:
            new_row = dict(row)
            new_row["Document Name"] = license_match
            new_row["Type"] = LICENSE_NAME_TO_TYPE.get(license_match.upper(), row.get("Type", ""))
            if "Issuing Authority" not in new_row: new_row["Issuing Authority"] = ""
            if "Place of Issue"    not in new_row: new_row["Place of Issue"]    = ""
            final_licenses.append(new_row)

        elif course_match:
            new_row = dict(row)
            new_row["Document Name"] = course_match
            new_row["Type"] = COURSE_NAME_TO_TYPE.get(course_match.upper(), row.get("Type", ""))
            final_courses.append(new_row)

        else:
            if origin == "Courses":
                final_courses.append(row)
            else:
                final_licenses.append(row)

    rows_dict["Courses"] = final_courses
    rows_dict["Licenses and Certification"] = final_licenses
    return rows_dict


# ── FIELD VALIDATION ─────────────────────────────────────────

def validate_field_value(col_name, value):
    """
    Returns (confidence_penalty, flag_message) for illogical or malformed values.
    confidence_penalty is subtracted from base confidence (0.0 = no issue).
    """
    if not value or not value.strip():
        return 0.0, None

    v = value.strip()
    col_lower = col_name.lower()

    if "date" in col_lower:
        if v.upper() == "LIFETIME":
            return 0.0, None
        if not DATE_PATTERN.match(v):
            return 0.3, f"{col_name}: '{v}' is not in DD-MMM-YYYY format"

    if col_lower == "gender":
        if v.lower() not in VALID_GENDERS:
            return 0.3, f"Gender: '{v}' is not a recognised value (expected Male/Female)"

    if col_lower in {"grt", "dwt", "bhp", "main engine kw", "height", "weight"}:
        clean = v.replace(",", "").replace(".", "").replace(" ", "")
        if not clean.isdigit():
            return 0.2, f"{col_name}: '{v}' expected a numeric value"

    if "certificate number" in col_lower or "account number" in col_lower:
        if len(v) < 3:
            return 0.2, f"{col_name}: '{v}' seems too short to be valid"

    return 0.0, None


def check_dropdown_validity(section, field_name, value):
    """Returns True if value is valid for dropdown field, False if invalid."""
    drops = DROPDOWN_FIELDS.get(section, {})
    if field_name not in drops:
        return True
    valid_options = drops[field_name]
    return any(value.strip().upper() == opt.upper() for opt in valid_options)


def get_dropdown_options(section, field_name):
    drops = DROPDOWN_FIELDS.get(section, {})
    return drops.get(field_name, [])


# ── FIELD COHERENCE SCORE ─────────────────────────────────────

def field_coherence(field_label, value):
    """Return 0.0-1.0 coherence score for a free-text value given its field label."""
    if not value:
        return 0.0
    v    = value.strip()
    fl   = field_label.lower()
    _len = max(len(v), 1)
    if fl in ("full name", "surname"):
        valid = len(re.findall(r"[a-zA-Z \-\.\']", v))
        return round(valid / _len, 2)
    if fl in ("height", "weight"):
        return 1.0 if re.match(r"^\d+(\.\d+)?\s*(cm|kg|lbs?|m)?$", v, re.IGNORECASE) else 0.3
    if fl == "email":
        return 1.0 if re.match(r"^[\w.+\-]+@[\w\-]+\.[\w.]+$", v) else 0.2
    if fl == "contact number":
        valid = len(re.findall(r"[\d\+\-\(\) ]", v))
        return round(min(1.0, valid / _len), 2)
    if fl in ("passport number", "cdc / seaman's book number"):
        valid = len(re.findall(r"[a-zA-Z0-9\-\/]", v))
        return round(min(1.0, valid / _len), 2)
    if fl in ("nationality", "country", "state", "address"):
        valid = len(re.findall(r"[a-zA-Z0-9 \-\/\.,]", v))
        return round(min(1.0, valid / _len), 2)
    if "date" in fl:
        _, flag = validate_field_value(field_label, v)
        return 0.7 if flag else 1.0
    bad = len(re.findall(r"[$#@!%^&*=|<>~`]", v))
    if bad == 0:
        return 1.0
    return max(0.1, round(1.0 - (bad / _len) * 3, 2))


# ── FIELD MAPPING HELPERS ─────────────────────────────────────

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
    return {v: p[k] for k, v in m.items() if p.get(k)}


def result_to_rows(result):
    rows = {s: [] for s in SECTION_COLUMNS}
    for e in result.get("sea_service", []):
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
    for e in result.get("travel_documents", []):
        if any(v for v in e.values()):
            rows["Travel Documents"].append({
                "Document Type":e.get("document_type",""),"Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Place of Issue":e.get("place_of_issue",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("medical", []):
        if any(v for v in e.values()):
            rows["Medical"].append({
                "Document Type":e.get("document_type",""),"Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("education", []):
        if any(v for v in e.values()):
            rows["Education"].append({
                "Institution Name":e.get("institution_name",""),"Country":e.get("country",""),
                "Level of Education":e.get("education_level",""),
                "Field of Study":e.get("field_of_study",""),
                "Start Date":e.get("start_date",""),"Completion Date":e.get("completion_date",""),
            })
    for e in result.get("courses", []):
        if any(v for v in e.values()):
            rows["Courses"].append({
                "Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Type":e.get("type",""),"Place of Issue":e.get("place_of_issue",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("licenses_certifications", []):
        if any(v for v in e.values()):
            rows["Licenses and Certification"].append({
                "Document Name":e.get("document_name",""),
                "Certificate Number":e.get("certificate_number",""),
                "Type":e.get("type",""),"Issuing Authority":e.get("issuing_authority",""),
                "Place of Issue":e.get("place_of_issue",""),
                "Date of Issue":e.get("date_of_issue",""),"Date of Expiry":e.get("date_of_expiry",""),
            })
    for e in result.get("bank_details", []):
        if any(v for v in e.values()):
            rows["Bank Details"].append({
                "Bank Name":e.get("bank_name",""),"Branch Name":e.get("branch_name",""),
                "Account Type":e.get("account_type",""),"Account Name":e.get("account_name",""),
                "Account Number":e.get("account_number",""),"IFSC / SWIFT Code":e.get("ifsc_swift",""),
            })
    rows = reclassify_courses_and_licenses(rows)
    return rows


# ── VESSEL RESULT MAPPING ─────────────────────────────────────

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
    if cert_category in ("Flag State", "Statutory"):
        return raw_text
    return raw_text


def vessel_result_to_row(result):
    vc = result.get("vessel_certificate", {})
    if not vc or not any(v for v in vc.values()): return None

    cert_name_raw   = vc.get("certificate_name", "")
    cert_name_fuzzy = fuzzy_match_dropdown(cert_name_raw, ALL_VESSEL_NAMES) or cert_name_raw
    cert_category   = infer_cert_category(cert_name_raw) or vc.get("certificate_category", "")
    cert_type       = infer_cert_type(vc.get("cert_type", ""))
    class_society   = vc.get("classification_society", "")
    issuing_auth    = infer_issuing_authority(vc.get("issuing_authority", ""), cert_category, class_society)

    return {
        "Certificate Name":       cert_name_fuzzy,
        "Certificate Category":   cert_category,
        "Vessel":                 vc.get("vessel_name", ""),
        "IMO Number":             vc.get("imo_number", ""),
        "Certificate Number":     vc.get("certificate_number", ""),
        "Type":                   cert_type,
        "Issuing Authority":      issuing_auth,
        "Classification Society": class_society,
        "Date of Issue":          vc.get("date_of_issue", ""),
        "Date of Expiry":         vc.get("date_of_expiry", ""),
        "Date of Survey":         vc.get("date_of_survey", ""),
    }
