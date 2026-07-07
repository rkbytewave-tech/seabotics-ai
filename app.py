import streamlit as st
import json, re, os, tempfile, copy
from difflib import SequenceMatcher
from pathlib import Path
from datetime import datetime
import pandas as pd

# ── MODULE IMPORTS ────────────────────────────────────────────
from data import (
    PERSONAL_FIELDS, SECTION_COLUMNS, VESSEL_COLUMNS, DROPDOWN_FIELDS,
    ALL_ORG_OPTIONS, ORG_PREMAP,
)
from mapping import (
    fuzzy_match_dropdown, validate_field_value, check_dropdown_validity,
    get_dropdown_options, field_coherence,
    personal_to_profile, result_to_rows, vessel_result_to_row,
)
from pipeline import extract_text, classify_document, call_llm
from smart_upload_ui import render_ai_smart_upload_tab

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

/* ── Doc Tank reskin (matches ui_mockups/02-04) ─────────────── */
.dt-banner {
  background: linear-gradient(90deg,#4338ca 0%,#4f46e5 60%,#6366f1 100%);
  border-radius: 10px; padding: 18px 24px; margin-bottom: 14px;
  display: flex; align-items: center; gap: 14px;
}
.dt-banner-icon {
  width: 42px; height: 42px; border-radius: 10px; flex-shrink: 0;
  background: rgba(255,255,255,0.18);
  display: flex; align-items: center; justify-content: center; font-size: 20px;
}
.dt-banner-title { color:#ffffff; font-size:1.15rem; font-weight:700; }
.dt-banner-sub   { color:#dbeafe; font-size:0.78rem; margin-top:2px; }

div[data-testid="stVerticalBlockBorderWrapper"]:has(.dt-upload-marker) {
  border: 1.5px dashed #93c5fd !important; border-radius: 10px !important; background:#ffffff !important;
}
.dt-upload-title { font-size:0.9rem; font-weight:700; color:#1e3a8a; }
.dt-upload-desc  { font-size:0.72rem; color:#64748b; margin-top:2px; }

.st-key-dt_tab_crew button, .st-key-dt_tab_vessel button {
  border-radius: 8px !important; font-weight: 600 !important; font-size: 0.82rem !important;
}

.dt-stats-card { display:flex; border:1px solid #e2e8f0; border-radius:10px; overflow:hidden; background:#fff; }
.dt-stat { flex:1; padding:12px 18px; border-right:1px solid #eef2f7; }
.dt-stat:last-child { border-right:none; }
.dt-stat-label { font-size:0.62rem; font-weight:700; letter-spacing:0.05em; color:#94a3b8; text-transform:uppercase; margin-bottom:4px; }
.dt-stat-value { font-size:1.35rem; font-weight:700; color:#0f172a; }
.dt-stat.pending  { border-left:4px solid #f59e0b; }
.dt-stat.draft    { border-left:4px solid #d97706; }
.dt-stat.approved { border-left:4px solid #22c55e; }
.dt-stat.pending  .dt-stat-value { color:#d97706; }
.dt-stat.draft    .dt-stat-value { color:#b45309; }
.dt-stat.approved .dt-stat-value { color:#16a34a; }

.dt-type-pill { display:inline-block; font-size:0.65rem; font-weight:600; padding:2px 9px; border-radius:999px; }
.dt-type-cv      { background:#ede9fe; color:#6d28d9; }
.dt-type-license { background:#dbeafe; color:#1d4ed8; }
.dt-type-travel  { background:#cffafe; color:#0e7490; }
.dt-type-medical { background:#dcfce7; color:#15803d; }
.dt-type-vessel  { background:#ffedd5; color:#c2410c; }
.dt-type-other   { background:#f1f5f9; color:#475569; }

.dt-status-dot { display:inline-flex; align-items:center; gap:5px; font-size:0.75rem; font-weight:600; }
.dt-status-dot::before { content:""; width:7px; height:7px; border-radius:50%; }
.dt-status-approved { color:#16a34a; } .dt-status-approved::before { background:#22c55e; }
.dt-status-draft    { color:#b45309; } .dt-status-draft::before    { background:#f59e0b; }
.dt-status-pending  { color:#64748b; } .dt-status-pending::before  { background:#94a3b8; }

/* ── Personal Info reskin (matches ui_mockups/01) ───────────── */
.pi-banner {
  background: linear-gradient(90deg,#4338ca 0%,#6366f1 55%,#818cf8 100%);
  border-radius: 10px; padding: 20px 26px; margin-bottom: 16px;
  display: flex; align-items: center; justify-content: space-between;
}
.pi-banner-left { display:flex; align-items:center; gap:16px; }
.pi-avatar-circle {
  width:56px; height:56px; border-radius:50%; background:#c7d2fe; flex-shrink:0;
  display:flex; align-items:center; justify-content:center; font-size:1.2rem; font-weight:700; color:#4338ca;
}
.pi-name { color:#ffffff; font-size:1.15rem; font-weight:700; letter-spacing:0.3px; }
.pi-sub  { color:#dbeafe; font-size:0.78rem; margin-top:2px; }
.pi-badge { font-size:0.68rem; font-weight:700; padding:5px 12px; border-radius:6px; color:#ffffff; display:inline-block; }
.pi-badge-complete { background:#22c55e; }
.pi-badge-partial  { background:#f59e0b; }
.pi-badge-empty    { background:#94a3b8; }

.pi-overview-card { background:#fff; border:1px solid #e2e8f0; border-radius:10px; padding:16px 20px; margin-bottom:16px; }
.pi-overview-title { font-size:0.95rem; font-weight:700; color:#0f172a; }
.pi-overview-sub   { font-size:0.75rem; color:#64748b; margin-top:2px; margin-bottom:10px; }
.pi-pill { display:inline-block; font-size:0.68rem; font-weight:600; padding:4px 11px; border-radius:999px; margin-right:6px; }
.pi-pill-pending  { background:#fef3c7; color:#92400e; }
.pi-pill-progress { background:#dbeafe; color:#1d4ed8; }
.pi-pill-complete { background:#dcfce7; color:#166534; }

.pi-section-card    { background:#fff; border:1px solid #e2e8f0; border-radius:10px; margin-bottom:14px; overflow:hidden; }
.pi-section-heading { font-size:0.82rem; font-weight:700; color:#4338ca; padding:12px 18px; border-bottom:1px solid #eef2f7; }
.pi-field-row        { padding:12px 18px; border-bottom:1px solid #f8fafc; }
.pi-field-row:last-child { border-bottom:none; }
</style>
""", unsafe_allow_html=True)


# ── SESSION STATE ─────────────────────────────────────────────

def init_state():
    if "profile"  not in st.session_state: st.session_state["profile"]  = {}
    if "sections" not in st.session_state: st.session_state["sections"] = {s:[] for s in SECTION_COLUMNS}
    if "vessels"  not in st.session_state: st.session_state["vessels"]  = []
    if "doc_tank" not in st.session_state: st.session_state["doc_tank"] = []
    if "popup"    not in st.session_state: st.session_state["popup"]    = None
    if "subpopup" not in st.session_state: st.session_state["subpopup"] = None
    if "page"     not in st.session_state: st.session_state["page"]     = "Crew"
    if "dt_popup" not in st.session_state: st.session_state["dt_popup"] = None

init_state()


def save_to_doc_tank(result, filename, source, doc_type, is_vessel, is_resume, section_name=None):
    """Save any uploaded doc to doc tank immediately as draft."""
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


# ── UI HELPERS ────────────────────────────────────────────────

def conf_color(c):
    if c >= 0.90: return "#22c55e"
    if c >= 0.60: return "#f59e0b"
    if c > 0.00:  return "#ef4444"
    return "#e2e8f0"


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
        df = pd.DataFrame(rows)
        for c in cols_def:
            if c not in df.columns:
                df[c] = ""
        df = df[cols_def].fillna("")
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        for i in range(0, len(rows), 2):
            card_cols = st.columns(2)
            for j in range(2):
                idx = i + j
                if idx >= len(rows): break
                row = rows[idx]
                with card_cols[j]:
                    title_val = next((row.get(f,"") for f in ["Document Name","Certificate Name","Company","Institution Name","Bank Name","Vessel"] if row.get(f)), f"Record {idx+1}")
                    invalid_fields = [cn for cn,val in row.items() if val and cn in drops and not check_dropdown_validity(section_name,cn,val)]

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


# ── APPROVE AND POPULATE ──────────────────────────────────────

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


# ── SUB-POPUP (edit individual record) ───────────────────────

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
                    for fk, fl in PERSONAL_FIELDS:
                        if fl in edited_row:
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
        for _wk in list(st.session_state.keys()):
            if _wk.startswith(("pp_popup_", "vc_popup_", "v_sp_")):
                del st.session_state[_wk]
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

        # ── VESSEL FLOW ───────────────────────────────────────
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

        # ── CREW FLOW ─────────────────────────────────────────
        if not is_vessel:
            active_sec = st.session_state.get("popup_active_section")

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

            elif active_sec == "Personal Information":
                if st.button("< Back", key="popup_back_personal"):
                    st.session_state["popup_active_section"] = None
                    st.rerun()
                st.markdown('<div class="section-label">Personal Information</div>', unsafe_allow_html=True)
                if popup.get("source") == "section":
                    st.info("Personal info found in this document - review below before approving.")
                if _dt_approved:
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
                    ph1, ph2, ph3, ph4, ph5 = st.columns([2, 2.5, 1.2, 1.2, 2.5])
                    ph1.markdown("**Field**")
                    ph2.markdown("**Value**")
                    ph3.markdown("**Confidence**")
                    ph4.markdown("**Status**")
                    ph5.markdown("**System Fields**")
                    st.markdown("<hr style='margin:4px 0 8px 0;border-color:#e2e8f0;'>", unsafe_allow_html=True)
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

                        if result.get("is_resume", False) and not is_vessel:
                            st.session_state["profile"] = {}
                            for _wipe_sec in SECTION_COLUMNS:
                                st.session_state["sections"][_wipe_sec] = []

                        approve_and_populate(result, is_vessel, pending_rows, pending_vessel, personal,
                                             selected_sections=_selected_set if _selected_set else None)

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

    _crew_tabs = st.tabs([
        "Personal Info", "Experience", "Licenses and Certifications",
        "Medical", "Education", "Courses", "Bank Details",
        "Travel Documents", "Doc Check-List", "Travel Details", "AI Smart Upload",
    ])

    with _crew_tabs[0]:
        profile = st.session_state["profile"]
        ORG_MAP = {"person_name":"Full Name","surname":"Surname","email":"Email Address","gender":"Gender","contact_number":"Phone Number","date_of_birth":"Date of Birth","nationality":"Nationality","blood_group":"Blood Group","rank":"Rank / Designation","identification_mark":"Identification Mark","availability_date":"Availability Date","address":"Residential Address","country":"Country","state":"State","zip_code":"ZIP / PIN Code","domestic_airport":"Domestic Airport","international_airport":"International Airport","passport_number":"Passport Number","cdc_number":"CDC / Seaman's Book Number","height":"Height","weight":"Weight","boiler_suit_size":"Boiler Suit Size","safety_shoe_size":"Safety Shoe Size","shirt_size":"Shirt Size","trouser_size":"Trouser Size"}

        _pi_name     = profile.get("Full Name", "") or "Unnamed Crew"
        _pi_cdc      = profile.get("CDC / Seaman's Book Number", "")
        _pi_initials = "".join(w[0] for w in _pi_name.split()[:2]).upper() if profile.get("Full Name") else "?"
        _pi_total    = len(PERSONAL_FIELDS)
        _pi_filled   = sum(1 for fk, fl in PERSONAL_FIELDS if profile.get(ORG_MAP.get(fk, fl), ""))

        if _pi_filled == 0:
            _pi_badge_cls, _pi_badge_txt = "pi-badge-empty", "Not Started"
        elif _pi_filled < _pi_total:
            _pi_badge_cls, _pi_badge_txt = "pi-badge-partial", "In Progress"
        else:
            _pi_badge_cls, _pi_badge_txt = "pi-badge-complete", "Complete"

        st.markdown(
            f'<div class="pi-banner">'
            f'<div class="pi-banner-left">'
            f'<div class="pi-avatar-circle">{_pi_initials}</div>'
            f'<div><div class="pi-name">{_pi_name.upper()}</div>'
            f'<div class="pi-sub">{_pi_cdc if _pi_cdc else "No CDC on file"}</div></div>'
            f'</div>'
            f'<div class="pi-badge {_pi_badge_cls}">{_pi_badge_txt}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

        _pi_pill_cls = "pi-pill-pending" if _pi_filled == 0 else ("pi-pill-complete" if _pi_filled == _pi_total else "pi-pill-progress")
        _pi_pill_txt = "No fields captured yet" if _pi_filled == 0 else f"{_pi_filled} of {_pi_total} fields captured"
        st.markdown(
            '<div class="pi-overview-card">'
            '<div class="pi-overview-title">Profile Overview</div>'
            '<div class="pi-overview-sub">Personal and professional information extracted from uploaded documents</div>'
            f'<span class="pi-pill {_pi_pill_cls}">{_pi_pill_txt}</span>'
            '</div>',
            unsafe_allow_html=True
        )

        if "pi_show_upload" not in st.session_state:
            st.session_state["pi_show_upload"] = not bool(profile)
        if st.button("Edit Profile", key="pi_edit_toggle", use_container_width=False):
            st.session_state["pi_show_upload"] = not st.session_state["pi_show_upload"]
            st.rerun()

        if st.session_state["pi_show_upload"]:
            st.markdown(
                '<div style="display:flex;align-items:center;gap:8px;margin:10px 0 8px 0;">'
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
                            result = call_llm(text, is_resume=True, path=tmp_path)
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

        _PI_GROUPS = [
            ("Personal Details", ["person_name","surname","rank","date_of_birth","gender","blood_group","identification_mark","marital_status","availability_date"]),
            ("Contact Details", ["email","contact_number","nationality","address","country","state","zip_code"]),
            ("Identity & Travel Documents", ["passport_number","cdc_number","domestic_airport","international_airport","indos_number","pan_number","aadhar_number"]),
            ("Physical & Sizing", ["height","weight","boiler_suit_size","safety_shoe_size","shirt_size","trouser_size"]),
        ]
        _pi_by_key = {fk: fl for fk, fl in PERSONAL_FIELDS}
        _pi_grouped_keys = set()
        for _grp_name, _grp_keys in _PI_GROUPS:
            _pi_grouped_keys.update(_grp_keys)
        _pi_other = [fk for fk, fl in PERSONAL_FIELDS if fk not in _pi_grouped_keys]
        if _pi_other:
            _PI_GROUPS.append(("Other Details", _pi_other))

        for _grp_name, _grp_keys in _PI_GROUPS:
            _grp_fields = [(fk, _pi_by_key[fk]) for fk in _grp_keys if fk in _pi_by_key]
            if not _grp_fields:
                continue
            with st.container(border=True):
                st.markdown(f'<div class="pi-section-heading" style="padding:0 0 8px 0;border-bottom:1px solid #eef2f7;margin-bottom:10px;">{_grp_name}</div>', unsafe_allow_html=True)
                for i in range(0, len(_grp_fields), 3):
                    row_f = _grp_fields[i:i+3]
                    cols = st.columns(3)
                    for j, (fk, fl) in enumerate(row_f):
                        with cols[j]:
                            org_key = ORG_MAP.get(fk, fl); val = profile.get(org_key, "")
                            st.markdown(f'<div class="field-label">{fl}</div>', unsafe_allow_html=True)
                            if val: st.markdown(f'<div class="field-value">{val}</div>', unsafe_allow_html=True)
                            else:   st.markdown(f'<div class="field-empty">-</div>', unsafe_allow_html=True)

        if profile:
            if st.button("Clear Personal Info", key="clear_personal"):
                st.session_state["profile"] = {}; st.rerun()

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
        rows = st.session_state["sections"].get(section_name, [])
        with _tab_sec:
            st.markdown(
                f'<div style="font-size:0.8rem;color:#64748b;margin-bottom:8px;">'
                f'{len(rows)} record{"s" if len(rows)!=1 else ""}</div>',
                unsafe_allow_html=True
            )
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
                                result = call_llm(text, target_section=section_name, path=tmp_path)
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

            st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)

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
                st.session_state["sections"][section_name] = []; st.rerun()

    with _crew_tabs[8]:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">Coming soon.</div>', unsafe_allow_html=True)
    with _crew_tabs[9]:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">Coming soon.</div>', unsafe_allow_html=True)
    with _crew_tabs[10]:
        render_ai_smart_upload_tab(save_to_doc_tank)

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

    vc1, vc2 = st.columns([4, 1])
    with vc1:
        v_file = st.file_uploader("Upload vessel certificate", type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                  key="up_vessel", label_visibility="collapsed")
    with vc2:
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
        if v_file and st.button("Extract", key="run_vessel", use_container_width=True):
            with st.spinner("Extracting vessel certificate..."):
                ext = Path(v_file.name).suffix.lower()
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(v_file.read()); tmp_path = tmp.name
                try:
                    text = extract_text(tmp_path)
                    result = call_llm(text, is_vessel=True)
                    if result:
                        result["is_vessel_document"] = True
                        result["is_resume"] = False
                        save_to_doc_tank(result, v_file.name, "vessel",
                                         result.get("document_type", "Vessel Certificate"),
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

    st.markdown("<hr style='margin:8px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)

    vessels = st.session_state.get("vessels", [])
    if not vessels:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No vessel certificates yet.</div>', unsafe_allow_html=True)
    else:
        view_key = "view_vessels"
        if view_key not in st.session_state: st.session_state[view_key] = "grid"
        vt1, vt2, _ = st.columns([1, 1, 6])
        with vt1:
            if st.button("Grid", key="btn_grid_vessels",
                         type="primary" if st.session_state[view_key]=="grid" else "secondary"):
                st.session_state[view_key] = "grid"; st.rerun()
        with vt2:
            if st.button("Table", key="btn_table_vessels",
                         type="primary" if st.session_state[view_key]=="table" else "secondary"):
                st.session_state[view_key] = "table"; st.rerun()

        render_section_content("Vessels", vessels, key_prefix="vessel_grid", view_mode=st.session_state[view_key])

# ════════════════════════════════════════════════════════════
#  DOC TANK
# ════════════════════════════════════════════════════════════
elif st.session_state["page"] == "Doc Tank":
    st.markdown(
        '<div class="dt-banner">'
        '<div class="dt-banner-icon">&#128196;</div>'
        '<div><div class="dt-banner-title">Doc Tank</div>'
        '<div class="dt-banner-sub">Global document log &nbsp;|&nbsp; All crew and vessel uploads</div>'
        '</div></div>',
        unsafe_allow_html=True
    )

    _dt_upload_box = st.container(border=True, key="dt_upload_card")
    with _dt_upload_box:
        st.markdown('<span class="dt-upload-marker"></span>', unsafe_allow_html=True)
        st.markdown(
            '<div class="dt-upload-title">Master Upload</div>'
            '<div class="dt-upload-desc">One drop point for every document &mdash; crew or vessel. '
            'The pipeline identifies, classifies, and routes each file automatically.</div>',
            unsafe_allow_html=True
        )
        mu1, mu2 = st.columns([4, 1])
        with mu1:
            master_file = st.file_uploader("master", type=["pdf","png","jpg","jpeg","tiff","docx","txt"],
                                           label_visibility="collapsed", key="master_uploader")
        with mu2:
            st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
            if master_file and st.button("Upload & Extract", use_container_width=True, type="primary", key="run_master"):
                with st.spinner("Classifying document..."):
                    ext = Path(master_file.name).suffix.lower()
                    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                        tmp.write(master_file.read()); tmp_path = tmp.name
                    try:
                        text = extract_text(tmp_path)

                        clf = classify_document(text)
                        if not clf:
                            st.error("Classification failed."); st.stop()

                        is_vessel  = clf.get("is_vessel", False)
                        is_resume  = clf.get("is_resume", False)
                        target_sec = clf.get("target_section", "")
                        doc_type   = clf.get("doc_type", "")

                        with st.spinner(f"Extracting: {doc_type or target_sec}..."):
                            result = call_llm(text,
                                              target_section=target_sec,
                                              is_vessel=is_vessel,
                                              is_resume=is_resume,
                                              path=tmp_path)

                        if result:
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
    _crew_n   = len([d for d in tank if not d.get("is_vessel")])
    _vessel_n = len([d for d in tank if d.get("is_vessel")])
    _dtb1, _dtb2, _ = st.columns([2, 2, 8])
    with _dtb1:
        with st.container(key="dt_tab_crew"):
            if st.button(f"Crew  {_crew_n}", key="dt_btn_crew", use_container_width=True,
                         type="primary" if st.session_state["dt_view"] == "crew" else "secondary"):
                st.session_state["dt_view"] = "crew"; st.rerun()
    with _dtb2:
        with st.container(key="dt_tab_vessel"):
            if st.button(f"Vessel  {_vessel_n}", key="dt_btn_vessel", use_container_width=True,
                         type="primary" if st.session_state["dt_view"] == "vessel" else "secondary"):
                st.session_state["dt_view"] = "vessel"; st.rerun()
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    _show_vessel = (st.session_state["dt_view"] == "vessel")

    if not tank:
        st.markdown('<div class="field-empty" style="text-align:center;padding:16px;">No documents processed yet.</div>', unsafe_allow_html=True)
    else:
        _n_total    = len(tank)
        _n_pending  = len([d for d in tank if d.get("status")=="Pending"])
        _n_draft    = len([d for d in tank if d.get("status")=="Draft"])
        _n_approved = len([d for d in tank if d.get("status") in ("Approved","Partial")])
        st.markdown(
            '<div class="dt-stats-card">'
            f'<div class="dt-stat"><div class="dt-stat-label">Total</div><div class="dt-stat-value">{_n_total}</div></div>'
            f'<div class="dt-stat pending"><div class="dt-stat-label">Pending</div><div class="dt-stat-value">{_n_pending}</div></div>'
            f'<div class="dt-stat draft"><div class="dt-stat-label">Draft</div><div class="dt-stat-value">{_n_draft}</div></div>'
            f'<div class="dt-stat approved"><div class="dt-stat-label">Approved</div><div class="dt-stat-value">{_n_approved}</div></div>'
            '</div>',
            unsafe_allow_html=True
        )
        st.caption(f"{_vessel_n} vessel cert(s)  ·  "
                   f"{len([d for d in tank if d.get('is_resume')])} CV/Resume upload(s)")
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
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
                approved_fields = doc.get("approved_fields", [])

                if doc.get("is_vessel"):
                    _type_cls, _type_label = "dt-type-vessel", "Vessel Cert"
                elif doc.get("is_resume"):
                    _type_cls, _type_label = "dt-type-cv", "CV/Resume"
                else:
                    _sec_l = (doc.get("section_name","") or doc.get("document_type","")).lower()
                    if "medical" in _sec_l:
                        _type_cls, _type_label = "dt-type-medical", "Medical"
                    elif "travel" in _sec_l or "passport" in _sec_l or "cdc" in _sec_l:
                        _type_cls, _type_label = "dt-type-travel", "Travel Doc"
                    elif "licen" in _sec_l or "certif" in _sec_l:
                        _type_cls, _type_label = "dt-type-license", "License"
                    else:
                        _type_cls, _type_label = "dt-type-other", (doc.get("section_name") or "Document")

                if status == "Approved" or status == "Partial":
                    _dot_cls = "dt-status-approved"
                elif status == "Draft":
                    _dot_cls = "dt-status-draft"
                else:
                    _dot_cls = "dt-status-pending"

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
                    elif _src == "smart_upload" and doc.get("is_resume"):
                        _src_label = "Crew / AI Smart Upload (Resume)"
                        _src_color = "#8b5cf6"
                    elif _src == "smart_upload":
                        _src_label = "Crew / AI Smart Upload"
                        _src_color = "#8b5cf6"
                    else:
                        _src_label = "Doc Tank / Smart Import"
                        _src_color = "#64748b"
                    st.markdown(
                        f'<div style="padding:8px 0;">'
                        f'<div style="font-size:0.85rem;font-weight:600;color:#1e293b;">{doc["filename"]}'
                        f'&nbsp;&nbsp;<span class="dt-type-pill {_type_cls}">{_type_label}</span></div>'
                        f'<div style="font-size:0.72rem;color:#64748b;margin-top:2px;">{doc["timestamp"]} &nbsp;|&nbsp; {doc["document_type"]}</div>'
                        f'<div style="font-size:0.68rem;color:{_src_color};margin-top:2px;">&#x2022; {_src_label}</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                with col_status:
                    st.markdown(
                        f'<div style="padding-top:14px;">'
                        f'<span class="dt-status-dot {_dot_cls}">{status}</span>'
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
