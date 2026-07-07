"""
AI Smart Upload tab - new UI surface, additive only.

Per the OCR Extraction Pipeline walkthrough doc: a dedicated tab (bottom of the
crew profile) with two sub-tabs:
  - Resume: single-file upload, no classification step (call_llm(is_resume=True)).
  - Personal Documents: multi-file upload. Each file is classified and extracted
    independently (classify_document -> call_llm), fail-soft - one bad file does
    not stop the rest. Each successful file creates its own save_to_doc_tank() row.

This module does not touch app.py's existing uploaders (Personal Info tab resume
upload, per-section AI Autofill uploaders, Doc Tank AI Smart Import). Those are
confirmed-working per CLAUDE.md and are left exactly as they are. This tab is an
additional entry point, matching the target architecture in the walkthrough doc.

save_to_doc_tank is passed in from app.py (not imported) so this module has no
dependency on app.py and there is no circular import.
"""

import copy
import tempfile
from pathlib import Path

import streamlit as st

from pipeline import extract_text, classify_document, call_llm


def render_ai_smart_upload_tab(save_to_doc_tank_fn):
    """Render the AI Smart Upload tab: Resume sub-tab + Personal Documents sub-tab,
    followed by this crew member's upload log.

    save_to_doc_tank_fn: the save_to_doc_tank() function from app.py, passed in
    so every row lands in the same st.session_state["doc_tank"] list used by
    Doc Tank and the rest of the app.
    """
    st.markdown(
        '<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">'
        '<span style="font-size:1.1rem;">✨</span>'
        '<span style="font-size:0.9rem;font-weight:700;color:#0f172a;">AI Smart Upload</span>'
        '<span style="font-size:0.75rem;color:#64748b;">'
        '&nbsp;Dedicated AI-assisted document uploads for this crew member'
        '</span></div>',
        unsafe_allow_html=True
    )

    _su_tabs = st.tabs(["Resume", "Personal Documents"])

    with _su_tabs[0]:
        _render_resume_subtab(save_to_doc_tank_fn)

    with _su_tabs[1]:
        _render_personal_documents_subtab(save_to_doc_tank_fn)

    st.markdown("<hr style='margin:10px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)
    _render_smart_upload_log()


# ── RESUME SUB-TAB (single upload, no classification) ──────────

def _render_resume_subtab(save_to_doc_tank_fn):
    st.caption("Uploading resume. A new CV overwrites previously approved resume data on approval.")
    resume_file = st.file_uploader(
        "Upload CV or Resume", type=["pdf", "png", "jpg", "jpeg", "tiff", "docx", "txt"],
        key="su_up_resume", label_visibility="collapsed", accept_multiple_files=False
    )
    if resume_file and st.button("Upload & Extract", key="su_run_resume"):
        with st.spinner(f"Uploading {resume_file.name}..."):
            ext = Path(resume_file.name).suffix.lower()
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(resume_file.read()); tmp_path = tmp.name
            try:
                text = extract_text(tmp_path)
                result = call_llm(text, is_resume=True, path=tmp_path)
                if result:
                    result["is_vessel_document"] = False
                    result["is_resume"] = True
                    save_to_doc_tank_fn(result, resume_file.name, "smart_upload",
                                         result.get("document_type", "CV / Resume"),
                                         False, True)
                    st.success(f"{resume_file.name} extracted and added as Pending.")
                else:
                    st.error(f"{resume_file.name}: extraction failed.")
            except Exception as e:
                st.error(f"{resume_file.name}: {e}")


# ── PERSONAL DOCUMENTS SUB-TAB (batch upload, auto-classify) ───

def _render_personal_documents_subtab(save_to_doc_tank_fn):
    st.caption("Upload one or more documents for this crew member. Each is classified and "
               "routed automatically. If one document fails, the others are still processed.")
    files = st.file_uploader(
        "Upload documents", type=["pdf", "png", "jpg", "jpeg", "tiff", "docx", "txt"],
        key="su_up_personal", label_visibility="collapsed", accept_multiple_files=True
    )
    if files and st.button(f"Upload & Extract ({len(files)})", key="su_run_personal"):
        successes, failures = [], []
        progress = st.progress(0.0, text="Starting...")

        for i, f in enumerate(files):
            progress.progress(i / len(files), text=f"Processing {f.name} ({i+1}/{len(files)})...")
            try:
                ext = Path(f.name).suffix.lower()
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(f.read()); tmp_path = tmp.name

                text = extract_text(tmp_path)

                clf = classify_document(text)
                if not clf:
                    failures.append((f.name, "Classification failed"))
                    continue

                is_vessel  = clf.get("is_vessel", False)
                is_resume  = clf.get("is_resume", False)
                target_sec = clf.get("target_section", "")
                doc_type   = clf.get("doc_type", "")

                result = call_llm(text, target_section=target_sec, is_vessel=is_vessel,
                                   is_resume=is_resume, path=tmp_path)
                if not result:
                    failures.append((f.name, "Extraction failed"))
                    continue

                if not result.get("document_type") and doc_type:
                    result["document_type"] = doc_type
                result["is_vessel_document"] = is_vessel
                result["is_resume"] = is_resume

                # Each file gets its own, independent doc tank row - same as
                # if it had been uploaded on its own.
                save_to_doc_tank_fn(result, f.name, "smart_upload",
                                     result.get("document_type", doc_type),
                                     is_vessel, is_resume)
                successes.append(f.name)
            except Exception as e:
                # Fail-soft: log this file's failure and keep going with the rest.
                failures.append((f.name, str(e)))
                continue

        progress.progress(1.0, text="Done.")

        if successes:
            st.success(
                f"{len(successes)} of {len(files)} document(s) processed and added as Pending: "
                + ", ".join(successes)
            )
        if failures:
            st.warning(f"{len(failures)} document(s) failed - the rest were processed normally.")
            for name, reason in failures:
                st.markdown(f'<span class="invalid-flag">{name}: {reason}</span>', unsafe_allow_html=True)


# ── UPLOAD LOG (this crew member's documents) ───────────────────

def _render_smart_upload_log():
    """Crew-scoped upload log. The prototype has a single implicit crew profile
    in session state (no multi-crew routing yet - see api_mapping.py notes in
    CLAUDE.md), so 'this crew member's documents' = every non-vessel doc_tank
    entry. Opens the same render_popup() as Doc Tank, so review/approve behaves
    identically from either location, per the walkthrough doc."""
    st.markdown('<div class="section-label">This Crew Member - Upload Log</div>', unsafe_allow_html=True)
    tank = st.session_state.get("doc_tank", [])
    _crew_docs = [(len(tank) - 1 - i, d) for i, d in enumerate(reversed(tank)) if not d.get("is_vessel")]

    if not _crew_docs:
        st.markdown(
            '<div class="field-empty" style="text-align:center;padding:12px;">'
            'No documents uploaded yet for this crew member.</div>',
            unsafe_allow_html=True
        )
        return

    for idx, doc in _crew_docs:
        status = doc.get("status", "Pending")
        if status == "Approved":
            status_cls = "status-matched"
        elif status in ("Partial", "Draft"):
            status_cls = "status-verify"
        else:
            status_cls = "status-manual"

        c1, c2, c3 = st.columns([5, 1, 1])
        with c1:
            st.markdown(
                f'<div style="padding:6px 0;">'
                f'<div style="font-size:0.85rem;font-weight:600;color:#1e293b;">{doc["filename"]}</div>'
                f'<div style="font-size:0.72rem;color:#64748b;">{doc["timestamp"]} &nbsp;|&nbsp; {doc["document_type"]}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
        with c2:
            st.markdown(
                f'<div style="padding-top:10px;"><span class="{status_cls}">{status}</span></div>',
                unsafe_allow_html=True
            )
        with c3:
            if st.button("Open", key=f"su_open_{idx}", use_container_width=True):
                raw = doc.get("raw_result", {})
                st.session_state["popup"] = {
                    "key":             f"su_{idx}",
                    "result":          copy.deepcopy(raw),
                    "filename":        doc["filename"],
                    "source":          "doc_tank",
                    "target_section":  doc.get("target_section", ""),
                    "tank_id":         doc.get("tank_id", ""),
                    "tank_idx":        idx,
                    "tank_status":     status,
                    "approved_fields": doc.get("approved_fields", []),
                    "draft_state":     doc.get("draft_state"),
                }
                st.session_state["popup_confirm_approve"] = False
                st.rerun()
        st.markdown("<hr style='margin:4px 0;border-color:#f1f5f9;'>", unsafe_allow_html=True)
