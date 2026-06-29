import json, re
from pathlib import Path

import fitz
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from docx import Document as DocxDocument
import openpyxl
from google import genai
from google.genai import types
import streamlit as st

from config import GEMINI_API_KEY, MODEL_NAME, IMAGE_EXT
from prompts import CLASSIFIER_PROMPT, VESSEL_PROMPT, RESUME_PROMPT, SECTION_PROMPTS


def preprocess(img):
    if img.mode != "RGB": img = img.convert("RGB")
    w, h = img.size
    if w < 1000: img = img.resize((1000, int(h * 1000 / w)), Image.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    img = ImageEnhance.Contrast(img).enhance(1.5)
    return img.convert("L")


def extract_text(path):
    ext = Path(path).suffix.lower()
    if ext in IMAGE_EXT:
        img = preprocess(Image.open(path))
        return max(
            [pytesseract.image_to_string(img, config=c) for c in ["--psm 3", "--psm 6", "--psm 11"]],
            key=len
        )
    if ext == ".pdf":
        doc = fitz.open(path)
        pages = []
        for p in doc:
            t = p.get_text("text").strip()
            if len(t) > 50:
                pages.append(t)
            else:
                pix = p.get_pixmap(dpi=250)
                img = preprocess(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
                pages.append(pytesseract.image_to_string(img))
        doc.close()
        return "\n--- PAGE BREAK ---\n".join(pages)
    if ext == ".docx":
        return "\n".join(p.text for p in DocxDocument(path).paragraphs if p.text.strip())
    if ext == ".txt":
        with open(path) as f:
            return f.read()
    if ext in [".xlsx", ".xls"]:
        wb = openpyxl.load_workbook(path, data_only=True)
        return "\n".join(
            "  |  ".join(str(c) for c in r if c)
            for s in wb.worksheets
            for r in s.iter_rows(values_only=True)
        )
    raise ValueError(f"Unsupported: {ext}")


def _llm_call(system_prompt, user_text, max_tokens=4000, text_limit=30000):
    """Core LLM caller - shared by all pipelines."""
    if len(user_text) > text_limit:
        user_text = user_text[:text_limit]
    if not user_text.strip():
        return None

    client = genai.Client(api_key=GEMINI_API_KEY)

    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=f"Extract all information:\n\n{user_text}",
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.1,
                    max_output_tokens=max_tokens,
                )
            )
            raw = response.text.strip()
            if not raw: continue
            raw = re.sub(r"^```[a-z]*\n?", "", raw)
            raw = re.sub(r"\n?```$", "", raw)
            start = raw.find("{"); end = raw.rfind("}") + 1
            if start == -1 or end == 0: continue
            return json.loads(raw[start:end])
        except json.JSONDecodeError:
            continue
        except Exception as e:
            st.error(f"API error: {e}")
            return None
    st.error("Pipeline failed after 3 attempts.")
    return None


def classify_document(text):
    """Step 1 for Master Upload - lightweight classifier, ~200 tokens."""
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
        return _llm_call(VESSEL_PROMPT, text, max_tokens=2000)
    if is_resume:
        # Gemini 2.5 Flash: no TPM constraint - feed full text
        return _llm_call(RESUME_PROMPT, text, max_tokens=8000, text_limit=50000)
    if target_section and target_section in SECTION_PROMPTS:
        return _llm_call(SECTION_PROMPTS[target_section], text, max_tokens=4000)
    return _llm_call(RESUME_PROMPT, text, max_tokens=8000)