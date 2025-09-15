#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import tempfile
import subprocess

import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image, ImageFilter, ImageOps, ImageEnhance
import pytesseract

# optional imports (graceful fallback)
try:
    from pdf2docx import Converter
    HAS_PDF2DOCX = True
except Exception:
    HAS_PDF2DOCX = False

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------- UI setup ----------------
st.set_page_config(page_title="PDF → Word (Kurdish-Optimized OCR)", page_icon="📄", layout="centered")
st.title("📄 PDF → Word Converter — Kurdish Optimized (Sorani/Kurmanji)")
st.caption("Convert any PDF (including scans) to editable Word. Best support for Sorani **ckb** and Kurmanji **kmr**.")

with st.expander("ℹ️ Tips for best Kurdish OCR"):
    st.markdown(
        "- Use Tesseract language codes: **ckb** (Sorani, Arabic script), **kmr** (Kurmanji, Latin script).  \n"
        "- For mixed content add **ara** and **eng**, e.g. `ckb+ara+eng`.  \n"
        "- If OCRmyPDF isn’t installed, the app falls back to a Kurdish-aware OCR pipeline."
    )

# ---------------- Sidebar ----------------
st.sidebar.header("OCR Settings")
lang_options = [
    "ckb", "kmr", "ara", "eng",
    "ck
