# PDF → Word (Kurdish-Optimized OCR)

Convert PDFs (including scans) to editable Word (.docx) with strong support for Kurdish:
- Sorani (Arabic script) **ckb**
- Kurmanji (Latin script) **kmr**
Also works with **ara** (Arabic) and **eng**.

## Features
- Detects text layer → **direct PDF→DOCX** (best layout) if `pdf2docx` is installed
- Otherwise → **OCR pipeline**:
  - Tries **OCRmyPDF** first (if installed): deskew, clean, remove background, optimize
  - If that’s not available or fails → **fallback OCR**:
    - Preprocess pages (grayscale, denoise, contrast/sharpness)
    - Reconstruct lines from Tesseract TSV
    - Mark paragraphs as **RTL** for Sorani (ckb)

## Run locally
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
