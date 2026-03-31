import re
from pathlib import Path

import fitz  # PyMuPDF
import easyocr
import io

import os
import numpy as np
from PIL import Image
import certifi


SAMPLE_SUFFIXES = ("SAMPLE", "SAMPL", "SMPL", "SMP", "SM", "SMPL", "S")


def is_model_like(s: str) -> bool:
    x = re.sub(r"[^A-Z0-9]", "", s.upper())
    if not (len(x) >= 6 and any(ch.isdigit() for ch in x) and any(ch.isalpha() for ch in x)):
        return False
    # Exclude order IDs-like prefixes.
    if x.startswith(("WN", "WK", "WH")):
        return False
    return True


def is_sample_code(code: str) -> bool:
    c = code.upper()
    # OCR can sometimes read 'A' instead of 'S' at the end; treat "...A" like "...S"
    if c.endswith("A") and not any(c.endswith(suf) for suf in SAMPLE_SUFFIXES):
        if len(c) >= 2 and c[-2] != "M":
            c = c[:-1] + "S"
    return any(c.endswith(suf) for suf in SAMPLE_SUFFIXES)


def ocr_page_model_column(pdf_path: Path, page_number: int, scale: float = 2.0):
    doc = fitz.open(str(pdf_path))
    page = doc[page_number - 1]

    # Try to find table column bounds via text search.
    # If search fails, we fallback to whole page OCR.
    rect = None
    try:
        model_boxes = page.search_for("Model Number")
        internet_boxes = page.search_for("Internet Number")
        qty_shipped_boxes = page.search_for("Qty Shipped")
        qty_returned_boxes = page.search_for("Qty Returned")

        page_r = page.rect
        y0 = page_r.y0
        y1 = page_r.y1
        if qty_shipped_boxes:
            y0 = qty_shipped_boxes[0].y0 - 20
        if qty_returned_boxes:
            y1 = qty_returned_boxes[0].y0 - 20

        # Select the model/internet header boxes closest to the shipped table range.
        chosen_model = None
        chosen_internet = None
        if model_boxes:
            chosen_model = min(model_boxes, key=lambda b: abs((b.y0 + b.y1) / 2 - (y0 + y1) / 2))
        if internet_boxes:
            chosen_internet = min(internet_boxes, key=lambda b: abs((b.y0 + b.y1) / 2 - (y0 + y1) / 2))

        if chosen_model and chosen_internet:
            m = chosen_model
            i = chosen_internet
            x0 = m.x0 - 15
            x1 = i.x0 + 0  # model column ends near the internet column start

            x0 = max(page_r.x0, x0)
            x1 = min(page_r.x1, x1)
            y0 = max(page_r.y0, y0)
            y1 = min(page_r.y1, y1)

            rect_candidate = fitz.Rect(x0, y0, x1, y1)
            if rect_candidate.width > 40 and rect_candidate.height > 80:
                rect = rect_candidate
    except Exception:
        rect = None

    def run_ocr(clip_rect):
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip_rect, alpha=False)
        img_bytes = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        img_np = np.array(img)
        return reader_global.readtext(img_np)

    results = run_ocr(rect)

    # Collect potential model codes from OCR results.
    # easyocr gives text snippets; we filter to model-like tokens.
    codes = []
    for (_bbox, text, _conf) in results:
        t = text.strip().upper()
        # Split by non-alnum to recover joined tokens.
        parts = re.split(r"[^A-Z0-9]+", t)
        for p in parts:
            if is_model_like(p):
                # Normalize common OCR confusion
                p = p.replace("O", "0") if "O" in p and p.count("0") == 0 else p
                codes.append(p)

    # de-duplicate preserving order
    seen = set()
    uniq = []
    for c in codes:
        if c not in seen:
            seen.add(c)
            uniq.append(c)

    sample_like = [c for c in uniq if is_sample_code(c)]

    # Fallback: if nothing plausible found, run OCR over the whole page.
    if len(uniq) < 2 and not sample_like:
        results2 = run_ocr(None)
        codes2 = []
        for (_bbox, text, _conf) in results2:
            t = text.strip().upper()
            parts = re.split(r"[^A-Z0-9]+", t)
            for p in parts:
                if is_model_like(p):
                    codes2.append(p)
        seen = set()
        uniq2 = []
        for c in codes2:
            if c not in seen:
                seen.add(c)
                uniq2.append(c)
        uniq = uniq2
        sample_like = [c for c in uniq if is_sample_code(c)]

    return uniq, sample_like


def main():
    pdf_path = Path("C:/Users/wdevi/Downloads/\u042d\u043a\u0437 5.pdf")
    pages = [16, 22, 30]

    global reader_global
    os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    reader_global = easyocr.Reader(["en"], gpu=False)

    for p in pages:
        codes, sample_codes = ocr_page_model_column(pdf_path, p, scale=2.5)
        print(f"--- OCR page {p} ---")
        print("codes:", codes[:30])
        print("sample_codes:", sample_codes[:30])


if __name__ == "__main__":
    main()

