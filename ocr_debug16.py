import os
import io
import re
from pathlib import Path

import certifi
import fitz  # PyMuPDF
import easyocr
import numpy as np
from PIL import Image


def main():
    os.environ.setdefault("SSL_CERT_FILE", certifi.where())

    pdf = Path("C:/Users/wdevi/Downloads/\u042d\u043a\u0437 5.pdf")
    page_number = 16
    scale = 2.2

    reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    doc = fitz.open(str(pdf))
    page = doc[page_number - 1]

    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    arr = np.array(img)

    results = reader.readtext(arr)
    print("num results:", len(results))
    texts = [t for (_b, t, _c) in results]
    print("first snippets:", texts[:80])

    hits = []
    for t in texts:
        u = t.upper()
        if (
            any(suf in u for suf in ["SMPL", "SMP", "SM", "SAMPLE", "SAMPL"])
            or re.search(r"SM\\d|SM$", u)
        ):
            hits.append(u)

    print("sample-ish hits:", hits[:80])
    doc.close()


if __name__ == "__main__":
    main()

