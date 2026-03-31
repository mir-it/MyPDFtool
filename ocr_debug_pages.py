import os
import io
import re
from pathlib import Path

import certifi
import fitz
import easyocr
import numpy as np
from PIL import Image


def main():
    os.environ.setdefault("SSL_CERT_FILE", certifi.where())

    pdf = Path("C:/Users/wdevi/Downloads/\u042d\u043a\u0437 5.pdf")
    pages = [22, 30]
    scale = 2.2

    reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    doc = fitz.open(str(pdf))

    for page_number in pages:
        page = doc[page_number - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        arr = np.array(img)
        results = reader.readtext(arr)
        texts = [t for (_b, t, _c) in results]

        # Extract sample-ish tokens (with relaxed matching).
        hits = []
        for t in texts:
            u = t.upper()
            if any(suf in u for suf in ["SMPL", "SMP", "SM", "SAMPLE", "SAMPL"]) or re.search(r"SM\\d|SM$", u):
                hits.append(u)

        print(f"--- page {page_number} ---")
        print("hits:", hits[:40])

    doc.close()


if __name__ == "__main__":
    main()

