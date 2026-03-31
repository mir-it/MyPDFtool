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
    page_number = 22
    scale = 2.5

    reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    doc = fitz.open(str(pdf))
    page = doc[page_number - 1]

    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    arr = np.array(img)

    results = reader.readtext(arr)
    tokens = []
    for (_b, t, _c) in results:
        u = t.strip().upper()
        if not u:
            continue
        # keep short-ish tokens too
        if any(s in u for s in ["SMPL", "SMP", "SM", "SAMPLE", "SAMPL"]):
            tokens.append(u)

    print("found snippets:", len(tokens))
    for t in tokens[:120]:
        print(t)

    # Also dump any token-like sequences with letters+digits that include 'SM'
    modelish = []
    for (_b, t, _c) in results:
        u = t.strip().upper()
        parts = re.split(r"[^A-Z0-9]+", u)
        for p in parts:
            if len(p) >= 6 and any(ch.isdigit() for ch in p) and "SM" in p:
                modelish.append(p)

    # de-dupe
    seen = set()
    uniq = []
    for m in modelish:
        if m not in seen:
            seen.add(m)
            uniq.append(m)

    print("\nmodelish containing SM:", len(uniq))
    for m in uniq[:80]:
        print(m)

    doc.close()


if __name__ == "__main__":
    main()

