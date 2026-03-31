import argparse
import sys
from pathlib import Path

# Cyrillic filename in repo, written as escapes so this file stays ASCII-only.
_DEFAULT_SAMPLES = "\u0441\u0435\u043c\u043f\u043b\u044b.pdf"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract and print text from a PDF using PyMuPDF (fitz)."
    )
    parser.add_argument(
        "pdf",
        nargs="?",
        default=_DEFAULT_SAMPLES,
        help="Path to the PDF file (default: Cyrillic-named samples PDF in this folder).",
    )
    parser.add_argument(
        "--no-page-headers",
        action="store_true",
        help="Do not print page separators/headers.",
    )
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"File not found: {pdf_path}", file=sys.stderr)
        return 2
    if pdf_path.is_dir():
        print(f"Expected a PDF file, got a directory: {pdf_path}", file=sys.stderr)
        return 2

    try:
        import fitz  # PyMuPDF
    except Exception as e:
        print(
            "PyMuPDF is not installed. Install it with:\n"
            "  python -m pip install pymupdf",
            file=sys.stderr,
        )
        print(f"\nImport error: {e}", file=sys.stderr)
        return 3

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        print(f"Failed to open PDF: {pdf_path}\n{e}", file=sys.stderr)
        return 4

    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            text = page.get_text("text")
            if not args.no_page_headers:
                print(f"\n===== PAGE {page_index + 1}/{len(doc)} =====\n")
            print(text.rstrip())
    finally:
        doc.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

