"""
Entry point for double-click / pythonw: same Python as the icon gets PyMuPDF, then starts the UI.
"""
from __future__ import annotations

import subprocess
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
REQ = ROOT / "requirements.txt"
LOG = DATA / "last_startup_error.txt"


def _pip_kw() -> dict:
    if sys.platform == "win32" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        return {"creationflags": subprocess.CREATE_NO_WINDOW}
    return {}


def _ensure_pymupdf() -> None:
    last_err: BaseException | None = None
    for attempt in range(4):
        try:
            import fitz  # noqa: F401

            return
        except ImportError as e:
            last_err = e
            if attempt >= 3:
                hint = ""
                msg = str(e).lower()
                if "dll load failed" in msg or "specified module could not be found" in msg:
                    hint = (
                        " On Windows, install Microsoft Visual C++ Redistributable (x64): "
                        "https://aka.ms/vs/17/release/vc_redist.x64.exe"
                    )
                raise ImportError(
                    "PyMuPDF (fitz) could not be imported after pip install attempts." + hint
                ) from last_err
            cmd = [sys.executable, "-m", "pip", "install", "-q", "-r", str(REQ)]
            subprocess.check_call(cmd, cwd=str(ROOT), **_pip_kw())


def main() -> None:
    try:
        _ensure_pymupdf()
        from pdf_splitter_app import App

        App().mainloop()
    except Exception:
        DATA.mkdir(parents=True, exist_ok=True)
        LOG.write_text(traceback.format_exc(), encoding="utf-8")
        sys.exit(1)


if __name__ == "__main__":
    main()
