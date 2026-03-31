# PDF Tools — Split, Merge, Packing + Labels

## When it works

**Right now**, as long as you have **Python 3.10+** and **PyMuPDF** installed.

## How to run

1. **No black console:** double-click **`START_NO_CONSOLE.vbs`**.  
   It runs **`launch_gui.py`**, which uses **the same `pythonw.exe` as the icon**, runs **`pip install -r requirements.txt`** if `fitz` (PyMuPDF) is missing, then starts the UI.  
   If something still fails, open **`data\last_startup_error.txt`** (inside the `MyAI` folder).
2. **With console** (easier to see errors): **`run_pdf_tools.bat`**
3. **Manual:**
   ```bat
   pip install -r requirements.txt
   python pdf_splitter_app.py
   ```

## Desktop shortcut (no console)

1. Find **`pythonw.exe`** (next to `python.exe`), e.g.  
   `C:\Users\YOUR_NAME\AppData\Local\Programs\Python\Python3xx\pythonw.exe`  
   (in `cmd`: `where pythonw`).
2. Right-click desktop → **New** → **Shortcut**.
3. **Target:**  
   `"full\path\to\pythonw.exe" "C:\Users\...\Desktop\MyAI\pdf_splitter_app.py"`
4. **Start in:** `C:\Users\...\Desktop\MyAI`
5. Name the shortcut e.g. **The Magnificent Creation of Konstantin**

If **`pyw`** is missing, install Python from python.org with **Add python.exe to PATH**, or fix the path to `pythonw.exe` manually.

## Tabs

| Tab | What it does |
|-----|----------------|
| **Split Orders / Samples** | Splits one PDF into `_orders.pdf` and `_samples.pdf` |
| **Merge PDFs** | Concatenates PDFs in list order |
| **Packing Slips + Labels** | All slips first, then labels matched by **REF 1:** vs 8-digit ref after **Ship Via** / **PO #** on the slip |

## Dependencies

- **tkinter** — usually included with Python on Windows  
- **PyMuPDF** (`pip install PyMuPDF`)
