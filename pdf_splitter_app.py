import re
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF

# Home Depot packing slip: sometimes "Date:" and "3/27/26" are on different lines
# (Ship Via / order # lines in between). Match same-line first, else first M/D/YY after "Date:".
_HD_SLIP_DATE_INLINE = re.compile(
    r"Date:\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\b",
    re.IGNORECASE,
)
_HD_SLIP_DATE_AFTER_LABEL = re.compile(
    r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b",
)
_DATE_LABEL = re.compile(r"Date\s*:", re.IGNORECASE)
_DATE_SORT_MISSING = date(9999, 12, 31)

from classify_pages import classify_pdf
from home_depot_thd_merger import (
    default_hd_exclusions_path,
    load_hd_exclusions,
    merge_hd_packing_with_thd_labels,
    save_hd_exclusions,
)
from tileclub_sample_merger import (
    default_exclusions_path,
    load_exclusions,
    merge_packing_with_sample_labels,
    save_exclusions,
)


def write_subset_pdf(src_path: Path, pages_1_based: list[int], output_path: Path) -> None:
    src = fitz.open(str(src_path))
    out = fitz.open()
    try:
        for p in pages_1_based:
            idx = p - 1
            if 0 <= idx < len(src):
                out.insert_pdf(src, from_page=idx, to_page=idx)
        out.save(str(output_path))
    finally:
        out.close()
        src.close()


def _date_from_mdy_groups(m: re.Match[str]) -> date | None:
    month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if year < 100:
        year += 2000
    try:
        return date(year, month, day)
    except ValueError:
        return None


def extract_hd_packing_slip_date(text: str) -> date | None:
    """
    Order date on homedepot.com packing slips: 'Date: M/D/YY' on one line, or label 'Date:'
    with the numeric date several lines below (PDF text order quirk).
    """
    if not text:
        return None
    m = _HD_SLIP_DATE_INLINE.search(text)
    if m:
        return _date_from_mdy_groups(m)
    m_label = _DATE_LABEL.search(text)
    if not m_label:
        return None
    # First calendar date after the "Date:" label (skip Ship Via / WN… / PO lines).
    window = text[m_label.end() : m_label.end() + 3000]
    m2 = _HD_SLIP_DATE_AFTER_LABEL.search(window)
    if m2:
        return _date_from_mdy_groups(m2)
    return None


def _sort_pages_by_hd_slip_date(pdf_path: Path, pages_1_based: list[int]) -> tuple[list[int], list[str]]:
    """
    Reorder page numbers by slip Date: on each page. Pages with no parseable date sort last,
    preserving ascending page order among those.
    """
    if not pages_1_based:
        return [], []
    doc = fitz.open(str(pdf_path))
    log: list[str] = []
    try:
        rows: list[tuple[date, int]] = []
        for p in pages_1_based:
            idx = p - 1
            if not (0 <= idx < len(doc)):
                continue
            text = doc[idx].get_text("text") or ""
            d = extract_hd_packing_slip_date(text)
            rows.append((d if d is not None else _DATE_SORT_MISSING, p))
        rows.sort(key=lambda t: (t[0], t[1]))
        out_pages = [p for _, p in rows]
        for d, p in rows:
            label = d.isoformat() if d != _DATE_SORT_MISSING else "(no Date:)"
            log.append(f"  page {p} -> {label}")
        return out_pages, log
    finally:
        doc.close()


def _sort_pdf_files_by_hd_first_page_date(paths: list[Path]) -> tuple[list[Path], list[str]]:
    """Merge PDFs: order whole documents by Date: on page 1. Same/unknown dates keep list order."""
    if not paths:
        return [], []
    decorated: list[tuple[date, int, Path]] = []
    for orig_i, path in enumerate(paths):
        d_opt: date | None = None
        try:
            doc = fitz.open(str(path))
            try:
                if len(doc) > 0:
                    d_opt = extract_hd_packing_slip_date(doc[0].get_text("text") or "")
            finally:
                doc.close()
        except Exception:
            d_opt = None
        d_key = d_opt if d_opt is not None else _DATE_SORT_MISSING
        decorated.append((d_key, orig_i, path))
    decorated.sort(key=lambda t: (t[0], t[1]))
    sorted_paths = [t[2] for t in decorated]
    log: list[str] = []
    for d_key, _oi, path in decorated:
        label = d_key.isoformat() if d_key != _DATE_SORT_MISSING else "(no Date: on page 1)"
        log.append(f"  {path.name}: {label}")
    return sorted_paths, log


def merge_pdfs(input_paths: list[Path], output_path: Path) -> None:
    out = fitz.open()
    try:
        for p in input_paths:
            src = fitz.open(str(p))
            try:
                out.insert_pdf(src)
            finally:
                src.close()
        out.save(str(output_path))
    finally:
        out.close()


def _extract_slip_match_id(text: str) -> str | None:
    """
    Packing slip: 8-digit id that matches UPS 'REF 1:' on the label.

    On Home Depot slips the header order is often:
      Purchase Order # / Date / Ship Via / WN12345678 / 04566596
    The WN... line is *not* the ship ref; the next 8-digit line is. The same
    number appears as 'PO # 04566596' on the return stub.
    """
    if not text:
        return None
    patterns = [
        # Ship Via -> internal order line -> 8-digit ship ref (most reliable)
        r"Ship\s+Via:\s*\n\S+\n(\d{8})\b",
        # Return form stub
        r"PO\s*#\s*(\d{8})\b",
        # Plain header when PDF text order is simple (no WN... line before ref)
        r"Purchase\s+Order\s*#?\s*:\s*(\d{8})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.MULTILINE)
        if m:
            return m.group(1)
    return None


def _extract_ref1_id(text: str) -> str | None:
    """Shipping label: REF 1: 04566596 or REF 1:02798429 (space after colon optional)."""
    if not text:
        return None
    m = re.search(r"REF\s*1\s*:\s*(\d{8})\b", text, re.IGNORECASE | re.MULTILINE)
    return m.group(1) if m else None


# -----------------------------------------------------------------------------
# Packing slips + labels — OPTIONAL multi-page slip grouping (fallback mode)
#
# When False (CURRENT): each slip page with a matched ref gets its OWN copy of
# the label page (even if several consecutive slip pages belong to one order).
#
# When True: consecutive slip pages with the SAME ref (and a match in label_map)
# are merged into one block: all those slip pages in a row, then ONE label page
# for the whole block. A ref break / page without ref starts a new block.
#
# NOTE: False means the mode is OFF; the tab has no toggle for this yet.
# To enable later: set True and run QA on test PDFs.
# -----------------------------------------------------------------------------
PACKING_SLIPS_GROUP_CONSECUTIVE_SAME_REF = False


def merge_slips_with_labels(
    packing_path: Path,
    labels_path: Path,
    output_path: Path,
) -> tuple[list[str], int, int, Path | None]:
    """
    Entry point for tab «Packing slips + labels». Dispatches to per-page or grouped implementation.
    """
    if PACKING_SLIPS_GROUP_CONSECUTIVE_SAME_REF:
        return _merge_slips_with_labels_grouped_consecutive_same_ref(
            packing_path, labels_path, output_path
        )
    return _merge_slips_with_labels_per_page(packing_path, labels_path, output_path)


def _merge_slips_with_labels_per_page(
    packing_path: Path,
    labels_path: Path,
    output_path: Path,
) -> tuple[list[str], int, int, Path | None]:
    """
    Main PDF: only packing pages that found a label match (preserving slip order), then those
    label pages in the same order (one label page per matched slip page).

    Second PDF (same folder, only if needed): packing pages with no label match, named
    {stem}_packing_slips_without_labels_YYYY-MM-DD.pdf

    Returns (log_lines, num_matched_slip_pages, num_labels_appended, unmatched_path_or_none).
    """
    log: list[str] = []
    slips = fitz.open(str(packing_path))
    labels = fitz.open(str(labels_path))
    out = fitz.open()
    unmatched_path: Path | None = None
    try:
        label_map: dict[str, int] = {}
        for li in range(len(labels)):
            t = labels[li].get_text("text") or ""
            rid = _extract_ref1_id(t)
            if rid:
                label_map[rid] = li

        matched_slip_indices: list[int] = []
        matched_label_indices: list[int] = []
        unmatched_slip_indices: list[int] = []
        summary_no_match: list[tuple[int, str | None]] = []

        for si in range(len(slips)):
            t = slips[si].get_text("text") or ""
            mid = _extract_slip_match_id(t)
            p1 = si + 1
            if mid and mid in label_map:
                li = label_map[mid]
                matched_slip_indices.append(si)
                matched_label_indices.append(li)
                log.append(f"Packing page {p1}: ref {mid} -> label page {li + 1} (REF 1)")
            else:
                unmatched_slip_indices.append(si)
                if mid:
                    log.append(f"Packing page {p1}: ref {mid} -> no matching REF 1:")
                    summary_no_match.append((p1, mid))
                else:
                    log.append(f"Packing page {p1}: slip ref not found -> skip label")
                    summary_no_match.append((p1, None))

        for si in matched_slip_indices:
            out.insert_pdf(slips, from_page=si, to_page=si)
        for li in matched_label_indices:
            out.insert_pdf(labels, from_page=li, to_page=li)

        out.save(str(output_path))

        n_matched = len(matched_slip_indices)
        n_lbl = len(matched_label_indices)
        log.append(
            f"--- Main output: {n_matched} packing page(s) with label + {n_lbl} label page(s) ---"
        )
        if n_matched:
            log.append(f"Labels start at PDF page {n_matched + 1}")
        if summary_no_match:
            log.append("--- No label match (slip page / order ref) ---")
            for p1, ref in summary_no_match:
                log.append(f"  page {p1}: {ref if ref is not None else 'ref not found on slip'}")

        if unmatched_slip_indices:
            day = date.today().isoformat()
            stem = output_path.stem
            unmatched_path = (
                output_path.parent / f"{stem}_packing_slips_without_labels_{day}.pdf"
            )
            orphan = fitz.open()
            try:
                for si in unmatched_slip_indices:
                    orphan.insert_pdf(slips, from_page=si, to_page=si)
                orphan.save(str(unmatched_path))
            finally:
                orphan.close()
            log.append(
                f"--- Orphans: {len(unmatched_slip_indices)} packing page(s) -> {unmatched_path.name} ---"
            )
        else:
            log.append("--- No unmatched packing pages; second file not written. ---")

        return log, n_matched, n_lbl, unmatched_path
    finally:
        out.close()
        labels.close()
        slips.close()


def _merge_slips_with_labels_grouped_consecutive_same_ref(
    packing_path: Path,
    labels_path: Path,
    output_path: Path,
) -> tuple[list[str], int, int, Path | None]:
    """
    Fallback mode (only when PACKING_SLIPS_GROUP_CONSECUTIVE_SAME_REF = True).

    Consecutive slip pages with the same ref (and a found label) → one slip block,
    then one label page. Second file name and no-match summary — same as per-page.

    Not used while the constant above is False.
    """
    log: list[str] = []
    slips = fitz.open(str(packing_path))
    labels = fitz.open(str(labels_path))
    out = fitz.open()
    unmatched_path: Path | None = None
    try:
        label_map: dict[str, int] = {}
        for li in range(len(labels)):
            t = labels[li].get_text("text") or ""
            rid = _extract_ref1_id(t)
            if rid:
                label_map[rid] = li

        unmatched_slip_indices: list[int] = []
        summary_no_match: list[tuple[int, str | None]] = []
        matched_slip_pages: list[int] = []
        n_label_pages = 0
        i = 0
        n = len(slips)

        while i < n:
            t = slips[i].get_text("text") or ""
            mid = _extract_slip_match_id(t)
            p1 = i + 1

            if not mid or mid not in label_map:
                unmatched_slip_indices.append(i)
                if mid:
                    log.append(f"Packing page {p1}: ref {mid} -> no matching REF 1:")
                    summary_no_match.append((p1, mid))
                else:
                    log.append(f"Packing page {p1}: slip ref not found -> skip label")
                    summary_no_match.append((p1, None))
                i += 1
                continue

            li = label_map[mid]
            group: list[int] = [i]
            j = i + 1
            while j < n:
                t2 = slips[j].get_text("text") or ""
                mid2 = _extract_slip_match_id(t2)
                if mid2 == mid and mid in label_map:
                    group.append(j)
                    j += 1
                else:
                    break

            pa, pb = group[0] + 1, group[-1] + 1
            if len(group) == 1:
                log.append(f"Packing page {pa}: ref {mid} -> label page {li + 1} (REF 1)")
            else:
                log.append(
                    f"Packing pages {pa}-{pb} ({len(group)} pp), ref {mid} -> "
                    f"label page {li + 1} (REF 1) [one label after block]"
                )
            for si in group:
                out.insert_pdf(slips, from_page=si, to_page=si)
                matched_slip_pages.append(si)
            out.insert_pdf(labels, from_page=li, to_page=li)
            n_label_pages += 1
            i = j

        out.save(str(output_path))

        n_matched = len(matched_slip_pages)
        log.append(
            f"--- Main output (grouped consecutive refs): {n_matched} packing page(s) + "
            f"{n_label_pages} label page(s) ---"
        )
        if n_matched:
            log.append(f"Labels start at PDF page {n_matched + 1}")
        if summary_no_match:
            log.append("--- No label match (slip page / order ref) ---")
            for pg, ref in summary_no_match:
                log.append(f"  page {pg}: {ref if ref is not None else 'ref not found on slip'}")

        if unmatched_slip_indices:
            day = date.today().isoformat()
            stem = output_path.stem
            unmatched_path = (
                output_path.parent / f"{stem}_packing_slips_without_labels_{day}.pdf"
            )
            orphan = fitz.open()
            try:
                for si in unmatched_slip_indices:
                    orphan.insert_pdf(slips, from_page=si, to_page=si)
                orphan.save(str(unmatched_path))
            finally:
                orphan.close()
            log.append(
                f"--- Orphans: {len(unmatched_slip_indices)} packing page(s) -> {unmatched_path.name} ---"
            )
        else:
            log.append("--- No unmatched packing pages; second file not written. ---")

        return log, n_matched, n_label_pages, unmatched_path
    finally:
        out.close()
        labels.close()
        slips.close()


APP_DISPLAY_NAME = "The Magnificent Creation of Konstantin"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_DISPLAY_NAME)
        self.geometry("1100x720")
        self.resizable(False, False)

        # Split tab state
        self.pdf_path_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.split_status_var = tk.StringVar(value="Ready")
        self.use_overrides_var = tk.BooleanVar(value=False)

        # Merge tab state
        self.merge_out_file_var = tk.StringVar()
        self.merge_status_var = tk.StringVar(value="Ready")
        self.merge_files: list[Path] = []

        # Packing slips + labels tab
        self.logistics_slips_var = tk.StringVar()
        self.logistics_labels_var = tk.StringVar()
        self.logistics_out_var = tk.StringVar()
        self.logistics_status_var = tk.StringVar(value="Ready")

        # Tile Club -smpl + labels
        self.tile_slip_var = tk.StringVar()
        self.tile_labels_var = tk.StringVar()
        self.tile_out_var = tk.StringVar()
        self.tile_status_var = tk.StringVar(value="Ready")
        self.tile_excl_add_var = tk.StringVar()
        self.tile_excl_search_var = tk.StringVar()
        self.tile_exclusions_path = default_exclusions_path()
        self.tile_exclusions: list[str] = load_exclusions(self.tile_exclusions_path)

        # Home Depot THD SMPL + barcode labels
        self.hd_slip_var = tk.StringVar()
        self.hd_labels_var = tk.StringVar()
        self.hd_out_var = tk.StringVar()
        self.hd_status_var = tk.StringVar(value="Ready")
        self.hd_excl_add_var = tk.StringVar()
        self.hd_excl_search_var = tk.StringVar()
        self.hd_exclusions_path = default_hd_exclusions_path()
        self.hd_exclusions: list[str] = load_hd_exclusions(self.hd_exclusions_path)

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        title_wrap = 820
        # Do not pass pady=(...) into tk.Label on Python 3.14+/some Tk builds — TclError: bad screen distance "0 10"
        title_lbl = tk.Label(
            root,
            text=APP_DISPLAY_NAME,
            font=("Segoe UI", 20, "bold"),
            fg="#1a1a1a",
            justify="center",
            wraplength=title_wrap,
        )
        title_lbl.pack(fill="x", pady=(0, 10))

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        split_tab = ttk.Frame(notebook, padding=14)
        merge_tab = ttk.Frame(notebook, padding=14)
        logistics_tab = ttk.Frame(notebook, padding=14)
        tile_tab = ttk.Frame(notebook, padding=0)
        hd_tab = ttk.Frame(notebook, padding=0)
        notebook.add(split_tab, text="Split Orders / Samples")
        notebook.add(merge_tab, text="Merge PDFs")
        notebook.add(logistics_tab, text="Packing Slips + Labels")
        notebook.add(tile_tab, text="Tile Club Samples (-smpl)")
        notebook.add(hd_tab, text="Home Depot THD (SMPL)")

        self._build_split_tab(split_tab)
        self._build_merge_tab(merge_tab)
        self._build_logistics_tab(logistics_tab)
        self._build_tileclub_tab(tile_tab)
        self._build_hd_thd_tab(hd_tab)

    def _build_split_tab(self, root: ttk.Frame):
        ttk.Label(root, text="Input PDF").grid(row=0, column=0, sticky="w")
        ttk.Entry(root, textvariable=self.pdf_path_var, width=88).grid(row=1, column=0, sticky="we", padx=(0, 8))
        ttk.Button(root, text="Browse...", command=self.pick_pdf).grid(row=1, column=1, sticky="e")

        ttk.Label(root, text="Output Folder").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(root, textvariable=self.out_dir_var, width=88).grid(row=3, column=0, sticky="we", padx=(0, 8))
        ttk.Button(root, text="Browse...", command=self.pick_out_dir).grid(row=3, column=1, sticky="e")

        ttk.Checkbutton(
            root,
            text="Use legacy overrides (for old approved test files)",
            variable=self.use_overrides_var,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))

        btns = ttk.Frame(root)
        btns.grid(row=5, column=0, columnspan=2, sticky="w", pady=(14, 0))
        ttk.Button(btns, text="Process PDF", command=self.process_split).pack(side="left")

        ttk.Separator(root).grid(row=6, column=0, columnspan=2, sticky="we", pady=(14, 10))
        ttk.Label(root, textvariable=self.split_status_var, foreground="#1f6f43").grid(row=7, column=0, columnspan=2, sticky="w")

        self.result_box = tk.Text(root, height=13, width=105, wrap="word")
        self.result_box.grid(row=8, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        self.result_box.configure(state="disabled")

        root.columnconfigure(0, weight=1)

    def _build_merge_tab(self, root: ttk.Frame):
        ttk.Label(root, text="Input PDFs (order matters)").grid(row=0, column=0, sticky="w")

        list_frame = ttk.Frame(root)
        list_frame.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        self.merge_listbox = tk.Listbox(list_frame, width=95, height=14, selectmode=tk.EXTENDED)
        self.merge_listbox.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.merge_listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.merge_listbox.configure(yscrollcommand=scroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        controls = ttk.Frame(root)
        controls.grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Button(controls, text="Add PDFs...", command=self.merge_add_files).pack(side="left")
        ttk.Button(controls, text="Remove Selected", command=self.merge_remove_selected).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Move Up", command=self.merge_move_up).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Move Down", command=self.merge_move_down).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Clear", command=self.merge_clear).pack(side="left", padx=(8, 0))

        ttk.Label(root, text="Output File").grid(row=3, column=0, sticky="w", pady=(12, 0))
        out_row = ttk.Frame(root)
        out_row.grid(row=4, column=0, sticky="we")
        ttk.Entry(out_row, textvariable=self.merge_out_file_var, width=88).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_row, text="Browse...", command=self.merge_pick_out_file).pack(side="left")

        run_row = ttk.Frame(root)
        run_row.grid(row=5, column=0, sticky="w", pady=(14, 0))
        ttk.Button(run_row, text="Merge PDFs", command=self.process_merge).pack(side="left")

        ttk.Separator(root).grid(row=6, column=0, sticky="we", pady=(14, 10))
        ttk.Label(root, textvariable=self.merge_status_var, foreground="#1f6f43").grid(row=7, column=0, sticky="w")

        self.merge_result_box = tk.Text(root, height=6, width=105, wrap="word")
        self.merge_result_box.grid(row=8, column=0, sticky="nsew", pady=(8, 0))
        self.merge_result_box.configure(state="disabled")

        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

    def _build_logistics_tab(self, root: ttk.Frame):
        ttk.Label(
            root,
            text=(
                "Match: 8-digit ref on slip (after Ship Via / WN line, or \"PO #\" on return stub) "
                "<-> REF 1: on label. Output (1): matched packing pages in order, then their labels. "
                "Output (2) if needed: {your_output_name}_packing_slips_without_labels_YYYY-MM-DD.pdf "
                "next to the main file."
            ),
            wraplength=820,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(root, text="Packing slips PDF").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(root, textvariable=self.logistics_slips_var, width=78).grid(
            row=2, column=0, sticky="we", padx=(0, 8)
        )
        ttk.Button(root, text="Browse...", command=self.logistics_pick_slips).grid(row=2, column=1, sticky="e")

        ttk.Label(root, text="Shipping labels PDF").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(root, textvariable=self.logistics_labels_var, width=78).grid(
            row=4, column=0, sticky="we", padx=(0, 8)
        )
        ttk.Button(root, text="Browse...", command=self.logistics_pick_labels).grid(row=4, column=1, sticky="e")

        ttk.Label(root, text="Output PDF").grid(row=5, column=0, sticky="w", pady=(10, 0))
        out_row = ttk.Frame(root)
        out_row.grid(row=6, column=0, columnspan=2, sticky="we")
        ttk.Entry(out_row, textvariable=self.logistics_out_var, width=78).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_row, text="Browse...", command=self.logistics_pick_out).pack(side="left")

        run_row = ttk.Frame(root)
        run_row.grid(row=7, column=0, columnspan=2, sticky="w", pady=(14, 0))
        ttk.Button(run_row, text="Merge packing slips + labels", command=self.process_logistics).pack(side="left")

        ttk.Separator(root).grid(row=8, column=0, columnspan=2, sticky="we", pady=(14, 10))
        ttk.Label(root, textvariable=self.logistics_status_var, foreground="#1f6f43").grid(
            row=9, column=0, columnspan=2, sticky="w"
        )

        self.logistics_log = tk.Text(root, height=14, width=105, wrap="word")
        self.logistics_log.grid(row=10, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        self.logistics_log.configure(state="disabled")

        root.columnconfigure(0, weight=1)
        root.rowconfigure(10, weight=1)

    def _build_tileclub_tab(self, root: ttk.Frame):
        BG = "#1e1e1e"
        FG = "#d4d4d4"
        SUB = "#9d9d9d"
        ACCENT = "#3794ff"
        PANEL = "#252526"

        shell = tk.Frame(root, bg=BG)
        shell.pack(fill="both", expand=True)

        main = tk.Frame(shell, bg=BG)
        main.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        right = tk.Frame(shell, bg=PANEL, width=280)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        tk.Label(
            main,
            text="PDF Merger Pro — Tile Club samples",
            font=("Segoe UI", 16, "bold"),
            bg=BG,
            fg=FG,
        ).pack(anchor="w")
        tk.Label(
            main,
            text="Packing slip: *-smpl lines → find label sheet in catalog → append to end of PDF.",
            font=("Segoe UI", 9),
            bg=BG,
            fg=SUB,
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(4, 14))

        def row_file(label_txt: str, var: tk.StringVar, browse_cmd, border: str):
            fr = tk.Frame(main, bg=BG)
            fr.pack(fill="x", pady=(0, 10))
            tk.Label(fr, text=label_txt, font=("Segoe UI", 10), bg=BG, fg=FG).pack(anchor="w")
            box = tk.Frame(fr, bg=border, padx=2, pady=2)
            box.pack(fill="x", pady=(4, 0))
            inner = tk.Frame(box, bg=PANEL)
            inner.pack(fill="x")
            e = tk.Entry(
                inner,
                textvariable=var,
                font=("Consolas", 10),
                bg="#3c3c3c",
                fg=FG,
                insertbackground=FG,
                relief="flat",
            )
            e.pack(side="left", fill="x", expand=True, padx=8, pady=8)
            tk.Button(
                inner,
                text="Browse…",
                command=browse_cmd,
                font=("Segoe UI", 9),
                bg="#3c3c3c",
                fg=FG,
                activebackground="#505050",
                activeforeground=FG,
                relief="flat",
            ).pack(side="right", padx=6, pady=6)

        row_file("1. Packing slip PDF", self.tile_slip_var, self.tile_pick_slip, "#3a7bd5")
        row_file("2. Label catalog (New_4x5_TC_ALL…)", self.tile_labels_var, self.tile_pick_labels, "#7c4dff")
        row_file("3. Output", self.tile_out_var, self.tile_pick_out, "#2e9d5c")

        run_fr = tk.Frame(main, bg=BG)
        run_fr.pack(fill="x", pady=(16, 8))
        tk.Button(
            run_fr,
            text="MERGE",
            command=self.tile_run_merge,
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT,
            fg="white",
            activebackground="#1f6feb",
            activeforeground="white",
            relief="flat",
            padx=18,
            pady=10,
        ).pack(side="left")
        tk.Button(
            run_fr,
            text="Open output folder",
            command=self.tile_open_out_dir,
            font=("Segoe UI", 9),
            bg="#3c3c3c",
            fg=FG,
            activebackground="#505050",
            relief="flat",
            padx=12,
            pady=8,
        ).pack(side="left", padx=(12, 0))

        tk.Label(main, textvariable=self.tile_status_var, font=("Segoe UI", 9), bg=BG, fg="#3ecb7a").pack(
            anchor="w", pady=(10, 4)
        )

        log_fr = tk.Frame(main, bg=PANEL)
        log_fr.pack(fill="both", expand=True, pady=(6, 0))
        self.tile_log = tk.Text(
            log_fr,
            height=14,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg=FG,
            insertbackground=FG,
            relief="flat",
            wrap="word",
        )
        self.tile_log.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        tscr = tk.Scrollbar(log_fr, command=self.tile_log.yview)
        tscr.pack(side="right", fill="y")
        self.tile_log.config(yscrollcommand=tscr.set, state="disabled")

        tk.Label(
            right,
            text="SKU exclusions",
            font=("Segoe UI", 11, "bold"),
            bg=PANEL,
            fg=FG,
        ).pack(anchor="w", padx=12, pady=(14, 6))
        tk.Label(
            right,
            text="If the code contains this substring — skip printing (e.g. tcmod → TCMODOLV258).",
            font=("Segoe UI", 8),
            bg=PANEL,
            fg=SUB,
            wraplength=250,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        tk.Entry(
            right,
            textvariable=self.tile_excl_search_var,
            font=("Segoe UI", 9),
            bg="#3c3c3c",
            fg=FG,
            insertbackground=FG,
            relief="flat",
        ).pack(fill="x", padx=12, pady=(0, 6))
        self.tile_excl_search_var.trace_add("write", lambda *_: self._tile_refresh_excl_list())

        add_fr = tk.Frame(right, bg=PANEL)
        add_fr.pack(fill="x", padx=12, pady=(0, 8))
        tk.Entry(
            add_fr,
            textvariable=self.tile_excl_add_var,
            font=("Segoe UI", 9),
            bg="#3c3c3c",
            fg=FG,
            insertbackground=FG,
            relief="flat",
        ).pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=4)
        tk.Button(
            add_fr,
            text="+",
            command=self.tile_excl_add,
            font=("Segoe UI", 11, "bold"),
            width=3,
            bg=ACCENT,
            fg="white",
            relief="flat",
        ).pack(side="right")

        tk.Label(
            right,
            text="List (saved automatically)",
            font=("Segoe UI", 8),
            bg=PANEL,
            fg=SUB,
        ).pack(anchor="w", padx=12, pady=(0, 4))

        list_fr = tk.Frame(right, bg=PANEL)
        list_fr.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.tile_excl_list = tk.Listbox(
            list_fr,
            font=("Consolas", 10),
            bg="#1e1e1e",
            fg=FG,
            selectbackground=ACCENT,
            selectforeground="white",
            relief="flat",
            highlightthickness=0,
            activestyle="none",
        )
        self.tile_excl_list.pack(side="left", fill="both", expand=True)
        lscr = tk.Scrollbar(list_fr, command=self.tile_excl_list.yview)
        lscr.pack(side="right", fill="y")
        self.tile_excl_list.config(yscrollcommand=lscr.set)
        self.tile_excl_list.bind("<Delete>", lambda e: self.tile_excl_remove_selected())
        self.tile_excl_list.bind("<Return>", lambda e: self.tile_excl_remove_selected())

        self._tile_refresh_excl_list()

    def _tile_refresh_excl_list(self):
        if not hasattr(self, "tile_excl_list"):
            return
        q = (self.tile_excl_search_var.get() or "").strip().casefold()
        self.tile_excl_list.delete(0, tk.END)
        for s in sorted(self.tile_exclusions, key=str.casefold):
            if not q or q in s.casefold():
                self.tile_excl_list.insert(tk.END, s)

    def _tile_set_log(self, text: str):
        self.tile_log.configure(state="normal")
        self.tile_log.delete("1.0", tk.END)
        self.tile_log.insert("1.0", text)
        self.tile_log.configure(state="disabled")

    def _tile_persist_exclusions(self):
        save_exclusions(self.tile_exclusions, self.tile_exclusions_path)

    def tile_excl_add(self):
        raw = (self.tile_excl_add_var.get() or "").strip()
        if not raw:
            return
        if raw not in self.tile_exclusions:
            self.tile_exclusions.append(raw)
            self._tile_persist_exclusions()
        self.tile_excl_add_var.set("")
        self._tile_refresh_excl_list()

    def tile_excl_remove_selected(self):
        sel = list(self.tile_excl_list.curselection())
        if not sel:
            return
        label = self.tile_excl_list.get(sel[0])
        self.tile_exclusions = [x for x in self.tile_exclusions if x != label]
        self._tile_persist_exclusions()
        self._tile_refresh_excl_list()

    def tile_pick_slip(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.tile_slip_var.set(p)
            if not self.tile_out_var.get():
                self.tile_out_var.set(str(Path(p).parent / "Final_Merge.pdf"))

    def tile_pick_labels(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.tile_labels_var.set(p)

    def tile_pick_out(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="Final_Merge.pdf",
        )
        if p:
            self.tile_out_var.set(p)

    def tile_open_out_dir(self):
        p = (self.tile_out_var.get() or "").strip()
        if not p:
            messagebox.showinfo("Folder", "Specify an output file first.", parent=self)
            return
        folder = Path(p).parent
        if folder.exists():
            try:
                import os

                os.startfile(folder)  # type: ignore[attr-defined]
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=self)
        else:
            messagebox.showinfo("Folder", f"Folder does not exist yet:\n{folder}", parent=self)

    def tile_run_merge(self):
        slip = Path(self.tile_slip_var.get().strip())
        labels = Path(self.tile_labels_var.get().strip())
        out_s = self.tile_out_var.get().strip()
        if not slip.exists() or slip.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid packing slip PDF.", parent=self)
            return
        if not labels.exists() or labels.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid label catalog PDF.", parent=self)
            return
        if not out_s:
            messagebox.showerror("Error", "Specify an output path.", parent=self)
            return
        out = Path(out_s)
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")
            self.tile_out_var.set(str(out))

        self.tile_status_var.set("Processing…")
        self.update_idletasks()
        try:
            log_lines = merge_packing_with_sample_labels(slip, labels, out, self.tile_exclusions)
            self._tile_set_log("\n".join(log_lines) + f"\n\nFile:\n{out}")
            self.tile_status_var.set("Completed")
        except Exception as e:
            self.tile_status_var.set("Error")
            messagebox.showerror("Error", str(e), parent=self)

    def _build_hd_thd_tab(self, root: ttk.Frame):
        BG = "#1e1e1e"
        FG = "#d4d4d4"
        SUB = "#9d9d9d"
        ACCENT = "#f57c00"
        PANEL = "#252526"

        shell = tk.Frame(root, bg=BG)
        shell.pack(fill="both", expand=True)

        main = tk.Frame(shell, bg=BG)
        main.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        right = tk.Frame(shell, bg=PANEL, width=280)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        tk.Label(
            main,
            text="PDF Merger Pro — Home Depot (THD)",
            font=("Segoe UI", 16, "bold"),
            bg=BG,
            fg=FG,
        ).pack(anchor="w")
        tk.Label(
            main,
            text="Packing slip: Model Number column (SMPL/ASMPL) + Qty Shipped; "
            "labels: Product Sku in THD Barcode Labels.",
            font=("Segoe UI", 9),
            bg=BG,
            fg=SUB,
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(4, 14))

        def row_file(label_txt: str, var: tk.StringVar, browse_cmd, border: str):
            fr = tk.Frame(main, bg=BG)
            fr.pack(fill="x", pady=(0, 10))
            tk.Label(fr, text=label_txt, font=("Segoe UI", 10), bg=BG, fg=FG).pack(anchor="w")
            box = tk.Frame(fr, bg=border, padx=2, pady=2)
            box.pack(fill="x", pady=(4, 0))
            inner = tk.Frame(box, bg=PANEL)
            inner.pack(fill="x")
            e = tk.Entry(
                inner,
                textvariable=var,
                font=("Consolas", 10),
                bg="#3c3c3c",
                fg=FG,
                insertbackground=FG,
                relief="flat",
            )
            e.pack(side="left", fill="x", expand=True, padx=8, pady=8)
            tk.Button(
                inner,
                text="Browse…",
                command=browse_cmd,
                font=("Segoe UI", 9),
                bg="#3c3c3c",
                fg=FG,
                activebackground="#505050",
                activeforeground=FG,
                relief="flat",
            ).pack(side="right", padx=6, pady=6)

        row_file("1. Packing slip PDF (Home Depot)", self.hd_slip_var, self.hd_pick_slip, "#3a7bd5")
        row_file("2. THD Barcode Labels PDF", self.hd_labels_var, self.hd_pick_labels, "#c66900")
        row_file("3. Output", self.hd_out_var, self.hd_pick_out, "#2e9d5c")

        run_fr = tk.Frame(main, bg=BG)
        run_fr.pack(fill="x", pady=(16, 8))
        tk.Button(
            run_fr,
            text="MERGE",
            command=self.hd_run_merge,
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT,
            fg="white",
            activebackground="#e65100",
            activeforeground="white",
            relief="flat",
            padx=18,
            pady=10,
        ).pack(side="left")
        tk.Button(
            run_fr,
            text="Open output folder",
            command=self.hd_open_out_dir,
            font=("Segoe UI", 9),
            bg="#3c3c3c",
            fg=FG,
            activebackground="#505050",
            relief="flat",
            padx=12,
            pady=8,
        ).pack(side="left", padx=(12, 0))

        tk.Label(main, textvariable=self.hd_status_var, font=("Segoe UI", 9), bg=BG, fg="#3ecb7a").pack(
            anchor="w", pady=(10, 4)
        )

        log_fr = tk.Frame(main, bg=PANEL)
        log_fr.pack(fill="both", expand=True, pady=(6, 0))
        self.hd_log = tk.Text(
            log_fr,
            height=14,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg=FG,
            insertbackground=FG,
            relief="flat",
            wrap="word",
        )
        self.hd_log.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        hscr = tk.Scrollbar(log_fr, command=self.hd_log.yview)
        hscr.pack(side="right", fill="y")
        self.hd_log.config(yscrollcommand=hscr.set, state="disabled")

        tk.Label(
            right,
            text="SKU exclusions",
            font=("Segoe UI", 11, "bold"),
            bg=PANEL,
            fg=FG,
        ).pack(anchor="w", padx=12, pady=(14, 6))
        tk.Label(
            right,
            text="If the code contains this substring — the line is skipped (separate list from Tile Club).",
            font=("Segoe UI", 8),
            bg=PANEL,
            fg=SUB,
            wraplength=250,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        tk.Entry(
            right,
            textvariable=self.hd_excl_search_var,
            font=("Segoe UI", 9),
            bg="#3c3c3c",
            fg=FG,
            insertbackground=FG,
            relief="flat",
        ).pack(fill="x", padx=12, pady=(0, 6))
        self.hd_excl_search_var.trace_add("write", lambda *_: self._hd_refresh_excl_list())

        add_fr = tk.Frame(right, bg=PANEL)
        add_fr.pack(fill="x", padx=12, pady=(0, 8))
        tk.Entry(
            add_fr,
            textvariable=self.hd_excl_add_var,
            font=("Segoe UI", 9),
            bg="#3c3c3c",
            fg=FG,
            insertbackground=FG,
            relief="flat",
        ).pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=4)
        tk.Button(
            add_fr,
            text="+",
            command=self.hd_excl_add,
            font=("Segoe UI", 11, "bold"),
            width=3,
            bg=ACCENT,
            fg="white",
            relief="flat",
        ).pack(side="right")

        tk.Label(
            right,
            text="List (saved automatically)",
            font=("Segoe UI", 8),
            bg=PANEL,
            fg=SUB,
        ).pack(anchor="w", padx=12, pady=(0, 4))

        list_fr = tk.Frame(right, bg=PANEL)
        list_fr.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.hd_excl_list = tk.Listbox(
            list_fr,
            font=("Consolas", 10),
            bg="#1e1e1e",
            fg=FG,
            selectbackground=ACCENT,
            selectforeground="white",
            relief="flat",
            highlightthickness=0,
            activestyle="none",
        )
        self.hd_excl_list.pack(side="left", fill="both", expand=True)
        lscr = tk.Scrollbar(list_fr, command=self.hd_excl_list.yview)
        lscr.pack(side="right", fill="y")
        self.hd_excl_list.config(yscrollcommand=lscr.set)
        self.hd_excl_list.bind("<Delete>", lambda e: self.hd_excl_remove_selected())
        self.hd_excl_list.bind("<Return>", lambda e: self.hd_excl_remove_selected())

        self._hd_refresh_excl_list()

    def _hd_refresh_excl_list(self):
        if not hasattr(self, "hd_excl_list"):
            return
        q = (self.hd_excl_search_var.get() or "").strip().casefold()
        self.hd_excl_list.delete(0, tk.END)
        for s in sorted(self.hd_exclusions, key=str.casefold):
            if not q or q in s.casefold():
                self.hd_excl_list.insert(tk.END, s)

    def _hd_set_log(self, text: str):
        self.hd_log.configure(state="normal")
        self.hd_log.delete("1.0", tk.END)
        self.hd_log.insert("1.0", text)
        self.hd_log.configure(state="disabled")

    def _hd_persist_exclusions(self):
        save_hd_exclusions(self.hd_exclusions, self.hd_exclusions_path)

    def hd_excl_add(self):
        raw = (self.hd_excl_add_var.get() or "").strip()
        if not raw:
            return
        if raw not in self.hd_exclusions:
            self.hd_exclusions.append(raw)
            self._hd_persist_exclusions()
        self.hd_excl_add_var.set("")
        self._hd_refresh_excl_list()

    def hd_excl_remove_selected(self):
        sel = list(self.hd_excl_list.curselection())
        if not sel:
            return
        label = self.hd_excl_list.get(sel[0])
        self.hd_exclusions = [x for x in self.hd_exclusions if x != label]
        self._hd_persist_exclusions()
        self._hd_refresh_excl_list()

    def hd_pick_slip(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.hd_slip_var.set(p)
            if not self.hd_out_var.get():
                self.hd_out_var.set(str(Path(p).parent / "THD_Final_Merge.pdf"))

    def hd_pick_labels(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.hd_labels_var.set(p)

    def hd_pick_out(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="THD_Final_Merge.pdf",
        )
        if p:
            self.hd_out_var.set(p)

    def hd_open_out_dir(self):
        p = (self.hd_out_var.get() or "").strip()
        if not p:
            messagebox.showinfo("Folder", "Specify an output file first.", parent=self)
            return
        folder = Path(p).parent
        if folder.exists():
            try:
                import os

                os.startfile(folder)  # type: ignore[attr-defined]
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=self)
        else:
            messagebox.showinfo("Folder", f"Folder does not exist yet:\n{folder}", parent=self)

    def hd_run_merge(self):
        slip = Path(self.hd_slip_var.get().strip())
        labels = Path(self.hd_labels_var.get().strip())
        out_s = self.hd_out_var.get().strip()
        if not slip.exists() or slip.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid packing slip PDF.", parent=self)
            return
        if not labels.exists() or labels.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid THD Barcode Labels PDF.", parent=self)
            return
        if not out_s:
            messagebox.showerror("Error", "Specify an output path.", parent=self)
            return
        out = Path(out_s)
        if out.suffix.lower() != ".pdf":
            out = out.with_suffix(".pdf")
            self.hd_out_var.set(str(out))

        self.hd_status_var.set("Processing…")
        self.update_idletasks()
        try:
            log_lines = merge_hd_packing_with_thd_labels(slip, labels, out, self.hd_exclusions)
            self._hd_set_log("\n".join(log_lines) + f"\n\nFile:\n{out}")
            self.hd_status_var.set("Completed")
        except Exception as e:
            self.hd_status_var.set("Error")
            messagebox.showerror("Error", str(e), parent=self)

    def pick_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p:
            self.pdf_path_var.set(p)
            if not self.out_dir_var.get():
                self.out_dir_var.set(str(Path(p).parent))

    def pick_out_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.out_dir_var.set(p)

    def _set_text(self, widget: tk.Text, text: str):
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def process_split(self):
        pdf_path = Path(self.pdf_path_var.get().strip())
        out_dir = Path(self.out_dir_var.get().strip())

        if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid input PDF file.")
            return
        if not out_dir.exists() or not out_dir.is_dir():
            messagebox.showerror("Error", "Select a valid output folder.")
            return

        self.split_status_var.set("Processing...")
        self.update_idletasks()

        try:
            samples, orders = classify_pdf(pdf_path, use_overrides=self.use_overrides_var.get())

            samples_sorted, s_log = _sort_pages_by_hd_slip_date(pdf_path, samples)
            orders_sorted, o_log = _sort_pages_by_hd_slip_date(pdf_path, orders)

            stem = pdf_path.stem
            samples_out = out_dir / f"{stem}_samples.pdf"
            orders_out = out_dir / f"{stem}_orders.pdf"

            if samples_sorted:
                write_subset_pdf(pdf_path, samples_sorted, samples_out)
            else:
                if samples_out.exists():
                    samples_out.unlink()

            if orders_sorted:
                write_subset_pdf(pdf_path, orders_sorted, orders_out)
            else:
                if orders_out.exists():
                    orders_out.unlink()

            result = (
                f"Done.\n\n"
                f"Classification (unchanged): samples vs orders by existing rules.\n"
                f"Output page order: chronological by slip Date: (Home Depot header) within each file.\n\n"
                f"Samples pages ({len(samples_sorted)}): {samples_sorted}\n"
                f"Orders pages ({len(orders_sorted)}): {orders_sorted}\n\n"
                f"Samples date order:\n" + "\n".join(s_log) + "\n\n"
                f"Orders date order:\n" + "\n".join(o_log) + "\n\n"
                f"Output files:\n"
                f"- {samples_out if samples_sorted else '(no samples pages)'}\n"
                f"- {orders_out if orders_sorted else '(no orders pages)'}"
            )
            self._set_text(self.result_box, result)
            self.split_status_var.set("Completed")
        except Exception as e:
            self.split_status_var.set("Failed")
            messagebox.showerror("Error", str(e))

    # ---------------- Merge tab actions ----------------
    def _refresh_merge_listbox(self):
        self.merge_listbox.delete(0, tk.END)
        for p in self.merge_files:
            self.merge_listbox.insert(tk.END, str(p))

    def merge_add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        existing = {str(p).lower() for p in self.merge_files}
        for raw in paths:
            p = Path(raw)
            if str(p).lower() not in existing:
                self.merge_files.append(p)
                existing.add(str(p).lower())
        self._refresh_merge_listbox()

        if not self.merge_out_file_var.get() and self.merge_files:
            base = self.merge_files[0].parent / "merged_output.pdf"
            self.merge_out_file_var.set(str(base))

    def merge_remove_selected(self):
        selected = list(self.merge_listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            del self.merge_files[idx]
        self._refresh_merge_listbox()

    def merge_move_up(self):
        selected = list(self.merge_listbox.curselection())
        if not selected or selected[0] == 0:
            return
        for idx in selected:
            self.merge_files[idx - 1], self.merge_files[idx] = self.merge_files[idx], self.merge_files[idx - 1]
        self._refresh_merge_listbox()
        for idx in [i - 1 for i in selected]:
            self.merge_listbox.selection_set(idx)

    def merge_move_down(self):
        selected = list(self.merge_listbox.curselection())
        if not selected or selected[-1] == len(self.merge_files) - 1:
            return
        for idx in reversed(selected):
            self.merge_files[idx + 1], self.merge_files[idx] = self.merge_files[idx], self.merge_files[idx + 1]
        self._refresh_merge_listbox()
        for idx in [i + 1 for i in selected]:
            self.merge_listbox.selection_set(idx)

    def merge_clear(self):
        self.merge_files = []
        self._refresh_merge_listbox()

    def merge_pick_out_file(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="merged_output.pdf",
        )
        if p:
            self.merge_out_file_var.set(p)

    def process_merge(self):
        if len(self.merge_files) < 2:
            messagebox.showerror("Error", "Add at least two PDF files to merge.")
            return

        out_file = Path(self.merge_out_file_var.get().strip())
        if not out_file:
            messagebox.showerror("Error", "Select an output PDF file.")
            return
        if out_file.suffix.lower() != ".pdf":
            out_file = out_file.with_suffix(".pdf")
            self.merge_out_file_var.set(str(out_file))

        self.merge_status_var.set("Merging...")
        self.update_idletasks()

        try:
            sorted_paths, sort_log = _sort_pdf_files_by_hd_first_page_date(self.merge_files)
            merge_pdfs(sorted_paths, out_file)
            result = (
                f"Done.\n\nMerged {len(sorted_paths)} files into:\n{out_file}\n\n"
                f"Chronological order (by Date: on page 1 of each PDF):\n"
                + "\n".join(sort_log)
                + "\n\nOriginal listbox order was:\n"
                + "\n".join(f"- {p}" for p in self.merge_files)
            )
            self._set_text(self.merge_result_box, result)
            self.merge_status_var.set("Completed")
        except Exception as e:
            self.merge_status_var.set("Failed")
            messagebox.showerror("Error", str(e))

    # ---------------- Packing slips + labels tab ----------------
    def logistics_pick_slips(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p:
            self.logistics_slips_var.set(p)
            if not self.logistics_out_var.get():
                self.logistics_out_var.set(str(Path(p).parent / "FINAL_MERGE_BLOCKS.pdf"))

    def logistics_pick_labels(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p:
            self.logistics_labels_var.set(p)

    def logistics_pick_out(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="FINAL_MERGE_BLOCKS.pdf",
        )
        if p:
            self.logistics_out_var.set(p)

    def process_logistics(self):
        slips_path = Path(self.logistics_slips_var.get().strip())
        labels_path = Path(self.logistics_labels_var.get().strip())
        out_path = Path(self.logistics_out_var.get().strip())

        if not slips_path.exists() or slips_path.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid packing slips PDF.")
            return
        if not labels_path.exists() or labels_path.suffix.lower() != ".pdf":
            messagebox.showerror("Error", "Select a valid labels PDF.")
            return
        if not out_path:
            messagebox.showerror("Error", "Select an output PDF path.")
            return
        if out_path.suffix.lower() != ".pdf":
            out_path = out_path.with_suffix(".pdf")
            self.logistics_out_var.set(str(out_path))

        self.logistics_status_var.set("Merging...")
        self.update_idletasks()

        try:
            log_lines, n_match, n_lbl, orphan_path = merge_slips_with_labels(
                slips_path, labels_path, out_path
            )
            body = "\n".join(log_lines) + f"\n\nMain output:\n{out_path}"
            if orphan_path:
                body += f"\n\nUnmatched slips only:\n{orphan_path}"
            self._set_text(self.logistics_log, body)
            msg = f"Completed — matched: {n_match} slips + {n_lbl} labels"
            if orphan_path:
                msg += f"; orphans -> {orphan_path.name}"
            self.logistics_status_var.set(msg)
        except Exception as e:
            self.logistics_status_var.set("Failed")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    App().mainloop()

