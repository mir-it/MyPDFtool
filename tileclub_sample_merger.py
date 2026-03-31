"""
Tile Club sample packing slip -> find label pages in New_4x5 catalog by SKU (-smpl) or product name.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import fitz  # PyMuPDF

_QTY_LINE = re.compile(r"^\s*(\d+)\s+of\s+\d+\s*$", re.IGNORECASE)
_SKU_FIELD = re.compile(r"Product\s+Sku:\s*([A-Z0-9]+)", re.IGNORECASE)
_PAGE_MARKER = re.compile(r"^--\s*\d+\s+of\s+\d+\s+--\s*$")
_ORDER_MARK = re.compile(r"^Order\s*#", re.IGNORECASE)


def default_exclusions_path() -> Path:
    base = Path.home() / ".tileclub_sample_merger"
    base.mkdir(parents=True, exist_ok=True)
    return base / "exclusions.json"


def load_exclusions(path: Path | None = None) -> list[str]:
    p = path or default_exclusions_path()
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [str(x).strip() for x in data if str(x).strip()]
    except (json.JSONDecodeError, OSError):
        pass
    return []


def save_exclusions(items: list[str], path: Path | None = None) -> None:
    p = path or default_exclusions_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    clean = sorted({str(x).strip() for x in items if str(x).strip()})
    p.write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8")


def is_excluded(base_sku: str, exclusions: list[str]) -> bool:
    u = base_sku.upper()
    for ex in exclusions:
        needle = ex.strip().upper()
        if needle and needle in u:
            return True
    return False


def _norm_slug(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def _extract_base_sku_before_smpl(line: str) -> str | None:
    low = line.lower()
    if "-smpl" not in low:
        return None
    idx = low.index("-smpl")
    left = line[:idx]
    base = re.sub(r"\s+", "", left)
    if not base or not re.fullmatch(r"[A-Z0-9]+", base, re.IGNORECASE):
        return None
    return base.upper()


def _slip_title_from_lines(lines: list[str], smpl_index: int) -> str:
    parts: list[str] = []
    j = smpl_index - 1
    while j >= 0:
        prev = lines[j].strip()
        if not prev:
            j -= 1
            continue
        if _extract_base_sku_before_smpl(prev):
            break
        if _QTY_LINE.match(prev):
            break
        low = prev.casefold()
        if low in ("items", "quantity"):
            break
        if low == "items quantity" or (
            prev.upper().startswith("ITEMS") and "QUANTITY" in prev.upper()
        ):
            break
        if _ORDER_MARK.match(prev):
            break
        if _PAGE_MARKER.match(prev):
            break
        if prev.upper().startswith("SHIP TO") or prev.upper().startswith("BILL TO"):
            break
        if prev.upper().startswith("NOTES"):
            break
        if low == "sample":
            j -= 1
            continue
        parts.insert(0, prev)
        j -= 1
    joined = " ".join(parts)
    joined = re.sub(r"^\d+\s+of\s+\d+\s+", "", joined, flags=re.IGNORECASE)
    joined = re.sub(r"\s+Sample\s*$", "", joined, flags=re.IGNORECASE)
    joined = re.sub(r"\s+Sample\s+", " ", joined, flags=re.IGNORECASE)
    return joined.strip()


def parse_smpl_items_from_packing_doc(slips_doc: fitz.Document) -> list[dict]:
    chunks = [(slips_doc[i].get_text("text") or "") for i in range(len(slips_doc))]
    full = "\n".join(chunks)
    lines = full.splitlines()
    items: list[dict] = []
    for i, line in enumerate(lines):
        base = _extract_base_sku_before_smpl(line)
        if not base:
            continue
        qty = 1
        if i + 1 < len(lines):
            qm = _QTY_LINE.match(lines[i + 1])
            if qm:
                qty = int(qm.group(1))
        title = _slip_title_from_lines(lines, i)
        items.append({"sku": base, "qty": max(1, qty), "title": title, "slip_line": line[:80]})
    return items


def _label_header_before_sku(page_text: str) -> str:
    m = _SKU_FIELD.search(page_text)
    if not m:
        return ""
    head = page_text[: m.start()]
    return " ".join(head.split())


def build_label_index(labels_doc: fitz.Document) -> tuple[dict[str, int], list[tuple[int, str]]]:
    sku_to_page: dict[str, int] = {}
    name_rows: list[tuple[int, str]] = []
    for idx in range(len(labels_doc)):
        text = labels_doc[idx].get_text("text") or ""
        sm = _SKU_FIELD.search(text)
        if sm:
            sku = sm.group(1).upper()
            if sku not in sku_to_page:
                sku_to_page[sku] = idx
        head = _label_header_before_sku(text)
        slug = _norm_slug(head) if head else ""
        if slug and len(slug) >= 8:
            name_rows.append((idx, slug))
    return sku_to_page, name_rows


def _find_page_by_name(title: str, name_rows: list[tuple[int, str]]) -> int | None:
    slip_slug = _norm_slug(title)
    if len(slip_slug) < 8:
        return None
    exact: list[int] = []
    contains: list[int] = []
    for idx, slug in name_rows:
        if slip_slug == slug:
            exact.append(idx)
        elif len(slip_slug) >= 12 and (slip_slug in slug or slug in slip_slug):
            contains.append(idx)
    if exact:
        return exact[0]
    if contains:
        return contains[0]
    return None


def find_label_page_index(
    sku: str,
    title: str,
    sku_map: dict[str, int],
    name_rows: list[tuple[int, str]],
) -> tuple[int | None, str]:
    s = sku.upper()
    if s in sku_map:
        return sku_map[s], "sku"
    ni = _find_page_by_name(title, name_rows)
    if ni is not None:
        return ni, "name"
    return None, "none"


def merge_packing_with_sample_labels(
    slips_path: Path,
    labels_path: Path,
    out_path: Path,
    exclusions: list[str],
) -> list[str]:
    log: list[str] = []
    slips = fitz.open(str(slips_path))
    labels = fitz.open(str(labels_path))
    out = fitz.open()
    try:
        sku_map, name_rows = build_label_index(labels)
        log.append(f"Labels index: {len(sku_map)} SKUs, {len(name_rows)} name rows, {len(labels)} pages.")

        items = parse_smpl_items_from_packing_doc(slips)
        log.append(f"Packing slip: parsed {len(items)} sample lines (-smpl).")

        for si in range(len(slips)):
            out.insert_pdf(slips, from_page=si, to_page=si)

        appended = 0
        skipped_excl = 0
        missing = 0
        for it in items:
            sku = it["sku"]
            qty = it["qty"]
            title = it["title"]
            if is_excluded(sku, exclusions):
                log.append(f"SKIP (exclusion) smpl/{sku} qty={qty} title={title!r}")
                skipped_excl += 1
                continue
            li, how = find_label_page_index(sku, title, sku_map, name_rows)
            if li is None:
                log.append(f"MISS smpl/{sku} qty={qty} title={title!r} — no label page")
                missing += 1
                continue
            log.append(f"OK smpl/{sku} -> label page {li + 1} ({how}) x{qty} | {title!r}")
            for _ in range(qty):
                out.insert_pdf(labels, from_page=li, to_page=li)
                appended += 1

        out.save(str(out_path))
        log.append(
            f"--- Done: {len(slips)} packing pages + {appended} label pages "
            f"(excluded {skipped_excl}, missing {missing}) ---"
        )
        log.append(f"Labels start at PDF page {len(slips) + 1}")
        return log
    finally:
        out.close()
        labels.close()
        slips.close()
