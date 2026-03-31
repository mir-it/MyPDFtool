"""
Home Depot packing slip (Model Number + sample tail) -> THD barcode label PDF (Product Sku).

Only this module is edited for the Home Depot THD tab; other app tabs stay unchanged.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import fitz  # PyMuPDF

# ---------------------------------------------------------------------------
# THD matching rules — edit independently; disable by setting to None / empty.
# ---------------------------------------------------------------------------

# If slip SKU (after glue + suffix strip) casefold-startswith "mod", search labels
# using only the first N characters. Set to None to disable.
MOD_SKU_LOOKUP_PREFIX_LEN: int | None = 10

# Exact slip-key -> label SKU (keys are compared on "compact" uppercase: no spaces).
# Add rows here for systematic rewrites (easy to maintain in one place).
SLIP_TO_LABEL_SKU: dict[str, str] = {
    "MODDECRED258": "MOD88DERE258",
    "MODDECWHT258": "MOD88DEWH258",
    "MODDECSKY258": "MOD88DESK258",
    "MODDECPNK258": "MOD88DEPN258",
    "MODDECOLV258": "MOD88DEOL258",
    "MOD88DECOLV44": "MOD88DECOL44",
    "MOD88DECPNK44": "MOD88DEPN44",
    "MOD88DECRED44": "MOD88DERE44",
    "MOD88DECSKY44": "MOD88DESK44",
    "MOD88DEWHT44": "MOD88DEWH44",
    # slip often has extra "C" vs label Product Sku (see MOD88DEWH44 on label)
    "MOD88DECWHT44": "MOD88DEWH44",
}

# Optional: thd_sku_overrides.txt — user-editable slip fragment -> label search string (see load_thd_sku_overrides).
# APLA880… prefix list was removed; use that file for APLA… patterns like A881434X16=APLA88143 4X16A.

# When slip code is APLEC61..65 (compact), search label *page text* for these phrases.
APLEC_LABEL_TEXT_HINT: dict[str, str] = {
    "APLEC61": "La Riviera Blanc 2.5x8",
    "APLEC62": "La Riviera Lavanda Blue 2.5x8",
    "APLEC63": "La Riviera Quetzal 2.5x8",
    "APLEC64": "La Riviera Blue Reef 2.5x8",
    "APLEC65": "La Riviera Rose 2.5x8",
}

# ---------------------------------------------------------------------------


def _thd_override_search_paths(slips_path: Path | str | None = None) -> list[Path]:
    """
    Search order (later files win on duplicate KEY):
    - MyAI/thd_sku_overrides.txt
    - MyAI/data/thd_sku_overrides.txt, or if missing, MyAI/data/thd_sku_overrides.example.txt
    - cwd/thd_sku_overrides.txt
    - folder of the packing PDF / thd_sku_overrides.txt (so rules can sit next to Downloads/…pdf)
    """
    here = Path(__file__).resolve().parent
    paths: list[Path] = [here / "thd_sku_overrides.txt"]
    _data_txt = here / "data" / "thd_sku_overrides.txt"
    _data_ex = here / "data" / "thd_sku_overrides.example.txt"
    if _data_txt.is_file():
        paths.append(_data_txt)
    elif _data_ex.is_file():
        paths.append(_data_ex)
    paths.append(Path.cwd() / "thd_sku_overrides.txt")
    if slips_path is not None:
        try:
            parent = Path(slips_path).expanduser().resolve().parent
            paths.append(parent / "thd_sku_overrides.txt")
        except OSError:
            pass
    seen: set[Path] = set()
    out: list[Path] = []
    for p in paths:
        try:
            rp = p.resolve()
        except OSError:
            rp = p
        if rp in seen:
            continue
        seen.add(rp)
        out.append(p)
    return out


def load_thd_sku_overrides(
    slips_path: Path | str | None = None,
) -> tuple[dict[str, str], list[Path]]:
    """
    Lines: KEY=VALUE
    - KEY is matched against compact slip SKU (after sample suffix strip + built-in SLIP_TO_LABEL_SKU).
    - VALUE is the label search string as printed on Product Sku (spaces allowed).
    - # starts a comment; empty lines ignored.
    Later files in the search list override earlier keys if the same KEY appears.
    """
    merged: dict[str, str] = {}
    loaded: list[Path] = []
    for p in _thd_override_search_paths(slips_path):
        if not p.is_file():
            continue
        loaded.append(p)
        for line in p.read_text(encoding="utf-8", errors="replace").splitlines():
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            if "=" not in s:
                continue
            k, v = s.split("=", 1)
            kc = _compact_sku(k)
            vv = v.strip()
            if kc and vv:
                merged[kc] = vv
    return merged, loaded


def _apply_optional_file_override(compact_or_raw: str, ov: dict[str, str] | None) -> str:
    """Replace slip compact key using overrides file; VALUE may contain spaces."""
    if not ov:
        return compact_or_raw
    c = _compact_sku(compact_or_raw)
    if c in ov:
        return ov[c].strip()
    for k in sorted(ov.keys(), key=len, reverse=True):
        if len(k) < 5:
            continue
        if k in c:
            return ov[k].strip()
    return compact_or_raw


def dedupe_consecutive_hd_items(items: list[dict]) -> list[dict]:
    """Drop consecutive duplicate rows (same SKU after suffix strip) from PDF text quirks."""
    out: list[dict] = []
    for it in items:
        c = _compact_sku(_strip_hd_sku_suffixes_looped(it["sku"]))
        if out:
            prev = _compact_sku(_strip_hd_sku_suffixes_looped(out[-1]["sku"]))
            if prev == c:
                continue
        out.append(it)
    return out


# Capture full Product Sku value (may contain spaces, e.g. "LR Rose 5x5").
_SKU_FIELD = re.compile(r"Product\s+Sku:\s*(.+)", re.IGNORECASE)
_COLLECTION = re.compile(r"Collection\s+Name:\s*([^\n]+)", re.IGNORECASE)
_COLOR = re.compile(r"^Color:\s*([^\n]+)", re.IGNORECASE | re.MULTILINE)


def default_hd_exclusions_path() -> Path:
    base = Path.home() / ".tileclub_sample_merger"
    base.mkdir(parents=True, exist_ok=True)
    return base / "home_depot_exclusions.json"


def load_hd_exclusions(path: Path | None = None) -> list[str]:
    p = path or default_hd_exclusions_path()
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [str(x).strip() for x in data if str(x).strip()]
    except (json.JSONDecodeError, OSError):
        pass
    return []


def save_hd_exclusions(items: list[str], path: Path | None = None) -> None:
    p = path or default_hd_exclusions_path()
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


def _compact_sku(s: str) -> str:
    return re.sub(r"\s+", "", s.strip().upper())


def _spaced_sku(s: str) -> str:
    return " ".join(s.split()).strip().upper()


def _is_sample_marker_line(line: str) -> bool:
    s = line.strip().upper()
    if not s:
        return False
    if s in ("SMPL", "ASMPL", "TASMP", "ASMP"):
        return True
    if "ASMPL" in s:
        return True
    if len(s) <= 14 and "SMPL" in s:
        return True
    if s.endswith("SMPL") and len(s) <= 14:
        return True
    return False


def _looks_like_model_line(line: str) -> bool:
    s = line.strip()
    if not s or s.isdigit():
        return False
    if _is_sample_marker_line(s):
        return False
    low = s.casefold()
    if low in {
        "model number",
        "internet number",
        "item description",
        "qty shipped",
        "qty returned",
        "return code",
    }:
        return False
    if len(s) > 48:
        return False
    if re.fullmatch(r"\d{6,}", s):
        return False
    # Classic models: ORB8801..., PS99HX01...
    if bool(re.fullmatch(r"[A-Z0-9][A-Z0-9\-]*", s, re.IGNORECASE)):
        return bool(any(ch.isdigit() for ch in s))
    # Names with spaces: "LR Rose" (no digits on this line) or "LR Rose 5x5"
    if bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\s.\-xX]*", s)):
        if re.search(r"[A-Za-z]", s):
            if any(ch.isdigit() for ch in s):
                return True
            # "LR Rose" + next line carries "5x5SMPL"
            return len(s) <= 28 and len(s.split()) >= 2
    return False


def _looks_like_sku_tail_fragment(line: str) -> bool:
    t = line.strip().upper()
    if not t:
        return False
    if _is_sample_marker_line(t):
        return True
    if t.isdigit():
        return False
    return bool(
        re.fullmatch(
            r"[A-Z0-9\-xX]*?(SAMPLE|SMPL|ASMPL|ASMP|TASMP|SMP|SM|S)$",
            t,
        )
    )


def _looks_like_split_sample_tail(prev_line: str, next_line: str) -> bool:
    """Handle OCR breaks like MOD88CEL44S + MPL, PS99HX01SMP + L."""
    p = prev_line.strip().upper()
    n = next_line.strip().upper()
    if n == "L" and p.endswith("SMP"):
        return True
    if n == "MPL" and p.endswith("S"):
        return True
    if n == "PL" and p.endswith("SM"):
        return True
    return False


def _glue_model_and_tail(first: str, tail: str) -> str:
    """Preserve spaces inside the first line (e.g. LR Rose 5x5)."""
    return (first.rstrip() + tail.strip()).strip()


def _strip_hd_sku_suffixes_looped(full_sku: str) -> str:
    """Remove trailing sample markers repeatedly (compact form, no spaces)."""
    s = _compact_sku(full_sku)
    if not s:
        return s
    while True:
        old = s
        for suf in ("SAMPLE", "SMPL", "SMP", "SM", "S"):
            if s.endswith(suf) and len(s) > len(suf):
                s = s[: -len(suf)]
                break
        if s == old:
            break
    return s


def _clean_slip_description(desc: str) -> str:
    s = " ".join(desc.split())
    s = re.sub(r"\bSample\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\bSMPL\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\bTASMP\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _is_internet_number_line(line: str) -> bool:
    s = line.strip()
    return s.isdigit() and len(s) >= 6


def _is_qty_line(line: str) -> bool:
    s = line.strip()
    if not s.isdigit():
        return False
    n = int(s)
    return 0 < n <= 9999 and len(s) <= 4


def _looks_like_classic_model_single_line(line: str) -> bool:
    """Tight match for Home Depot style codes (no spaces); avoids prose in Item Description."""
    s = line.strip()
    if not s or len(s) > 28:
        return False
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9\-]*", s, re.IGNORECASE):
        return False
    return any(ch.isdigit() for ch in s)


def _starts_following_item_block(lj: str, lines: list[str], j: int, n: int) -> bool:
    """
    PDF sometimes merges the next row into Item Description. Stop before a line that
    clearly begins another Model Number row (same signals as item start, or model then Internet #).
    """
    if not _looks_like_model_line(lj):
        return False
    if j + 1 >= n:
        return False
    nxt = lines[j + 1].strip()
    if _looks_like_sku_tail_fragment(nxt):
        return True
    if _looks_like_split_sample_tail(lj, nxt):
        return True
    # Next line is Internet # with no SMPL line between (text extract quirk).
    if _is_internet_number_line(nxt) and _looks_like_classic_model_single_line(lj):
        return True
    return False


def parse_hd_sample_items_from_packing_doc(slips_doc: fitz.Document) -> list[dict]:
    chunks = [(slips_doc[i].get_text("text") or "") for i in range(len(slips_doc))]
    full = "\n".join(chunks)
    lines = full.splitlines()
    items: list[dict] = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        started = False
        model_glued = ""
        j = i

        # Same-line sample suffix, then Internet # on the next line:
        #   APLEC75SAMPLE
        #   322715629
        em = re.match(
            r"^(.+?)(SAMPLE|SMPL|ASMPL|ASMP|TASMP|SMP|SM)$",
            line,
            flags=re.IGNORECASE,
        )
        if (
            em
            and _looks_like_model_line(em.group(1).strip())
            and i + 1 < n
            and _is_internet_number_line(lines[i + 1].strip())
        ):
            model_glued = em.group(1).strip() + em.group(2).strip()
            j = i + 1
            started = True

        if not started and _looks_like_model_line(line) and i + 1 < n:
            n1 = lines[i + 1].strip()
            if _looks_like_sku_tail_fragment(n1):
                model_glued = _glue_model_and_tail(line, n1)
                j = i + 2
                started = True
            elif i + 1 < n and _looks_like_split_sample_tail(line, n1):
                model_glued = _glue_model_and_tail(line, n1)
                j = i + 2
                started = True

        if not started:
            i += 1
            continue

        if j < n and _is_internet_number_line(lines[j]):
            j += 1

        desc_parts: list[str] = []
        qty = 1
        while j < n:
            lj = lines[j].strip()
            if not lj:
                j += 1
                continue
            next_ok = j + 1 < n and _looks_like_sku_tail_fragment(lines[j + 1])
            next_split = j + 1 < n and _looks_like_split_sample_tail(lj, lines[j + 1])
            if _looks_like_model_line(lj) and (next_ok or next_split):
                break
            if _starts_following_item_block(lj, lines, j, n):
                break
            if lj == "Model Number":
                break
            if lj.startswith("Thank you for shopping"):
                break
            if lj.startswith("Page:"):
                break
            if _is_qty_line(lj):
                qty = max(1, int(lj))
                j += 1
                break
            desc_parts.append(lj)
            j += 1
        desc = _clean_slip_description(" ".join(desc_parts))
        items.append({"sku": model_glued, "qty": qty, "title": desc})
        i = j
    return items


def _register_sku_aliases(m: dict[str, int], sku_raw: str, idx: int) -> None:
    if not sku_raw.strip():
        return
    variants: set[str] = set()
    s0 = sku_raw.strip().upper()
    variants.add(s0)
    variants.add(_spaced_sku(s0))
    variants.add(_compact_sku(s0))
    for v in list(variants):
        if v.endswith("A") and len(v) > 2:
            variants.add(v[:-1])
    for v in variants:
        if v and v not in m:
            m[v] = idx


def _thd_spec_slug(text: str) -> str:
    parts: list[str] = []
    cm = _COLLECTION.search(text)
    if cm:
        parts.append(cm.group(1).strip())
    colm = _COLOR.search(text)
    if colm:
        parts.append(colm.group(1).strip())
    return _norm_slug(" ".join(parts))


def build_thd_label_index(
    labels_doc: fitz.Document,
) -> tuple[dict[str, int], list[tuple[int, str]], dict[str, int], list[str]]:
    """
    sku_to_page: many keys per label Product Sku (spaced / compact / no trailing A)
    name_rows: (page_idx, collection+color slug)
    text_hint_to_page: first page index containing substring (lowercased hint)
    raw_pages_lower: per-page lowercased text for hint search
    """
    sku_to_page: dict[str, int] = {}
    name_rows: list[tuple[int, str]] = []
    text_hint_to_page: dict[str, int] = {}
    raw_pages_lower: list[str] = []

    for idx in range(len(labels_doc)):
        text = labels_doc[idx].get_text("text") or ""
        low = text.casefold()
        raw_pages_lower.append(low)

        sm = _SKU_FIELD.search(text)
        if sm:
            _register_sku_aliases(sku_to_page, sm.group(1).strip(), idx)

        slug = _thd_spec_slug(text)
        if slug and len(slug) >= 6:
            name_rows.append((idx, slug))

    for k, hint in APLEC_LABEL_TEXT_HINT.items():
        h = hint.casefold()
        if h in text_hint_to_page:
            continue
        for pi, pl in enumerate(raw_pages_lower):
            if h in pl:
                text_hint_to_page[k] = pi
                break

    return sku_to_page, name_rows, text_hint_to_page, raw_pages_lower


def _find_sku_prefix_extension(
    slip_sku: str, sku_map: dict[str, int]
) -> tuple[int, str] | None:
    s = slip_sku.upper().replace(" ", "")
    if len(s) < 10:
        return None
    cand: list[tuple[str, int]] = []
    for k, page_idx in sku_map.items():
        ku = k.upper()
        if ku.startswith(s) and len(ku) > len(s):
            extra = len(ku) - len(s)
            if 1 <= extra <= 5:
                cand.append((ku, page_idx))
    if not cand:
        return None
    cand.sort(key=lambda x: (len(x[0]) - len(s), x[0]))
    pages = {p for _, p in cand}
    if len(pages) != 1:
        return None
    k, page_idx = cand[0]
    return page_idx, f"sku+ext({k})"


def _find_page_by_name(title: str, name_rows: list[tuple[int, str]]) -> int | None:
    slip_slug = _norm_slug(title)
    if len(slip_slug) < 6:
        return None
    exact: list[int] = []
    contains: list[int] = []
    for idx, slug in name_rows:
        if slip_slug == slug:
            exact.append(idx)
        elif len(slip_slug) >= 10 and (slip_slug in slug or slug in slip_slug):
            contains.append(idx)
    if exact:
        return exact[0]
    if contains:
        return contains[0]
    return None


def _apply_mod_prefix_lookup(compact: str) -> list[str]:
    if MOD_SKU_LOOKUP_PREFIX_LEN is None:
        return []
    if not compact.casefold().startswith("mod"):
        return []
    n = MOD_SKU_LOOKUP_PREFIX_LEN
    if len(compact) <= n:
        return []
    return [compact[:n]]


def _apply_slip_alias(compact: str) -> str:
    return SLIP_TO_LABEL_SKU.get(compact, compact)


def _apla_spaced_search_keys(compact: str) -> list[str]:
    out: list[str] = []
    cup = compact.upper()
    # Generic APLA88… codes like APLA881434x16A ->
    #   APLA88 + (88143) + (4)x(16)  =>  "APLA88143 4x16A"
    # Implemented as: APLA88 + digits + single_digit + x + height + optional A.
    if cup.startswith("APLA88"):
        m = re.match(r"^APLA88(\d+)(\d)X(\d{1,3})(A)?$", cup, flags=re.IGNORECASE)
        if m:
            prefix = f"APLA88{m.group(1)}".upper()
            dim = f"{m.group(2)}x{m.group(3)}".upper()
            spaced = f"{prefix} {dim}".upper()
            # Labels commonly store the trailing "A" (e.g. "APLA88143 4x16A").
            if m.group(4):
                sa = f"{spaced}{m.group(4).upper()}"
            else:
                sa = f"{spaced}A"
            out.append(sa)
            out.append(_compact_sku(sa))
            out.append(spaced)
            out.append(_compact_sku(spaced))
    return out


def _lookup_in_map(candidates: list[str], sku_map: dict[str, int]) -> tuple[int, str] | None:
    seen: set[str] = set()
    ordered: list[str] = []
    for c in candidates:
        u = c.strip().upper()
        variants = {u, _compact_sku(u), _spaced_sku(u)}
        for v in variants:
            if v and v not in seen:
                seen.add(v)
                ordered.append(v)

    for cand in ordered:
        if cand in sku_map:
            return sku_map[cand], f"hit({cand})"
        if not cand.endswith("A") and f"{cand}A" in sku_map:
            return sku_map[f"{cand}A"], f"hit({cand}+A)"
        if cand.endswith("A") and len(cand) > 2 and cand[:-1] in sku_map:
            return sku_map[cand[:-1]], f"hit({cand}-A)"
    return None


def find_thd_label_page_index(
    sku: str,
    title: str,
    sku_map: dict[str, int],
    name_rows: list[tuple[int, str]],
    text_hint_to_page: dict[str, int],
    file_overrides: dict[str, str] | None = None,
) -> tuple[int | None, str]:
    raw = sku.strip()
    base_compact = _strip_hd_sku_suffixes_looped(raw)
    base_compact = _apply_slip_alias(base_compact)
    base_compact = _apply_optional_file_override(base_compact, file_overrides)
    cup = _compact_sku(base_compact)

    candidates: list[str] = []
    # Override / spaced label string first, then heuristic APLA88… split, then raw compact.
    candidates.append(base_compact)
    candidates.append(cup)
    candidates.extend(_apla_spaced_search_keys(cup))
    candidates.append(_compact_sku(raw))
    candidates.extend(_apply_mod_prefix_lookup(cup))
    if " " in raw.strip():
        candidates.append(_strip_hd_sku_suffixes_looped(_compact_sku(raw)))

    hit = _lookup_in_map(candidates, sku_map)
    if hit is not None:
        return hit

    ckey = _compact_sku(base_compact)
    if ckey in APLEC_LABEL_TEXT_HINT and ckey in text_hint_to_page:
        return text_hint_to_page[ckey], f"aplec-text({APLEC_LABEL_TEXT_HINT[ckey]})"

    ext = _find_sku_prefix_extension(base_compact, sku_map)
    if ext is not None:
        page_idx, tag = ext
        return page_idx, tag

    ni = _find_page_by_name(title, name_rows)
    if ni is not None:
        return ni, "name"

    for pl, hint in APLEC_LABEL_TEXT_HINT.items():
        if ckey.startswith(pl):
            if pl in text_hint_to_page:
                return text_hint_to_page[pl], f"aplec-prefix-text({hint})"

    return None, "none"


def merge_hd_packing_with_thd_labels(
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
        sku_map, name_rows, text_hints, _raw_low = build_thd_label_index(labels)
        log.append(f"Labels index: {len(sku_map)} SKU keys, {len(name_rows)} name rows, {len(labels)} pages.")

        ov, ov_paths = load_thd_sku_overrides(slips_path)
        if ov_paths:
            log.append(
                f"THD thd_sku_overrides.txt: {len(ov)} entr(y/ies) from "
                + ", ".join(str(p) for p in ov_paths)
            )

        items = parse_hd_sample_items_from_packing_doc(slips)
        n_before = len(items)
        items = dedupe_consecutive_hd_items(items)
        if len(items) != n_before:
            log.append(f"Deduped consecutive duplicate SKU rows: {n_before} -> {len(items)}")
        log.append(f"Packing slip: parsed {len(items)} HD sample rows (Model + sample tail).")

        for si in range(len(slips)):
            out.insert_pdf(slips, from_page=si, to_page=si)

        appended = 0
        skipped_excl = 0
        missing = 0
        for it in items:
            sku = it["sku"]
            qty = it["qty"]
            title = it["title"]
            excl_key = _compact_sku(sku)
            if is_excluded(excl_key, exclusions):
                log.append(f"SKIP (exclusion) {sku} qty={qty} title={title!r}")
                skipped_excl += 1
                continue
            li, how = find_thd_label_page_index(
                sku, title, sku_map, name_rows, text_hints, file_overrides=ov
            )
            if li is None:
                log.append(f"MISS {sku} qty={qty} title={title!r} — no label page")
                missing += 1
                continue
            log.append(f"OK {sku} -> label page {li + 1} ({how}) x{qty} | {title!r}")
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
