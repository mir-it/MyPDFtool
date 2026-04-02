"""UI strings for pdf_splitter_app (English / Russian)."""

from __future__ import annotations

import json
from pathlib import Path

LANG_PATH = Path(__file__).resolve().parent / "data" / "ui_language.json"

UI: dict[str, dict[str, str]] = {
    "en": {
        "app_title": "MyPDF Tools",
        "language": "Language",
        "lang_en": "English",
        "lang_ru": "Русский",
        "tab_split": "Split Orders / Samples",
        "tab_merge": "Merge PDFs",
        "tab_logistics": "Packing Slips + Labels",
        "tab_tile": "Tile Club Samples (-smpl)",
        "tab_hd": "Home Depot THD (SMPL)",
        # Split
        "split_input_pdf": "Input PDF",
        "split_output_folder": "Output Folder",
        "browse": "Browse...",
        "split_legacy": "Use legacy overrides (for old approved test files)",
        "split_process": "Process PDF",
        # Merge
        "merge_input_pdfs": "Input PDFs (order matters)",
        "merge_add": "Add PDFs...",
        "merge_remove": "Remove Selected",
        "merge_up": "Move Up",
        "merge_down": "Move Down",
        "merge_clear": "Clear",
        "merge_output_file": "Output File",
        "merge_run": "Merge PDFs",
        # Logistics
        "logistics_help": (
            "Match: 8-digit ref on slip (after Ship Via / WN line, or \"PO #\" on return stub) "
            "<-> REF 1: on label. Output (1): matched packing pages in order, then their labels. "
            "Output (2) if needed: {your_output_name}_packing_slips_without_labels_YYYY-MM-DD.pdf "
            "next to the main file."
        ),
        "logistics_slips": "Packing slips PDF",
        "logistics_labels": "Shipping labels PDF",
        "logistics_out": "Output PDF",
        "logistics_run": "Merge packing slips + labels",
        # Tile Club tab
        "tile_heading": "PDF Merger Pro — Tile Club samples",
        "tile_desc": "Packing slip: *-smpl lines → find label sheet in catalog → append to end of PDF.",
        "tile_row1": "1. Packing slip PDF",
        "tile_row2": "2. Label catalog (New_4x5_TC_ALL…)",
        "tile_row3": "3. Output",
        "tile_merge": "MERGE",
        "tile_open_folder": "Open output folder",
        "tile_sku_title": "SKU exclusions",
        "tile_sku_hint": "If the code contains this substring — skip printing (e.g. tcmod → TCMODOLV258).",
        "tile_list_caption": "List (saved automatically)",
        # Home Depot tab
        "hd_heading": "PDF Merger Pro — Home Depot (THD)",
        "hd_desc": (
            "Packing slip: Model Number column (SMPL/ASMPL) + Qty Shipped; "
            "labels: Product Sku in THD Barcode Labels."
        ),
        "hd_row1": "1. Packing slip PDF (Home Depot)",
        "hd_row2": "2. THD Barcode Labels PDF",
        "hd_row3": "3. Output",
        "hd_merge": "MERGE",
        "hd_open_folder": "Open output folder",
        "hd_sku_title": "SKU exclusions",
        "hd_sku_hint": "If the code contains this substring — the line is skipped (separate list from Tile Club).",
        "hd_list_caption": "List (saved automatically)",
        # Dialogs
        "dlg_folder": "Folder",
        "dlg_error": "Error",
        "dlg_folder_need_out": "Specify an output file first.",
        "dlg_folder_missing": "Folder does not exist yet:\n{path}",
        "err_tile_slip": "Select a valid packing slip PDF.",
        "err_tile_labels": "Select a valid label catalog PDF.",
        "err_out_path": "Specify an output path.",
        "err_hd_slip": "Select a valid packing slip PDF.",
        "err_hd_labels": "Select a valid THD Barcode Labels PDF.",
        "err_split_pdf": "Select a valid input PDF file.",
        "err_split_dir": "Select a valid output folder.",
        "err_merge_count": "Add at least two PDF files to merge.",
        "err_merge_out": "Select an output PDF file.",
        "err_log_slips": "Select a valid packing slips PDF.",
        "err_log_labels": "Select a valid labels PDF.",
        "err_log_out": "Select an output PDF path.",
        # Status semantics
        "st_ready": "Ready",
        "st_processing": "Processing...",
        "st_processing_ellipsis": "Processing…",
        "st_merging": "Merging...",
        "st_completed": "Completed",
        "st_failed": "Failed",
        "st_error": "Error",
        # Logistics dynamic status
        "logistics_done": "Completed — matched: {n} slips + {n_lbl} labels",
        "logistics_done_orphans": "; orphans -> {name}",
        # PDF file dialog
        "ft_pdf": "PDF files",
        # Split result (template)
        "split_result": (
            "Done.\n\n"
            "Classification (unchanged): samples vs orders by existing rules.\n"
            "Output page order: chronological by slip Date: (Home Depot header) within each file.\n\n"
            "Samples pages ({ns}): {samples}\n"
            "Orders pages ({no}): {orders}\n\n"
            "Samples date order:\n{slog}\n\n"
            "Orders date order:\n{olog}\n\n"
            "Output files:\n"
            "- {sout}\n"
            "- {oout}"
        ),
        "split_no_samples": "(no samples pages)",
        "split_no_orders": "(no orders pages)",
        # Merge result
        "merge_result": (
            "Done.\n\nMerged {n} files into:\n{out}\n\n"
            "Chronological order (by Date: on page 1 of each PDF):\n"
            "{sort_log}\n\nOriginal listbox order was:\n{orig}"
        ),
        "merge_orig_line": "- {path}",
        "log_file_label": "File:",
        "log_main_output": "Main output:",
        "log_unmatched": "Unmatched slips only:",
    },
    "ru": {
        "app_title": "MyPDF Tools",
        "language": "Язык",
        "lang_en": "English",
        "lang_ru": "Русский",
        "tab_split": "Разделение: заказы / образцы",
        "tab_merge": "Объединение PDF",
        "tab_logistics": "Упаковочные листы + этикетки",
        "tab_tile": "Tile Club Samples (-smpl)",
        "tab_hd": "Home Depot THD (SMPL)",
        "split_input_pdf": "Входной PDF",
        "split_output_folder": "Папка вывода",
        "browse": "Обзор…",
        "split_legacy": "Устаревшие переопределения (для старых тестовых PDF)",
        "split_process": "Обработать PDF",
        "merge_input_pdfs": "Входные PDF (порядок важен)",
        "merge_add": "Добавить PDF…",
        "merge_remove": "Удалить выбранные",
        "merge_up": "Вверх",
        "merge_down": "Вниз",
        "merge_clear": "Очистить",
        "merge_output_file": "Файл результата",
        "merge_run": "Объединить PDF",
        "logistics_help": (
            "Сопоставление: 8-значный ref на листе (после Ship Via / WN или «PO #» на отрывной части) "
            "<-> REF 1: на этикетке. Выход (1): совпавшие страницы листов по порядку, затем этикетки. "
            "Выход (2) при необходимости: {your_output_name}_packing_slips_without_labels_YYYY-MM-DD.pdf "
            "рядом с основным файлом."
        ),
        "logistics_slips": "PDF упаковочных листов",
        "logistics_labels": "PDF этикеток доставки",
        "logistics_out": "Выходной PDF",
        "logistics_run": "Объединить листы и этикетки",
        "tile_heading": "PDF Merger Pro — образцы Tile Club",
        "tile_desc": "Упаковочный лист: строки *-smpl → поиск листа в каталоге этикеток → в конец PDF.",
        "tile_row1": "1. PDF упаковочного листа",
        "tile_row2": "2. Каталог этикеток (New_4x5_TC_ALL…)",
        "tile_row3": "3. Результат",
        "tile_merge": "ОБЪЕДИНИТЬ",
        "tile_open_folder": "Открыть папку результата",
        "tile_sku_title": "Исключения SKU",
        "tile_sku_hint": "Если код содержит подстроку — не печатаем (напр. tcmod → TCMODOLV258).",
        "tile_list_caption": "Список (сохраняется автоматически)",
        "hd_heading": "PDF Merger Pro — Home Depot (THD)",
        "hd_desc": (
            "Упаковочный лист: колонка Model Number (SMPL/ASMPL) + Qty Shipped; "
            "этикетки: Product Sku в THD Barcode Labels."
        ),
        "hd_row1": "1. PDF упаковочного листа (Home Depot)",
        "hd_row2": "2. PDF THD Barcode Labels",
        "hd_row3": "3. Результат",
        "hd_merge": "ОБЪЕДИНИТЬ",
        "hd_open_folder": "Открыть папку результата",
        "hd_sku_title": "Исключения SKU",
        "hd_sku_hint": "Если код содержит подстроку — строка пропускается (отдельный список от Tile Club).",
        "hd_list_caption": "Список (сохраняется автоматически)",
        "dlg_folder": "Папка",
        "dlg_error": "Ошибка",
        "dlg_folder_need_out": "Сначала укажите файл результата.",
        "dlg_folder_missing": "Папка ещё не существует:\n{path}",
        "err_tile_slip": "Укажите корректный PDF упаковочного листа.",
        "err_tile_labels": "Укажите корректный PDF каталога этикеток.",
        "err_out_path": "Укажите путь для результата.",
        "err_hd_slip": "Укажите корректный PDF упаковочного листа.",
        "err_hd_labels": "Укажите корректный PDF THD Barcode Labels.",
        "err_split_pdf": "Выберите корректный входной PDF.",
        "err_split_dir": "Выберите корректную папку вывода.",
        "err_merge_count": "Добавьте не менее двух PDF для объединения.",
        "err_merge_out": "Укажите выходной PDF.",
        "err_log_slips": "Выберите корректный PDF упаковочных листов.",
        "err_log_labels": "Выберите корректный PDF этикеток.",
        "err_log_out": "Укажите путь выходного PDF.",
        "st_ready": "Готово",
        "st_processing": "Обработка…",
        "st_processing_ellipsis": "Обработка…",
        "st_merging": "Объединение…",
        "st_completed": "Готово",
        "st_failed": "Ошибка",
        "st_error": "Ошибка",
        "logistics_done": "Готово — совпало: {n} листов + {n_lbl} этикеток",
        "logistics_done_orphans": "; без пары → {name}",
        "ft_pdf": "PDF-файлы",
        "split_result": (
            "Готово.\n\n"
            "Классификация (без изменений): образцы и заказы по существующим правилам.\n"
            "Порядок страниц: по дате Date: в шапке Home Depot внутри каждого файла.\n\n"
            "Страницы образцов ({ns}): {samples}\n"
            "Страницы заказов ({no}): {orders}\n\n"
            "Порядок дат (образцы):\n{slog}\n\n"
            "Порядок дат (заказы):\n{olog}\n\n"
            "Выходные файлы:\n"
            "- {sout}\n"
            "- {oout}"
        ),
        "split_no_samples": "(нет страниц образцов)",
        "split_no_orders": "(нет страниц заказов)",
        "merge_result": (
            "Готово.\n\nОбъединено {n} файлов в:\n{out}\n\n"
            "Хронологический порядок (по Date: на стр. 1 каждого PDF):\n"
            "{sort_log}\n\nИсходный порядок в списке:\n{orig}"
        ),
        "merge_orig_line": "- {path}",
        "log_file_label": "Файл:",
        "log_main_output": "Основной выход:",
        "log_unmatched": "Только листы без пары:",
    },
}


def load_lang() -> str:
    try:
        if LANG_PATH.exists():
            data = json.loads(LANG_PATH.read_text(encoding="utf-8"))
            lang = data.get("lang", "en")
            if lang in UI:
                return lang
    except Exception:
        pass
    return "en"


def save_lang(lang: str) -> None:
    if lang not in UI:
        return
    try:
        LANG_PATH.parent.mkdir(parents=True, exist_ok=True)
        LANG_PATH.write_text(json.dumps({"lang": lang}, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def t(lang: str, key: str) -> str:
    return UI.get(lang, UI["en"]).get(key, UI["en"].get(key, key))
