import re
import io
import time
import os
import tempfile
from copy import copy as _copy
from collections import defaultdict
from typing import Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string


# =========================
# Helpers
# =========================
def safe_sheet_name(name: str) -> str:
    """Excel: max 31 chars, cannot contain : \\ / ? * [ ]"""
    banned = [":", "\\", "/", "?", "*", "[", "]"]
    for ch in banned:
        name = name.replace(ch, "")
    name = (name or "").strip() or "лист"
    return name[:31]


def safe_filename(name: str) -> str:
    banned = ["<", ">", ":", "\"", "/", "\\", "|", "?", "*"]
    for ch in banned:
        name = (name or "").replace(ch, "")
    name = (name or "").strip() or "output.xlsx"
    return name


def split_prefix_suffix4(sheet_name: str) -> Tuple[str, str]:
    """prefix + last 4 chars (lowercased)"""
    if len(sheet_name) < 4:
        return sheet_name, ""
    return sheet_name[:-4], sheet_name[-4:].lower()


def split_prefix_suffix2(sheet_name: str) -> Tuple[str, str]:
    """prefix + last 2 chars (lowercased)"""
    if len(sheet_name) < 2:
        return sheet_name, ""
    return sheet_name[:-2], sheet_name[-2:].lower()


def normalize_prefix(prefix: str) -> str:
    return (prefix or "").strip()


def make_unique_sheet_title(wb, desired_title: str) -> str:
    base = safe_sheet_name(desired_title)
    if base not in wb.sheetnames:
        return base
    i = 1
    while True:
        suffix = f" ({i})"
        trimmed = base
        if len(trimmed) + len(suffix) > 31:
            trimmed = trimmed[: 31 - len(suffix)]
        candidate = f"{trimmed}{suffix}"
        if candidate not in wb.sheetnames:
            return candidate
        i += 1


def make_unique_with_fixed_suffix(wb, prefix: str, suffix: str) -> str:
    """
    Returns a unique title that always ends with `suffix` (e.g. '1210' or 'wd').
    If `prefix+suffix` exists, inserts a counter before the suffix: f'{prefix}{i}{suffix}'.
    """
    prefix = (prefix or "").strip()
    suffix = (suffix or "").strip()
    base = safe_sheet_name(prefix + suffix)
    if base.endswith(suffix) and base not in wb.sheetnames:
        return base

    i = 1
    while True:
        mid = str(i)
        # ensure total length <= 31 and ends with suffix
        max_prefix_len = 31 - len(mid) - len(suffix)
        p = prefix[:max_prefix_len] if max_prefix_len > 0 else ""
        candidate = safe_sheet_name(f"{p}{mid}{suffix}")
        if candidate.endswith(suffix) and candidate not in wb.sheetnames:
            return candidate
        i += 1


def load_wb_from_bytes(file_bytes: bytes, filename: str):
    ext = (filename or "").lower().rsplit(".", 1)[-1]
    keep_vba = ext == "xlsm"
    return load_workbook(io.BytesIO(file_bytes), keep_vba=keep_vba), keep_vba


def is_xls_filename(filename: str) -> bool:
    return (filename or "").lower().endswith(".xls") and not (filename or "").lower().endswith(".xlsx")


def convert_xls_to_xlsx_via_excel(xls_bytes: bytes, original_name: str) -> Tuple[str, bytes]:
    """
    Converts .xls bytes to .xlsx bytes using installed Microsoft Excel (COM).
    Works only on Windows with Excel installed and pywin32 available.
    """
    if os.name != "nt":
        raise RuntimeError("Конвертация .xls поддерживается только на Windows с установленным Excel.")

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        raise RuntimeError("Для .xls нужно установить pywin32: pip install pywin32")

    pythoncom.CoInitialize()
    excel = None
    xls_path = None
    xlsx_path = None
    try:
        fd_xls, xls_path = tempfile.mkstemp(suffix=".xls")
        os.close(fd_xls)
        with open(xls_path, "wb") as f:
            f.write(xls_bytes)

        fd_xlsx, xlsx_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd_xlsx)

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(xls_path, ReadOnly=True)
        try:
            # 51 = xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(xlsx_path, FileFormat=51)
        finally:
            wb.Close(SaveChanges=False)
        excel.Quit()
        excel = None

        with open(xlsx_path, "rb") as f:
            out_bytes = f.read()

        base = (original_name or "input.xls")
        base = base[:-4] if base.lower().endswith(".xls") else base
        out_name = base + ".xlsx"
        return out_name, out_bytes
    finally:
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        for p in (xls_path, xlsx_path):
            if p and os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


def ensure_openpyxl_bytes(filename: str, file_bytes: bytes, cache: Optional[Dict[Tuple[str, int], Tuple[str, bytes]]] = None) -> Tuple[str, bytes]:
    """
    Ensures we return bytes loadable by openpyxl (xlsx/xlsm). Converts xls -> xlsx if needed.
    `cache` is keyed by (filename, len(bytes)).
    """
    if not is_xls_filename(filename):
        return filename, file_bytes
    key = (filename, len(file_bytes))
    if cache is not None and key in cache:
        return cache[key]
    out = convert_xls_to_xlsx_via_excel(file_bytes, filename)
    if cache is not None:
        cache[key] = out
    return out


def copy_sheet(src_ws, dst_wb, new_title: str):
    dst_ws = dst_wb.create_sheet(title=new_title)
    dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
    dst_ws.freeze_panes = src_ws.freeze_panes

    for key, dim in src_ws.column_dimensions.items():
        dst = dst_ws.column_dimensions[key]
        dst.width = dim.width
        dst.hidden = dim.hidden
        dst.outlineLevel = dim.outlineLevel

    for idx, dim in src_ws.row_dimensions.items():
        dst = dst_ws.row_dimensions[idx]
        dst.height = dim.height
        dst.hidden = dim.hidden
        dst.outlineLevel = dim.outlineLevel

    for merged in list(src_ws.merged_cells.ranges):
        dst_ws.merge_cells(str(merged))

    # Fast path: copy only instantiated cells (values/styles), avoids scanning huge blank rectangles.
    for (_r, _c), cell in getattr(src_ws, "_cells", {}).items():
        if isinstance(cell, MergedCell):
            continue
        col_idx = cell.column if isinstance(cell.column, int) else column_index_from_string(str(cell.column))
        dst_cell = dst_ws.cell(row=cell.row, column=col_idx, value=cell.value)
        if cell.has_style:
            dst_cell.font = _copy(cell.font)
            dst_cell.border = _copy(cell.border)
            dst_cell.fill = _copy(cell.fill)
            dst_cell.number_format = cell.number_format
            dst_cell.protection = _copy(cell.protection)
            dst_cell.alignment = _copy(cell.alignment)
    return dst_ws


def merge_wh_m_into_analysis(analysis_wb, wh_bytes: Optional[bytes], wh_name: str, m_bytes: Optional[bytes], m_name: str) -> Dict[str, List[str]]:
    report = {"missing_wh": [], "missing_m": [], "copied": []}

    if m_bytes:
        m_wb, _ = load_wb_from_bytes(m_bytes, m_name)
        mapping_m = {
            "Итог": "M",
            "Реестр": "Mr",
            "Контрагенты": "Mt",
            "Договоры": "Md",
        }
        for src_title, dst_title in mapping_m.items():
            if src_title not in m_wb.sheetnames:
                report["missing_m"].append(src_title)
                continue
            new_title = make_unique_sheet_title(analysis_wb, dst_title)
            copy_sheet(m_wb[src_title], analysis_wb, new_title)
            report["copied"].append(f"{m_name}:{src_title} -> {new_title}")

    if wh_bytes:
        wh_wb, _ = load_wb_from_bytes(wh_bytes, wh_name)
        mapping_wh = {
            "Итог": "W",
            "Реестр": "Wr",
            "Таблицы": "Wt",
            "Договоры": "Wd",
        }
        for src_title, dst_title in mapping_wh.items():
            if src_title not in wh_wb.sheetnames:
                report["missing_wh"].append(src_title)
                continue
            new_title = make_unique_sheet_title(analysis_wb, dst_title)
            copy_sheet(wh_wb[src_title], analysis_wb, new_title)
            report["copied"].append(f"{wh_name}:{src_title} -> {new_title}")

        if "Кредиты" in wh_wb.sheetnames:
            ws_cred = wh_wb["Кредиты"]
            if ws_cred["AC6"].value not in (None, 0, 0.0, "0"):
                new_title = make_unique_sheet_title(analysis_wb, "кред")
                copy_sheet(ws_cred, analysis_wb, new_title)
                report["copied"].append(f"{wh_name}:Кредиты -> {new_title}")

    return report


def merge_wh_m_into_analysis_with_prefix(
    analysis_wb,
    wh_bytes: Optional[bytes],
    wh_name: str,
    m_bytes: Optional[bytes],
    m_name: str,
    prefix: str,
) -> Dict[str, List[str]]:
    """
    Same as merge_wh_m_into_analysis, but forces the provided prefix on imported sheet names.
    Ensures names end with expected suffixes (Wd/Md etc.) so contracts detection keeps working.
    """
    report = {"missing_wh": [], "missing_m": [], "copied": []}
    p = (prefix or "").strip()

    if m_bytes:
        m_wb, _ = load_wb_from_bytes(m_bytes, m_name)
        # "Реестр" больше не копируем (Mr/Wr не должно быть).
        mapping_m = {"Итог": "M", "Контрагенты": "Mt", "Договоры": "Md"}
        for src_title, base in mapping_m.items():
            if src_title not in m_wb.sheetnames:
                report["missing_m"].append(src_title)
                continue
            new_title = make_unique_sheet_title(analysis_wb, f"{p}{base}")
            # Keep Md suffix intact for contracts pairing.
            if base.lower().endswith("md"):
                new_title = make_unique_with_fixed_suffix(analysis_wb, p + base[:-2], "Md")
            copy_sheet(m_wb[src_title], analysis_wb, new_title)
            report["copied"].append(f"{m_name}:{src_title} -> {new_title}")

    if wh_bytes:
        wh_wb, _ = load_wb_from_bytes(wh_bytes, wh_name)
        # "Реестр" больше не копируем (Mr/Wr не должно быть).
        mapping_wh = {"Итог": "W", "Таблицы": "Wt", "Договоры": "Wd"}
        for src_title, base in mapping_wh.items():
            if src_title not in wh_wb.sheetnames:
                report["missing_wh"].append(src_title)
                continue
            new_title = make_unique_sheet_title(analysis_wb, f"{p}{base}")
            if base.lower().endswith("wd"):
                new_title = make_unique_with_fixed_suffix(analysis_wb, p + base[:-2], "Wd")
            copy_sheet(wh_wb[src_title], analysis_wb, new_title)
            report["copied"].append(f"{wh_name}:{src_title} -> {new_title}")

        if "Кредиты" in wh_wb.sheetnames:
            ws_cred = wh_wb["Кредиты"]
            if ws_cred["AC6"].value not in (None, 0, 0.0, "0"):
                new_title = make_unique_sheet_title(analysis_wb, f"{p}кред")
                copy_sheet(ws_cred, analysis_wb, new_title)
                report["copied"].append(f"{wh_name}:Кредиты -> {new_title}")

    return report


# =========================
# OSV Cleaning (auto for "random named" OSV files)
# =========================
MAX_CELLS_PER_SHEET = 300_000
SORT_ACCOUNTS = {"1310", "1320", "1330"}  # only these accounts sorted by column G
OSV_BAD_WORDS = [
    "Договор", "Догов", "Д/р о государственных закупках",
    "KZT", "RUB", "EUR", "USD", "Жетысу",
    "Головное подразделение",
    "Основной склад", "Основное подразделение", "Склад",
    "Обороты за",
    "Соглашение от", "Соглашение об", ",,,.", 
    "Б/н", " от ",
    "01.", "02.", "03.", "04.", "05.", "06.",
    "07.", "08.", "09.", "10.", "11.", "12.", "<...>", "<..>", "<.>",
    "основной склад",
    "вспомогательный склад",
    "резервный склад",
    "торговый зал",
    "подразделение",
    "филиал",
    "представительство",
    "департамент",
    "управление",
    "отдел",
    "сектор",
    "группа",
    "бюро",
    "служба",
    "цех",
    "участок",
    "бригада",
    "лаборатория",
    "архив",
    "канцелярия",
    "бухгалтерия",
    "кадры",
    "юрист",
    "охрана",
    "администрация",
    "руководство",
    "дирекция",
    "секретариат",
    "хозчасть",
    "хозяйственная часть",
    "административно-хозяйственный",
    "ахч",
    "хозу",
    "хоз отдел",
    "хозяйственный отдел",
    "материальный склад",
    "продуктовый склад",
    "товарный склад",
    "оптовый склад",
    "розничный склад",
    "центральный склад",
    "региональный склад",
    "распределительный центр",
    "логистический центр",
    "транспортный отдел",
    "экспедиция",
    "доставка",
    "перемещение",
    "инвентаризация",
    "пересчет",
    "списание",
    "оприходование",
    "поступление",
    "реализация",
    "возврат",
    "брак",
    "пересортица",
    "недостача",
    "излишек",
    "корректировка",
    "переоценка",
    "уценка",
    "наценка",
    "маркировка",
    "перемаркировка",
    "упаковка",
    "переупаковка",
    "фасовка",
    "перефасовка",
    "комплектация",
    "раскомплектация",
    "сборка",
    "разборка",
    "ремонт",
    "обслуживание",
    "техническое обслуживание",
    "то-1",
    "то-2",
    "сезонное обслуживание",
    "консервация",
    "расконсервация",
    "монтаж",
    "демонтаж",
    "наладка",
    "регулировка",
    "калибровка",
    "поверка",
    "испытание",
    "проверка",
    "контроль",
    "аудит",
    "ревизия",
    "мониторинг",
    "наблюдение",
    "анализ",
    "исследование",
    "экспертиза",
    "оценка",
    "сертификация",
    "лицензирование",
    "аккредитация",
    "аттестация",
    "рационализация",
    "модернизация",
    "реконструкция",
    "строительство",
    "ремонт",
    "ремонтно-строительные",
    "отделочные работы",
    "сантехнические работы",
    "электромонтажные работы",
    "сварочные работы",
    "погрузочно-разгрузочные",
    "такелажные работы",
    "уборочные работы", "Основной",
    "клининговые услуги",
    "дезинфекция",
    "дератизация",
    "дезинсекция",
    "вывоз мусора",
    "утилизация",
    "переработка",
    "хранение",
    "ответственное хранение",
    "аренда",
    "субаренда",
    "лизинг",
    "прокат",
    "наем",
    "поднаем",
    "пользование",
    "эксплуатация",
    "содержание",
    "обслуживание зданий",
    "обслуживание сооружений",
    "обслуживание оборудования",
    "обслуживание техники",
    "обслуживание транспорта",
    "обслуживание инвентаря",
    "обслуживание инструмента",
    "обслуживание оснастки",
    "обслуживание приспособлений",
    "смазка",
    "заправка",
    "зарядка",
    "подзарядка",
    "замена",
    "замена масла",
    "замена фильтров",
    "замена расходников",
    "замена запчастей",
    "замена комплектующих",
    "ремонт оборудования",
    "ремонт техники",
    "ремонт транспорта",
    "ремонт инвентаря",
    "ремонт инструмента",
    "ремонт оснастки",
    "ремонт приспособлений",
    "текущий ремонт",
    "капитальный ремонт",
    "плановый ремонт",
    "аварийный ремонт",
    "восстановительный ремонт",
    "профилактический ремонт",
]


def find_first_row_with_value(ws, value, col: int = 1) -> Optional[int]:
    target_raw = str(value).strip()
    target = target_raw.lower()
    target_digits = target_raw.isdigit()
    for r in range(1, (ws.max_row or 1) + 1):
        v = ws.cell(row=r, column=col).value
        if v is None:
            continue
        if target_digits and isinstance(v, (int, float)) and float(v).is_integer():
            if int(v) == int(target_raw):
                return r
            continue
        s = str(v).replace("\u00A0", " ").replace("\u202F", " ").strip().lower()
        if s == target:
            return r
    return None


def find_first_row_contains(ws, substring: str, col: int = 1) -> Optional[int]:
    sub = (substring or "").lower()
    for r in range(1, (ws.max_row or 1) + 1):
        v = ws.cell(row=r, column=col).value
        if v and sub in str(v).lower():
            return r
    return None


def clear_outline_for_sheet(ws):
    last_row = max(ws.max_row or 1, 1)
    last_col = get_column_letter(max(ws.max_column or 1, 1))
    try:
        ws.ungroup_rows(1, last_row)
    except Exception:
        pass
    try:
        ws.ungroup_columns("A", last_col)
    except Exception:
        pass
    for r in range(1, last_row + 1):
        ws.row_dimensions[r].outlineLevel = 0
    for dim in ws.column_dimensions.values():
        dim.outlineLevel = 0


def remove_duplicate_rows(ws, start_row: int, end_row: int) -> bool:
    rows_data: Dict[Tuple[str, ...], List[int]] = {}
    rows_to_delete: Set[int] = set()
    for r in range(start_row, end_row + 1):
        if all((c.value is None or str(c.value).strip() == "") for c in ws[r]):
            continue
        row_values = tuple(str(cell.value).strip() if cell.value else "" for cell in ws[r])
        rows_data.setdefault(row_values, []).append(r)
    for rows in rows_data.values():
        if len(rows) >= 2:
            rows_to_delete.update(rows)
    if rows_to_delete:
        for r in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(r)
        return True
    return False


def get_account_number(ws) -> Optional[str]:
    """
    Extracts account number from the 2nd row.
    Some exports put the text in a merged cell that doesn't start at A2,
    so we scan multiple columns on row 2 and pick the first 4-digit match.
    """
    candidates: List[str] = []

    max_col = min(ws.max_column or 1, 60)
    for c in range(1, max_col + 1):
        v = ws.cell(row=2, column=c).value
        if v is None:
            continue
        if isinstance(v, (int, float)) and float(v).is_integer():
            candidates.append(str(int(v)))
        else:
            candidates.append(str(v))

    # If row 2 is part of a merged header where the value sits in row 1, grab the top-left value.
    for r in ws.merged_cells.ranges:
        min_col, min_row, max_col2, max_row2 = r.bounds
        if not (min_row <= 2 <= max_row2):
            continue
        tl = ws.cell(row=min_row, column=min_col).value
        if tl is None:
            continue
        candidates.append(str(tl))

    def _normalize(s: str) -> str:
        s = (s or "").replace("\u00A0", " ").replace("\u202F", " ").strip()
        return s

    def _digits_compact(s: str) -> str:
        s = _normalize(s).replace(" ", "")
        return "".join(ch for ch in s if ch.isdigit())

    def _pick_from_text(s: str) -> Optional[str]:
        s_norm = _normalize(s)
        s_low = s_norm.lower()
        # Prefer the 4 digits that follow the word "счет".
        idx = s_low.find("счет")
        if idx != -1:
            tail = s_norm[idx:]
            tail_compact = tail.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")
            m = re.search(r"(\\d{4})", tail_compact)
            if m:
                return m.group(1)
        # Fallback: any 4 consecutive digits (after compacting spaces between digits)
        compact = s_norm.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")
        m = re.search(r"(\\d{4})", compact)
        if m:
            return m.group(1)
        # Last resort: take first 4 digits from digits-only string.
        d = _digits_compact(s_norm)
        if len(d) >= 4:
            return d[:4]
        return None

    for s in candidates:
        acc = _pick_from_text(s)
        if acc:
            return acc
    return None


def set_all_rows_height(ws, height: float = 12):
    for r in range(1, (ws.max_row or 1) + 1):
        ws.row_dimensions[r].height = height


_num_re = re.compile(r"[-+]?\\d+(?:[.,]\\d+)?")


def to_number(v) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return None
    s = s.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    m = _num_re.search(s)
    if not m:
        return None
    num = m.group(0).replace(",", ".")
    try:
        return float(num)
    except Exception:
        return None


def _cell_text(v) -> str:
    if v is None:
        return ""
    try:
        s = str(v)
    except Exception:
        return ""
    s = s.replace("\u00A0", " ").replace("\u202F", " ").strip()
    return s


def extract_company_label_from_a1(ws) -> str:
    """
    For UX: show something human-readable per OSV sheet.
    Prefer A1; if empty, use first non-empty cell in row 1 (A..J).
    """
    s = _cell_text(ws.cell(row=1, column=1).value)
    if s:
        return s
    for c in range(1, 11):
        s = _cell_text(ws.cell(row=1, column=c).value)
        if s:
            return s
    return ""


def _short(s: str, n: int = 90) -> str:
    s = _cell_text(s)
    if len(s) <= n:
        return s
    return s[: n - 1].rstrip() + "…"


def sort_block_by_column(ws, start_row: int, end_row: int, col_index: int, descending: bool = True):
    if start_row >= end_row:
        return
    max_col = ws.max_column or 1
    rows = []
    for r in range(start_row, end_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        key = to_number(ws.cell(row=r, column=col_index).value)
        rows.append((key, row_vals))
    rows.sort(key=lambda x: (x[0] is not None, x[0] if x[0] is not None else 0.0), reverse=descending)
    for i, (_, row_vals) in enumerate(rows):
        rr = start_row + i
        for c in range(1, max_col + 1):
            ws.cell(row=rr, column=c).value = row_vals[c - 1]


def clean_osv_sheet_inplace(ws) -> Optional[str]:
    account_number = get_account_number(ws)
    if not account_number:
        set_all_rows_height(ws, 12)
        return None

    clear_outline_for_sheet(ws)
    if 5 in ws.row_dimensions:
        ws.row_dimensions[5].hidden = False

    total_cells = (ws.max_row or 1) * (ws.max_column or 1)
    if total_cells > MAX_CELLS_PER_SHEET:
        set_all_rows_height(ws, 12)
        return account_number

    last_col = ws.max_column or 1
    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))

    start_row = find_first_row_with_value(ws, account_number, col=1)
    end_row = find_first_row_with_value(ws, "Итого", col=1)
    if end_row is None:
        # Some OSV exports use "Итого:" or "Итого ..." in column A.
        for r in range(1, (ws.max_row or 1) + 1):
            v = ws.cell(row=r, column=1).value
            if v is None:
                continue
            s = str(v).replace("\u00A0", " ").replace("\u202F", " ").strip().lower()
            if s.startswith("итого"):
                end_row = r
                break
    if not start_row or not end_row:
        set_all_rows_height(ws, 12)
        # Still keep the sheet: we at least know the account number.
        return account_number

    for r in range(ws.max_row or 1, start_row - 1, -1):
        v = ws.cell(row=r, column=1).value
        if v is None or str(v).strip() == "":
            ws.delete_rows(r)

    bad_words_lower = [bw.lower() for bw in OSV_BAD_WORDS]
    for r in range(ws.max_row or 1, start_row - 1, -1):
        v = ws.cell(row=r, column=1).value
        if v and any(bw in str(v).lower() for bw in bad_words_lower):
            ws.delete_rows(r)

    start_row = find_first_row_with_value(ws, account_number, col=1)
    end_row = find_first_row_with_value(ws, "Итого", col=1)
    if end_row is None:
        for r in range(1, (ws.max_row or 1) + 1):
            v = ws.cell(row=r, column=1).value
            if v is None:
                continue
            s = str(v).replace("\u00A0", " ").replace("\u202F", " ").strip().lower()
            if s.startswith("итого"):
                end_row = r
                break
    if not start_row or not end_row:
        set_all_rows_height(ws, 12)
        return account_number

    if start_row + 1 <= end_row - 1:
        remove_duplicate_rows(ws, start_row + 1, end_row - 1)

    if account_number in SORT_ACCOUNTS:
        data_start = start_row + 1
        data_end = end_row - 1
        if data_start <= data_end:
            sort_block_by_column(ws, data_start, data_end, col_index=7, descending=True)

    for r in range(start_row + 1, start_row + 2000):
        if r > (ws.max_row or 1):
            break
        cell = ws.cell(row=r, column=1)
        if cell.value is None:
            continue
        al = cell.alignment or Alignment()
        cell.alignment = Alignment(horizontal="left", vertical=al.vertical)

    for r in range(ws.max_row or 1, 0, -1):
        if all((c.value is None or str(c.value).strip() == "") for c in ws[r]):
            ws.delete_rows(r)

    ws.column_dimensions["A"].width = 50
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws.column_dimensions[col].width = 18

    if last_col == 8 and (ws.max_column or 1) >= 9:
        ws.delete_cols(9)
    elif last_col == 9 and (ws.max_column or 1) >= 8:
        ws.delete_cols(8)

    countword = find_first_row_contains(ws, "Счет", col=1)
    if countword:
        for r in range(1, countword + 1):
            for cell in ws[r]:
                al = cell.alignment or Alignment()
                cell.alignment = Alignment(horizontal=al.horizontal, vertical=al.vertical, wrap_text=False)

    set_all_rows_height(ws, 12)
    return account_number


def add_cleaned_osv_files_to_analysis(analysis_wb, osv_files: Iterable[Tuple[str, bytes]]) -> Dict[str, List[str]]:
    report = {"added": [], "skipped": []}
    for fname, fbytes in osv_files:
        osv_wb, _ = load_wb_from_bytes(fbytes, fname)
        for ws in osv_wb.worksheets:
            clear_outline_for_sheet(ws)
        for ws in osv_wb.worksheets:
            acc = clean_osv_sheet_inplace(ws)
            if not acc:
                # Do not drop the sheet: still add it, but name it generically.
                fallback = make_unique_sheet_title(analysis_wb, "OSV")
                copy_sheet(ws, analysis_wb, fallback)
                report["skipped"].append(f"{fname}:{ws.title} (не найден счет во 2-й строке)")
                report["added"].append(f"{fname}:{ws.title} -> {fallback}")
            else:
                # Keep account suffix intact (important for saldo detection on *1210/*1710/*3310/*3510).
                if acc.isdigit() and len(acc) == 4:
                    new_title = make_unique_with_fixed_suffix(analysis_wb, "", acc)
                else:
                    new_title = make_unique_sheet_title(analysis_wb, acc)
                copy_sheet(ws, analysis_wb, new_title)
                report["added"].append(f"{fname}:{ws.title} -> {new_title}")
    return report


def compute_availability_from_wb(wb) -> Dict[str, object]:
    inventory_accounts = ["1310", "1320", "1330"]
    saldo_suffixes = {"1210", "1710", "3310", "3510"}

    inv_map = {acc: any(acc in sh for sh in wb.sheetnames) for acc in inventory_accounts}
    saldo_ok = any(split_prefix_suffix4(sh)[1] in saldo_suffixes for sh in wb.sheetnames)

    prefix_to_pair: Dict[str, Dict[str, str]] = defaultdict(dict)
    for sh in wb.sheetnames:
        prefix, suf2 = split_prefix_suffix2(sh)
        if suf2 in ("wd", "md"):
            prefix = normalize_prefix(prefix)
            if suf2 not in prefix_to_pair[prefix]:
                prefix_to_pair[prefix][suf2] = sh
    contracts_prefixes = [p for p, d in prefix_to_pair.items() if "wd" in d and "md" in d]
    contracts_ok = bool(contracts_prefixes)

    return {
        "inventory_map": inv_map,
        "saldo_ok": saldo_ok,
        "contracts_ok": contracts_ok,
        "contracts_prefixes": contracts_prefixes,
    }


def find_existing_saldo_prefixes(wb) -> Dict[str, Set[str]]:
    """For each saldo account (1210/1710/3310/3510), returns existing prefixes found in sheetnames."""
    accounts = {"1210", "1710", "3310", "3510"}
    out: Dict[str, Set[str]] = {a: set() for a in accounts}
    for sh in wb.sheetnames:
        prefix, suf = split_prefix_suffix4(sh)
        if suf in accounts:
            out[suf].add(normalize_prefix(prefix))
    return out


def list_existing_saldo_sheets_with_a1(wb) -> List[Tuple[str, str, str]]:
    """Returns tuples: (sheetname, account_suffix, A1_text) for saldo accounts."""
    accounts = {"1210", "1710", "3310", "3510"}
    out = []
    for sh in wb.sheetnames:
        prefix, suf = split_prefix_suffix4(sh)
        if suf in accounts:
            try:
                a1 = extract_company_label_from_a1(wb[sh])
            except Exception:
                a1 = ""
            out.append((sh, suf, _short(a1)))
    return out


def build_analysis_workbook(
    analysis_file: Optional[Tuple[str, bytes]],
    wh_files: List[Tuple[str, bytes, str]],
    m_files: List[Tuple[str, bytes, str]],
    osv_files: List[Tuple[str, bytes]],
    analysis_name_for_new: str,
    osv_prefix_by_sheet: Optional[Dict[Tuple[str, str], str]] = None,
) -> Tuple[bytes, str, Dict[str, object], Dict[str, List[str]]]:
    report: Dict[str, List[str]] = {"warnings": [], "copied": []}

    if analysis_file:
        analysis_name, analysis_bytes = analysis_file
        analysis_wb, keep_vba = load_wb_from_bytes(analysis_bytes, analysis_name)
        out_name = analysis_name
        placeholder_title = None
    else:
        analysis_wb = Workbook()
        # Keep a visible placeholder sheet so saving never fails with
        # "At least one sheet must be visible" even if all inputs are skipped.
        placeholder_ws = analysis_wb.active
        placeholder_title = "Сборка"
        placeholder_ws.title = placeholder_title
        placeholder_ws["A1"] = "Служебный лист. Удалится автоматически, если добавятся другие листы."
        keep_vba = False
        base = (analysis_name_for_new or "").strip()
        out_name = safe_filename(f"_Анализ {base}".strip() + ".xlsx")

    if wh_files or m_files:
        missing_wh_all: List[str] = []
        missing_m_all: List[str] = []

        for wh_name, wh_bytes, pref in wh_files:
            merge_report = merge_wh_m_into_analysis_with_prefix(analysis_wb, wh_bytes, wh_name, None, "", prefix=pref)
            missing_wh_all.extend([f"{wh_name}:{x}" for x in merge_report["missing_wh"]])
            report["copied"].extend(merge_report["copied"])

        for m_name, m_bytes, pref in m_files:
            merge_report = merge_wh_m_into_analysis_with_prefix(analysis_wb, None, "", m_bytes, m_name, prefix=pref)
            missing_m_all.extend([f"{m_name}:{x}" for x in merge_report["missing_m"]])
            report["copied"].extend(merge_report["copied"])

        if missing_wh_all or missing_m_all:
            msg = []
            if missing_wh_all:
                msg.append("WH: " + ", ".join(missing_wh_all[:6]) + (" ..." if len(missing_wh_all) > 6 else ""))
            if missing_m_all:
                msg.append("M: " + ", ".join(missing_m_all[:6]) + (" ..." if len(missing_m_all) > 6 else ""))
            report["warnings"].append("WH/M: не найдены листы. " + " | ".join(msg))

    if osv_files:
        osv_prefix_by_sheet = osv_prefix_by_sheet or {}
        # We need per-account prefix control for saldo accounts to avoid collisions while keeping suffixes intact.
        report_local = {"added": [], "skipped": []}
        for fname, fbytes in osv_files:
            osv_wb, _ = load_wb_from_bytes(fbytes, fname)
            for ws in osv_wb.worksheets:
                clear_outline_for_sheet(ws)
            for idx, ws in enumerate(osv_wb.worksheets):
                original_title = ws.title
                acc = clean_osv_sheet_inplace(ws)
                if not acc:
                    fallback = make_unique_sheet_title(analysis_wb, "OSV")
                    copy_sheet(ws, analysis_wb, fallback)
                    report_local["skipped"].append(f"{fname}:{ws.title} (не найден счет во 2-й строке)")
                    report_local["added"].append(f"{fname}:{ws.title} -> {fallback}")
                    continue

                if acc.isdigit() and len(acc) == 4:
                    p = (osv_prefix_by_sheet.get((fname, original_title, idx)) or "").strip()
                    new_title = make_unique_with_fixed_suffix(analysis_wb, p, acc)
                else:
                    new_title = make_unique_sheet_title(analysis_wb, acc)
                copy_sheet(ws, analysis_wb, new_title)
                report_local["added"].append(f"{fname}:{ws.title} -> {new_title}")

        osv_report = report_local
        if osv_report["skipped"]:
            details = "; ".join(osv_report["skipped"][:8])
            more = f" (+{len(osv_report['skipped'])-8})" if len(osv_report["skipped"]) > 8 else ""
            report["warnings"].append(f"ОСВ: пропущены листы: {len(osv_report['skipped'])}. {details}{more}")
        report["copied"].extend(osv_report["added"])

    # Drop placeholder if we successfully added at least one real sheet.
    if placeholder_title and len(analysis_wb.sheetnames) > 1 and placeholder_title in analysis_wb.sheetnames:
        del analysis_wb[placeholder_title]

    # Ensure at least one visible sheet exists (openpyxl requirement on save).
    if not analysis_wb.sheetnames:
        ws = analysis_wb.create_sheet("Сборка")
        ws["A1"] = "Пустой файл: не удалось добавить ни одного листа."
    visible = [ws for ws in analysis_wb.worksheets if getattr(ws, "sheet_state", "visible") == "visible"]
    if not visible:
        analysis_wb.worksheets[0].sheet_state = "visible"

    availability = compute_availability_from_wb(analysis_wb)

    out = io.BytesIO()
    analysis_wb.save(out)
    return out.getvalue(), out_name, availability, report


# =========================
# CODE 1 (Saldo) — multi-company by prefix + ####
# Finds sheets where last 4 chars are 1210/1710/3310/3510
# Creates output sheet per prefix: "<prefix>сальд" or "сальд" if no prefix
# =========================
def run_code_1(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    target_suffixes = {"1210": 6, "1710": 6, "3310": 7, "3510": 7}

    prefix_to_sheets: Dict[str, Dict[str, str]] = defaultdict(dict)
    for sh in xls.sheet_names:
        prefix, suf = split_prefix_suffix4(sh)
        if suf in target_suffixes:
            prefix = normalize_prefix(prefix)
            if suf not in prefix_to_sheets[prefix]:
                prefix_to_sheets[prefix][suf] = sh

    if not prefix_to_sheets:
        raise ValueError("Код 1: не найдено листов, заканчивающихся на 1210/1710/3310/3510.")

    for prefix, suf_map in prefix_to_sheets.items():
        sheet_data = {}

        for suf, col_idx in target_suffixes.items():
            if suf not in suf_map:
                continue

            df = pd.read_excel(xls, sheet_name=suf_map[suf], header=0)
            if df.shape[1] <= col_idx:
                continue

            temp = df.iloc[:, [0, col_idx]].copy()
            temp.columns = ["Контрагент", "value"]
            temp["value"] = pd.to_numeric(temp["value"], errors="coerce").fillna(0)
            temp = temp[temp["value"] != 0]
            temp["Контрагент"] = temp["Контрагент"].astype(str).str.strip()
            temp = temp[temp["Контрагент"] != ""]
            temp = temp[~temp["Контрагент"].isin(["1210", "1710", "3310", "3510"])]
            temp = temp[~temp["Контрагент"].str.lower().str.startswith("итого")]

            if temp.empty:
                continue

            sheet_data[suf] = temp.groupby("Контрагент")["value"].sum()

        s1210 = sheet_data.get("1210", pd.Series(dtype=float))
        s3510 = sheet_data.get("3510", pd.Series(dtype=float))
        s1710 = sheet_data.get("1710", pd.Series(dtype=float))
        s3310 = sheet_data.get("3310", pd.Series(dtype=float))

        cust_set = set(s1210.index).union(set(s3510.index))
        supp_set = set(s1710.index).union(set(s3310.index))
        all_set = cust_set.union(supp_set)

        if not cust_set and not supp_set:
            continue

        if cust_set:
            df_cust = pd.DataFrame(sorted(cust_set), columns=["Контрагент"])
            df_cust["1210"] = df_cust["Контрагент"].map(s1210).fillna(0) / 1000
            df_cust["3510"] = df_cust["Контрагент"].map(s3510).fillna(0) / 1000
            df_cust["сальдо заказчики"] = df_cust["1210"] - df_cust["3510"]
            df_cust = df_cust.sort_values(by="сальдо заказчики", ascending=False).reset_index(drop=True)
        else:
            df_cust = pd.DataFrame(columns=["Контрагент", "1210", "3510", "сальдо заказчики"])

        if supp_set:
            df_supp = pd.DataFrame(sorted(supp_set), columns=["Контрагент"])
            df_supp["1710"] = df_supp["Контрагент"].map(s1710).fillna(0) / 1000
            df_supp["3310"] = df_supp["Контрагент"].map(s3310).fillna(0) / 1000
            df_supp["сальдо поставщики"] = df_supp["1710"] - df_supp["3310"]
            df_supp = df_supp.sort_values(by="сальдо поставщики", ascending=False).reset_index(drop=True)
        else:
            df_supp = pd.DataFrame(columns=["Контрагент", "1710", "3310", "сальдо поставщики"])

        if all_set:
            df_total = pd.DataFrame(sorted(all_set), columns=["Контрагент"])
            df_total["1210"] = df_total["Контрагент"].map(s1210).fillna(0) / 1000
            df_total["1710"] = df_total["Контрагент"].map(s1710).fillna(0) / 1000
            df_total["3310"] = df_total["Контрагент"].map(s3310).fillna(0) / 1000
            df_total["3510"] = df_total["Контрагент"].map(s3510).fillna(0) / 1000
            df_total["общее сальдо"] = df_total["1210"] + df_total["1710"] - df_total["3310"] - df_total["3510"]
            df_total = df_total.sort_values(by="общее сальдо", ascending=False).reset_index(drop=True)
        else:
            df_total = pd.DataFrame(columns=["Контрагент", "общее сальдо"])

        out_sheet_name = safe_sheet_name(f"{prefix}сальд" if prefix else "сальд")
        if out_sheet_name in wb.sheetnames:
            wb.remove(wb[out_sheet_name])
        ws = wb.create_sheet(out_sheet_name)

        ws["A1"] = "Все значения указаны в тысячах тенге"
        ws["A1"].font = Font(name="Arial", size=10, bold=True)

        start_row = 2
        start_col = 2  # B

        font_header = Font(name="Arial", size=10, bold=True)
        font_body = Font(name="Arial", size=10)
        font_bold_body = Font(name="Arial", size=10, bold=True)
        align_center = Alignment(horizontal="center")
        align_left = Alignment(horizontal="left")
        number_format_acc = "#,##0;[Red](#,##0)"

        col_cust_contr = start_col
        col_cust_1210 = start_col + 1
        col_cust_3510 = start_col + 2
        col_cust_saldo = start_col + 3

        col_supp_contr = start_col + 5
        col_supp_1710 = start_col + 6
        col_supp_3310 = start_col + 7
        col_supp_saldo = start_col + 8

        col_total_contr = start_col + 10
        col_total_saldo = start_col + 11

        if not df_cust.empty:
            headers = {
                col_cust_contr: "Контрагент",
                col_cust_1210: "1210",
                col_cust_3510: "3510",
                col_cust_saldo: "сальдо с заказчиками",
            }
            for col, text in headers.items():
                c = ws.cell(row=start_row, column=col, value=text)
                c.font = font_header
                c.alignment = align_center

            for i, (_, row) in enumerate(df_cust.iterrows(), start=start_row + 1):
                r = i
                c_contr = ws.cell(row=r, column=col_cust_contr, value=row["Контрагент"])
                c_contr.font = font_body
                c_contr.alignment = align_left

                for col, val, style in [
                    (col_cust_1210, row["1210"], font_body),
                    (col_cust_3510, row["3510"], font_body),
                    (col_cust_saldo, row["сальдо заказчики"], font_bold_body),
                ]:
                    cell = ws.cell(row=r, column=col, value=val)
                    cell.font = style
                    cell.alignment = align_center
                    cell.number_format = number_format_acc

        if not df_supp.empty:
            headers = {
                col_supp_contr: "Контрагент",
                col_supp_1710: "1710",
                col_supp_3310: "3310",
                col_supp_saldo: "сальдо с поставщиками",
            }
            for col, text in headers.items():
                c = ws.cell(row=start_row, column=col, value=text)
                c.font = font_header
                c.alignment = align_center

            for i, (_, row) in enumerate(df_supp.iterrows(), start=start_row + 1):
                r = i
                c_contr = ws.cell(row=r, column=col_supp_contr, value=row["Контрагент"])
                c_contr.font = font_body
                c_contr.alignment = align_left

                for col, val, style in [
                    (col_supp_1710, row["1710"], font_body),
                    (col_supp_3310, row["3310"], font_body),
                    (col_supp_saldo, row["сальдо поставщики"], font_bold_body),
                ]:
                    cell = ws.cell(row=r, column=col, value=val)
                    cell.font = style
                    cell.alignment = align_center
                    cell.number_format = number_format_acc

        if not df_total.empty:
            headers = {col_total_contr: "Контрагент", col_total_saldo: "общее сальдо"}
            for col, text in headers.items():
                c = ws.cell(row=start_row, column=col, value=text)
                c.font = font_header
                c.alignment = align_center

            for i, (_, row) in enumerate(df_total.iterrows(), start=start_row + 1):
                r = i
                c_contr = ws.cell(row=r, column=col_total_contr, value=row["Контрагент"])
                c_contr.font = font_body
                c_contr.alignment = align_left

                cell = ws.cell(row=r, column=col_total_saldo, value=row["общее сальдо"])
                cell.font = font_bold_body
                cell.alignment = align_center
                cell.number_format = number_format_acc

        WIDTH_CONTR = 30
        WIDTH_NUM = 18
        for col in [col_cust_contr, col_supp_contr, col_total_contr]:
            ws.column_dimensions[get_column_letter(col)].width = WIDTH_CONTR
        for col in [
            col_cust_1210, col_cust_3510, col_cust_saldo,
            col_supp_1710, col_supp_3310, col_supp_saldo,
            col_total_saldo
        ]:
            ws.column_dimensions[get_column_letter(col)].width = WIDTH_NUM

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 2 (Contracts) — multi-company by prefix + (md/wd) ignoring case
# Creates output sheet per prefix: "<prefix>контр" or "контр" if no prefix
# =========================
def run_code_2(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))

    prefix_to_pair: Dict[str, Dict[str, str]] = defaultdict(dict)
    for sh in wb.sheetnames:
        prefix, suf2 = split_prefix_suffix2(sh)
        if suf2 in ("wd", "md"):
            prefix = normalize_prefix(prefix)
            if suf2 not in prefix_to_pair[prefix]:
                prefix_to_pair[prefix][suf2] = sh

    valid_prefixes = [p for p, d in prefix_to_pair.items() if "wd" in d and "md" in d]
    if not valid_prefixes:
        raise ValueError("Код 2: не найдено пар листов (Wd/Md) по префиксам.")

    source_sheets_to_delete = set()
    for prefix in valid_prefixes:
        md_name = prefix_to_pair[prefix]["md"]
        wd_name = prefix_to_pair[prefix]["wd"]
        md_ws = wb[md_name]
        wd_ws = wb[wd_name]
        source_sheets_to_delete.add(md_name)
        source_sheets_to_delete.add(wd_name)

        payments_year = defaultdict(lambda: [0.0, 0.0, 0.0])
        performance_year = defaultdict(lambda: [0.0, 0.0, 0.0])
        payments_2025_monthly = defaultdict(lambda: [0.0] * 12)
        performance_2025_monthly = defaultdict(lambda: [0.0] * 12)

        def collect_yearly(sheet, target_dict):
            for row in range(2, sheet.max_row + 1):
                n = sheet[f"A{row}"].value
                c = sheet[f"B{row}"].value
                if not n and not c:
                    continue
                key = (str(n).strip() if n else "", str(c).strip() if c else "")
                for idx, col in enumerate(["C", "D", "E"]):
                    v = sheet[f"{col}{row}"].value
                    if v is None:
                        continue
                    try:
                        target_dict[key][idx] += float(v)
                    except:
                        pass

        def collect_monthly_2025(sheet, target_dict):
            start_col = column_index_from_string("AE")
            for row in range(2, sheet.max_row + 1):
                n = sheet[f"A{row}"].value
                c = sheet[f"B{row}"].value
                if not n and not c:
                    continue
                key = (str(n).strip() if n else "", str(c).strip() if c else "")
                for i in range(12):
                    v = sheet.cell(row=row, column=start_col + i).value
                    if v is None:
                        continue
                    try:
                        target_dict[key][i] += float(v)
                    except:
                        pass

        collect_yearly(wd_ws, payments_year)
        collect_yearly(md_ws, performance_year)
        collect_monthly_2025(wd_ws, payments_2025_monthly)
        collect_monthly_2025(md_ws, performance_2025_monthly)

        all_keys = sorted(
            set(payments_year.keys())
            | set(performance_year.keys())
            | set(payments_2025_monthly.keys())
            | set(performance_2025_monthly.keys()),
            key=lambda x: (x[0], x[1]),
        )

        out_sheet_name = safe_sheet_name(f"{prefix}контр" if prefix else "контр")
        if out_sheet_name in wb.sheetnames:
            del wb[out_sheet_name]
        ws = wb.create_sheet(out_sheet_name)

        ws["A1"] = "ИТОГО в тыс тенге"
        ws["A2"] = "Контрагент"
        ws["B2"] = "Договор"

        ws["C1"] = "оплата"
        ws["C2"] = 2023
        ws["D2"] = 2024
        ws["E2"] = 2025
        ws["F2"] = "Total"

        ws["G1"] = "выполнения с ндс"
        ws["G2"] = 2023
        ws["H2"] = 2024
        ws["I2"] = 2025
        ws["J2"] = "Total"

        ws["K2"] = "дз/(аванс)"

        ws["M1"] = "оплата"
        months = [f"2025_{str(i).zfill(2)}" for i in range(1, 13)]
        start_col_pay = column_index_from_string("M")
        for i, label in enumerate(months):
            ws[f"{get_column_letter(start_col_pay + i)}2"] = label

        ws["Y1"] = "выполнения с ндс"
        start_col_perf = column_index_from_string("Y")
        for i, label in enumerate(months):
            ws[f"{get_column_letter(start_col_perf + i)}2"] = label

        start_row = 3
        for idx, key in enumerate(all_keys):
            row = start_row + idx
            name, contract = key

            ws[f"A{row}"] = name
            ws[f"B{row}"] = contract

            py = payments_year.get(key, [0, 0, 0])
            ws[f"C{row}"] = py[0]
            ws[f"D{row}"] = py[1]
            ws[f"E{row}"] = py[2]

            pf = performance_year.get(key, [0, 0, 0])
            ws[f"G{row}"] = pf[0] * 1.12
            ws[f"H{row}"] = pf[1] * 1.12
            ws[f"I{row}"] = pf[2] * 1.12

            ws[f"F{row}"] = f"=SUM(C{row}:E{row})"
            ws[f"J{row}"] = f"=SUM(G{row}:I{row})"
            ws[f"K{row}"] = f"=J{row}-F{row}"

            mp = payments_2025_monthly.get(key, [0] * 12)
            for i in range(12):
                ws[f"{get_column_letter(start_col_pay + i)}{row}"] = mp[i]

            mf = performance_2025_monthly.get(key, [0] * 12)
            for i in range(12):
                ws[f"{get_column_letter(start_col_perf + i)}{row}"] = mf[i] * 1.12

        last_row = start_row + len(all_keys) - 1 if all_keys else 2

        regular = Font(name="Arial", size=10)
        bold = Font(name="Arial", size=10, bold=True)

        for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=column_index_from_string("AJ")):
            for c in row:
                c.font = regular

        for col in range(1, column_index_from_string("AJ") + 1):
            ws[f"{get_column_letter(col)}2"].font = bold
        for addr in ["A1", "C1", "G1", "M1", "Y1", "K2"]:
            ws[addr].font = bold

        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")
        for col in range(1, column_index_from_string("AJ") + 1):
            addr = f"{get_column_letter(col)}2"
            ws[addr].alignment = left if addr in ["A2", "B2"] else center

        num_format = "#,##0;[Red](#,##0)"
        numeric_cols = list("CDEFGHIJK")
        numeric_cols += [get_column_letter(c) for c in range(start_col_pay, start_col_pay + 12)]
        numeric_cols += [get_column_letter(c) for c in range(start_col_perf, start_col_perf + 12)]

        for col in numeric_cols:
            for r in range(3, last_row + 1):
                cell = ws[f"{col}{r}"]
                cell.alignment = center
                cell.number_format = num_format

        for r in range(1, last_row + 1):
            ws[f"A{r}"].alignment = left
            ws[f"B{r}"].alignment = left

        for addr in ["C1", "G1", "M1", "Y1"]:
            ws[addr].alignment = left

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 38
        for col in numeric_cols:
            ws.column_dimensions[col].width = 12.2 if column_index_from_string(col) >= start_col_pay else 12.6

        thin = Side(border_style="thin", color="000000")
        border_cols = ["C", "G", "K", "M", "Y"]
        for r in range(1, last_row + 1):
            for col in border_cols:
                cell = ws[f"{col}{r}"]
                cell.border = Border(
                    left=thin,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )

    # After building the contract sheet(s), drop source Wd/Md sheets from the output file.
    for sh in sorted(source_sheets_to_delete):
        if sh in wb.sheetnames:
            del wb[sh]

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 3 (Inventory / Запасы)
# For each selected account (1310/1320/1330):
# - Find sheets whose name contains the account substring.
# - In each sheet, find the row in column A where the account appears.
# - Write threshold % values into J..N above that row.
# - Fill formulas in J..N for rows below until column G becomes empty.
# Does not create any extra report sheets; UI uses returned report info.
# =========================
def run_code_3_inventory(file_bytes: bytes, accounts: List[str]) -> Tuple[bytes, Dict[str, List[str]]]:
    wb = load_workbook(io.BytesIO(file_bytes))

    processed: List[str] = []
    missing_sheets: List[str] = []
    missing_markers: List[str] = []

    def _is_blank(v) -> bool:
        return v is None or (isinstance(v, str) and v.strip() == "")

    def _find_account_row(ws, account: str) -> Optional[int]:
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if _is_blank(v):
                continue
            if isinstance(v, (int, float)) and float(v).is_integer():
                s = str(int(v))
            else:
                s = str(v).strip()
            if s == account:
                return r
        return None

    def _set_percent(ws, row: int, col: int, value: float):
        cell = ws.cell(row=row, column=col, value=value)
        cell.number_format = "0%"
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def _set_number_style(cell, bold: bool = False, fill=None):
        cell.number_format = "#,##0;[Red](#,##0)"
        cell.alignment = Alignment(horizontal="center", vertical="center")
        # Force consistent font sizing for all J..N numbers (including totals).
        cell.font = Font(name="Arial", size=9, bold=bold)
        if fill is not None:
            cell.fill = fill

    def _fill_formulas(ws, start_row: int, stop_row: int, thr_low: int, thr_high: int):
        for r in range(start_row, stop_row + 1):
            ratio = f"IFERROR($G{r}/$F{r},100)"
            ws.cell(row=r, column=10).value = f"=IF({ratio}>J${thr_low},$G{r},0)"
            ws.cell(row=r, column=11).value = (
                f"=IF({ratio}>K${thr_low},IF({ratio}<=K${thr_high},$G{r},0),0)"
            )
            ws.cell(row=r, column=12).value = (
                f"=IF({ratio}>L${thr_low},IF({ratio}<=L${thr_high},$G{r},0),0)"
            )
            ws.cell(row=r, column=13).value = (
                f"=IF({ratio}>M${thr_low},IF({ratio}<=M${thr_high},$G{r},0),0)"
            )
            ws.cell(row=r, column=14).value = (
                f"=IF({ratio}>N${thr_low},IF({ratio}<=N${thr_high},$G{r},0),0)"
            )
            for c in range(10, 15):
                _set_number_style(ws.cell(row=r, column=c))

    def _apply_dotted_grid(ws, row_from: int, row_to: int, col_from: int, col_to: int):
        side = Side(border_style="dotted", color="000000")
        border = Border(left=side, right=side, top=side, bottom=side)
        for r in range(row_from, row_to + 1):
            for c in range(col_from, col_to + 1):
                ws.cell(row=r, column=c).border = border

    for account in accounts:
        matched_sheets = [sh for sh in wb.sheetnames if account in sh]
        if not matched_sheets:
            missing_sheets.append(account)
            continue

        for sh in matched_sheets:
            ws = wb[sh]
            found_row = _find_account_row(ws, account)
            if found_row is None:
                missing_markers.append(f"{account}: {sh}")
                continue

            thr_low = found_row - 3
            thr_high = found_row - 2
            if thr_low < 1 or thr_high < 1:
                missing_markers.append(f"{account}: {sh} (слишком близко к началу листа)")
                continue

            # Thresholds (percent values stored as decimals).
            _set_percent(ws, thr_low, 10, 2.0)   # J: 200%
            _set_percent(ws, thr_low, 11, 1.0)   # K: 100%
            _set_percent(ws, thr_high, 11, 2.0)  # K: 200%
            _set_percent(ws, thr_low, 12, 0.5)   # L: 50%
            _set_percent(ws, thr_high, 12, 1.0)  # L: 100%
            _set_percent(ws, thr_low, 13, 0.25)  # M: 25%
            _set_percent(ws, thr_high, 13, 0.5)  # M: 50%
            _set_percent(ws, thr_low, 14, 0.0)   # N: 0%
            _set_percent(ws, thr_high, 14, 0.25) # N: 25%

            # Fill formulas for inventory rows: from row after the marker until first blank in column G.
            start_row = found_row + 1
            last = start_row - 1
            for r in range(start_row, ws.max_row + 1):
                g = ws.cell(row=r, column=7).value
                if _is_blank(g):
                    break
                last = r

            if last >= start_row:
                _fill_formulas(ws, start_row, last, thr_low, thr_high)
                # Totals on the same row where the account marker is found.
                total_fill = PatternFill(fill_type="solid", fgColor="C6EFCE")
                for col_letter, col_idx in [("J", 10), ("K", 11), ("L", 12), ("M", 13), ("N", 14)]:
                    cell = ws.cell(row=found_row, column=col_idx)
                    cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{last})"
                    _set_number_style(cell, bold=True, fill=total_fill)

                # Dotted grid around the mini-table (thresholds + totals + data).
                _apply_dotted_grid(ws, thr_low, last, 10, 14)
                processed.append(f"{account}: {sh} (строки {start_row}-{last})")
            else:
                # No data rows: still show dotted frame around thresholds + total row.
                _apply_dotted_grid(ws, thr_low, found_row, 10, 14)
                processed.append(f"{account}: {sh} (нет строк с данными в G ниже маркера)")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), {"processed": processed, "missing_sheets": missing_sheets, "missing_markers": missing_markers}


# =========================
# UI — Theme toggle (Dark/Light), no gradients, default font
# =========================
st.set_page_config(page_title="", page_icon=None, layout="wide", initial_sidebar_state="collapsed")

# Dark theme only (deep navy)
BG = "#1D2F4E"
TEXT = "#EAF0FF"
CARD = "#162844"
BORDER = "#2B426A"
MUTED = "#B6C4E3"
BTN_BG = "#EADFCB"   # beige buttons
BTN_TEXT = "#111111" # black text/icons
PROG = "#A9C7FF"

st.markdown(
    f"""
    <style>
      #MainMenu {{visibility: hidden;}}
      footer {{visibility: hidden;}}
      header {{visibility: hidden;}}

      .stApp {{
        background: {BG};
        color: {TEXT};
      }}

      [data-testid="stSidebar"] {{
        background: {BG};
      }}

      /* Keep default fonts; only set sizes and colors */
      html, body, [class*="css"], .stApp {{
        font-size: 15px !important;
        font-family: "Cambria Math", "STIX Two Math", "Latin Modern Math", "Times New Roman", serif !important;
      }}

      button, input, textarea, select, option {{
        font-family: "Cambria Math", "STIX Two Math", "Latin Modern Math", "Times New Roman", serif !important;
      }}

      /* Headings/markdown containers sometimes override font-family */
      h1, h2, h3, h4, h5, h6,
      [data-testid="stMarkdownContainer"],
      [data-testid="stMarkdownContainer"] * {{
        font-family: "Cambria Math", "STIX Two Math", "Latin Modern Math", "Times New Roman", serif !important;
      }}

      .block-container {{
        max-width: 900px;
        padding-top: 1.0rem;
        padding-bottom: 1.2rem;
      }}

      /* Markdown headings used for section titles (####) */
      .stMarkdown h4 {{
        font-size: 18px !important;
        margin: 0 0 0.35rem 0 !important;
      }}

      .card {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 14px;
        padding: 18px;
      }}

      .title {{
        font-size: 24px;
        font-weight: 700;
        margin: 0 0 12px 0;
        color: {TEXT};
      }}

      .sub {{
        margin: 0 0 14px 0;
        color: {MUTED};
        font-size: 16px;
      }}

      /* Inputs */
      [data-testid="stFileUploader"] section {{
        border-radius: 12px;
        padding: 12px;
        border: 1px solid {BORDER};
        background: {CARD};
      }}

      /* File uploader text + button contrast */
      [data-testid="stFileUploader"] * {{
        color: {TEXT} !important;
      }}
      [data-testid="stFileUploader"] small {{
        color: {MUTED} !important;
      }}
      [data-testid="stFileUploader"] button {{
        background: {BTN_BG} !important;
        color: {BTN_TEXT} !important;
        border: 1px solid {BORDER} !important;
        border-radius: 12px !important;
        padding: 0.55rem 0.85rem !important;
        font-weight: 700 !important;
      }}
      /* File uploader "remove" (X) button */
      [data-testid="stFileUploaderDeleteBtn"] button {{
        background: {BTN_BG} !important;
        color: {BTN_TEXT} !important;
        border: 1px solid {BORDER} !important;
      }}
      [data-testid="stFileUploaderDeleteBtn"] svg {{
        fill: {BTN_TEXT} !important;
        color: {BTN_TEXT} !important;
      }}

      [role="radiogroup"] {{
        border-radius: 12px;
        padding: 12px 12px 8px 12px;
        border: 1px solid {BORDER};
        background: transparent;
      }}

      input, textarea {{
        background: {CARD} !important;
        color: {TEXT} !important;
        border-color: {BORDER} !important;
      }}
      input::placeholder, textarea::placeholder {{
        color: {MUTED} !important;
      }}

      /* Buttons */
      div.stButton > button {{
        width: 100%;
        border-radius: 12px !important;
        padding: 0.60rem 0.90rem !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        background: {BTN_BG} !important;
        color: {BTN_TEXT} !important;
        border: 1px solid {BORDER} !important;
      }}

      div.stDownloadButton > button {{
        width: 100%;
        border-radius: 12px !important;
        padding: 0.60rem 0.90rem !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        background: {BTN_BG} !important;
        color: {BTN_TEXT} !important;
        border: 1px solid {BORDER} !important;
      }}

      /* Progress bar */
      div[data-testid="stProgress"] > div > div {{
        background-color: {PROG} !important;
      }}

      /* Text colors for markdown */
      .stMarkdown, .stMarkdown p, .stCaption, label {{
        color: {TEXT} !important;
      }}
      .stCaption {{
        color: {MUTED} !important;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)

def classify_uploads(uploads) -> Dict[str, object]:
    analysis = []
    wh = []
    mkz = []
    osv = []
    other = []

    for u in uploads:
        name = u.name or ""
        lower = name.lower()
        if lower.startswith("_анализ"):
            analysis.append(u)
        elif lower.startswith("wh_kz"):
            wh.append(u)
        elif lower.startswith("m_kz"):
            mkz.append(u)
        else:
            ext = lower.rsplit(".", 1)[-1] if "." in lower else ""
            if ext in ("xlsx", "xlsm", "xls"):
                osv.append(u)
            else:
                other.append(u)

    return {"analysis": analysis, "wh": wh, "m": mkz, "osv": osv, "other": other}


st.markdown("#### Загрузка")
upload_types = ["xlsx", "xlsm"] + (["xls"] if os.name == "nt" else [])
uploads = st.file_uploader(
    "Загрузите файлы Excel",
    type=upload_types,
    accept_multiple_files=True,
    label_visibility="collapsed",
    key="uploads",
)

if not uploads:
    st.stop()

if os.name != "nt":
    st.caption(".xls на Streamlit Cloud (Linux) не поддерживается. Используй .xlsx/.xlsm или запускай локально на Windows для авто-конвертации.")

if any(is_xls_filename(u.name) for u in uploads):
    st.info("Обнаружены .xls. Перед обработкой они будут конвертированы в .xlsx через установленный Microsoft Excel (может занять время).")

sig = tuple(sorted((u.name, getattr(u, "size", None)) for u in uploads))
if st.session_state.get("upload_sig") != sig:
    st.session_state["upload_sig"] = sig
    for k in ["prepared_bytes", "prepared_name", "availability", "prep_report", "processed_bytes"]:
        st.session_state.pop(k, None)

cls = classify_uploads(uploads)
analysis_list: List = cls["analysis"]
wh_list: List = cls["wh"]
m_list: List = cls["m"]
osv_list: List = cls["osv"]

if cls["other"]:
    st.warning("Некоторые файлы не подходят (нужны .xlsx/.xlsm). Они будут проигнорированы.")

if len(analysis_list) > 1:
    st.error("Файл _Анализ должен быть один.")
    st.stop()

analysis_u = analysis_list[0] if analysis_list else None

if (wh_list and not m_list) or (m_list and not wh_list):
    st.error("Для сценариев с WH_KZ/M_KZ нужно загрузить оба файла: WH_KZ и M_KZ.")
    st.stop()

needs_new_name = analysis_u is None
analysis_title = ""
if needs_new_name:
    st.markdown("#### Имя _Анализ")
    analysis_title = st.text_input("Название", value="", placeholder="Например: Алтын Мунай ЛТД 170740018427")

out_name_preview = analysis_u.name if analysis_u else safe_filename(f"_Анализ {analysis_title}".strip() + ".xlsx")
st.caption("Выходной файл: " + out_name_preview)

prep_disabled = needs_new_name and not analysis_title.strip()

# Prefix questions for OSV duplicates (per uploaded OSV sheet)
osv_prefix_by_sheet: Dict[Tuple[str, str, int], str] = {}
wh_prefix_by_file: Dict[str, str] = {}
m_prefix_by_file: Dict[str, str] = {}

xls_cache = st.session_state.setdefault("xls_cache", {})

analysis_wb_tmp = None
if analysis_u is not None:
    try:
        an_name, an_bytes = ensure_openpyxl_bytes(analysis_u.name, analysis_u.getvalue(), cache=xls_cache)
        analysis_wb_tmp, _ = load_wb_from_bytes(an_bytes, an_name)
    except Exception:
        analysis_wb_tmp = None

existing_suffixes: Set[str] = set()
if analysis_wb_tmp is not None:
    for sh in analysis_wb_tmp.sheetnames:
        existing_suffixes.add(split_prefix_suffix4(sh)[1])

    existing_saldo_info = list_existing_saldo_sheets_with_a1(analysis_wb_tmp)
    if existing_saldo_info:
        st.markdown("#### _Анализ: найденные ОСВ")
        for sh, suf, a1 in existing_saldo_info[:12]:
            tail = f" | {_short(a1)}" if a1 else ""
            st.caption(f"{sh} → {suf}{tail}")
        if len(existing_saldo_info) > 12:
            st.caption(f"... и еще {len(existing_saldo_info) - 12}")

# Detect all uploaded OSV sheets with 4-digit account numbers.
osv_items: List[Tuple[str, str, int, str, str]] = []
for u in osv_list:
    try:
        u_name, u_bytes = ensure_openpyxl_bytes(u.name, u.getvalue(), cache=xls_cache)
        wb_tmp, _ = load_wb_from_bytes(u_bytes, u_name)
        for idx, ws in enumerate(wb_tmp.worksheets):
            acc = get_account_number(ws)
            if acc and acc.isdigit() and len(acc) == 4:
                company = _short(extract_company_label_from_a1(ws))
                osv_items.append((u.name, ws.title, idx, acc, company))
    except Exception:
        pass

acc_to_items: Dict[str, List[Tuple[str, str, int]]] = defaultdict(list)
for fname, sheet_title, idx, acc, _company in osv_items:
    acc_to_items[acc].append((fname, sheet_title, idx))

need_prefix_items: List[Tuple[str, str, int, str]] = []
for acc, items in acc_to_items.items():
    if len(items) > 1 or (analysis_wb_tmp is not None and acc in existing_suffixes):
        for fname, sheet_title, idx in items:
            need_prefix_items.append((fname, sheet_title, idx, acc))

if need_prefix_items:
    st.markdown("#### Префиксы ОСВ (повторы/конфликт)")
    st.caption("Обнаружены повторяющиеся ОСВ по счетам или такие счета уже есть в _Анализ. Для КАЖДОГО ОСВ укажи префикс или отметь «без префикса».")

    # Validate: within same account, the same prefix cannot be used twice; also disallow multiple 'без префикса'.
    chosen_by_acc: Dict[str, List[str]] = defaultdict(list)
    for fname, sheet_title, idx, acc in need_prefix_items:
        company = ""
        for f2, s2, i2, a2, comp2 in osv_items:
            if f2 == fname and s2 == sheet_title and i2 == idx and a2 == acc:
                company = comp2
                break
        if company:
            st.caption(f"{fname} / {sheet_title} → {acc} | {company}")
        else:
            st.caption(f"{fname} / {sheet_title} → {acc}")
        no_pref = st.checkbox("без префикса", value=False, key=f"osv_nopref::{fname}::{idx}")
        pref = st.text_input("префикс", value="", key=f"osv_pref::{fname}::{idx}", disabled=no_pref)
        if no_pref:
            osv_prefix_by_sheet[(fname, sheet_title, idx)] = ""
            chosen_by_acc[acc].append("")
        else:
            if not pref.strip():
                prep_disabled = True
                st.caption("Нужен префикс или отметь «без префикса».")
            else:
                p = pref.strip()
                osv_prefix_by_sheet[(fname, sheet_title, idx)] = p
                chosen_by_acc[acc].append(p)

    for acc, prefs in chosen_by_acc.items():
        if len(prefs) != len(set(prefs)):
            prep_disabled = True
            st.caption(f"Для счета {acc} префиксы должны быть уникальными (включая «без префикса»).")

# WH/M prefixes when multiple files are uploaded
if len(wh_list) > 1:
    st.markdown("#### Префиксы WH_KZ")
    st.caption("Загружено несколько WH_KZ. Для каждого укажи префикс или «без префикса» (префиксы должны быть уникальными).")
    seen: Set[str] = set()
    for u in wh_list:
        no_pref = st.checkbox(f"{u.name}: без префикса", value=False, key=f"wh_nopref::{u.name}")
        pref = st.text_input(f"{u.name}: префикс", value="", key=f"wh_pref::{u.name}", disabled=no_pref)
        p = "" if no_pref else pref.strip()
        if not no_pref and not p:
            prep_disabled = True
            st.caption(f"{u.name}: нужен префикс или «без префикса».")
        if p in seen:
            prep_disabled = True
            st.caption("Префиксы WH_KZ должны быть уникальными.")
        seen.add(p)
        wh_prefix_by_file[u.name] = p

if len(m_list) > 1:
    st.markdown("#### Префиксы M_KZ")
    st.caption("Загружено несколько M_KZ. Для каждого укажи префикс или «без префикса» (префиксы должны быть уникальными).")
    seen: Set[str] = set()
    for u in m_list:
        no_pref = st.checkbox(f"{u.name}: без префикса", value=False, key=f"m_nopref::{u.name}")
        pref = st.text_input(f"{u.name}: префикс", value="", key=f"m_pref::{u.name}", disabled=no_pref)
        p = "" if no_pref else pref.strip()
        if not no_pref and not p:
            prep_disabled = True
            st.caption(f"{u.name}: нужен префикс или «без префикса».")
        if p in seen:
            prep_disabled = True
            st.caption("Префиксы M_KZ должны быть уникальными.")
        seen.add(p)
        m_prefix_by_file[u.name] = p

prep_btn = st.button("Собрать _Анализ", disabled=prep_disabled)

if prep_btn:
    status = st.empty()
    status.info("Сборка…")
    try:
        analysis_file = None
        if analysis_u:
            an_name, an_bytes = ensure_openpyxl_bytes(analysis_u.name, analysis_u.getvalue(), cache=xls_cache)
            analysis_file = (an_name, an_bytes)

        wh_files = []
        for u in wh_list:
            n, b = ensure_openpyxl_bytes(u.name, u.getvalue(), cache=xls_cache)
            wh_files.append((n, b, wh_prefix_by_file.get(u.name, "")))

        m_files = []
        for u in m_list:
            n, b = ensure_openpyxl_bytes(u.name, u.getvalue(), cache=xls_cache)
            m_files.append((n, b, m_prefix_by_file.get(u.name, "")))

        osv_files = []
        for u in osv_list:
            n, b = ensure_openpyxl_bytes(u.name, u.getvalue(), cache=xls_cache)
            osv_files.append((u.name, b))

        out_bytes, out_name, availability, prep_report = build_analysis_workbook(
            analysis_file=analysis_file,
            wh_files=wh_files,
            m_files=m_files,
            osv_files=osv_files,
            analysis_name_for_new=analysis_title,
            osv_prefix_by_sheet=osv_prefix_by_sheet,
        )

        st.session_state["prepared_bytes"] = out_bytes
        st.session_state["prepared_name"] = out_name
        st.session_state["availability"] = availability
        st.session_state["prep_report"] = prep_report
        st.session_state.pop("processed_bytes", None)
        status.success("Готово.")
    except Exception as e:
        status.error(f"Ошибка сборки: {e}")

if "prepared_bytes" in st.session_state:
    prep_report = st.session_state.get("prep_report") or {}
    if prep_report.get("warnings"):
        for w in prep_report["warnings"]:
            st.warning(w)

    st.write("")
    st.markdown("#### Обработка")

    availability = st.session_state.get("availability") or {}
    inv_map = availability.get("inventory_map") or {"1310": False, "1320": False, "1330": False}
    saldo_ok = bool(availability.get("saldo_ok"))
    contracts_ok = bool(availability.get("contracts_ok"))

    st.checkbox("Очистить ОСВ", value=False, disabled=True, help="Скоро")

    opt_saldo = st.checkbox("Сальдо", value=False, disabled=(not saldo_ok))
    if not saldo_ok:
        st.caption("Сальдо недоступно: не найдены листы, заканчивающиеся на 1210/1710/3310/3510.")

    opt_contracts = st.checkbox("Контракты", value=False, disabled=(not contracts_ok))
    if not contracts_ok:
        st.caption("Контракты недоступны: не найдены пары листов *Wd/*Md.")

    inv_available_any = any(bool(v) for v in inv_map.values())
    opt_inventory = st.checkbox("Запасы", value=False, disabled=(not inv_available_any))
    if not inv_available_any:
        st.caption("Запасы недоступны: не найдены листы по счетам 1310/1320/1330.")

    inventory_accounts: List[str] = []
    if opt_inventory:
        st.caption("Счета запасов")
        inv_1310 = st.checkbox("1310", value=False, disabled=(not inv_map.get("1310", False)))
        inv_1320 = st.checkbox("1320", value=False, disabled=(not inv_map.get("1320", False)))
        inv_1330 = st.checkbox("1330", value=False, disabled=(not inv_map.get("1330", False)))

        if inv_1310:
            inventory_accounts.append("1310")
        if inv_1320:
            inventory_accounts.append("1320")
        if inv_1330:
            inventory_accounts.append("1330")

        if not inventory_accounts:
            st.caption("Выберите минимум один счет: 1310 / 1320 / 1330.")

    selected_modes: List[str] = []
    if opt_saldo:
        selected_modes.append("Сальдо")
    if opt_contracts:
        selected_modes.append("Контракты")

    st.write("")
    st.markdown("#### Запуск")

    has_any_mode = bool(selected_modes) or opt_inventory
    inventory_ok = (not opt_inventory) or bool(inventory_accounts)
    run_btn = st.button("Обработать", disabled=((not has_any_mode) or (not inventory_ok)))

    status_box = st.empty()
    progress = st.progress(0)

    if run_btn:
        try:
            status_box.info("Подготовка…")
            progress.progress(10)
            time.sleep(0.05)

            out_bytes = st.session_state["prepared_bytes"]

            if "Сальдо" in selected_modes:
                status_box.info("Обработка: Сальдо…")
                progress.progress(35)
                out_bytes = run_code_1(out_bytes)
                progress.progress(55)

            if "Контракты" in selected_modes:
                status_box.info("Обработка: Контракты…")
                progress.progress(60)
                out_bytes = run_code_2(out_bytes)
                progress.progress(80)

            if opt_inventory:
                status_box.info("Обработка: Запасы…")
                progress.progress(82)
                out_bytes, inv_report = run_code_3_inventory(out_bytes, inventory_accounts)
                progress.progress(95)
                if inv_report.get("missing_sheets") or inv_report.get("missing_markers"):
                    parts = []
                    if inv_report.get("missing_sheets"):
                        parts.append("не найдены листы: " + ", ".join(inv_report["missing_sheets"]))
                    if inv_report.get("missing_markers"):
                        parts.append("в листах не найден счет в колонке A")
                    st.warning("Запасы: " + "; ".join(parts))

            st.session_state["processed_bytes"] = out_bytes
            status_box.success("Готово.")
            progress.progress(100)
        except Exception as e:
            progress.progress(0)
            status_box.error(f"Ошибка: {e}")

    st.write("")
    st.markdown("#### Скачать")
    download_bytes = st.session_state.get("processed_bytes") or st.session_state.get("prepared_bytes")
    download_name = st.session_state.get("prepared_name") or "output.xlsx"

    if not has_any_mode:
        confirm_merge = st.checkbox("Только объединить (без обработок)", value=False)
        st.download_button(
            label="Скачать",
            data=download_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=(not confirm_merge),
        )
    else:
        st.download_button(
            label="Скачать",
            data=download_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
