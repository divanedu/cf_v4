import re
import io
import time
import os
import tempfile
from copy import copy as _copy
from collections import defaultdict
from typing import Callable, Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string


# =========================
# Вспомогательные функции
# =========================
def safe_sheet_name(name: str) -> str:
    """Правила Excel для имени листа: максимум 31 символ, нельзя : \\ / ? * [ ]"""
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
    """(префикс, последние 4 символа)"""
    if len(sheet_name) < 4:
        return sheet_name, ""
    return sheet_name[:-4], sheet_name[-4:].lower()


def split_prefix_suffix2(sheet_name: str) -> Tuple[str, str]:
    """(префикс, последние 2 символа)"""
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
    Возвращает уникальное имя листа, которое ВСЕГДА заканчивается на `suffix` (например, '1210' или 'Wd').
    Если `prefix+suffix` уже существует, вставляет счётчик перед suffix: f'{prefix}{i}{suffix}'.
    """
    prefix = (prefix or "").strip()
    suffix = (suffix or "").strip()
    base = safe_sheet_name(prefix + suffix)
    if base.endswith(suffix) and base not in wb.sheetnames:
        return base

    i = 1
    while True:
        mid = str(i)
        # Следим за ограничениями Excel: общая длина <= 31 и окончание = суффикс (важно для логики поиска листов).
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
    Конвертирует байты .xls в байты .xlsx с помощью установленного Microsoft Excel (COM).
    Работает только на Windows, где установлен Excel и доступен pywin32.
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
            # 51 = xlOpenXMLWorkbook (.xlsx) — формат сохранения для Excel COM
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
    Гарантирует, что на выходе будут байты, которые читает openpyxl (xlsx/xlsm).
    Если загружен .xls — конвертируем в .xlsx (локально на Windows через Excel COM).
    `cache` ключуется по (filename, len(bytes)), чтобы не конвертировать один и тот же файл повторно.
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

    # Быстрый путь: копируем только реально существующие ячейки (значения/стили),
    # чтобы не сканировать огромные "пустые" прямоугольники.
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


def excel_col_width_from_pixels(px: float) -> float:
    # Приблизительный перевод пикселей в "ширину колонки" Excel (по умолчанию).
    # Часто используют оценку: пиксели ~= width*7 + 5  => width ~= (px-5)/7
    try:
        px = float(px)
    except Exception:
        px = 48.0
    return max(0.0, round((px - 5.0) / 7.0, 2))


def set_all_columns_width(ws, px: float):
    width = excel_col_width_from_pixels(px)
    max_col = ws.max_column or 1
    # Ограничение на случай, если max_column "раздулся" из-за форматирования.
    max_col = min(max_col, 2000)
    for i in range(1, max_col + 1):
        col = get_column_letter(i)
        dim = ws.column_dimensions[col]
        dim.width = width


REGISTRY_COL_WIDTH_PX = 48.0
# Высота строк в Excel задаётся в пунктах (points). Требование: фиксированная высота (без авто-роста).
REGISTRY_ROW_HEIGHT_PT = 28.8


def format_registry_sheet(ws) -> None:
    """
    Реестр (Wr/Mr): фиксируем ширину всех колонок и высоту всех строк,
    чтобы Excel не "раздувал" строки из-за wrap_text.
    """
    set_all_columns_width(ws, REGISTRY_COL_WIDTH_PX)
    try:
        ws.sheet_format.defaultRowHeight = REGISTRY_ROW_HEIGHT_PT
    except Exception:
        pass

    rows: Set[int] = set(getattr(ws, "row_dimensions", {}).keys())
    for (r, _c), cell in getattr(ws, "_cells", {}).items():
        rows.add(int(r))
        # Принудительно отключаем перенос текста, иначе Excel может визуально увеличивать высоту строки.
        try:
            al = cell.alignment
            if al and getattr(al, "wrap_text", None):
                cell.alignment = Alignment(
                    horizontal=al.horizontal,
                    vertical=al.vertical,
                    textRotation=al.textRotation,
                    wrapText=False,
                    shrinkToFit=al.shrinkToFit,
                    indent=al.indent,
                    relativeIndent=al.relativeIndent,
                    justifyLastLine=al.justifyLastLine,
                    readingOrder=al.readingOrder,
                )
        except Exception:
            pass

    # Ставим высоту только тем строкам, которые реально существуют (в row_dimensions или есть ячейки),
    # чтобы не гонять цикл 1..max_row на больших листах.
    for r in rows:
        try:
            ws.row_dimensions[r].height = REGISTRY_ROW_HEIGHT_PT
        except Exception:
            pass


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
            dst_ws = copy_sheet(m_wb[src_title], analysis_wb, new_title)
            if dst_title.lower() == "mr":
                format_registry_sheet(dst_ws)
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
            dst_ws = copy_sheet(wh_wb[src_title], analysis_wb, new_title)
            if dst_title.lower() == "wr":
                format_registry_sheet(dst_ws)
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
    То же, что merge_wh_m_into_analysis, но принудительно добавляет префикс к импортируемым листам.
    Важно: сохраняем ожидаемые суффиксы (Wd/Md и т.п.), чтобы блок "Контракты" корректно находил пары листов.
    """
    report = {"missing_wh": [], "missing_m": [], "copied": []}
    p = (prefix or "").strip()

    if m_bytes:
        m_wb, _ = load_wb_from_bytes(m_bytes, m_name)
        mapping_m = {"Итог": "M", "Контрагенты": "Mt", "Реестр": "Mr", "Договоры": "Md"}
        for src_title, base in mapping_m.items():
            if src_title not in m_wb.sheetnames:
                report["missing_m"].append(src_title)
                continue
            new_title = make_unique_sheet_title(analysis_wb, f"{p}{base}")
            # Суффикс Md должен остаться Md, иначе не соберём пары Wd/Md при создании "контр".
            if base.lower().endswith("md"):
                new_title = make_unique_with_fixed_suffix(analysis_wb, p + base[:-2], "Md")
            dst_ws = copy_sheet(m_wb[src_title], analysis_wb, new_title)
            if base.lower() == "mr":
                format_registry_sheet(dst_ws)
            report["copied"].append(f"{m_name}:{src_title} -> {new_title}")

    if wh_bytes:
        wh_wb, _ = load_wb_from_bytes(wh_bytes, wh_name)
        mapping_wh = {"Итог": "W", "Таблицы": "Wt", "Реестр": "Wr", "Договоры": "Wd"}
        for src_title, base in mapping_wh.items():
            if src_title not in wh_wb.sheetnames:
                report["missing_wh"].append(src_title)
                continue
            new_title = make_unique_sheet_title(analysis_wb, f"{p}{base}")
            if base.lower().endswith("wd"):
                new_title = make_unique_with_fixed_suffix(analysis_wb, p + base[:-2], "Wd")
            dst_ws = copy_sheet(wh_wb[src_title], analysis_wb, new_title)
            if base.lower() == "wr":
                format_registry_sheet(dst_ws)
            report["copied"].append(f"{wh_name}:{src_title} -> {new_title}")

        if "Кредиты" in wh_wb.sheetnames:
            ws_cred = wh_wb["Кредиты"]
            if ws_cred["AC6"].value not in (None, 0, 0.0, "0"):
                new_title = make_unique_sheet_title(analysis_wb, f"{p}кред")
                copy_sheet(ws_cred, analysis_wb, new_title)
                report["copied"].append(f"{wh_name}:Кредиты -> {new_title}")

    return report


# =========================
# Очистка ОСВ (автоматически для ОСВ с "рандомным" именем файла)
# =========================
MAX_CELLS_PER_SHEET = 300_000
SORT_ACCOUNTS = {"1310", "1320", "1330"}  # только эти счета сортируем по колонке G
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

    # Если 2-я строка часть объединённой шапки, а значение лежит в 1-й строке — берём левую-верхнюю ячейку merge-диапазона.
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
        # Приоритет: 4 цифры сразу после слова "счет".
        idx = s_low.find("счет")
        if idx != -1:
            tail = s_norm[idx:]
            tail_compact = tail.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")
            m = re.search(r"(\\d{4})", tail_compact)
            if m:
                return m.group(1)
        # Фолбэк: любые 4 подряд идущие цифры (предварительно "сжимаем" пробелы между цифрами).
        compact = s_norm.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")
        m = re.search(r"(\\d{4})", compact)
        if m:
            return m.group(1)
        # Последний шанс: берём первые 4 цифры из строки, где оставили только цифры.
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
        # Некоторые ОСВ выгрузки пишут "Итого:" или "Итого ..." в колонке A.
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
        # Лист не выкидываем: номер счета мы всё равно знаем, и дальше он нужен в _Анализ.
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
                # Не выбрасываем лист: добавляем в _Анализ, но называем нейтрально, чтобы не ломать сборку.
                fallback = make_unique_sheet_title(analysis_wb, "OSV")
                copy_sheet(ws, analysis_wb, fallback)
                report["skipped"].append(f"{fname}:{ws.title} (не найден счет во 2-й строке)")
                report["added"].append(f"{fname}:{ws.title} -> {fallback}")
            else:
                # Суффикс счета должен остаться (важно: Сальдо ищет *1210/*1710/*3310/*3510 по названию листа).
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

    # ОСВ "общ" (казахстанская ОСВ) — для неё нужен справочник "Счета каз"
    kaz_ok = ("общ" in wb.sheetnames) and ("Счета каз" in wb.sheetnames)

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
        "kaz_ok": kaz_ok,
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
    """Возвращает кортежи (имя_листа, счет_суффикс, A1_текст) для ОСВ-счетов сальдо."""
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
    progress_cb: Optional[Callable[[float, str], None]] = None,
) -> Tuple[bytes, str, Dict[str, object], Dict[str, List[str]]]:
    report: Dict[str, List[str]] = {"warnings": [], "copied": []}

    def _progress(frac: float, msg: str) -> None:
        if progress_cb:
            try:
                progress_cb(max(0.0, min(1.0, float(frac))), msg)
            except Exception:
                pass

    total_steps = 2 + (1 if analysis_file else 0) + len(wh_files) + len(m_files) + len(osv_files)
    done_steps = 0

    def _step(msg: str) -> None:
        nonlocal done_steps
        done_steps += 1
        _progress(done_steps / max(1, total_steps), f"Сборка ({done_steps}/{total_steps}): {msg}")

    _progress(0.0, "Сборка: старт…")

    if analysis_file:
        analysis_name, analysis_bytes = analysis_file
        analysis_wb, keep_vba = load_wb_from_bytes(analysis_bytes, analysis_name)
        out_name = analysis_name
        placeholder_title = None
        _step(f"Открываю {analysis_name}")
    else:
        analysis_wb = Workbook()
        # Держим видимый "служебный" лист, чтобы сохранение не падало с
        # "At least one sheet must be visible" даже если все входные листы пропущены/пустые.
        placeholder_ws = analysis_wb.active
        placeholder_title = "Сборка"
        placeholder_ws.title = placeholder_title
        placeholder_ws["A1"] = "Служебный лист. Удалится автоматически, если добавятся другие листы."
        keep_vba = False
        base = (analysis_name_for_new or "").strip()
        out_name = safe_filename(f"_Анализ {base}".strip() + ".xlsx")
        _step("Создаю новый _Анализ")

    if wh_files or m_files:
        missing_wh_all: List[str] = []
        missing_m_all: List[str] = []

        for wh_name, wh_bytes, pref in wh_files:
            merge_report = merge_wh_m_into_analysis_with_prefix(analysis_wb, wh_bytes, wh_name, None, "", prefix=pref)
            missing_wh_all.extend([f"{wh_name}:{x}" for x in merge_report["missing_wh"]])
            report["copied"].extend(merge_report["copied"])
            _step(f"WH_KZ: {wh_name} (листов: {len(merge_report.get('copied') or [])})")

        for m_name, m_bytes, pref in m_files:
            merge_report = merge_wh_m_into_analysis_with_prefix(analysis_wb, None, "", m_bytes, m_name, prefix=pref)
            missing_m_all.extend([f"{m_name}:{x}" for x in merge_report["missing_m"]])
            report["copied"].extend(merge_report["copied"])
            _step(f"M_KZ: {m_name} (листов: {len(merge_report.get('copied') or [])})")

        if missing_wh_all or missing_m_all:
            msg = []
            if missing_wh_all:
                msg.append("WH: " + ", ".join(missing_wh_all[:6]) + (" ..." if len(missing_wh_all) > 6 else ""))
            if missing_m_all:
                msg.append("M: " + ", ".join(missing_m_all[:6]) + (" ..." if len(missing_m_all) > 6 else ""))
            report["warnings"].append("WH/M: не найдены листы. " + " | ".join(msg))

    if osv_files:
        osv_prefix_by_sheet = osv_prefix_by_sheet or {}
        # Для ОСВ-счетов (1210/1710/3310/3510) префиксы задаём по каждому листу, чтобы избежать коллизий,
        # но при этом сохранить суффикс счета в названии (это критично для дальнейшей обработки).
        report_local = {"added": [], "skipped": []}
        for fname, fbytes in osv_files:
            _step(f"ОСВ: очищаю {fname}")
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

    # Удаляем служебный лист, если в книге уже появился хотя бы один "нормальный" лист.
    if placeholder_title and len(analysis_wb.sheetnames) > 1 and placeholder_title in analysis_wb.sheetnames:
        del analysis_wb[placeholder_title]

    # Гарантируем, что есть хотя бы один видимый лист (требование openpyxl при сохранении).
    if not analysis_wb.sheetnames:
        ws = analysis_wb.create_sheet("Сборка")
        ws["A1"] = "Пустой файл: не удалось добавить ни одного листа."
    visible = [ws for ws in analysis_wb.worksheets if getattr(ws, "sheet_state", "visible") == "visible"]
    if not visible:
        analysis_wb.worksheets[0].sheet_state = "visible"

    availability = compute_availability_from_wb(analysis_wb)

    _step("Сохраняю файл…")
    out = io.BytesIO()
    analysis_wb.save(out)
    return out.getvalue(), out_name, availability, report


# =========================
# CODE 1 (Сальдо) — несколько компаний через префиксы
# Ищем листы, у которых последние 4 символа = 1210/1710/3310/3510
# Создаём лист на каждый префикс: "<префикс>сальд" или "сальд" (если префикса нет)
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

    def _excel_sheet_ref(sheet_name: str) -> str:
        # Если нужно — оборачиваем имя листа в кавычки и экранируем апострофы (правила Excel).
        if sheet_name is None:
            return "''"
        s = str(sheet_name)
        needs = any(ch in s for ch in " []()!,'\"") or (" " in s)
        if "'" in s:
            s = s.replace("'", "''")
            needs = True
        return f"'{s}'" if needs else s

    def _find_prefixed_sheetname(prefix_norm: str, base2: str) -> Optional[str]:
        """
        Ищет лист вида "<префикс>Wr" / "<префикс>Mr" (регистр не важен),
        также допускается суффикс " (n)". Возвращает лучший кандидат или None.
        """
        base2_l = base2.lower()
        best = None
        best_key = None
        for name in wb.sheetnames:
            nm = str(name)
            nm2 = re.sub(r"\s*\(\d+\)$", "", nm).strip()
            if not nm2.lower().endswith(base2_l):
                continue
            pref_part = nm2[: -len(base2_l)]
            if normalize_prefix(pref_part) != prefix_norm:
                continue
            m = re.search(r"\((\d+)\)\s*$", nm)
            n = int(m.group(1)) if m else 0
            key = (1 if m else 0, n, len(nm))  # предпочитаем точное имя без " (n)"
            if best is None or key < best_key:
                best = nm
                best_key = key
        return best

    def _top15_pos_neg(df: pd.DataFrame, col: str) -> pd.DataFrame:
        if df is None or df.empty or col not in df.columns:
            return df.iloc[0:0].copy()
        pos = df[df[col] > 0].nlargest(15, col) if (df[col] > 0).any() else df.iloc[0:0]
        neg = df[df[col] < 0].nsmallest(15, col) if (df[col] < 0).any() else df.iloc[0:0]
        out = pd.concat([pos, neg], ignore_index=True)
        return out.reset_index(drop=True)

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

        # Переименование по требованию:
        # - "сальд" (базовый лист) теперь называется "сальд (2)"
        # - "сальд (2)" (расширенный лист) теперь называется "сальд"
        out_sheet_name = safe_sheet_name(f"{prefix}сальд (2)" if prefix else "сальд (2)")
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

        # =========================
        # Лист "сальд (2)" — топ-15 (плюс/минус) + блоки оплат/выполнений с формулами по месяцам
        # =========================
        out2_name = safe_sheet_name(f"{prefix}сальд" if prefix else "сальд")
        if out2_name in wb.sheetnames:
            wb.remove(wb[out2_name])
        ws2 = wb.create_sheet(out2_name)

        ws2["A1"] = "Все значения указаны в тысячах тенге"
        ws2["A1"].font = Font(name="Calibri", size=10, bold=True)

        # На "сальд (2)" везде используем Calibri.
        font_h = Font(name="Calibri", size=10, bold=True)
        font_b = Font(name="Calibri", size=10)
        font_bb = Font(name="Calibri", size=10, bold=True)
        align_c = Alignment(horizontal="center")
        align_l = Alignment(horizontal="left")
        num_fmt = "#,##0;[Red](#,##0)"

        # A2 = 6 (влево + мягкая серая заливка)
        ws2["A2"] = 6
        ws2["A2"].alignment = Alignment(horizontal="left")
        ws2["A2"].fill = PatternFill("solid", fgColor="D9D9D9")

        ws2["F2"] = "коммент"

        # Используем конкатенацию через &, чтобы Excel не добавлял неявное пересечение ("@")
        ws2["G2"] = '="Опл L"&$A$2&"M"'
        ws2["H2"] = '="Вып L"&$A$2&"M"'

        ws2["P1"] = "Оплаты"
        ws2["AO1"] = "Выполнения"
        ws2["P1"].font = font_h
        ws2["AO1"].font = font_h

        # Заголовки месяцев
        months = [f"2025_{m:02d}" for m in range(1, 13)] + [f"2026_{m:02d}" for m in range(1, 13)]
        for i, m in enumerate(months):
            ws2.cell(row=2, column=10 + i, value=m).font = font_h  # J..AG
            ws2.cell(row=2, column=10 + i, value=m).alignment = align_c
            ws2.cell(row=2, column=35 + i, value=m).font = font_h  # AI..BF
            ws2.cell(row=2, column=35 + i, value=m).alignment = align_c

        # Заголовки таблиц
        ws2["B2"] = "Контрагент"
        ws2["C2"] = "1210"
        ws2["D2"] = "3510"
        ws2["E2"] = "сальдо с заказчиками"
        for c in ("B2", "C2", "D2", "E2", "F2", "G2", "H2"):
            ws2[c].font = font_h
            ws2[c].alignment = align_c if c != "B2" else align_l

        # Блоки "Поставщики" и "Общее сальдо" размещаем ниже динамически (так как добавляем строки сумм/итого).

        # Подбираем Wr/Mr для текущего префикса (если такие листы есть) — имена нужны внутри формул.
        pref_norm = normalize_prefix(prefix)
        wr_name = _find_prefixed_sheetname(pref_norm, "Wr") or (f"{prefix}Wr" if prefix else "Wr")
        mr_name = _find_prefixed_sheetname(pref_norm, "Mr") or (f"{prefix}Mr" if prefix else "Mr")
        wr_ref = _excel_sheet_ref(wr_name)
        mr_ref = _excel_sheet_ref(mr_name)

        def _comment_formula_for_row(r: int, which: str) -> str:
            """
            which: 'cust' or 'supp'
            Возвращает Excel-формулу для колонки F на основе:
            - оплат J:AG
            - выполнений AI:BF (нужно только для правила по 3510)
            """
            last_pos = f"IFERROR(LOOKUP(2,1/($J{r}:$AG{r}<>0),COLUMN($J{r}:$AG{r}))-COLUMN($J{r})+1,0)"
            months_since_last_pay = (
                f"IF(({last_pos})=0,999,"
                f"DATEDIF("
                f"DATE(LEFT(INDEX($J$2:$AG$2,1,({last_pos})),4),RIGHT(INDEX($J$2:$AG$2,1,({last_pos})),2),1),"
                f"DATE(YEAR(TODAY()),MONTH(TODAY()),1),\"m\"))"
            )

            def _rule_by_last_pay(t_ok: int, t_warn: int, warn_label: str) -> str:
                return (
                    f"IF(({last_pos})=0,\"списание\","
                    f"IF(({months_since_last_pay})<={t_ok},\"ОК\","
                    f"IF(({months_since_last_pay})<={t_warn},\"{warn_label}\",\"списание\")))"
                )

            idx_now_pay = "IFERROR(MATCH(TEXT(TODAY(),\"yyyy\")&\"_\"&TEXT(TODAY(),\"mm\"),$J$2:$AG$2,0),COUNTA($J$2:$AG$2))"
            idx_now_perf = "IFERROR(MATCH(TEXT(TODAY(),\"yyyy\")&\"_\"&TEXT(TODAY(),\"mm\"),$AI$2:$BF$2,0),COUNTA($AI$2:$BF$2))"
            sum_pay_0_3 = f"SUM(INDEX($J{r}:$AG{r},1,MAX(1,({idx_now_pay})-2)):INDEX($J{r}:$AG{r},1,({idx_now_pay})))"
            sum_perf_0_3 = f"SUM(INDEX($AI{r}:$BF{r},1,MAX(1,({idx_now_perf})-2)):INDEX($AI{r}:$BF{r},1,({idx_now_perf})))"
            sum_pay_0_12 = f"SUM(INDEX($J{r}:$AG{r},1,MAX(1,({idx_now_pay})-11)):INDEX($J{r}:$AG{r},1,({idx_now_pay})))"
            sum_perf_0_12 = f"SUM(INDEX($AI{r}:$BF{r},1,MAX(1,({idx_now_perf})-11)):INDEX($AI{r}:$BF{r},1,({idx_now_perf})))"
            rule_3510 = (
                f"IF(AND(({sum_pay_0_3})<>0,({sum_perf_0_3})<>0),\"ОК\","
                f"IF(OR(({sum_pay_0_12})<>0,({sum_perf_0_12})<>0),\"сомнительно\",\"списание\"))"
            )

            if which == "cust":
                rule_1210 = _rule_by_last_pay(3, 6, "сомнительный")
                return f"=IF($B{r}=\"\",\"\",IF($E{r}>0,{rule_1210},{rule_3510}))"
            # поставщики
            rule_1710 = _rule_by_last_pay(3, 6, "сомнительный")
            rule_3310 = _rule_by_last_pay(3, 12, "сомнительный")
            return f"=IF($B{r}=\"\",\"\",IF($E{r}>0,{rule_1710},{rule_3310}))"

        def _rollup_pay_formula(r: int) -> str:
            return (
                f'=IF($B{r}="","",SUM('
                f'INDEX($J{r}:$AG{r},1,MAX(1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),$J$2:$AG$2,0),COUNTA($J$2:$AG$2))-$A$2+1)):'
                f'INDEX($J{r}:$AG{r},1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),$J$2:$AG$2,0),COUNTA($J$2:$AG$2)))))'
            )

        def _rollup_perf_formula(r: int) -> str:
            return (
                f'=IF($B{r}="","",SUM('
                f'INDEX($AI{r}:$BF{r},1,MAX(1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),$AI$2:$BF$2,0),COUNTA($AI$2:$BF$2))-$A$2+1)):'
                f'INDEX($AI{r}:$BF{r},1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),$AI$2:$BF$2,0),COUNTA($AI$2:$BF$2)))))'
            )

        def _write_customer_row(r: int, contr: str, v1210: int, v3510: int, vsaldo: int) -> None:
            ws2.cell(row=r, column=2, value=contr).font = font_b
            ws2.cell(row=r, column=2).alignment = align_l

            ws2.cell(row=r, column=3, value=v1210).number_format = num_fmt
            ws2.cell(row=r, column=4, value=v3510).number_format = num_fmt
            ws2.cell(row=r, column=5, value=vsaldo).number_format = num_fmt
            ws2.cell(row=r, column=3).font = font_b
            ws2.cell(row=r, column=4).font = font_b
            # Для строк с контрагентами (топ-15) числа НЕ делаем жирными.
            ws2.cell(row=r, column=5).font = font_b
            for c in (3, 4, 5):
                ws2.cell(row=r, column=c).alignment = align_c

            ws2.cell(row=r, column=6, value=_comment_formula_for_row(r, "cust")).alignment = align_l
            ws2.cell(row=r, column=6).font = font_b

            ws2.cell(row=r, column=7, value=_rollup_pay_formula(r)).alignment = align_c
            ws2.cell(row=r, column=7).font = font_b
            ws2.cell(row=r, column=7).number_format = num_fmt

            ws2.cell(row=r, column=8, value=_rollup_perf_formula(r)).alignment = align_c
            ws2.cell(row=r, column=8).font = font_b
            ws2.cell(row=r, column=8).number_format = num_fmt

            # J..AG: оплаты из Wr
            for ci in range(10, 34):  # J..AG
                col_letter = get_column_letter(ci)
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",'
                        f'SUMIFS({wr_ref}!$F:$F,{wr_ref}!$G:$G,"1210",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$R:$R,$B{r})'
                        f'+SUMIFS({wr_ref}!$F:$F,{wr_ref}!$G:$G,"3510",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$R:$R,$B{r})'
                        f')'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

            # AI..AT (2025) *1.12 и AU..BF (2026) *1.16 — выполнения из Mr
            for ci in range(35, 47):  # AI..AT
                col_letter = get_column_letter(ci)
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",SUMIFS({mr_ref}!$H:$H,{mr_ref}!$P:$P,{col_letter}$2,{mr_ref}!$Q:$Q,$B{r},{mr_ref}!$G:$G,"6010")*1.12)'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

            for ci in range(47, 59):  # AU..BF
                col_letter = get_column_letter(ci)
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",SUMIFS({mr_ref}!$H:$H,{mr_ref}!$P:$P,{col_letter}$2,{mr_ref}!$Q:$Q,$B{r},{mr_ref}!$G:$G,"6010")*1.16)'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

        def _write_supplier_row(r: int, contr: str, v1710: int, v3310: int, vsaldo: int) -> None:
            ws2.cell(row=r, column=2, value=contr).font = font_b
            ws2.cell(row=r, column=2).alignment = align_l

            ws2.cell(row=r, column=3, value=v1710).number_format = num_fmt
            ws2.cell(row=r, column=4, value=v3310).number_format = num_fmt
            ws2.cell(row=r, column=5, value=vsaldo).number_format = num_fmt
            ws2.cell(row=r, column=3).font = font_b
            ws2.cell(row=r, column=4).font = font_b
            # Для строк с контрагентами (топ-15) числа НЕ делаем жирными.
            ws2.cell(row=r, column=5).font = font_b
            for c in (3, 4, 5):
                ws2.cell(row=r, column=c).alignment = align_c

            ws2.cell(row=r, column=6, value=_comment_formula_for_row(r, "supp")).alignment = align_l
            ws2.cell(row=r, column=6).font = font_b

            ws2.cell(row=r, column=7, value=_rollup_pay_formula(r)).alignment = align_c
            ws2.cell(row=r, column=7).font = font_b
            ws2.cell(row=r, column=7).number_format = num_fmt

            # Для поставщиков: только оплаты J..AG (выполнений нет)
            for ci in range(10, 34):  # J..AG
                col_letter = get_column_letter(ci)
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",'
                        f'SUMIFS({wr_ref}!$H:$H,{wr_ref}!$E:$E,"1710",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$Q:$Q,$B{r})'
                        f'+SUMIFS({wr_ref}!$H:$H,{wr_ref}!$E:$E,"3310",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$Q:$Q,$B{r})'
                        f')'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

        def _write_summary_rows(kind: str, pos_start: int, pos_end: int, sum_row: int, other_row: int, total_row: int) -> None:
            """
            Writes TOP-15 sum, прочее, ИТОГО rows for a block.
            kind: 'cust' or 'supp'
            """
            # Подписи строк
            # Жирным выделяем только строки "ТОП-15" и "ИТОГО"; "прочее" оставляем обычным.
            ws2.cell(row=sum_row, column=2, value="ТОП-15").font = font_h
            ws2.cell(row=other_row, column=2, value="прочее").font = font_b
            ws2.cell(row=total_row, column=2, value="ИТОГО").font = font_h
            for rr in (sum_row, other_row, total_row):
                ws2.cell(row=rr, column=2).alignment = align_l

            def _sum_formula(col_letter: str) -> str:
                return f"=SUM({col_letter}{pos_start}:{col_letter}{pos_end})"

            # Суммы по ТОП-15 (C/D/E, G/H и по месяцам)
            for col_letter in ("C", "D", "E", "G"):
                c = ws2[f"{col_letter}{sum_row}"]
                c.value = _sum_formula(col_letter)
                c.font = font_bb
                c.alignment = align_c
                c.number_format = num_fmt

            if kind == "cust":
                c = ws2[f"H{sum_row}"]
                c.value = _sum_formula("H")
                c.font = font_bb
                c.alignment = align_c
                c.number_format = num_fmt

            for ci in range(10, 34):  # J..AG
                col_letter = get_column_letter(ci)
                c = ws2.cell(row=sum_row, column=ci, value=_sum_formula(col_letter))
                c.alignment = align_c
                c.number_format = num_fmt
                c.font = font_bb
            if kind == "cust":
                for ci in range(35, 59):  # AI..BF
                    col_letter = get_column_letter(ci)
                    c = ws2.cell(row=sum_row, column=ci, value=_sum_formula(col_letter))
                    c.alignment = align_c
                    c.number_format = num_fmt
                    c.font = font_bb

            # ИТОГО:
            # - по сальдо (C/D/E) пишем числом (посчитано в pandas)
            # - оплаты/выполнения по месяцам считаем формулами без критерия контрагента
            if kind == "cust":
                tot_1210 = int(round(float(df_cust["1210"].sum() if not df_cust.empty else 0.0)))
                tot_3510 = int(round(float(df_cust["3510"].sum() if not df_cust.empty else 0.0)))
                tot_saldo = int(round(float(df_cust["сальдо заказчики"].sum() if not df_cust.empty else 0.0)))
                ws2[f"C{total_row}"] = tot_1210
                ws2[f"D{total_row}"] = tot_3510
                ws2[f"E{total_row}"] = tot_saldo
                for addr in (f"C{total_row}", f"D{total_row}", f"E{total_row}"):
                    ws2[addr].font = font_bb
                    ws2[addr].alignment = align_c
                    ws2[addr].number_format = num_fmt

                # ИТОГО по месяцам (оплаты Wr, выполнения Mr) без фильтра по контрагенту
                for ci in range(10, 34):  # J..AG
                    col_letter = get_column_letter(ci)
                    ws2.cell(
                        row=total_row,
                        column=ci,
                        value=(
                            f'=SUMIFS({wr_ref}!$F:$F,{wr_ref}!$G:$G,"1210",{wr_ref}!$P:$P,{col_letter}$2)'
                            f'+SUMIFS({wr_ref}!$F:$F,{wr_ref}!$G:$G,"3510",{wr_ref}!$P:$P,{col_letter}$2)'
                        ),
                    ).number_format = num_fmt
                    ws2.cell(row=total_row, column=ci).alignment = align_c
                    ws2.cell(row=total_row, column=ci).font = font_bb

                for ci in range(35, 47):  # AI..AT
                    col_letter = get_column_letter(ci)
                    ws2.cell(
                        row=total_row,
                        column=ci,
                        value=f'=SUMIFS({mr_ref}!$H:$H,{mr_ref}!$P:$P,{col_letter}$2,{mr_ref}!$G:$G,"6010")*1.12',
                    ).number_format = num_fmt
                    ws2.cell(row=total_row, column=ci).alignment = align_c
                    ws2.cell(row=total_row, column=ci).font = font_bb
                for ci in range(47, 59):  # AU..BF
                    col_letter = get_column_letter(ci)
                    ws2.cell(
                        row=total_row,
                        column=ci,
                        value=f'=SUMIFS({mr_ref}!$H:$H,{mr_ref}!$P:$P,{col_letter}$2,{mr_ref}!$G:$G,"6010")*1.16',
                    ).number_format = num_fmt
                    ws2.cell(row=total_row, column=ci).alignment = align_c
                    ws2.cell(row=total_row, column=ci).font = font_bb

                # G/H суммы за "последние A2 месяцев" для строки ИТОГО
                ws2.cell(row=total_row, column=7, value=_rollup_pay_formula(total_row)).alignment = align_c
                ws2.cell(row=total_row, column=7).number_format = num_fmt
                ws2.cell(row=total_row, column=7).font = font_bb
                ws2.cell(row=total_row, column=8, value=_rollup_perf_formula(total_row)).alignment = align_c
                ws2.cell(row=total_row, column=8).number_format = num_fmt
                ws2.cell(row=total_row, column=8).font = font_bb

            else:
                tot_1710 = int(round(float(df_supp["1710"].sum() if not df_supp.empty else 0.0)))
                tot_3310 = int(round(float(df_supp["3310"].sum() if not df_supp.empty else 0.0)))
                tot_saldo = int(round(float(df_supp["сальдо поставщики"].sum() if not df_supp.empty else 0.0)))
                ws2[f"C{total_row}"] = tot_1710
                ws2[f"D{total_row}"] = tot_3310
                ws2[f"E{total_row}"] = tot_saldo
                for addr in (f"C{total_row}", f"D{total_row}", f"E{total_row}"):
                    ws2[addr].font = font_bb
                    ws2[addr].alignment = align_c
                    ws2[addr].number_format = num_fmt

                for ci in range(10, 34):  # J..AG
                    col_letter = get_column_letter(ci)
                    ws2.cell(
                        row=total_row,
                        column=ci,
                        value=(
                            f'=SUMIFS({wr_ref}!$H:$H,{wr_ref}!$E:$E,"1710",{wr_ref}!$P:$P,{col_letter}$2)'
                            f'+SUMIFS({wr_ref}!$H:$H,{wr_ref}!$E:$E,"3310",{wr_ref}!$P:$P,{col_letter}$2)'
                        ),
                    ).number_format = num_fmt
                    ws2.cell(row=total_row, column=ci).alignment = align_c
                    ws2.cell(row=total_row, column=ci).font = font_bb

                ws2.cell(row=total_row, column=7, value=_rollup_pay_formula(total_row)).alignment = align_c
                ws2.cell(row=total_row, column=7).number_format = num_fmt
                ws2.cell(row=total_row, column=7).font = font_bb

            # прочее = ИТОГО - ТОП-15 (для чисел и для помесячных колонок)
            for col_letter in ("C", "D", "E", "G"):
                ws2[f"{col_letter}{other_row}"] = f"={col_letter}{total_row}-{col_letter}{sum_row}"
                ws2[f"{col_letter}{other_row}"].font = font_b
                ws2[f"{col_letter}{other_row}"].alignment = align_c
                ws2[f"{col_letter}{other_row}"].number_format = num_fmt

            if kind == "cust":
                ws2[f"H{other_row}"] = f"=H{total_row}-H{sum_row}"
                ws2[f"H{other_row}"].font = font_b
                ws2[f"H{other_row}"].alignment = align_c
                ws2[f"H{other_row}"].number_format = num_fmt

            for ci in range(10, 34):  # J..AG
                col_letter = get_column_letter(ci)
                ws2.cell(row=other_row, column=ci, value=f"={col_letter}{total_row}-{col_letter}{sum_row}").number_format = num_fmt
                ws2.cell(row=other_row, column=ci).alignment = align_c
            if kind == "cust":
                for ci in range(35, 59):  # AI..BF
                    col_letter = get_column_letter(ci)
                    ws2.cell(row=other_row, column=ci, value=f"={col_letter}{total_row}-{col_letter}{sum_row}").number_format = num_fmt
                    ws2.cell(row=other_row, column=ci).alignment = align_c

        # === Заказчики (фиксированная раскладка + строки ТОП-15/прочее/ИТОГО)
        cust_pos = df_cust[df_cust["сальдо заказчики"] > 0].nlargest(15, "сальдо заказчики") if not df_cust.empty else df_cust
        cust_neg = df_cust[df_cust["сальдо заказчики"] < 0].nsmallest(15, "сальдо заказчики") if not df_cust.empty else df_cust

        cust_pos_start, cust_pos_end = 3, 17
        cust_pos_sum, cust_pos_other, cust_pos_total = 18, 19, 20
        cust_gap = 21
        cust_neg_start, cust_neg_end = 22, 36
        cust_neg_sum, cust_neg_other, cust_neg_total = 37, 38, 39
        cust_end_gap = 40

        for i in range(15):
            if cust_pos is None or i >= len(cust_pos.index):
                continue
            rr = cust_pos_start + i
            r0 = cust_pos.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_customer_row(
                rr,
                contr,
                int(round(float(r0.get("1210", 0) or 0))),
                int(round(float(r0.get("3510", 0) or 0))),
                int(round(float(r0.get("сальдо заказчики", 0) or 0))),
            )

        _write_summary_rows("cust", cust_pos_start, cust_pos_end, cust_pos_sum, cust_pos_other, cust_pos_total)

        for i in range(15):
            if cust_neg is None or i >= len(cust_neg.index):
                continue
            rr = cust_neg_start + i
            r0 = cust_neg.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_customer_row(
                rr,
                contr,
                int(round(float(r0.get("1210", 0) or 0))),
                int(round(float(r0.get("3510", 0) or 0))),
                int(round(float(r0.get("сальдо заказчики", 0) or 0))),
            )

        _write_summary_rows("cust", cust_neg_start, cust_neg_end, cust_neg_sum, cust_neg_other, cust_neg_total)

        # === Поставщики (ниже блока заказчиков)
        supp_header_row = cust_end_gap + 2  # пустая строка после блока заказчиков
        ws2[f"B{supp_header_row}"] = "Контрагент"
        ws2[f"C{supp_header_row}"] = "1710"
        ws2[f"D{supp_header_row}"] = "3310"
        ws2[f"E{supp_header_row}"] = "сальдо с поставщиками"
        for caddr in (f"B{supp_header_row}", f"C{supp_header_row}", f"D{supp_header_row}", f"E{supp_header_row}"):
            ws2[caddr].font = font_h
            ws2[caddr].alignment = align_l if caddr.startswith("B") else align_c

        supp_pos = df_supp[df_supp["сальдо поставщики"] > 0].nlargest(15, "сальдо поставщики") if not df_supp.empty else df_supp
        supp_neg = df_supp[df_supp["сальдо поставщики"] < 0].nsmallest(15, "сальдо поставщики") if not df_supp.empty else df_supp

        supp_pos_start, supp_pos_end = supp_header_row + 1, supp_header_row + 15
        supp_pos_sum, supp_pos_other, supp_pos_total = supp_pos_end + 1, supp_pos_end + 2, supp_pos_end + 3
        supp_neg_start, supp_neg_end = supp_pos_end + 5, supp_pos_end + 19
        supp_neg_sum, supp_neg_other, supp_neg_total = supp_neg_end + 1, supp_neg_end + 2, supp_neg_end + 3

        for i in range(15):
            if supp_pos is None or i >= len(supp_pos.index):
                continue
            rr = supp_pos_start + i
            r0 = supp_pos.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_supplier_row(
                rr,
                contr,
                int(round(float(r0.get("1710", 0) or 0))),
                int(round(float(r0.get("3310", 0) or 0))),
                int(round(float(r0.get("сальдо поставщики", 0) or 0))),
            )
        _write_summary_rows("supp", supp_pos_start, supp_pos_end, supp_pos_sum, supp_pos_other, supp_pos_total)

        for i in range(15):
            if supp_neg is None or i >= len(supp_neg.index):
                continue
            rr = supp_neg_start + i
            r0 = supp_neg.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_supplier_row(
                rr,
                contr,
                int(round(float(r0.get("1710", 0) or 0))),
                int(round(float(r0.get("3310", 0) or 0))),
                int(round(float(r0.get("сальдо поставщики", 0) or 0))),
            )
        _write_summary_rows("supp", supp_neg_start, supp_neg_end, supp_neg_sum, supp_neg_other, supp_neg_total)

        # === Общее сальдо (ниже поставщиков): только значение сальдо + строки ТОП-15/прочее/ИТОГО
        total_header_row = supp_neg_total + 5  # пустая строка после блока поставщиков
        ws2[f"B{total_header_row}"] = "Контрагент"
        ws2[f"E{total_header_row}"] = "общее сальдо"
        ws2[f"B{total_header_row}"].font = font_h
        ws2[f"E{total_header_row}"].font = font_h
        ws2[f"B{total_header_row}"].alignment = align_l
        ws2[f"E{total_header_row}"].alignment = align_c

        total_pos = df_total[df_total["общее сальдо"] > 0].nlargest(15, "общее сальдо") if not df_total.empty else df_total
        total_neg = df_total[df_total["общее сальдо"] < 0].nsmallest(15, "общее сальдо") if not df_total.empty else df_total

        total_pos_start, total_pos_end = total_header_row + 1, total_header_row + 15
        total_pos_sum, total_pos_other, total_pos_total = total_pos_end + 1, total_pos_end + 2, total_pos_end + 3
        total_neg_start, total_neg_end = total_pos_end + 5, total_pos_end + 19
        total_neg_sum, total_neg_other, total_neg_total = total_neg_end + 1, total_neg_end + 2, total_neg_end + 3

        def _write_total_row(r: int, contr: str, vtot: int) -> None:
            ws2.cell(row=r, column=2, value=contr).font = font_b
            ws2.cell(row=r, column=2).alignment = align_l
            # Для строк с контрагентами (топ-15) числа НЕ делаем жирными.
            ws2.cell(row=r, column=5, value=vtot).font = font_b
            ws2.cell(row=r, column=5).alignment = align_c
            ws2.cell(row=r, column=5).number_format = num_fmt

        for i in range(15):
            if total_pos is None or i >= len(total_pos.index):
                continue
            rr = total_pos_start + i
            r0 = total_pos.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_total_row(rr, contr, int(round(float(r0.get("общее сальдо", 0) or 0))))

        # Сводные строки (положительный блок)
        ws2.cell(row=total_pos_sum, column=2, value="ТОП-15").font = font_h
        ws2.cell(row=total_pos_other, column=2, value="прочее").font = font_h
        ws2.cell(row=total_pos_total, column=2, value="ИТОГО").font = font_h
        for rr in (total_pos_sum, total_pos_other, total_pos_total):
            ws2.cell(row=rr, column=2).alignment = align_l

        ws2[f"E{total_pos_sum}"] = f"=SUM(E{total_pos_start}:E{total_pos_end})"
        ws2[f"E{total_pos_sum}"].font = font_bb
        ws2[f"E{total_pos_sum}"].alignment = align_c
        ws2[f"E{total_pos_sum}"].number_format = num_fmt

        total_all = int(round(float(df_total["общее сальдо"].sum() if not df_total.empty else 0.0)))
        ws2[f"E{total_pos_total}"] = total_all
        ws2[f"E{total_pos_total}"].font = font_bb
        ws2[f"E{total_pos_total}"].alignment = align_c
        ws2[f"E{total_pos_total}"].number_format = num_fmt

        ws2[f"E{total_pos_other}"] = f"=E{total_pos_total}-E{total_pos_sum}"
        ws2[f"E{total_pos_other}"].font = font_b
        ws2[f"E{total_pos_other}"].alignment = align_c
        ws2[f"E{total_pos_other}"].number_format = num_fmt

        for i in range(15):
            if total_neg is None or i >= len(total_neg.index):
                continue
            rr = total_neg_start + i
            r0 = total_neg.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_total_row(rr, contr, int(round(float(r0.get("общее сальдо", 0) or 0))))

        # Сводные строки (отрицательный блок)
        ws2.cell(row=total_neg_sum, column=2, value="ТОП-15").font = font_h
        ws2.cell(row=total_neg_other, column=2, value="прочее").font = font_h
        ws2.cell(row=total_neg_total, column=2, value="ИТОГО").font = font_h
        for rr in (total_neg_sum, total_neg_other, total_neg_total):
            ws2.cell(row=rr, column=2).alignment = align_l

        ws2[f"E{total_neg_sum}"] = f"=SUM(E{total_neg_start}:E{total_neg_end})"
        ws2[f"E{total_neg_sum}"].font = font_bb
        ws2[f"E{total_neg_sum}"].alignment = align_c
        ws2[f"E{total_neg_sum}"].number_format = num_fmt

        ws2[f"E{total_neg_total}"] = total_all
        ws2[f"E{total_neg_total}"].font = font_bb
        ws2[f"E{total_neg_total}"].alignment = align_c
        ws2[f"E{total_neg_total}"].number_format = num_fmt

        ws2[f"E{total_neg_other}"] = f"=E{total_neg_total}-E{total_neg_sum}"
        ws2[f"E{total_neg_other}"].font = font_b
        ws2[f"E{total_neg_other}"].alignment = align_c
        ws2[f"E{total_neg_other}"].number_format = num_fmt

        # Минимальные ширины колонок для читаемости
        ws2.column_dimensions["A"].width = 6
        ws2.column_dimensions["B"].width = 35
        for col in ["C", "D", "E", "F", "G", "H"]:
            ws2.column_dimensions[col].width = 18

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 2 (Контракты) — несколько компаний через префиксы (Wd/Md, регистр не важен)
# Создаём лист на каждый префикс: "<префикс>контр" или "контр" (если префикса нет)
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
        row = start_row
        for key in all_keys:
            name, contract = key

            py = payments_year.get(key, [0, 0, 0])
            pf = performance_year.get(key, [0, 0, 0])

            # Убираем строки, где одновременно:
            # - F (итого оплаты) = 0
            # - J (итого выполнения) = 0
            # Проверяем по исходным суммам (а не по формуле), чтобы не зависеть от пересчёта Excel.
            if abs(float(py[0]) + float(py[1]) + float(py[2])) < 1e-9 and abs(float(pf[0]) + float(pf[1]) + float(pf[2])) < 1e-9:
                continue

            ws[f"A{row}"] = name
            ws[f"B{row}"] = contract

            ws[f"C{row}"] = py[0]
            ws[f"D{row}"] = py[1]
            ws[f"E{row}"] = py[2]

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

            row += 1

        last_row = row - 1 if row > start_row else 2

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

    # После формирования листа(ов) "контр" удаляем исходные листы Wd/Md из выходного файла.
    for sh in sorted(source_sheets_to_delete):
        if sh in wb.sheetnames:
            del wb[sh]

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 3 (Запасы)
# Для каждого выбранного счета (1310/1320/1330):
# - находим лист(ы), где в названии встречается номер счета
# - в листе ищем строку в колонке A, где ячейка равна номеру счета
# - над блоком записываем пороги в J..N
# - ниже заполняем формулы J..N до первой пустой ячейки в колонке G
# Отдельных "отчётных" листов не создаём: UI показывает предупреждения по результату.
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
        # Фиксируем размер шрифта для всех чисел J..N (включая суммы), чтобы Excel не "плясал" стилями.
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

            # Пороговые значения (в Excel проценты храним как доли: 2.0 = 200%).
            _set_percent(ws, thr_low, 10, 2.0)   # J: 200%
            _set_percent(ws, thr_low, 11, 1.0)   # K: 100%
            _set_percent(ws, thr_high, 11, 2.0)  # K: 200%
            _set_percent(ws, thr_low, 12, 0.5)   # L: 50%
            _set_percent(ws, thr_high, 12, 1.0)  # L: 100%
            _set_percent(ws, thr_low, 13, 0.25)  # M: 25%
            _set_percent(ws, thr_high, 13, 0.5)  # M: 50%
            _set_percent(ws, thr_low, 14, 0.0)   # N: 0%
            _set_percent(ws, thr_high, 14, 0.25) # N: 25%

            # Заполняем формулы по строкам запасов: от строки ниже маркера до первой пустой в колонке G.
            start_row = found_row + 1
            last = start_row - 1
            for r in range(start_row, ws.max_row + 1):
                g = ws.cell(row=r, column=7).value
                if _is_blank(g):
                    break
                last = r

            if last >= start_row:
                _fill_formulas(ws, start_row, last, thr_low, thr_high)
                # Суммы ставим в строке, где найден номер счета (та же строка, что A=1310/1320/1330).
                total_fill = PatternFill(fill_type="solid", fgColor="C6EFCE")
                for col_letter, col_idx in [("J", 10), ("K", 11), ("L", 12), ("M", 13), ("N", 14)]:
                    cell = ws.cell(row=found_row, column=col_idx)
                    cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{last})"
                    _set_number_style(cell, bold=True, fill=total_fill)

                # Пунктирная рамка вокруг мини-таблицы (пороги + суммы + строки данных).
                _apply_dotted_grid(ws, thr_low, last, 10, 14)
                processed.append(f"{account}: {sh} (строки {start_row}-{last})")
            else:
                # Если данных нет — рамку всё равно рисуем вокруг порогов + строки сумм.
                _apply_dotted_grid(ws, thr_low, found_row, 10, 14)
                processed.append(f"{account}: {sh} (нет строк с данными в G ниже маркера)")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), {"processed": processed, "missing_sheets": missing_sheets, "missing_markers": missing_markers}


# =========================
# CODE 4 (Обработка общей ОСВ / Казахстан)
# Работает по листу "общ" и использует справочник "Счета каз".
# ВАЖНО: мы записываем именно формулы (строки начинающиеся с "="),
# поэтому в выходном Excel они будут видны как формулы.
# =========================
def run_code_4_obsh_kaz(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))

    if "общ" not in wb.sheetnames:
        raise ValueError("Обработка общей ОСВ: не найден лист «общ».")
    if "Счета каз" not in wb.sheetnames:
        raise ValueError("Обработка общей ОСВ: не найден лист «Счета каз».")

    ws = wb["общ"]

    # На листе "общ" иногда есть объединённые ячейки (merged). openpyxl возвращает для них MergedCell,
    # и попытка записать .value в такую ячейку падает (read-only). Поэтому все записи делаем "безопасно":
    # если целевая ячейка входит в merged-диапазон — записываем в левую-верхнюю ячейку этого диапазона.
    merged_ranges = list(getattr(ws.merged_cells, "ranges", []) or [])

    def _merge_topleft(row: int, col: int) -> Optional[Tuple[int, int]]:
        for mr in merged_ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                return (mr.min_row, mr.min_col)
        return None

    def _safe_set_value(row: int, col: int, value) -> None:
        cell = ws.cell(row=row, column=col)
        if isinstance(cell, MergedCell):
            tl = _merge_topleft(row, col)
            if not tl:
                return
            ws.cell(row=tl[0], column=tl[1]).value = value
            return
        cell.value = value

    # Ищем строку "Итого" по колонке A.
    # Важно: в ОСВ могут встречаться промежуточные "Итого" внутри разделов,
    # поэтому берём ПОСЛЕДНЕЕ вхождение в пределах разумного диапазона.
    itogo_row = None
    max_scan = min(max(ws.max_row, 1), 5000)
    for r in range(8, max_scan + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        s = str(v).strip()
        if "Итого" in s:
            itogo_row = r
    if itogo_row is None:
        raise ValueError("Обработка общей ОСВ: не найдена строка «Итого» в колонке A (до 5000 строки).")

    last_data_row = itogo_row - 1
    if last_data_row < 8:
        raise ValueError("Обработка общей ОСВ: нет строк данных (ожидается минимум с 8 строки).")

    # Родительские строки определяем по жирному шрифту в колонке A (8..last_data_row).
    parent_rows: List[int] = []
    for r in range(8, last_data_row + 1):
        cell = ws.cell(row=r, column=1)
        if cell.font and bool(cell.font.bold):
            parent_rows.append(r)

    # Формулы: K, L, M, N, P (O пропускаем).
    for r in range(8, last_data_row + 1):
        _safe_set_value(r, 11, f"=G{r}-C{r}")  # K
        _safe_set_value(r, 12, f"=H{r}-D{r}")  # L
        _safe_set_value(r, 13, f"=L{r}-K{r}")  # M
        _safe_set_value(r, 14, f"=IFERROR(INDEX('Счета каз'!B:B,MATCH(LEFT(A{r},4),'Счета каз'!A:A,0)),0)")  # N
        _safe_set_value(r, 16, f"=IFERROR(INDEX('Счета каз'!C:C,MATCH(LEFT(A{r},4),'Счета каз'!A:A,0)),0)")  # P

    # Визуальный стиль как в шаблоне (обычно Aptos Display 9).
    base_font = Font(name="Aptos Display", size=9)
    base_font_bold = Font(name="Aptos Display", size=9, bold=True)
    right = Alignment(horizontal="right", vertical="center")

    # Заголовок P7
    _safe_set_value(7, 16, "ЧОК")
    p7 = ws.cell(row=7, column=16)
    if isinstance(p7, MergedCell):
        tl = _merge_topleft(7, 16)
        if tl:
            p7 = ws.cell(row=tl[0], column=tl[1])
    p7.font = base_font_bold
    p7.alignment = right

    # Форматирование диапазонов
    num_fmt = "#,##0"

    # K:M
    for row in ws.iter_rows(min_row=8, max_row=last_data_row, min_col=11, max_col=13):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            cell.font = base_font
            cell.alignment = right
            cell.number_format = num_fmt

    # N и P — шрифт + число + выравнивание (как в шаблоне)
    for row in ws.iter_rows(min_row=8, max_row=last_data_row, min_col=14, max_col=14):
        if not isinstance(row[0], MergedCell):
            row[0].font = base_font
            row[0].alignment = right
            row[0].number_format = num_fmt
    for row in ws.iter_rows(min_row=8, max_row=last_data_row, min_col=16, max_col=16):
        if not isinstance(row[0], MergedCell):
            row[0].font = base_font
            row[0].alignment = right
            row[0].number_format = num_fmt

    # Границы
    def _side(color: str) -> Side:
        return Side(style="thin", color=color)

    def _set_border(cell, left=None, right=None, top=None, bottom=None):
        cell.border = Border(
            left=_side(left) if left else None,
            right=_side(right) if right else None,
            top=_side(top) if top else None,
            bottom=_side(bottom) if bottom else None,
        )

    C_GREEN = "ACC8BD"
    C_GRAY = "A0A0A0"
    C_BLACK = "000000"

    # K8:P(last_data_row) — внешние + внутренние линии
    min_row, max_row = 8, last_data_row
    min_col, max_col = 11, 16
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            left_col = C_GREEN if c == min_col else C_GREEN  # inside vertical тоже зеленая
            right_col = C_GREEN if c == max_col else C_GREEN
            top_col = C_GREEN if r == min_row else C_GREEN   # inside horizontal тоже зеленая
            bottom_col = C_GREEN if r == max_row else C_GREEN
            cell.border = Border(
                left=_side(left_col),
                right=_side(right_col),
                top=_side(top_col),
                bottom=_side(bottom_col),
            )

    # В шаблоне границы на K7:P7 и колонке J могут быть заданы через темы/условное форматирование.
    # Здесь не навязываем эти линии, чтобы не "ломать" вид.

    # Сводная таблица: через одну строку после "Итого" (одна пустая строка между ними)
    ss = itogo_row + 2
    labels = [
        "ЧП", "Аморт", "Запасы", "Производство", "Налоги",
        "незакр прибыль", "незакр расходы", "ФОТ", "Прочее",
        "ДЗ", "Авансы выданные", "КЗ", "Авансы полученные", "прочее дз", "прочее кз",
        "ОС", "Финвложения", "Капитал", "Лизинг", "Дивики", "Долг",
        "CFO", "CFI", "CFF", "CF",
    ]
    bold_labels = {"ЧП", "CFO", "CFI", "CFF", "CF"}
    cf_labels = {"CFO", "CFI", "CFF", "CF"}

    row_of = {label: ss + i for i, label in enumerate(labels)}

    def _cfo_formula() -> str:
        cats = [
            "ЧП", "Аморт", "Запасы", "Производство", "Налоги",
            "ФОТ", "Прочее", "ДЗ", "Авансы выданные", "КЗ",
            "Авансы полученные", "прочее дз", "прочее кз",
        ]
        inner = ",".join(f"C{row_of[c]}" for c in cats)
        return f"=SUM({inner})+C{row_of['незакр прибыль']}+C{row_of['незакр расходы']}"

    cf_formulas = {
        "CFO": _cfo_formula(),
        "CFI": f"=C{row_of['ОС']}",
        "CFF": f"=C{row_of['Финвложения']}+C{row_of['Долг']}+C{row_of['Дивики']}+C{row_of['Лизинг']}+C{row_of['Капитал']}",
        "CF": f"=SUM(C{row_of['ЧП']}:C{row_of['Долг']})",
    }

    num_fmt_cf = "#,##0;(#,##0);-"
    # Сводная таблица в шаблоне также на Aptos Display 9
    arial9_bold = base_font_bold

    for label in labels:
        rr = row_of[label]
        b = ws.cell(row=rr, column=2)  # B
        c = ws.cell(row=rr, column=3)  # C

        _safe_set_value(rr, 2, label)
        b = ws.cell(row=rr, column=2)
        if isinstance(b, MergedCell):
            tl = _merge_topleft(rr, 2)
            if tl:
                b = ws.cell(row=tl[0], column=tl[1])
        b.font = base_font_bold if label in bold_labels else base_font
        b.number_format = num_fmt_cf

        if label in cf_labels:
            _safe_set_value(rr, 3, cf_formulas[label])
        else:
            _safe_set_value(rr, 3, f"=+SUMIFS($M:$M,$N:$N,$B{rr})")
        c = ws.cell(row=rr, column=3)
        if isinstance(c, MergedCell):
            tl = _merge_topleft(rr, 3)
            if tl:
                c = ws.cell(row=tl[0], column=tl[1])
        c.font = arial9_bold
        c.alignment = right
        c.number_format = num_fmt_cf

    # Линии: под CFF и над CF (B:C)
    r_cff = row_of["CFF"]
    r_cf = row_of["CF"]
    for col in (2, 3):
        ws.cell(row=r_cff, column=col).border = Border(bottom=_side(C_BLACK))
        ws.cell(row=r_cf, column=col).border = Border(top=_side(C_BLACK))

    # Высота строк: 12pt для 6..Итого
    # Требование: весь лист с высотой строк 12pt.
    try:
        ws.sheet_format.defaultRowHeight = 12
    except Exception:
        pass

    # Если в файле есть явно заданные высоты строк — переопределяем их на 12.
    for r in list(ws.row_dimensions.keys()):
        ws.row_dimensions[r].height = 12

    # Для типичных ОСВ размер листа небольшой, можно проставить высоту всем строкам до max_row.
    # Но если max_row раздулся форматированием, не делаем дорогой цикл на десятки тысяч строк.
    max_row = int(ws.max_row or 1)
    if max_row <= 8000:
        for r in range(1, max_row + 1):
            ws.row_dimensions[r].height = 12

    # Группировка/сворачивание: скрываем дочерние строки между "родителями"
    try:
        ws.sheet_format.outlineLevelRow = 1
        ws.sheet_properties.outlinePr.summaryBelow = False
    except Exception:
        pass

    for i, prow in enumerate(parent_rows):
        child_start = prow + 1
        child_end = (parent_rows[i + 1] if i + 1 < len(parent_rows) else itogo_row) - 1
        if child_start <= child_end:
            for rr in range(child_start, child_end + 1):
                ws.row_dimensions[rr].outlineLevel = 1
                ws.row_dimensions[rr].hidden = True

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# Интерфейс — стили и тема (без градиентов, фиксированный шрифт)
# =========================
st.set_page_config(page_title="", page_icon=None, layout="wide", initial_sidebar_state="collapsed")

# Переключатель темы: на странице, максимально ненавязчиво (справа сверху).
_tcol_l, _tcol_r = st.columns([0.93, 0.07])
with _tcol_r:
    # Без подписи: только тумблер. Подсказка при наведении.
    light_mode = st.toggle("", value=False, key="light_theme", help="Переключить тему (светлая/тёмная)", label_visibility="collapsed")

if light_mode:
    # Светлая тема (инверсия)
    BG = "#FFFFFF"
    TEXT = "#0B0B0B"
    CARD = "#FFFFFF"
    BORDER = "#D0D0D0"
    MUTED = "#4D4D4D"
    BTN_BG = "#FFFFFF"
    BTN_TEXT = "#0B0B0B"
    PROG = "#2F5FD7"
else:
    # Тёмная тема: чёрный фон, белый текст (как просили)
    BG = "#000000"
    TEXT = "#FFFFFF"
    CARD = "#000000"
    BORDER = "#3A3A3A"
    MUTED = "#CFCFCF"
    BTN_BG = "#000000"
    BTN_TEXT = "#FFFFFF"
    PROG = "#FFFFFF"

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

      /* Checkbox/toggle accents: make sure they are visible on both themes */
      [data-testid="stCheckbox"] * {{
        color: {TEXT} !important;
      }}
      input[type="checkbox"], input[type="radio"] {{
        accent-color: {TEXT} !important;
      }}

      /* Theme toggle: компактный, без лишних отступов */
      div[data-testid="stToggle"] {{
        padding-top: 0.15rem !important;
      }}
      div[data-testid="stToggle"] > label {{
        justify-content: flex-end !important;
        gap: 0.25rem !important;
      }}
      /* Track / knob colors */
      div[data-testid="stToggle"] svg {{
        color: {TEXT} !important;
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

# Запрос префиксов для дубликатов ОСВ (на каждый загруженный лист ОСВ)
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

# Находим все загруженные ОСВ-листы, у которых определился 4-значный номер счета.
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

    # Валидация: в рамках одного счета нельзя использовать один и тот же префикс дважды;
    # "без префикса" тоже должен быть максимум один раз на счет.
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

# Префиксы для WH_KZ / M_KZ, если загружено несколько файлов
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
    prog = st.progress(0)

    def _ui_progress(frac: float, msg: str) -> None:
        status.info(msg)
        try:
            prog.progress(int(round(frac * 100)))
        except Exception:
            try:
                prog.progress(frac)
            except Exception:
                pass

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
            progress_cb=_ui_progress,
        )

        st.session_state["prepared_bytes"] = out_bytes
        st.session_state["prepared_name"] = out_name
        st.session_state["availability"] = availability
        st.session_state["prep_report"] = prep_report
        st.session_state.pop("processed_bytes", None)
        _ui_progress(1.0, "Сборка: готово.")
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
    kaz_ok = bool(availability.get("kaz_ok"))

    opt_saldo = st.checkbox("Сальдо", value=False, disabled=(not saldo_ok))
    if not saldo_ok:
        st.caption("Сальдо недоступно: не найдены листы, заканчивающиеся на 1210/1710/3310/3510.")

    opt_contracts = st.checkbox("Контракты", value=False, disabled=(not contracts_ok))
    if not contracts_ok:
        st.caption("Контракты недоступны: не найдены пары листов *Wd/*Md.")

    opt_kaz_obsh = st.checkbox("Обработка общей ОСВ", value=False, disabled=(not kaz_ok))
    if not kaz_ok:
        st.caption("Обработка общей ОСВ недоступна: нужны листы «общ» и «Счета каз» в собранном файле.")

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

    has_any_mode = bool(selected_modes) or opt_inventory or opt_kaz_obsh
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

            if opt_kaz_obsh:
                status_box.info("Обработка: Общая ОСВ…")
                progress.progress(97)
                out_bytes = run_code_4_obsh_kaz(out_bytes)
                progress.progress(99)

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
