# -*- coding: utf-8 -*-
"""insights.py

Генерация листа `инсайты` — компактный кредитный анализ.

Вход: книга Excel (bytes) с листами W, M, Mt, Wt (обязательные) и листом `кред` (опционально).
Выход: те же bytes, но с добавленным/пересозданным листом `инсайты`.

Ключевые требования (сокращённо):
- Минимум текста, формальный стиль, без эмодзи.
- Форматирование: Calibri 11, числа #,##0, проценты 0.0%, цвета/заливки по ТЗ.
- Данные после «последнего месяца с реальными данными» игнорируем.

Важно: Листы могут иметь префикс (например, "МW", "МWt", "Мкред" и т.д.).
Мы выбираем один набор листов для анализа: без префикса, если он есть, иначе первый найденный префикс.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string


_MONTH_HDR_RE = re.compile(r"^\d{4}_(0[1-9]|1[0-2])$")


def _as_text(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _to_float(v) -> float:
    """Текст/пусто -> 0. Числа -> float. Скобки (1 234) -> -1234."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    s = s.replace("\u00A0", "").replace("\u202F", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _is_month_header(v) -> bool:
    return bool(_MONTH_HDR_RE.match(_as_text(v)))


def _ym(year: int, month: int) -> str:
    return f"{year:04d}_{month:02d}"


def _normalize_prefix(prefix: str) -> str:
    # В проекте префиксы обычно буквенные; нормализуем пробелы.
    return (prefix or "").strip()


def _load_wb_from_bytes(file_bytes: bytes) -> "load_workbook":
    bio = io.BytesIO(file_bytes)
    return load_workbook(bio)


def _pick_first_sheet_by_suffix(sheetnames: Sequence[str], prefix: str, suffix: str) -> str:
    """Выбираем лист по суффиксу (без учёта регистра), учитывая префикс."""
    pref = prefix or ""
    suf_l = suffix.lower()

    # 1) сначала пробуем точное имя
    for cand in (pref + suffix, pref + suffix.upper(), pref + suffix.lower(), pref + suffix.capitalize()):
        if cand in sheetnames:
            return cand

    # 2) иначе первый совпавший по окончанию
    for sh in sheetnames:
        low = sh.lower()
        if not low.endswith(suf_l):
            continue
        if pref and not sh.startswith(pref):
            continue
        return sh

    raise KeyError(f"Не найден лист с суффиксом {suffix}")


def _select_insights_prefix(sheetnames: Sequence[str]) -> str:
    """Если в книге несколько наборов W/M/Wt/Mt (по префиксам), берём без префикса, иначе первый."""
    required = {"w", "m", "wt", "mt"}
    pref_to: Dict[str, set] = {}

    for sh in sheetnames:
        low = sh.lower()
        for suf in ("wt", "mt", "w", "m"):
            if low.endswith(suf):
                pref = _normalize_prefix(sh[:-len(suf)])
                pref_to.setdefault(pref, set()).add(suf)

    prefixes = [p for p, s in pref_to.items() if required.issubset(s)]
    prefixes = sorted(set(prefixes), key=lambda x: (x != "", x))
    return prefixes[0] if prefixes else ""


def _delete_and_create_sheet_first(wb, title: str):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    try:
        wb._sheets.remove(ws)
        wb._sheets.insert(0, ws)
    except Exception:
        pass
    return ws


def _set_insights_column_widths(ws) -> None:
    # openpyxl.width — условные "символы", задаём приближённо.
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 60
    for col in range(column_index_from_string("C"), column_index_from_string("N") + 1):
        ws.column_dimensions[get_column_letter(col)].width = 12
    for col in ("C", "D", "E", "F"):
        ws.column_dimensions[col].width = 18


@dataclass
class Styles:
    # Цвета
    red: str = "CC0000"
    green: str = "006600"
    orange: str = "CC8800"
    gray: str = "666666"

    # Заливки
    fill_hdr: PatternFill = PatternFill("solid", fgColor="D9E1F2")
    fill_prob: PatternFill = PatternFill("solid", fgColor="FCE4EC")
    fill_risk: PatternFill = PatternFill("solid", fgColor="FFF2CC")
    fill_peak: PatternFill = PatternFill("solid", fgColor="E8F5E9")
    fill_low: PatternFill = PatternFill("solid", fgColor="FDE8E8")

    # Шрифты
    f_def: Font = Font(name="Calibri", size=11)
    f_hdr: Font = Font(name="Calibri", size=11, bold=True)
    f_title: Font = Font(name="Calibri", size=14, bold=True)
    f_italic_gray_small: Font = Font(name="Calibri", size=9, italic=True, color="666666")

    # Выравнивания
    a_left: Alignment = Alignment(horizontal="left", vertical="center")
    a_right: Alignment = Alignment(horizontal="right", vertical="center")
    a_center: Alignment = Alignment(horizontal="center", vertical="center")


def _write_row(ws, r: int, values: Sequence[object], *, font: Optional[Font] = None, fill: Optional[PatternFill] = None,
               align: Optional[Alignment] = None, num_fmt: Optional[str] = None, font_color: Optional[str] = None,
               start_col: int = 2) -> None:
    """Пишем строку, начиная с колонки B по умолчанию."""
    for i, v in enumerate(values):
        c = ws.cell(row=r, column=start_col + i, value=v)
        if font is not None:
            c.font = font
        if fill is not None:
            c.fill = fill
        if align is not None:
            c.alignment = align
        if num_fmt is not None:
            c.number_format = num_fmt
        if font_color is not None:
            base = c.font
            c.font = Font(name=base.name, size=base.size, bold=base.bold, italic=base.italic, color=font_color)


def _iter_month_rows(ws, start_row: int, year_col: int, month_col: int) -> Iterable[Tuple[int, int, int]]:
    """Итерируем помесячные строки: пока год остаётся числом (строки 12+)."""
    r = start_row
    max_r = int(ws.max_row or 1)
    while r <= max_r:
        y = ws.cell(row=r, column=year_col).value
        if y is None or _as_text(y) == "":
            break
        if not isinstance(y, (int, float)):
            break
        m = ws.cell(row=r, column=month_col).value
        if not isinstance(m, (int, float)):
            break
        yield r, int(y), int(m)
        r += 1


@dataclass
class WRow:
    ym: str
    cfo: float
    clients: float
    pers: float
    taxes: float


def parse_w_sheet(ws_w) -> List[WRow]:
    """Парсим W: строки 12+, фиксированные колонки (B,C,F,G,R,S)."""
    rows: List[WRow] = []
    for rr, y, m in _iter_month_rows(ws_w, start_row=12, year_col=2, month_col=3):
        cfo = _to_float(ws_w.cell(row=rr, column=6).value)
        clients = _to_float(ws_w.cell(row=rr, column=7).value)
        pers = _to_float(ws_w.cell(row=rr, column=18).value)
        taxes = _to_float(ws_w.cell(row=rr, column=19).value)
        rows.append(WRow(ym=_ym(y, m), cfo=cfo, clients=clients, pers=pers, taxes=taxes))

    # Последний месяц = последний, где CFO != 0 или Клиенты != 0
    last = -1
    for i in range(len(rows) - 1, -1, -1):
        if abs(rows[i].cfo) > 0 or abs(rows[i].clients) > 0:
            last = i
            break
    return rows[: last + 1] if last >= 0 else rows


@dataclass
class MRow:
    ym: str
    rev: float


def parse_m_sheet(ws_m) -> List[MRow]:
    """Парсим M: строки 12+, используем только реальные месяцы (A=1), выручка G."""
    tmp: List[Tuple[str, int, float]] = []
    for rr, y, m in _iter_month_rows(ws_m, start_row=12, year_col=2, month_col=3):
        flag = ws_m.cell(row=rr, column=1).value
        flag = int(flag) if isinstance(flag, (int, float)) else 0
        rev = _to_float(ws_m.cell(row=rr, column=7).value)
        tmp.append((_ym(y, m), flag, rev))

    real = [MRow(ym=ym, rev=rev) for (ym, flag, rev) in tmp if flag == 1]

    # Последний реальный месяц = последний, где rev != 0
    last = -1
    for i in range(len(real) - 1, -1, -1):
        if abs(real[i].rev) > 0:
            last = i
            break
    return real[: last + 1] if last >= 0 else real


@dataclass
class TableEntry:
    name: str
    total: float


@dataclass
class ParsedTable:
    name: str
    header_row: int
    total_row: Optional[int]
    month_cols: List[int]  # колонки месяцев (C..)
    grand_total_col: int
    entries: List[TableEntry]


def _find_grand_total_col(ws, header_row: int, start_col: int = 3, fallback: str = "AC") -> int:
    max_c = int(ws.max_column or 1)
    for c in range(start_col, min(max_c, 120) + 1):
        hv = _as_text(ws.cell(row=header_row, column=c).value).lower()
        if hv in ("итог", "итого"):
            return c
    return column_index_from_string(fallback)


def _read_month_cols(ws, header_row: int, first_month_col: int = 3, limit: int = 60) -> List[int]:
    cols: List[int] = []
    max_c = min(int(ws.max_column or 1), limit)
    c = first_month_col
    while c <= max_c and _is_month_header(ws.cell(row=header_row, column=c).value):
        cols.append(c)
        c += 1
    return cols


def parse_wt_like_tables(
    ws,
    start_row: int,
    *,
    allowed_names: Optional[set] = None,
    excluded_names: Optional[set] = None,
    skip_row_names: Optional[set] = None,
) -> List[ParsedTable]:
    """Универсальный парсер таблиц Wt/кред.

    Заголовок таблицы: B непусто и C = '2025_01'.
    Конец таблицы: строка, где B = 'Доля'.
    Контрагенты: строки между заголовком и 'Доля', кроме Топ/Всего/Доля.
    """
    allowed = allowed_names
    excluded = excluded_names or set()
    skip = skip_row_names or set()

    out: List[ParsedTable] = []

    r = start_row
    max_r = int(ws.max_row or 1)
    while r <= max_r:
        tname = _as_text(ws.cell(row=r, column=2).value)
        first_month = ws.cell(row=r, column=3).value
        if tname and _is_month_header(first_month):
            if (allowed is None or tname in allowed) and tname not in excluded:
                gt_col = _find_grand_total_col(ws, r)
                month_cols = _read_month_cols(ws, r)

                entries: List[TableEntry] = []
                total_row: Optional[int] = None

                rr = r + 1
                while rr <= max_r:
                    nm = _as_text(ws.cell(row=rr, column=2).value)
                    if nm == "Доля":
                        break
                    if nm == "Всего":
                        total_row = rr
                    if nm and nm not in skip and nm not in {"Топ", "Всего", "Доля"}:
                        tot = _to_float(ws.cell(row=rr, column=gt_col).value)
                        entries.append(TableEntry(name=nm, total=tot))
                    rr += 1

                out.append(
                    ParsedTable(
                        name=tname,
                        header_row=r,
                        total_row=total_row,
                        month_cols=month_cols,
                        grand_total_col=gt_col,
                        entries=entries,
                    )
                )

            r += 1
            continue

        r += 1

    return out


@dataclass
class MtParsed:
    total_all: float
    entries: List[TableEntry]
    header_row: int
    total_row: Optional[int]
    grand_total_col: int


def parse_mt(ws_mt) -> MtParsed:
    hdr = 5
    name_col = 2
    gt_col = _find_grand_total_col(ws_mt, hdr)

    entries: List[TableEntry] = []
    r = hdr + 1
    max_r = int(ws_mt.max_row or 1)
    stop = {"Топ", "Всего", "Доля"}
    while r <= max_r:
        nm = _as_text(ws_mt.cell(row=r, column=name_col).value)
        if not nm:
            r += 1
            continue
        if nm in stop:
            break
        entries.append(TableEntry(name=nm, total=_to_float(ws_mt.cell(row=r, column=gt_col).value)))
        r += 1

    total_all = 0.0
    total_row = None
    for rr in range(hdr + 1, min(max_r, hdr + 400) + 1):
        if _as_text(ws_mt.cell(row=rr, column=name_col).value) == "Всего":
            total_all = _to_float(ws_mt.cell(row=rr, column=gt_col).value)
            total_row = rr
            break

    return MtParsed(total_all=total_all, entries=entries, header_row=hdr, total_row=total_row, grand_total_col=gt_col)


def last_month_col_by_total_row(ws, header_row: int, total_row: int, month_cols: List[int]) -> Optional[int]:
    """Для таблиц: справа налево ищем последний месяц, где в строке 'Всего' != 0."""
    if not total_row or not month_cols:
        return None
    for c in reversed(month_cols):
        if abs(_to_float(ws.cell(row=total_row, column=c).value)) > 0:
            return c
    return month_cols[-1]


# =========================
# Генератор листа `инсайты`
# =========================


def generate_insights(file_bytes: bytes) -> bytes:
    """Точка входа: принимает bytes книги, возвращает bytes с листом `инсайты`."""
    wb = _load_wb_from_bytes(file_bytes)

    pref = _select_insights_prefix(wb.sheetnames)

    # Обязательные листы
    sh_w = _pick_first_sheet_by_suffix(wb.sheetnames, pref, "W")
    sh_m = _pick_first_sheet_by_suffix(wb.sheetnames, pref, "M")
    sh_wt = _pick_first_sheet_by_suffix(wb.sheetnames, pref, "Wt")
    sh_mt = _pick_first_sheet_by_suffix(wb.sheetnames, pref, "Mt")

    ws_w = wb[sh_w]
    ws_m = wb[sh_m]
    ws_wt = wb[sh_wt]
    ws_mt = wb[sh_mt]

    # Опциональный лист
    ws_cred = None
    try:
        sh_cred = _pick_first_sheet_by_suffix(wb.sheetnames, pref, "кред")
        ws_cred = wb[sh_cred]
    except Exception:
        ws_cred = None

    styles = Styles()

    ws = _delete_and_create_sheet_first(wb, "инсайты")
    _set_insights_column_widths(ws)

    NUM = "#,##0"
    PCT = "0.0%"

    row = 1
    _write_row(ws, row, ["ИНСАЙТЫ"], font=styles.f_title, align=styles.a_left)
    row += 2

    # ---- читаем входные данные ----
    w_rows = parse_w_sheet(ws_w)
    m_rows = parse_m_sheet(ws_m)

    wt_exclude = {"Прочая ДЗ/КЗ Inflow", "Прочая ДЗ/КЗ Outflow"}
    wt_tables = parse_wt_like_tables(ws_wt, start_row=3, excluded_names=wt_exclude)

    cred_tables: List[ParsedTable] = []
    if ws_cred is not None:
        cred_tables = parse_wt_like_tables(
            ws_cred,
            start_row=5,
            allowed_names={"Краткосрочные кредиты (Netto)", "Долгосрочные кредиты (Netto)"},
            skip_row_names={"", "Тело", "Проценты", "Платежи"},
        )

    mt = parse_mt(ws_mt)

    # =========================
    # БЛОК 1: Пересечения контрагентов
    # =========================

    _write_row(ws, row, ["1. ПЕРЕСЕЧЕНИЯ КОНТРАГЕНТОВ (Wt / кред)"], font=styles.f_hdr, align=styles.a_left)
    row += 1
    _write_row(ws, row, ["Контрагенты, присутствующие в 2+ таблицах"], font=styles.f_italic_gray_small, align=styles.a_left)
    row += 1
    _write_row(ws, row, ["Контрагент", "Лист", "Таблица", "Направление", "Итого"], font=styles.f_hdr, fill=styles.fill_hdr, align=styles.a_left)
    row += 1

    entries: List[Tuple[str, str, str, float]] = []
    for t in wt_tables:
        for e in t.entries:
            entries.append((e.name, "Wt", t.name, e.total))

    for t in cred_tables:
        for e in t.entries:
            entries.append((e.name, "кред", t.name, e.total))

    # карта: контрагент -> список уникальных (лист+таблица)
    mp: Dict[str, List[Tuple[str, str, float]]] = {}
    for name, sh, tbl, tot in entries:
        mp.setdefault(name, [])
        if (sh, tbl) not in {(a, b) for (a, b, _v) in mp[name]}:
            mp[name].append((sh, tbl, tot))

    intersect = {k: v for k, v in mp.items() if len(v) >= 2}

    company_tables = {"Клиенты", "Пост-ки", "Прочая ДЗ/КЗ", "Прочие займы", "Краткосрочные кредиты (Netto)", "Долгосрочные кредиты (Netto)"}

    def is_company(items: List[Tuple[str, str, float]]) -> bool:
        for sh, tbl, _tot in items:
            if tbl in company_tables:
                return True
        return False

    companies = sorted([k for k, v in intersect.items() if is_company(v)])
    persons = sorted([k for k, v in intersect.items() if not is_company(v)])

    used_kred_and_wt: set = set()

    def write_intersect_group(names: List[str]) -> None:
        nonlocal row
        for nm in names:
            items = intersect[nm]
            for i, (sh, tbl, tot) in enumerate(items):
                if sh == "кред":
                    used_kred_and_wt.add(nm)
                direction = "Приход" if tot >= 0 else "Расход"
                _write_row(ws, row, [nm if i == 0 else "", sh, tbl, direction, int(round(tot))], font=styles.f_def, align=styles.a_left)
                c_val = ws.cell(row=row, column=6)
                c_val.number_format = NUM
                c_val.font = Font(name="Calibri", size=11, color=(styles.green if tot >= 0 else styles.red))
                row += 1

    if intersect:
        write_intersect_group(companies)
        if persons:
            _write_row(ws, row, ["Физические лица (подотчет + персонал):"], font=Font(name="Calibri", size=9, italic=True, color=styles.gray), align=styles.a_left)
            row += 1
            write_intersect_group(persons)
    else:
        _write_row(ws, row, ["Пересечений не выявлено"], font=Font(name="Calibri", size=11, color=styles.green), align=styles.a_left)
        row += 1

    row += 2

    # =========================
    # БЛОК 2: Регулярность выручки (M)
    # =========================

    _write_row(ws, row, ["2. РЕГУЛЯРНОСТЬ ВЫРУЧКИ (M)"], font=styles.f_hdr, align=styles.a_left)
    row += 1

    # Среднее по месяцам с данными (не нулевым)
    rev_vals = [r.rev for r in m_rows if abs(r.rev) > 0]
    avg_rev = (sum(rev_vals) / len(rev_vals)) if rev_vals else 0.0
    thr = 0.1 * avg_rev

    # Перерывы: последовательности месяцев, где выручка < порога
    breaks: List[Tuple[str, str, int]] = []
    cur_start = None
    cur_n = 0
    prev_ym = None

    for r0 in m_rows:
        low = (r0.rev < thr) if avg_rev > 0 else False
        if low:
            if cur_start is None:
                cur_start = r0.ym
                cur_n = 1
            else:
                cur_n += 1
            prev_ym = r0.ym
        else:
            if cur_start is not None:
                breaks.append((cur_start, prev_ym or cur_start, cur_n))
                cur_start = None
                cur_n = 0
                prev_ym = None

    if cur_start is not None:
        breaks.append((cur_start, prev_ym or cur_start, cur_n))

    if breaks:
        for s, e, n in breaks:
            _write_row(ws, row, [f"Перерыв: {s} — {e} ({n} мес. выручка < порога)"], font=Font(name="Calibri", size=11, color=styles.red), align=styles.a_left)
            row += 1
    else:
        _write_row(ws, row, ["Перерывов не выявлено"], font=Font(name="Calibri", size=11, color=styles.green), align=styles.a_left)
        row += 1

    # Аномалии: падение относительно скользящей средней
    problems: List[Tuple[str, float, float, str]] = []
    for i in range(1, len(m_rows)):
        prev = m_rows[max(0, i - 3): i]
        avg3 = (sum(p.rev for p in prev) / len(prev)) if prev else 0.0
        cur = m_rows[i].rev
        if avg3 > 0:
            if cur == 0:
                problems.append((m_rows[i].ym, cur, avg3, "Нулевая выручка"))
            elif cur < 0.3 * avg3:
                пад = int(round((1 - (cur / avg3)) * 100))
                problems.append((m_rows[i].ym, cur, avg3, f"Падение {пад}%"))

    if problems:
        _write_row(ws, row, ["Период", "Выручка", "Ср. 3 мес.", "Статус"], font=styles.f_hdr, fill=styles.fill_hdr, align=styles.a_left)
        row += 1
        for ym, rev, a3, st in problems:
            _write_row(ws, row, [ym, int(round(rev)), int(round(a3)), st], font=styles.f_def, align=styles.a_left)
            ws.cell(row=row, column=3).number_format = NUM
            ws.cell(row=row, column=4).number_format = NUM
            ws.cell(row=row, column=5).font = Font(name="Calibri", size=11, bold=True, color=styles.red)
            row += 1
    else:
        _write_row(ws, row, ["Проблемных месяцев не выявлено"], font=Font(name="Calibri", size=11, color=styles.green), align=styles.a_left)
        row += 1

    row += 2

    # =========================
    # БЛОК 3: Регулярность персонала и налогов (W)
    # =========================

    _write_row(ws, row, ["3. РЕГУЛЯРНОСТЬ ПЕРСОНАЛА И НАЛОГОВ (W)"], font=styles.f_hdr, align=styles.a_left)
    row += 1
    _write_row(ws, row, ["Аномалии: падение >50% от скользящего среднего за 3 мес."], font=styles.f_italic_gray_small, align=styles.a_left)
    row += 1

    def rolling_anoms(values: List[Tuple[str, float]]) -> List[Tuple[str, float, float, str]]:
        out2: List[Tuple[str, float, float, str]] = []
        for i in range(1, len(values)):
            prev = values[max(0, i - 3): i]
            avg3 = sum(v for _ym0, v in prev) / len(prev) if prev else 0.0
            cur = values[i][1]
            if avg3 > 1000:
                if cur == 0:
                    out2.append((values[i][0], cur, avg3, "Нулевой платеж"))
                elif cur < 0.5 * avg3:
                    пад = int(round((1 - (cur / avg3)) * 100))
                    out2.append((values[i][0], cur, avg3, f"Падение {пад}%"))
        return out2

    pers_series = [(r.ym, r.pers) for r in w_rows]
    tax_series = [(r.ym, r.taxes) for r in w_rows]

    pers_an = rolling_anoms(pers_series)
    tax_an = rolling_anoms(tax_series)

    def write_subblock(title: str, items: List[Tuple[str, float, float, str]]) -> None:
        nonlocal row
        _write_row(ws, row, [title], font=styles.f_hdr, align=styles.a_left)
        row += 1
        if not items:
            _write_row(ws, row, ["Аномалий не выявлено"], font=Font(name="Calibri", size=11, color=styles.green), align=styles.a_left)
            row += 1
            return
        _write_row(ws, row, ["Период", "Сумма", "Ср. 3 мес.", "Статус"], font=styles.f_hdr, fill=styles.fill_prob, align=styles.a_left)
        row += 1
        for ym, val, a3, st in items:
            _write_row(ws, row, [ym, int(round(val)), int(round(a3)), st], font=styles.f_def, align=styles.a_left)
            ws.cell(row=row, column=3).number_format = NUM
            ws.cell(row=row, column=4).number_format = NUM
            ws.cell(row=row, column=5).font = Font(name="Calibri", size=11, bold=True, color=styles.red)
            row += 1

    write_subblock("Персонал", pers_an)
    write_subblock("Налоги", tax_an)

    row += 2

    # =========================
    # БЛОК 4: Концентрация клиентов (Mt + Wt Клиенты)
    # =========================

    _write_row(ws, row, ["4. КОНЦЕНТРАЦИЯ КЛИЕНТОВ"], font=styles.f_hdr, align=styles.a_left)
    row += 1

    def write_concentration(title: str, cps: List[TableEntry], total_all: float) -> Tuple[Optional[str], float, int]:
        nonlocal row
        _write_row(ws, row, [title], font=styles.f_hdr, align=styles.a_left)
        row += 1
        if not cps or total_all == 0:
            _write_row(ws, row, ["Данных нет"], font=Font(name="Calibri", size=11, color=styles.gray), align=styles.a_left)
            row += 1
            return None, 0.0, 0

        _write_row(ws, row, ["Клиент", "Сумма", "Доля"], font=styles.f_hdr, fill=styles.fill_hdr, align=styles.a_left)
        row += 1

        cps_sorted = sorted(cps, key=lambda x: abs(x.total), reverse=True)
        top1_name = cps_sorted[0].name
        top1_share = (cps_sorted[0].total / total_all) if total_all else 0.0

        for e in cps_sorted:
            share = (e.total / total_all) if total_all else 0.0
            _write_row(ws, row, [e.name, int(round(e.total)), share], font=styles.f_def, align=styles.a_left)
            ws.cell(row=row, column=4).number_format = NUM
            ws.cell(row=row, column=5).number_format = PCT
            if share > 0.70:
                ws.cell(row=row, column=5).font = Font(name="Calibri", size=11, bold=True, color=styles.red)
            row += 1

        if top1_share > 0.70:
            _write_row(
                ws,
                row,
                [f"ВЫСОКАЯ КОНЦЕНТРАЦИЯ: {top1_share*100:.1f}% — один клиент ({top1_name})"],
                font=Font(name="Calibri", size=11, bold=True, color=styles.red),
                align=styles.a_left,
            )
            row += 1
        elif top1_share > 0.50:
            _write_row(
                ws,
                row,
                [f"ЗНАЧИТЕЛЬНАЯ КОНЦЕНТРАЦИЯ: {top1_share*100:.1f}% — один клиент ({top1_name})"],
                font=Font(name="Calibri", size=11, bold=True, color=styles.orange),
                align=styles.a_left,
            )
            row += 1

        _write_row(ws, row, [f"Всего клиентов: {len(cps_sorted)}"], font=styles.f_def, align=styles.a_left)
        row += 1

        return top1_name, top1_share, len(cps_sorted)

    mt_top1_name, mt_top1_share, mt_n = write_concentration("Выручка (Mt)", mt.entries, mt.total_all)
    row += 1

    wt_clients_tbl = next((t for t in wt_tables if t.name == "Клиенты"), None)
    wt_total_all = 0.0
    wt_entries: List[TableEntry] = []
    if wt_clients_tbl is not None:
        wt_entries = wt_clients_tbl.entries
        if wt_clients_tbl.total_row:
            wt_total_all = _to_float(ws_wt.cell(row=wt_clients_tbl.total_row, column=wt_clients_tbl.grand_total_col).value)

    wt_top1_name, wt_top1_share, wt_n = write_concentration("Поступления от клиентов (Wt)", wt_entries, wt_total_all)

    row += 2

    # =========================
    # БЛОК 5: Сезонность поступлений от клиентов (Wt Клиенты / строка Всего)
    # =========================

    _write_row(ws, row, ["5. СЕЗОННОСТЬ ПОСТУПЛЕНИЙ ОТ КЛИЕНТОВ (Wt)"], font=styles.f_hdr, align=styles.a_left)
    row += 1

    months_ru = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]

    season_lows: List[int] = []
    season_spread = 0.0

    if wt_clients_tbl is None or not wt_clients_tbl.total_row or not wt_clients_tbl.month_cols:
        _write_row(ws, row, ["Данных нет"], font=Font(name="Calibri", size=11, color=styles.gray), align=styles.a_left)
        row += 1
    else:
        # Берём первые 24 месяца из month_cols: 2025 (12) + 2026 (12)
        mcols = wt_clients_tbl.month_cols[:24]
        vals_25 = [0.0] * 12
        vals_26 = [0.0] * 12

        for i, c in enumerate(mcols):
            v = _to_float(ws_wt.cell(row=wt_clients_tbl.total_row, column=c).value)
            if i < 12:
                vals_25[i] = v
            elif i < 24:
                vals_26[i - 12] = v

        # Последний месяц 2026 с данными
        last_26 = 12
        for i in range(11, -1, -1):
            if abs(vals_26[i]) > 0:
                last_26 = i + 1
                break

        avg25 = sum(vals_25) / 12.0 if vals_25 else 0.0
        peaks = [i + 1 for i, v in enumerate(vals_25) if avg25 > 0 and v > 1.5 * avg25]
        season_lows = [i + 1 for i, v in enumerate(vals_25) if avg25 > 0 and v < 0.3 * avg25]

        nonzero = [v for v in vals_25 if abs(v) > 0]
        mn = min(nonzero) if nonzero else 0.0
        mx = max(nonzero) if nonzero else 0.0
        season_spread = (mx / mn) if mn not in (0, 0.0) else 0.0

        _write_row(
            ws,
            row,
            [
                f"Среднее за 2025: {int(round(avg25))} / мес. "
                f"Пиковые: {', '.join(map(str, peaks)) or '-'}; "
                f"Провалы: {', '.join(map(str, season_lows)) or '-'}",
            ],
            font=styles.f_italic_gray_small,
            align=styles.a_left,
        )
        row += 1

        _write_row(ws, row, ["", *months_ru], font=styles.f_hdr, fill=styles.fill_hdr, align=styles.a_center)
        row += 1

        _write_row(ws, row, ["2025", *[int(round(v)) for v in vals_25]], font=styles.f_def, align=styles.a_center)
        for i in range(12):
            c = ws.cell(row=row, column=3 + i)
            c.number_format = NUM
            if (i + 1) in peaks:
                c.fill = styles.fill_peak
                c.font = Font(name="Calibri", size=11, color=styles.green)
            if (i + 1) in season_lows:
                c.fill = styles.fill_low
                c.font = Font(name="Calibri", size=11, color=styles.red)
        row += 1

        _write_row(ws, row, ["2026", *[int(round(v)) for v in vals_26]], font=styles.f_def, align=styles.a_center)
        for i in range(12):
            c = ws.cell(row=row, column=3 + i)
            c.number_format = NUM
            if i + 1 > last_26:
                c.font = Font(name="Calibri", size=11, color="CCCCCC")
        row += 1

        _write_row(
            ws,
            row,
            ["Ср. мес.", *[int(round(avg25)) for _ in range(12)]],
            font=Font(name="Calibri", size=11, bold=True, italic=True, color=styles.gray),
            align=styles.a_center,
        )
        for i in range(12):
            ws.cell(row=row, column=3 + i).number_format = NUM
        row += 1

        if season_spread and season_spread > 0:
            _write_row(
                ws,
                row,
                [f"Поступления крайне неравномерны: мин. {int(round(mn))}, макс. {int(round(mx))} — разброс в {season_spread:.1f}x"],
                font=Font(name="Calibri", size=11, color=styles.orange),
                align=styles.a_left,
            )
            row += 1
        if season_lows:
            _write_row(
                ws,
                row,
                ["Риск кассового разрыва в периоды низких поступлений при сохранении фиксированных расходов"],
                font=Font(name="Calibri", size=11, bold=True, color=styles.red),
                align=styles.a_left,
            )
            row += 1

    row += 2

    # =========================
    # ИТОГО: ключевые риски
    # =========================

    _write_row(ws, row, ["ИТОГО: КЛЮЧЕВЫЕ РИСКИ"], font=Font(name="Calibri", size=12, bold=True), fill=styles.fill_risk, align=styles.a_left)
    row += 1

    risks: List[str] = []

    # 1) Концентрация: топ-1 по Mt > 70%
    if mt_top1_share > 0.70 and mt_top1_name:
        x = mt_top1_share * 100
        y = wt_top1_share * 100
        risks.append(f"Критическая зависимость от одного клиента ({mt_top1_name}): {x:.1f}% выручки, {y:.1f}% поступлений")

    # 2) Перерывы выручки >= 2 мес.
    for s, e, n in breaks:
        if n >= 2:
            risks.append(f"Нерегулярная выручка: провал {s}—{e} ({n} мес. < порога)")
            break

    # 3) Сезонность: есть провальные месяцы
    if season_lows:
        risks.append(f"Сезонность поступлений: провальные месяцы, риск кассового разрыва")

    # 4) Аномалии в персонале/налогах
    if pers_an or tax_an:
        risks.append("Нестабильные платежи по налогам и персоналу: пропуски и резкие снижения")

    # 5) Пересечения с кредитами
    if used_kred_and_wt:
        names = ", ".join(sorted(list(used_kred_and_wt))[:5])
        risks.append(f"Пересечения с кредитами (кред + Wt): {names}")

    if not risks:
        _write_row(ws, row, ["Риски не выявлены"], font=Font(name="Calibri", size=11, color=styles.green), align=styles.a_left)
        row += 1
    else:
        for i, t in enumerate(risks, start=1):
            _write_row(ws, row, [f"{i}. {t}"], font=Font(name="Calibri", size=11, color=styles.red), align=styles.a_left)
            row += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


__all__ = ["generate_insights"]
