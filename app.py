import re
import io
import time
import os
import tempfile
import zipfile
from copy import copy as _copy
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Tuple

import requests as _requests_lib

import pandas as pd
import streamlit as st

# ── egov / ГБД ЮЛ ─────────────────────────────────────────────────────────
_EGOV_API_KEY = "ef214eb1c5684c86a7991c91caa93f5b"
_EGOV_BASE    = "https://data.egov.kz"
_egov_session = _requests_lib.Session()
_egov_session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120",
    "Referer": _EGOV_BASE,
})
_egov_session.cookies.update({
    "cookiesession1": "678B769B1B988DB877C03DCFCF48FEF4",
    "egovLang": "ru",
})
_egov_cache: Dict[str, dict] = {}

# Организационно-правовые формы для классификации
_RU_FORMS  = {'ООО','ОАО','ЗАО','ПАО','НАО','ФГУП','МУП','ГУП','АНО','НКО','АООТ','АОЗТ','ООО.','ОАО.'}
# LLC и LLP убраны — широко используются и KZ-компаниями (ТОО в англ. написании)
_FOR_FORMS = {'LTD','GMBH','INC','CORP','BV','AG','SAS','NV','AB','OY','SARL','LTD.','CO.'}
_KZ_FORMS  = {'ТОО','АО','РГП','ГКП','ТДО','ПК','АКБ','БРК','НАО','ГУ','КГП','ТК','ГП','ЧК','КФХ','СПК','МКО','МФО','НПФ','РГУ','КГУ'}
# Иностранные формы из стран СНГ (не KZ)
_FOR_FORMS.update({'ТОВ','ОАТ','ПАТ','ПРАТ','ФОП','ОДО','ЧТУП','ЧУП','ТЧУП','РУП'})
# ХХК — монгольская форма (Хязгаарлагдмал Хариуцлагатай Компани)
_FOR_FORMS.add('ХХК')

# Полные русские названия форм для обогащения поискового запроса (улучшает ранжирование в ES)
_KZ_FORM_FULL = {
    'ТОО': 'Товарищество с ограниченной ответственностью',
    'АО':  'Акционерное общество',
    'НАО': 'Некоммерческое акционерное общество',
    'ПК':  'Производственный кооператив',
    'РГП': 'Республиканское государственное предприятие',
    'ГП':  'Государственное предприятие',
    'ГУ':  'Государственное учреждение',
    'КГП': 'Коммунальное государственное предприятие',
    'ГКП': 'Государственное коммунальное предприятие',
    'ТДО': 'Товарищество с дополнительной ответственностью',
    'КФХ': 'Крестьянское фермерское хозяйство',
    'СПК': 'Сельскохозяйственный производственный кооператив',
    'МКО': 'Микрокредитная организация',
    'МФО': 'Микрофинансовая организация',
    'НПФ': 'Накопительный пенсионный фонд',
    'АКБ': 'Акционерный коммерческий банк',
    'РГУ': 'Республиканское государственное учреждение',
    'КГУ': 'Коммунальное государственное учреждение',
}

# Полные русские/казахские названия KZ организационно-правовых форм
_KZ_FULL_RE = re.compile(
    r'\bтоварищество\s+с\s+ограниченн'          # ТОО полная форма (с опечатками тоже)
    r'|\bтоварищество\s+с\s+дополнительной\s+ответственностью\b'
    r'|\bакционерное\s+общество\b'
    r'|\bнекоммерческое\s+акционерное\s+общество\b'
    r'|\bпроизводственный\s+кооператив\b'
    r'|\bполное\s+товарищество\b'
    r'|\bкоммандитное\s+товарищество\b'
    r'|\bобщественное\s+объединение\b'
    r'|\bобщественный\s+(?:благотворительный\s+)?фонд\b'
    r'|\bучреждение\b'
    r'|\bфилиал\b'
    r'|\bпредставительство\b'
    r'|\bгосударственное\s+(?:казенное\s+)?предприятие\b'
    r'|\bреспубликанское\s+государственное\s+предприятие\b'
    r'|\bчастная\s+компания\b',
    re.IGNORECASE,
)

# Весь кириллический блок Unicode U+0400-U+04FF — покрывает все буквы
# казахского алфавита включая Ħ/ħ (U+04BA/04BB), Ö/ö (U+04E8/04E9) и др.
_CYR = r'Ѐ-ӿ'

# ФИО: Фамилия Имя Отчество (патроним — обязателен для детекции)
_FIO_RE = re.compile(
    rf'^[{_CYR}][{_CYR}\-]+'
    rf'\s+[{_CYR}][{_CYR}\-]+'
    rf'\s+[{_CYR}][{_CYR}]*'
    r'(?:ович|евич|овна|евна|ьевич|ьевна|ынич|иевич|иевна|вич|вна'
    r'|улы|ұлы|кызы|қызы|уғли|ўғли)\s*$',
    re.IGNORECASE,
)
# ФИО с инициалами: "Юсупов Т.Т." / "Иванов А."
_FIO_INITIALS_RE = re.compile(
    rf'^[{_CYR}][{_CYR}\-]+'
    rf'\s+[{_CYR}A-Z]\.(?:[{_CYR}A-Z]\.)?\s*$',
)
# Патроним как первое слово: "АЛТЫНБЕКҰЛЫ МҰҚАН" (обратный порядок)
_FIO_REV_RE = re.compile(
    rf'^[{_CYR}][{_CYR}\-]*'
    r'(?:ұлы|улы|қызы|кызы|уғли|ўғли)'
    rf'\s+[{_CYR}][{_CYR}\-]*\s*$',
    re.IGNORECASE,
)

# Внутренние учётные позиции — не контрагенты
_INTERNAL_RE = re.compile(
    r'^(?:основн\w*\s+касс\w+'
    r'|основн\w*\s+счёт\b|основн\w*\s+счет\b'
    r'|корпоратив\w*\s+карточ\w+'
    r'|курсовые\s+разниц\w+'
    r'|комиссия\s+банк\w*'
    r'|расчет[нь]?[ыо][йе]?\s+счет|расчёт[нь]?[ыо][йе]?\s+счёт'
    r'|депозит\s+договор'
    r'|прочи[ех]?\s+(?:расход|выплат|затрат|поступлени|операци|доход)\w*'
    r'|доход\w*\s+по\s+вознаграждени\w*'
    r'|(?:прочи[ех]?\s+)?(?:финансов\w+\s+)?(?:расход|выплат)\w*\s+(?:по|за)\s+\w+)',
    re.IGNORECASE,
)

# Паттерны физлиц и заглушек — поиск не нужен
_PHYS_PATTERNS = [
    (r'физ\.?\s*лиц',               'Физическое лицо'),
    (r'физическое\s+лицо',          'Физическое лицо'),
    (r'частное\s+лиц',              'Физическое лицо'),
    (r'частн\.?\s*лиц',             'Физическое лицо'),
    (r'^частное$',                  'Физическое лицо'),
]
_SKIP_RE = re.compile(
    r'^[\s\-–—\.,:;/\\0-9]+$'                   # пусто / только спецсимволы / только цифры
    r'|^(?:н/?д|нет\s+данных?|не\s+определ|прочи[ех]?|разные?|итого|прочее)$',
    re.IGNORECASE,
)

def _egov_v4(dataset: str, source: dict) -> list:
    import json as _json
    try:
        r = _egov_session.get(
            f"{_EGOV_BASE}/api/v4/{dataset}/v1",
            params={"apiKey": _EGOV_API_KEY, "source": _json.dumps(source, ensure_ascii=False)},
            timeout=10,
        )
        return r.json() if r.ok else []
    except Exception:
        return []

def _egov_classify(name: str) -> Tuple[str, str]:
    """
    Классифицирует название контрагента.
    Возвращает (тип, поисковый запрос):
      'kz_ul'   — казахстанское юрлицо → поиск в gbd_ul
      'kz_ip'   — ИП (физлицо) → нет в ГБД ЮЛ
      'russian' — российская форма (ООО/ОАО/ЗАО…) → поиск не нужен
      'foreign' — иностранная форма (LLC/Ltd…) → поиск не нужен
    """
    up = name.upper().strip()

    # 1. Пустые и заглушки — не ищем
    if not name.strip() or _SKIP_RE.match(name.strip()):
        return ('skip', '')

    # 1а. IBAN / номер счёта (KZxx... или любой IBAN-формат: 2 буквы + 2 цифры)
    if re.match(r'^[A-Z]{2}\d{2}[A-Z0-9]', up):
        return ('bank_account', '')

    # 2. Внутренние учётные записи (Основная касса, Основной счёт и т.п.)
    if _INTERNAL_RE.match(name.strip()):
        return ('internal', '')

    # 3. Физические лица
    for pattern, _ in _PHYS_PATTERNS:
        if re.search(pattern, name, flags=re.IGNORECASE):
            return ('phys', '')
    # ФИО: Фамилия Имя Отчество (патроним последним)
    if _FIO_RE.match(name.strip()):
        return ('phys', '')
    # ФИО: Патроним(ұлы/қызы) Имя (патроним первым — АЛТЫНБЕКҰЛЫ МҰҚАН)
    if _FIO_REV_RE.match(name.strip()):
        return ('phys', '')
    # ФИО с инициалами: "Юсупов Т.Т." / "Иванов А."
    if _FIO_INITIALS_RE.match(name.strip()):
        return ('phys', '')

    # 4. ИП и аналоги — проверяем до всего остального
    _ip_markers = (
        r'ИНДИВИДУАЛЬНЫЙ\s+ПРЕДПРИНИМАТЕЛЬ',
        r'ЧАСТНЫЙ\s+ПРЕДПРИНИМАТЕЛЬ',
        r'ЧАСТНЫЙ\s+ПРЕДП',
    )
    _ip_abbr = r'(?:^|[\s«"\(\.])(?:ИП|ЧП|ФОП)(?:$|[\s»"\)\.])'
    if (any(re.search(p, up) for p in _ip_markers)
            or re.search(_ip_abbr, up)
            or up.startswith('ИП ') or up.startswith('ИП«') or up.startswith('ИП"')
            or up.startswith('ЧП ') or up.startswith('ФОП ')):
        clean = name
        for p in _ip_markers:
            clean = re.sub(p, '', clean, flags=re.IGNORECASE)
        clean = re.sub(r'(?:^|[\s«"\(\.])(?:ИП|ЧП|ФОП)(?:$|[\s»"\)\.])', ' ', clean, flags=re.IGNORECASE)
        clean = re.sub(r'\s*\(.*?\)\s*', ' ', clean).strip('«»"\' ()-').strip()
        return ('kz_ip', clean)

    # Убираем скобочные пояснения для поиска: "АО АРВИАЙ (РАШЕН ВЕНЧУР)" → "АО АРВИАЙ"
    core = re.sub(r'\s*\(.*?\)\s*', ' ', name).strip()

    # KZ юрлицо — проверяем ДО иностранных форм (ЧК/ТОО выигрывают у Ltd.)
    _Q_RE = re.compile(r'[«"""„]([^«"""»""]{2,})[»""""]')
    for f in _KZ_FORMS:
        if re.search(r'(?:^|[\s"«\.])' + re.escape(f) + r'(?:$|[\s"»\.\,])', up):
            # Если есть кавычки — берём имя оттуда: 'Специализированное "Кузет" АО' → 'Кузет'
            m_q = _Q_RE.search(name)
            if m_q:
                name_part = m_q.group(1).strip()
            else:
                name_part = re.sub(r'(?:^|[\s"«\.])' + re.escape(f) + r'(?:$|[\s"»\.\,])', ' ', core, flags=re.IGNORECASE)
                name_part = name_part.strip('«»"\' -.').strip()
            full_form = _KZ_FORM_FULL.get(f, '')
            query = f"{full_form} {name_part}".strip() if full_form and name_part else (name_part or core)
            return ('kz_ul', query)

    # Полные русские названия KZ форм (Товарищество с ограниченной..., Филиал, Учреждение...)
    if _KZ_FULL_RE.search(name):
        # Приоритет: содержимое кавычек («…», "…", "…")
        m = re.search(r'[«"""„]([^«"""»""]+)[»""""]', name)
        if m:
            query = m.group(1).strip()
        elif re.search(r'\bфилиал\b|\bпредставительство\b', name, re.IGNORECASE):
            # Нет кавычек — убираем "Филиал/Представительство" и "в РК/в Республике Казахстан"
            query = re.sub(r'\bфилиал\b|\bпредставительство\b', ' ', name, flags=re.IGNORECASE)
            query = re.sub(r'\s*в\s+(?:Республике\s+Казахстан|РК)\s*', ' ', query, flags=re.IGNORECASE)
            query = query.strip('«»"\' -.').strip()
        else:
            query = _KZ_FULL_RE.sub(' ', core).strip('«»"\' -.').strip()
        return ('kz_ul', query or core)

    # Проверяем российские формы
    for f in _RU_FORMS:
        if re.search(r'(?:^|[\s"«])' + re.escape(f) + r'(?:$|[\s"»\.])', up):
            return ('russian', '')

    # Проверяем иностранные формы
    for f in _FOR_FORMS:
        if re.search(r'\b' + re.escape(f) + r'\b', up):
            return ('foreign', core)  # передаём имя — иностр. компании могут иметь КЗ БИН

    # Форма не распознана — пробуем как KZ юрлицо с исходным именем
    return ('kz_ul', core)

def _egov_confidence(query: str, result_name: str) -> bool:
    """True если ключевые слова запроса достаточно совпадают с результатом."""
    skip = set([
        "товарищество","ограниченной","ответственностью","дополнительной",
        "акционерное","общество","некоммерческое",
        "государственное","республиканское","коммунальное","муниципальное",
        "предприятие","учреждение","производственный","кооператив",
        "сельскохозяйственный","крестьянское","фермерское","хозяйство",
        "накопительный","пенсионный","микрокредитная","микрофинансовая","организация",
        "компания","группа","объединение","холдинг","корпорация","фонд",
        "тоо","жск","гкп","рсо","мкк","мфо",
        "company","limited","liability","incorporated","corporation",
        "enterprise","holding","group","fund",
        "anonim","sirketi",
    ])
    _quote_re = re.compile(chr(91) + str().join(chr(c) for c in [171,187,8220,8221,8222,8216,8217,8218,8219]) + chr(93))
    _homoglyph_tbl = str.maketrans("СсАаЕеОоРрХхКкТтМмВв", "CcAaEeOoPpXxKkTtMmBb")
    def _norm_word(w: str) -> str:
        has_cyr = any(chr(1024) <= c <= chr(1279) for c in w)
        has_lat = any(chr(65) <= c <= chr(90) or chr(97) <= c <= chr(122) for c in w)
        return w.translate(_homoglyph_tbl) if (has_cyr and has_lat) else w
    _kw_re = re.compile("[^\\W\\d_]{3,}", re.UNICODE)
    def keywords(s: str) -> set:
        s = _quote_re.sub(chr(32), s)
        return {_norm_word(w).lower() for w in _kw_re.findall(s)
                if _norm_word(w).lower() not in skip}
    q_kw = keywords(query)
    r_kw = keywords(result_name)
    if not q_kw:
        return not bool(r_kw)
    def _is_latin(w): return all(ord(c) < 0x400 for c in w)
    def _is_cyr(w):   return all(0x400 <= ord(c) <= 0x4FF for c in w)
    # Запрос чисто Latin, результат чисто Cyrillic — транслитерация, доверяем ES
    if all(_is_latin(w) for w in q_kw) and all(_is_cyr(w) for w in r_kw):
        return True
    # Смешанный скрипт в запросе (напр. "KLV INSAAT (КЛВ ИНШААТ)"):
    # проверяем только кириллическую часть запроса против результата
    # или только латинскую часть — одно из должно быть подмножеством
    q_kw_cyr = {w for w in q_kw if _is_cyr(w)}
    q_kw_lat = {w for w in q_kw if _is_latin(w)}
    if q_kw_cyr and q_kw_lat:
        # оба скрипта в запросе → пробуем каждый независимо
        if q_kw_cyr and q_kw_cyr <= r_kw:
            return True
        if q_kw_lat and q_kw_lat <= r_kw:
            return True
        return False
    # Обычная проверка: все слова запроса есть в результате
    if not (q_kw <= r_kw):
        return False
    extra = len(r_kw - q_kw)
    if extra == 0:
        return True
    # Лишние слова в результате: разрешаем не более 1, но только если слова запроса
    # стоят в результате в ТОМ ЖЕ порядке (неполное название конторы в wt).
    # Если порядок перевёрнут — скорее всего другая компания с теми же словами.
    if extra > 1:
        return False
    _kw_re2 = re.compile("[^\\W\\d_]{3,}", re.UNICODE)
    def _ordered_kw(s):
        return [_norm_word(w).lower() for w in _kw_re2.findall(s)
                if _norm_word(w).lower() in q_kw]
    q_ord = _ordered_kw(query)
    r_ord = _ordered_kw(result_name)
    if len(q_ord) < 2 or len(r_ord) < 2:
        return True
    # Проверяем порядок: позиции слов запроса в результате должны быть монотонно возрастающими
    positions = []
    for w in q_ord:
        try:
            positions.append(r_ord.index(w))
        except ValueError:
            return True
    return positions == sorted(positions)

def _egov_company_summary(name: str) -> dict:
    if name in _egov_cache:
        return _egov_cache[name]

    kind, query = _egov_classify(name)

    if kind == 'skip':
        result = {"found": False, "note": "Пропущено (пустое/служебное значение)"}
    elif kind == 'bank_account':
        result = {"found": False, "note": "Банковский счёт (IBAN)"}
    elif kind == 'internal':
        result = {"found": False, "note": "Внутренняя позиция"}
    elif kind == 'phys':
        result = {"found": False, "note": "Физическое лицо"}
    elif kind == 'russian':
        result = {"found": False, "note": "Иностранная организация (РФ)"}
    elif kind == 'foreign':
        # Иностранные компании могут быть зарегистрированы в КЗ и иметь БИН
        if len(query.strip()) >= 3:
            rows = _egov_v4("gbd_ul", {
                "query": {"query_string": {"query": query, "fields": ["nameru"]}},
                "size": 5,
            })
            if rows:
                best, confident = rows[0], False
                for candidate in rows:
                    if _egov_confidence(query, candidate.get("nameru", "")):
                        best, confident = candidate, True
                        break
                result = {
                    "found":    confident,
                    "nameru":   best.get("nameru", ""),
                    "bin":      best.get("bin", ""),
                    "regdate":  best.get("regdate", ""),
                    "director": best.get("director", ""),
                    "note":     "" if confident else "Иностранная организация (БИН не подтверждён)",
                }
            else:
                result = {"found": False, "note": "Иностранная организация"}
        else:
            result = {"found": False, "note": "Иностранная организация"}
    elif kind == 'kz_ip':
        result = {"found": False, "note": "ИП / ЧП — нет в ГБД ЮЛ"}
    else:
        # kz_ul — ищем в ГБД ЮЛ
        if len(query.strip()) < 3:
            result = {"found": False, "note": "Слишком короткий запрос"}
        else:
            rows = _egov_v4("gbd_ul", {
                "query": {"query_string": {"query": query, "fields": ["nameru"]}},
                "size": 5,
            })
            if not rows:
                result = {"found": False, "note": "Не найдено в ГБД ЮЛ"}
            else:
                # Берём первый из 5 результатов, прошедший confidence
                best, confident = rows[0], False
                for candidate in rows:
                    if _egov_confidence(query, candidate.get("nameru", "")):
                        best, confident = candidate, True
                        break
                # Дефис-ретрай: если 2 значимых слова и уверенного совпадения нет →
                # пробуем соединить дефисом: "Ерболь ЭЛЬ" → "Ерболь-ЭЛЬ"
                if not confident:
                    _org_stop = {
                        'товарищество','ограниченной','ответственностью','дополнительной',
                        'акционерное','общество','некоммерческое','государственное',
                        'республиканское','коммунальное','предприятие','учреждение',
                        'производственный','кооператив','сельскохозяйственный',
                        'крестьянское','фермерское','хозяйство','фонд','корпорация',
                        'тоо','жск','гкп','рсо','мкк','мфо','ип','чп',
                        'с','в','и','на','для',
                    }
                    _meaningful = [w for w in query.split()
                                   if w.lower() not in _org_stop and len(w) >= 2]
                    if len(_meaningful) == 2:
                        hyphen_q = '-'.join(_meaningful)
                        rows2 = _egov_v4("gbd_ul", {
                            "query": {"query_string": {
                                "query": hyphen_q, "fields": ["nameru"]}},
                            "size": 5,
                        })
                        for candidate in (rows2 or []):
                            if _egov_confidence(hyphen_q, candidate.get("nameru", "")):
                                best, confident = candidate, True
                                break
                result = {
                    "found":    confident,
                    "nameru":   best.get("nameru", ""),
                    "bin":      best.get("bin", ""),
                    "director": best.get("director", ""),
                    "datereg":  best.get("datereg", ""),
                    "note":     "" if confident else "Возможно неверное совпадение",
                }

    _egov_cache[name] = result
    return result

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

def has_vba_project(file_bytes: bytes) -> bool:
    """True если в файле есть проект VBA (xlsm). Используем для правильного расширения выходного файла."""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            return "xl/vbaProject.bin" in set(zf.namelist())
    except Exception:
        return False


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
        s = str(v).replace(" ", " ").replace(" ", " ").strip().lower()
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
        s = (s or "").replace(" ", " ").replace(" ", " ").strip()
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
            tail_compact = tail.replace(" ", "").replace(" ", "").replace(" ", "")
            m = re.search(r"(\\d{4})", tail_compact)
            if m:
                return m.group(1)
        # Фолбэк: любые 4 подряд идущие цифры (предварительно "сжимаем" пробелы между цифрами).
        compact = s_norm.replace(" ", "").replace(" ", "").replace(" ", "")
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
    s = s.replace(" ", "").replace(" ", "").replace(" ", "")
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
    s = s.replace(" ", " ").replace(" ", " ").strip()
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
            s = str(v).replace(" ", " ").replace(" ", " ").strip().lower()
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
            s = str(v).replace(" ", " ").replace(" ", " ").strip().lower()
            if s.startswith("итого"):
                end_row = r
                break
    if not start_row or not end_row:
        set_all_rows_height(ws, 12)
        return account_number

    if start_row + 1 <= end_row - 1:
        remove_duplicate_rows(ws, start_row + 1, end_row - 1)

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

    # =========================
    # Сортировка + контроль суммы после очистки (UX/качество данных)
    # =========================
    try:
        # Для счетов, у которых первые 2 цифры < 30: "Итого" берём из колонки G, иначе из H.
        two = int(str(account_number)[:2]) if str(account_number)[:2].isdigit() else 99
        total_col = 7 if two < 30 else 8  # G/H (после нормализации колонок)

        # Пересчитываем границы после удаления строк/колонок.
        start_row2 = find_first_row_with_value(ws, account_number, col=1)
        end_row2 = find_first_row_with_value(ws, "Итого", col=1)
        if end_row2 is None:
            for r in range(1, (ws.max_row or 1) + 1):
                v = ws.cell(row=r, column=1).value
                if v is None:
                    continue
                s = str(v).replace(" ", " ").replace(" ", " ").strip().lower()
                if s.startswith("итого"):
                    end_row2 = r
                    break
        if not start_row2 or not end_row2:
            start_row2, end_row2 = start_row, end_row

        # Формат чисел в блоке счета (включая строку счета и строку "Итого"):
        # от колонки C до "рабочей" колонки (G или H), только ПОСЛЕ очистки/нормализации.
        num_fmt_osv = "#,##0; -#,##0"
        for rr in range(int(start_row2), int(end_row2) + 1):
            for cc in range(3, int(total_col) + 1):  # C..G/H
                ws.cell(row=rr, column=cc).number_format = num_fmt_osv

        # Сортировка по убыванию по "рабочей" колонке (G для дебетовых <30, H для кредитовых >=30).
        data_start = start_row2 + 1
        data_end = end_row2 - 1
        if data_start <= data_end:
            sort_block_by_column(ws, data_start, data_end, col_index=total_col, descending=True)
            # После сортировки стили (в т.ч. красный шрифт) могут "остаться" на старых строках,
            # потому что мы переносим только значения. Сбрасываем цвет и заново красим отрицательные.
            for rr in range(data_start, data_end + 1):
                cell = ws.cell(row=rr, column=total_col)
                v = to_number(cell.value)
                base = cell.font or Font()
                if v is not None and float(v) < 0:
                    cell.font = Font(
                        name=base.name,
                        size=base.size,
                        bold=base.bold,
                        italic=base.italic,
                        underline=base.underline,
                        color="FFCC0000",
                    )
                else:
                    cell.font = Font(
                        name=base.name,
                        size=base.size,
                        bold=base.bold,
                        italic=base.italic,
                        underline=base.underline,
                        color=None,
                    )

        expected_total = to_number(ws.cell(row=start_row2, column=total_col).value) or 0.0

        actual_total = 0.0
        for rr in range(start_row2 + 1, end_row2):
            a = ws.cell(row=rr, column=1).value
            if a is not None and str(a).replace(" ", " ").replace(" ", " ").strip().lower().startswith("итого"):
                break
            v = to_number(ws.cell(row=rr, column=total_col).value)
            if v is not None:
                actual_total += float(v)

        if expected_total == 0.0:
            diff_pct = 0.0 if abs(actual_total) < 1e-9 else 1.0
        else:
            diff_pct = abs(actual_total - expected_total) / abs(expected_total)

        # % расхождения: пишем, но НЕ заливаем
        out_cell = ws.cell(row=1, column=total_col)
        out_cell.value = float(diff_pct)
        out_cell.number_format = "0.0%"
        out_cell.font = Font(name="Arial", size=9, bold=True, color="FF4B5563")
        out_cell.alignment = Alignment(horizontal="center", vertical="center")
        out_cell.fill = PatternFill(fill_type=None)

        # Статус: F1 если % в G1, либо G1 если % в H1
        status_col = 6 if total_col == 7 else 7  # F/G
        status_cell = ws.cell(row=1, column=status_col)
        status_cell.font = Font(name="Arial", size=9, bold=True, color="FF111827")
        status_cell.alignment = Alignment(horizontal="center", vertical="center")

        fill_match = PatternFill(fill_type="solid", fgColor="FFE4F0DD")     # мягкий зелёный
        fill_warn = PatternFill(fill_type="solid", fgColor="FFFFF2CC")      # мягкий жёлтый
        fill_check = PatternFill(fill_type="solid", fgColor="FFFCE4EC")     # мягкий красный

        if diff_pct <= 1e-12:
            status_cell.value = "MATCH"
            status_cell.fill = fill_match
        elif diff_pct < 0.10:
            status_cell.value = "MISMATCH"
            status_cell.fill = fill_warn
        else:
            status_cell.value = "CHECK"
            status_cell.fill = fill_check
    except Exception:
        pass

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
    saldo_suffixes = {"1210", "1710", "3310", "3510"}

    # Запасы: считаем доступными, если есть хотя бы один ОСВ-лист с 4-значным счётом, начинающимся на "13"
    inventory_accounts_found: List[str] = []
    inv_set: Set[str] = set()
    for sh in wb.sheetnames:
        _p, suf4 = split_prefix_suffix4(sh)
        if suf4.isdigit() and len(suf4) == 4 and suf4.startswith("13"):
            inv_set.add(suf4)
    if inv_set:
        inventory_accounts_found = sorted(inv_set)

    saldo_ok = any(split_prefix_suffix4(sh)[1] in saldo_suffixes for sh in wb.sheetnames)

    # "общ*" (общая ОСВ) — для чистой прибыли достаточно наличия хотя бы одного такого листа
    obsh_sheets = [sh for sh in wb.sheetnames if "общ" in str(sh).lower()]
    profit_ok = bool(obsh_sheets)

    prefix_to_pair: Dict[str, Dict[str, str]] = defaultdict(dict)
    for sh in wb.sheetnames:
        prefix, suf2 = split_prefix_suffix2(sh)
        if suf2 in ("wd", "md"):
            prefix = normalize_prefix(prefix)
            if suf2 not in prefix_to_pair[prefix]:
                prefix_to_pair[prefix][suf2] = sh
    contracts_prefixes = [p for p, d in prefix_to_pair.items() if "wd" in d and "md" in d]
    contracts_ok = bool(contracts_prefixes)

    # Инсайты: нужен набор листов W/M/Wt/Mt (возможно с префиксом).
    required_ins = {'w','m','wt','mt'}
    prefix_to_ins: Dict[str, Set[str]] = defaultdict(set)
    for sh in wb.sheetnames:
        low = sh.lower()
        for suf in ('wt','mt','w','m','кред'):
            if low.endswith(suf):
                prefix_to_ins[normalize_prefix(sh[:-len(suf)])].add(suf)
    insights_prefixes = [p for p, s in prefix_to_ins.items() if required_ins.issubset(s)]

    # Листы инсайтов уже могут быть в файле и их нельзя перезатирать.
    # Также поддерживаем миграцию старого имени листа: "инсайты" -> "инс".
    # Поэтому считаем доступность как:
    # - есть ли хотя бы один префикс-набор, где "инс" ещё НЕ создан
    # - или есть ли хотя бы один старый лист "инсайты", который нужно переименовать.
    insights_existing_titles: List[str] = []
    insights_missing_titles: List[str] = []
    insights_legacy_titles: List[str] = []
    for pfx in insights_prefixes:
        base_new = 'инс' if not pfx else f'{pfx}инс'
        title_new = base_new[:31]
        base_old = 'инсайты' if not pfx else f'инсайты {pfx}'
        title_old = base_old[:31]

        if title_new in wb.sheetnames:
            insights_existing_titles.append(title_new)
        elif title_old in wb.sheetnames:
            insights_existing_titles.append(title_old)
            insights_legacy_titles.append(title_old)
        else:
            insights_missing_titles.append(title_new)

    insights_ok = bool(insights_missing_titles) or bool(insights_legacy_titles)

    # Госы: листы, содержащие "гос" в названии (без учёта регистра)
    gos_sheets = [sh for sh in wb.sheetnames if "гос" in str(sh).lower()]
    gos_ok = bool(gos_sheets)

    # Объединение ОСВ по группе: «*общ*» листы с ≥2 разными префиксами на один суффикс
    _grp_by_suffix: Dict[str, List[str]] = {}
    for sh in wb.sheetnames:
        sh_lower = sh.lower()
        idx = sh_lower.find("общ")
        if idx < 0:
            continue
        suffix = sh[idx + 3:]  # всё после «общ»
        _grp_by_suffix.setdefault(suffix, []).append(sh)

    valid_group_osv: Dict[str, List[str]] = {}
    for suf, sheets in _grp_by_suffix.items():
        # Проверяем, что у нас ≥2 разных префиксов
        prefixes = set(sh[:sh.lower().find("общ")] for sh in sheets)
        if len(prefixes) >= 2:
            valid_group_osv[suf] = sheets

    group_osv_ok = bool(valid_group_osv)

    # Контрагенты (egov): листы, заканчивающиеся на "wt" (регистр не важен)
    wt_sheets = [sh for sh in wb.sheetnames if sh.strip().lower().endswith("wt")]
    clients_ok = bool(wt_sheets)

    return {
        "inventory_accounts_found": inventory_accounts_found,
        "saldo_ok": saldo_ok,
        "contracts_ok": contracts_ok,
        "contracts_prefixes": contracts_prefixes,
        "profit_ok": profit_ok,
        "obsh_sheets": obsh_sheets,
        "insights_ok": insights_ok,
        "insights_prefixes": insights_prefixes,
        "insights_existing_titles": insights_existing_titles,
        "insights_missing_titles": insights_missing_titles,
        "insights_legacy_titles": insights_legacy_titles,
        "gos_ok": gos_ok,
        "gos_sheets": gos_sheets,
        "group_osv_ok": group_osv_ok,
        "valid_group_osv": valid_group_osv,
        "clients_ok": clients_ok,
        "wt_sheets": wt_sheets,
    }


def compute_mergeable_osv_from_wb(wb) -> Dict[str, List[str]]:
    """Возвращает {account: [sheet1, sheet2, ...]} для счетов, у которых ≥2 листа в одной книге."""
    osv_by_account: Dict[str, List[str]] = defaultdict(list)
    for sh in wb.sheetnames:
        _, suf4 = split_prefix_suffix4(sh)
        if suf4.isdigit() and len(suf4) == 4:
            osv_by_account[suf4].append(sh)
    return {acc: sheets for acc, sheets in osv_by_account.items() if len(sheets) >= 2}


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


def list_existing_osv_sheets_with_a1(wb) -> List[Tuple[str, str, str]]:
    """Возвращает кортежи (имя_листа, 4-значный суффикс, A1_текст) для всех ОСВ-листов."""
    out = []
    for sh in wb.sheetnames:
        _prefix, suf = split_prefix_suffix4(sh)
        if suf.isdigit() and len(suf) == 4:
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

    def _find_prefixed_sheetname_any(prefix_norm: str, base: str) -> Optional[str]:
        """
        Ищет лист вида "<префикс><base>" (регистр не важен),
        также допускается суффикс " (n)". Возвращает лучший кандидат или None.
        """
        base_l = (base or "").lower()
        best = None
        best_key = None
        for name in wb.sheetnames:
            nm = str(name)
            nm2 = re.sub(r"\s*\(\d+\)$", "", nm).strip()
            if not nm2.lower().endswith(base_l):
                continue
            pref_part = nm2[: -len(base_l)]
            if normalize_prefix(pref_part) != prefix_norm:
                continue
            m = re.search(r"\((\d+)\)\s*$", nm)
            n = int(m.group(1)) if m else 0
            key = (1 if m else 0, n, len(nm))  # предпочитаем точное имя без " (n)"
            if best is None or key < best_key:
                best = nm
                best_key = key
        return best

    TOP_N = 100

    def _topn_pos_neg(df: pd.DataFrame, col: str) -> pd.DataFrame:
        if df is None or df.empty or col not in df.columns:
            return df.iloc[0:0].copy()
        pos = df[df[col] > 0].nlargest(TOP_N, col) if (df[col] > 0).any() else df.iloc[0:0]
        neg = df[df[col] < 0].nsmallest(TOP_N, col) if (df[col] < 0).any() else df.iloc[0:0]
        out = pd.concat([pos, neg], ignore_index=True)
        return out.reset_index(drop=True)

    for prefix, suf_map in prefix_to_sheets.items():
        sheet_data = {}

        # Для расширенных колонок в "сальд": начало/обороты/конец по 1210/1710/3310/3510
        # (используем фиксированные позиции как в ОСВ: C..H -> индексы 2..7).
        sheet_metrics: Dict[str, Dict[str, pd.Series]] = {}

        def _clean_contr(v) -> str:
            if v is None:
                return ""
            s = str(v).strip()
            if not s:
                return ""
            if s.lower().startswith("итого"):
                return ""
            if s in ("1210", "1710", "3310", "3510"):
                return ""
            return s

        def _series_by_contr(df: pd.DataFrame, col_i: int) -> pd.Series:
            if df is None or df.empty or int(df.shape[1]) <= int(col_i):
                return pd.Series(dtype=float)
            temp = df.iloc[:, [0, col_i]].copy()
            temp.columns = ["Контрагент", "value"]
            temp["Контрагент"] = temp["Контрагент"].map(_clean_contr)
            temp = temp[temp["Контрагент"] != ""]
            temp["value"] = pd.to_numeric(temp["value"], errors="coerce").fillna(0)
            temp = temp[temp["value"] != 0]
            if temp.empty:
                return pd.Series(dtype=float)
            return temp.groupby("Контрагент")["value"].sum()

        for suf, col_idx in target_suffixes.items():
            if suf not in suf_map:
                continue

            df = pd.read_excel(xls, sheet_name=suf_map[suf], header=0)
            if df is None or df.empty:
                continue

            # Конец периода (как было раньше)
            s_end = _series_by_contr(df, col_idx)
            if not s_end.empty:
                sheet_data[suf] = s_end

            # Расширенные показатели (начало/обороты/конец по нужной стороне)
            m = {
                "bop_deb": _series_by_contr(df, 2),   # C
                "bop_cred": _series_by_contr(df, 3),  # D
                "turn_deb": _series_by_contr(df, 4),  # E
                "turn_cred": _series_by_contr(df, 5), # F
                "eop_deb": _series_by_contr(df, 6),   # G
                "eop_cred": _series_by_contr(df, 7),  # H
            }
            sheet_metrics[suf] = m

        m1210 = sheet_metrics.get("1210", {})
        m3510 = sheet_metrics.get("3510", {})
        m1710 = sheet_metrics.get("1710", {})
        m3310 = sheet_metrics.get("3310", {})

        def _ser(m: Dict[str, pd.Series], key: str) -> pd.Series:
            s = m.get(key)
            return s if isinstance(s, pd.Series) else pd.Series(dtype=float)

        # G(eop_deb) и H(eop_cred) с обоих счетов — в cust_set/supp_set
        cust_set = (
            set(_ser(m1210, "eop_deb").index)
            .union(set(_ser(m1210, "eop_cred").index))
            .union(set(_ser(m3510, "eop_deb").index))
            .union(set(_ser(m3510, "eop_cred").index))
        )
        supp_set = (
            set(_ser(m1710, "eop_deb").index)
            .union(set(_ser(m1710, "eop_cred").index))
            .union(set(_ser(m3310, "eop_deb").index))
            .union(set(_ser(m3310, "eop_cred").index))
        )
        all_set = cust_set.union(supp_set)

        if not cust_set and not supp_set:
            continue

        if cust_set:
            df_cust = pd.DataFrame(sorted(cust_set), columns=["Контрагент"])

            df_cust["1210"] = (
                df_cust["Контрагент"].map(_ser(m1210, "eop_deb")).fillna(0)
                + df_cust["Контрагент"].map(_ser(m3510, "eop_deb")).fillna(0)
            ) / 1000
            df_cust["3510"] = (
                df_cust["Контрагент"].map(_ser(m1210, "eop_cred")).fillna(0)
                + df_cust["Контрагент"].map(_ser(m3510, "eop_cred")).fillna(0)
            ) / 1000
            df_cust["1210_нач"] = (
                df_cust["Контрагент"].map(_ser(m1210, "bop_deb")).fillna(0)
                + df_cust["Контрагент"].map(_ser(m3510, "bop_deb")).fillna(0)
            ) / 1000
            df_cust["3510_нач"] = (
                df_cust["Контрагент"].map(_ser(m1210, "bop_cred")).fillna(0)
                + df_cust["Контрагент"].map(_ser(m3510, "bop_cred")).fillna(0)
            ) / 1000
            df_cust["нетто_нач"] = df_cust["1210_нач"] - df_cust["3510_нач"]
            df_cust["оборот_дт"] = (
                df_cust["Контрагент"].map(_ser(m1210, "turn_deb")).fillna(0)
                + df_cust["Контрагент"].map(_ser(m3510, "turn_deb")).fillna(0)
            ) / 1000
            df_cust["оборот_кт"] = (
                df_cust["Контрагент"].map(_ser(m1210, "turn_cred")).fillna(0)
                + df_cust["Контрагент"].map(_ser(m3510, "turn_cred")).fillna(0)
            ) / 1000

            df_cust["сальдо заказчики"] = df_cust["1210"] - df_cust["3510"]
            df_cust = df_cust.sort_values(by="сальдо заказчики", ascending=False).reset_index(drop=True)
        else:
            df_cust = pd.DataFrame(columns=[
                "Контрагент", "1210_нач", "3510_нач", "нетто_нач", "оборот_дт", "оборот_кт",
                "1210", "3510", "сальдо заказчики",
            ])

        if supp_set:
            df_supp = pd.DataFrame(sorted(supp_set), columns=["Контрагент"])

            df_supp["1710"] = (
                df_supp["Контрагент"].map(_ser(m1710, "eop_deb")).fillna(0)
                + df_supp["Контрагент"].map(_ser(m3310, "eop_deb")).fillna(0)
            ) / 1000
            df_supp["3310"] = (
                df_supp["Контрагент"].map(_ser(m1710, "eop_cred")).fillna(0)
                + df_supp["Контрагент"].map(_ser(m3310, "eop_cred")).fillna(0)
            ) / 1000
            df_supp["1710_нач"] = (
                df_supp["Контрагент"].map(_ser(m1710, "bop_deb")).fillna(0)
                + df_supp["Контрагент"].map(_ser(m3310, "bop_deb")).fillna(0)
            ) / 1000
            df_supp["3310_нач"] = (
                df_supp["Контрагент"].map(_ser(m1710, "bop_cred")).fillna(0)
                + df_supp["Контрагент"].map(_ser(m3310, "bop_cred")).fillna(0)
            ) / 1000
            df_supp["нетто_нач"] = df_supp["1710_нач"] - df_supp["3310_нач"]
            df_supp["оборот_дт"] = (
                df_supp["Контрагент"].map(_ser(m1710, "turn_deb")).fillna(0)
                + df_supp["Контрагент"].map(_ser(m3310, "turn_deb")).fillna(0)
            ) / 1000
            df_supp["оборот_кт"] = (
                df_supp["Контрагент"].map(_ser(m1710, "turn_cred")).fillna(0)
                + df_supp["Контрагент"].map(_ser(m3310, "turn_cred")).fillna(0)
            ) / 1000

            df_supp["сальдо поставщики"] = df_supp["1710"] - df_supp["3310"]
            df_supp = df_supp.sort_values(by="сальдо поставщики", ascending=False).reset_index(drop=True)
        else:
            df_supp = pd.DataFrame(columns=[
                "Контрагент", "1710_нач", "3310_нач", "нетто_нач", "оборот_дт", "оборот_кт",
                "1710", "3310", "сальдо поставщики",
            ])

        if all_set:
            df_total = pd.DataFrame(sorted(all_set), columns=["Контрагент"])
            df_total["1210"] = (
                df_total["Контрагент"].map(_ser(m1210, "eop_deb")).fillna(0)
                + df_total["Контрагент"].map(_ser(m3510, "eop_deb")).fillna(0)
            ) / 1000
            df_total["1710"] = (
                df_total["Контрагент"].map(_ser(m1710, "eop_deb")).fillna(0)
                + df_total["Контрагент"].map(_ser(m3310, "eop_deb")).fillna(0)
            ) / 1000
            df_total["3310"] = (
                df_total["Контрагент"].map(_ser(m1710, "eop_cred")).fillna(0)
                + df_total["Контрагент"].map(_ser(m3310, "eop_cred")).fillna(0)
            ) / 1000
            df_total["3510"] = (
                df_total["Контрагент"].map(_ser(m1210, "eop_cred")).fillna(0)
                + df_total["Контрагент"].map(_ser(m3510, "eop_cred")).fillna(0)
            ) / 1000
            df_total["общее сальдо"] = df_total["1210"] + df_total["1710"] - df_total["3310"] - df_total["3510"]
            df_total = df_total.sort_values(by="общее сальдо", ascending=False).reset_index(drop=True)
        else:
            df_total = pd.DataFrame(columns=["Контрагент", "общее сальдо"])

        # =========================
        # Лист "сальд" — топ-15 (плюс/минус) + блоки оплат/выполнений с формулами по месяцам
        # =========================
        out2_name = safe_sheet_name(f"{prefix}сальд" if prefix else "сальд")
        if out2_name in wb.sheetnames:
            wb.remove(wb[out2_name])

        # Порядок вкладок: "сальд" должен идти сразу после листа "общ" соответствующего префикса.
        # Если "общ" не найден — добавляем в конец.
        pref_norm = normalize_prefix(prefix)
        obsh_name = _find_prefixed_sheetname_any(pref_norm, "общ") or (f"{prefix}общ" if prefix else "общ")
        insert_index = None
        if obsh_name in wb.sheetnames:
            try:
                insert_index = wb.sheetnames.index(obsh_name) + 1
            except Exception:
                insert_index = None
        ws2 = wb.create_sheet(out2_name, insert_index) if insert_index is not None else wb.create_sheet(out2_name)

        ws2["A1"] = "Все значения указаны в тысячах тенге"
        ws2["A1"].font = Font(name="Arial", size=8, bold=True)

        # На "сальд" везде используем Arial 8.
        font_h  = Font(name="Arial", size=8, bold=True)
        font_b  = Font(name="Arial", size=8)
        font_bb = Font(name="Arial", size=8, bold=True)
        font_key = Font(name="Arial", size=8, color="B3B3B3")  # серый ключ
        align_c = Alignment(horizontal="center")

        def _make_key(name: str) -> str:
            """Ключ контрагента (Python): строчные, без пробелов и знаков препинания."""
            if not name:
                return ""
            s = name
            for ch in ["-", "\u2013", "\u2014", " ", "\u00a0", ">", "<", "_", "?", ".", ",", "(", ")", "\"", "/", "\\", "!", ":", ";", "@", "#", "$", "%", "^", "&", "*", "+", "=", "|", "~", "`", "[", "]", "{", "}", "\u00ab", "\u00bb", "\u2018", "\u2019", "\u201c", "\u201d"]:
                s = s.replace(ch, "")
            return s.lower()

        align_l = Alignment(horizontal="left")
        num_fmt = "#,##0;[Red](#,##0)"

        # A2 = 6 (влево + мягкая серая заливка)
        ws2["A2"] = 6
        ws2["A2"].alignment = Alignment(horizontal="left")
        ws2["A2"].fill = PatternFill("solid", fgColor="D9D9D9")

        # Используем конкатенацию через &, чтобы Excel не добавлял неявное пересечение ("@")
        ws2["L2"] = '="Опл L"&$A$2&"M"'
        ws2["M2"] = '="Вып L"&$A$2&"M"'
        ws2["K2"] = "коммент"

        # Заголовки месяцев (сдвинуты вправо из-за доп. колонок после "Контрагент")
        months = [f"2025_{m:02d}" for m in range(1, 13)] + [f"2026_{m:02d}" for m in range(1, 13)]
        start_col_pay = 14  # N
        start_col_perf = start_col_pay + len(months) + 1  # + разделитель
        last_used_col = start_col_perf + len(months) - 1

        ws2.cell(row=1, column=start_col_pay, value="Оплаты").font = font_h
        ws2.cell(row=1, column=start_col_perf, value="Выполнения").font = font_h

        for i, m in enumerate(months):
            ws2.cell(row=2, column=start_col_pay + i, value=m).font = font_h
            ws2.cell(row=2, column=start_col_pay + i).alignment = align_c
            ws2.cell(row=2, column=start_col_perf + i, value=m).font = font_h
            ws2.cell(row=2, column=start_col_perf + i).alignment = align_c

        # Заголовки таблиц (Заказчики)
        ws2["B2"] = "Контрагент"
        ws2["C2"] = "1210 нач"
        ws2["D2"] = "3510 нач"
        ws2["E2"] = "Нетто ДЗ (нач)"
        ws2["F2"] = "оборот ДТ"
        ws2["G2"] = "оборот КТ"
        ws2["H2"] = "1210"
        ws2["I2"] = "3510"
        ws2["J2"] = "Нетто ДЗ (кон)"
        for c in ("B2","C2","D2","E2","F2","G2","H2","I2","J2","K2","L2","M2"):
            ws2[c].font = font_h
            ws2[c].alignment = align_c if c != "B2" else align_l

        # Визуальные разделители: "ЗАКАЗЧИКИ" (до колонок L/M, не залезаем на блок месяцев)
        sec_fill = PatternFill(fill_type="solid", fgColor="1F2937")
        sec_font = Font(name="Arial", size=8, bold=True, color="FFFFFF")
        for col in range(2, last_used_col + 1):  # B..(конец таблицы)
            c = ws2.cell(row=1, column=col)
            c.fill = sec_fill
            c.font = sec_font
        ws2["B1"] = "ЗАКАЗЧИКИ"
        ws2["B1"].alignment = align_l

        # Блоки "Поставщики" и "Общее сальдо" размещаем ниже динамически (так как добавляем строки сумм/итого).

        # Подбираем Wr/Mr для текущего префикса (если такие листы есть) — имена нужны внутри формул.
        wr_name = _find_prefixed_sheetname(pref_norm, "Wr") or (f"{prefix}Wr" if prefix else "Wr")
        mr_name = _find_prefixed_sheetname(pref_norm, "Mr") or (f"{prefix}Mr" if prefix else "Mr")
        wr_ref = _excel_sheet_ref(wr_name)
        mr_ref = _excel_sheet_ref(mr_name)

        pay_start = get_column_letter(start_col_pay)
        pay_end = get_column_letter(start_col_pay + len(months) - 1)
        perf_start = get_column_letter(start_col_perf)
        perf_end = get_column_letter(start_col_perf + len(months) - 1)

        def _comment_formula_for_row(r: int, which: str) -> str:
            """
            which: 'cust' or 'supp'
            Возвращает Excel-формулу для колонки F на основе:
            - оплат J:AG
            - выполнений AI:BF (нужно только для правила по 3510)
            """
            last_pos = f"IFERROR(LOOKUP(2,1/(${pay_start}{r}:${pay_end}{r}<>0),COLUMN(${pay_start}{r}:${pay_end}{r}))-COLUMN(${pay_start}{r})+1,0)"
            months_since_last_pay = (
                f"IF(({last_pos})=0,999,"
                f"DATEDIF("
                f"DATE(LEFT(INDEX(${pay_start}$2:${pay_end}$2,1,({last_pos})),4),RIGHT(INDEX(${pay_start}$2:${pay_end}$2,1,({last_pos})),2),1),"
                f"DATE(YEAR(TODAY()),MONTH(TODAY()),1),\"m\"))"
            )

            def _rule_by_last_pay(t_ok: int, t_warn: int, warn_label: str) -> str:
                return (
                    f"IF(({last_pos})=0,\"списание\","
                    f"IF(({months_since_last_pay})<={t_ok},\"ОК\","
                    f"IF(({months_since_last_pay})<={t_warn},\"{warn_label}\",\"списание\")))"
                )

            idx_now_pay = (
                f'IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),${pay_start}$2:${pay_end}$2,0),'
                f'COUNTA(${pay_start}$2:${pay_end}$2))'
            )
            idx_now_perf = (
                f'IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),${perf_start}$2:${perf_end}$2,0),'
                f'COUNTA(${perf_start}$2:${perf_end}$2))'
            )
            sum_pay_0_3 = f"SUM(INDEX(${pay_start}{r}:${pay_end}{r},1,MAX(1,({idx_now_pay})-2)):INDEX(${pay_start}{r}:${pay_end}{r},1,({idx_now_pay})))"
            sum_perf_0_3 = f"SUM(INDEX(${perf_start}{r}:${perf_end}{r},1,MAX(1,({idx_now_perf})-2)):INDEX(${perf_start}{r}:${perf_end}{r},1,({idx_now_perf})))"
            sum_pay_0_12 = f"SUM(INDEX(${pay_start}{r}:${pay_end}{r},1,MAX(1,({idx_now_pay})-11)):INDEX(${pay_start}{r}:${pay_end}{r},1,({idx_now_pay})))"
            sum_perf_0_12 = f"SUM(INDEX(${perf_start}{r}:${perf_end}{r},1,MAX(1,({idx_now_perf})-11)):INDEX(${perf_start}{r}:${perf_end}{r},1,({idx_now_perf})))"
            rule_3510 = (
                f"IF(AND(({sum_pay_0_3})<>0,({sum_perf_0_3})<>0),\"ОК\","
                f"IF(OR(({sum_pay_0_12})<>0,({sum_perf_0_12})<>0),\"сомнительно\",\"списание\"))"
            )

            if which == "cust":
                # "ОК" only when BOTH recent payment AND non-zero turnover in last 3 months
                rule_1210 = (
                    f"IF(({last_pos})=0,\"списание\","
                    f"IF(AND(({months_since_last_pay})<=3,({sum_perf_0_3})<>0),\"ОК\","
                    f"IF(({months_since_last_pay})<=6,\"сомнительный\",\"списание\")))"
                )
                return f"=IF($B{r}=\"\",\"\",IF($J{r}>0,{rule_1210},{rule_3510}))"
            # поставщики
            rule_1710 = _rule_by_last_pay(3, 6, "сомнительный")
            rule_3310 = _rule_by_last_pay(3, 12, "сомнительный")
            return f"=IF($B{r}=\"\",\"\",IF($J{r}>0,{rule_1710},{rule_3310}))"

        def _rollup_pay_formula(r: int) -> str:
            return (
                f'=IF($B{r}="","",SUM('
                f'INDEX(${pay_start}{r}:${pay_end}{r},1,MAX(1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),${pay_start}$2:${pay_end}$2,0),COUNTA(${pay_start}$2:${pay_end}$2))-$A$2+1)):'
                f'INDEX(${pay_start}{r}:${pay_end}{r},1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),${pay_start}$2:${pay_end}$2,0),COUNTA(${pay_start}$2:${pay_end}$2)))))'
            )

        def _rollup_perf_formula(r: int) -> str:
            return (
                f'=IF($B{r}="","",SUM('
                f'INDEX(${perf_start}{r}:${perf_end}{r},1,MAX(1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),${perf_start}$2:${perf_end}$2,0),COUNTA(${perf_start}$2:${perf_end}$2))-$A$2+1)):'
                f'INDEX(${perf_start}{r}:${perf_end}{r},1,IFERROR(MATCH(TEXT(TODAY(),"yyyy")&"_"&TEXT(TODAY(),"mm"),${perf_start}$2:${perf_end}$2,0),COUNTA(${perf_start}$2:${perf_end}$2)))))'
            )

        def _write_customer_row(
            r: int,
            contr: str,
            v1210_bop: int,
            v3510_bop: int,
            vnet_bop: int,
            vturn_deb: int,
            vturn_cred: int,
            v1210: int,
            v3510: int,
            vsaldo: int,
        ) -> None:
            # Колонка A: ключ контрагента (значение, серый)
            ws2.cell(row=r, column=1, value=_make_key(contr)).font = font_key
            ws2.cell(row=r, column=1).alignment = align_l

            ws2.cell(row=r, column=2, value=contr).font = font_b
            ws2.cell(row=r, column=2).alignment = align_l

            # BOP/обороты/конец
            for c, val in [
                (3, v1210_bop),
                (4, v3510_bop),
                (5, vnet_bop),
                (6, vturn_deb),
                (7, vturn_cred),
                (8, v1210),
                (9, v3510),
                (10, vsaldo),
            ]:
                cell = ws2.cell(row=r, column=c, value=val)
                cell.number_format = num_fmt
                cell.font = font_b
                cell.alignment = align_c

            for c in range(3, 11):
                ws2.cell(row=r, column=c).alignment = align_c

            ws2.cell(row=r, column=11, value=_comment_formula_for_row(r, "cust")).alignment = align_l
            ws2.cell(row=r, column=11).font = font_b

            ws2.cell(row=r, column=12, value=_rollup_pay_formula(r)).alignment = align_c
            ws2.cell(row=r, column=12).font = font_b
            ws2.cell(row=r, column=12).number_format = num_fmt

            ws2.cell(row=r, column=13, value=_rollup_perf_formula(r)).alignment = align_c
            ws2.cell(row=r, column=13).font = font_b
            ws2.cell(row=r, column=13).number_format = num_fmt

            # J..AG: оплаты из Wr
            for i in range(len(months)):
                ci = start_col_pay + i
                col_letter = get_column_letter(ci)
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",'
                        f'SUMIFS({wr_ref}!$F:$F,{wr_ref}!$G:$G,"1210",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$Z:$Z,$A{r})'
                        f'+SUMIFS({wr_ref}!$F:$F,{wr_ref}!$G:$G,"3510",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$Z:$Z,$A{r})'
                        f')'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

            # AI..AT (2025) *1.12 и AU..BF (2026) *1.16 — выполнения из Mr
            for i in range(len(months)):
                ci = start_col_perf + i
                col_letter = get_column_letter(ci)
                coef = 1.12 if i < 12 else 1.16
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",SUMIFS({mr_ref}!$H:$H,{mr_ref}!$P:$P,{col_letter}$2,{mr_ref}!$X:$X,$A{r},{mr_ref}!$G:$G,"6010")*{coef})'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

        def _write_supplier_row(
            r: int,
            contr: str,
            v1710_bop: int,
            v3310_bop: int,
            vnet_bop: int,
            vturn_deb: int,
            vturn_cred: int,
            v1710: int,
            v3310: int,
            vsaldo: int,
        ) -> None:
            # Колонка A: ключ контрагента (значение, серый)
            ws2.cell(row=r, column=1, value=_make_key(contr)).font = font_key
            ws2.cell(row=r, column=1).alignment = align_l

            ws2.cell(row=r, column=2, value=contr).font = font_b
            ws2.cell(row=r, column=2).alignment = align_l

            for c, val in [
                (3, v1710_bop),
                (4, v3310_bop),
                (5, vnet_bop),
                (6, vturn_deb),
                (7, vturn_cred),
                (8, v1710),
                (9, v3310),
                (10, vsaldo),
            ]:
                cell = ws2.cell(row=r, column=c, value=val)
                cell.number_format = num_fmt
                cell.font = font_b
                cell.alignment = align_c

            ws2.cell(row=r, column=11, value=_comment_formula_for_row(r, "supp")).alignment = align_l
            ws2.cell(row=r, column=11).font = font_b

            ws2.cell(row=r, column=12, value=_rollup_pay_formula(r)).alignment = align_c
            ws2.cell(row=r, column=12).font = font_b
            ws2.cell(row=r, column=12).number_format = num_fmt

            # Для поставщиков: только оплаты J..AG (выполнений нет)
            for i in range(len(months)):
                ci = start_col_pay + i
                col_letter = get_column_letter(ci)
                ws2.cell(
                    row=r,
                    column=ci,
                    value=(
                        f'=IF($B{r}="","",'
                        f'SUMIFS({wr_ref}!$H:$H,{wr_ref}!$E:$E,"1710",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$Y:$Y,$A{r})'
                        f'+SUMIFS({wr_ref}!$H:$H,{wr_ref}!$E:$E,"3310",{wr_ref}!$P:$P,{col_letter}$2,{wr_ref}!$Y:$Y,$A{r})'
                        f')'
                    ),
                ).number_format = num_fmt
                ws2.cell(row=r, column=ci).alignment = align_c

        def _write_summary_rows(kind: str, df_total_src: pd.DataFrame, pos_start: int, pos_end: int, sum_row: int, other_row: int, total_row: int) -> None:
            """
            Writes TOP-N sum, прочее, ИТОГО rows for a block.
            kind: 'cust' or 'supp'
            """
            # Подписи строк
            # Жирным выделяем только строки "ТОП-15" и "ИТОГО"; "прочее" оставляем обычным.
            ws2.cell(row=sum_row, column=2, value=f"ТОП-{TOP_N}").font = font_h
            ws2.cell(row=other_row, column=2, value="прочее").font = font_b
            ws2.cell(row=total_row, column=2, value="ИТОГО").font = font_h
            for rr in (sum_row, other_row, total_row):
                ws2.cell(row=rr, column=2).alignment = align_l

            def _sum_formula(col_letter: str) -> str:
                return f"=SUM({col_letter}{pos_start}:{col_letter}{pos_end})"

            # Суммы по ТОП-N (числовые колонки + роллапы + помесячные)
            base_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "L"]
            if kind == "cust":
                base_cols.append("M")
            for col_letter in base_cols:
                c = ws2[f"{col_letter}{sum_row}"]
                c.value = _sum_formula(col_letter)
                c.font = font_bb
                c.alignment = align_c
                c.number_format = num_fmt

            for i in range(len(months)):
                ci = start_col_pay + i
                col_letter = get_column_letter(ci)
                c = ws2.cell(row=sum_row, column=ci, value=_sum_formula(col_letter))
                c.alignment = align_c
                c.number_format = num_fmt
                c.font = font_bb
            if kind == "cust":
                for i in range(len(months)):
                    ci = start_col_perf + i
                    col_letter = get_column_letter(ci)
                    c = ws2.cell(row=sum_row, column=ci, value=_sum_formula(col_letter))
                    c.alignment = align_c
                    c.number_format = num_fmt
                    c.font = font_bb

            # ИТОГО:
            # - по сальдо (C/D/E) пишем числом (посчитано в pandas для текущего блока: + или -)
            # - оплаты/выполнения по месяцам считаем формулами без критерия контрагента
            if kind == "cust":
                def _sum_df(col: str) -> int:
                    return int(round(float(df_total_src[col].sum() if df_total_src is not None and not df_total_src.empty and col in df_total_src.columns else 0.0)))

                ws2[f"C{total_row}"] = _sum_df("1210_нач")
                ws2[f"D{total_row}"] = _sum_df("3510_нач")
                ws2[f"E{total_row}"] = _sum_df("нетто_нач")
                ws2[f"F{total_row}"] = _sum_df("оборот_дт")
                ws2[f"G{total_row}"] = _sum_df("оборот_кт")
                ws2[f"H{total_row}"] = _sum_df("1210")
                ws2[f"I{total_row}"] = _sum_df("3510")
                tot_net = _sum_df("сальдо заказчики")
                if df_total_src is not None and not df_total_src.empty:
                    try:
                        if float(df_total_src["сальдо заказчики"].sum()) < 0:
                            tot_net = abs(int(tot_net))
                    except Exception:
                        pass
                ws2[f"J{total_row}"] = tot_net
                for addr in (f"C{total_row}", f"D{total_row}", f"E{total_row}", f"F{total_row}", f"G{total_row}", f"H{total_row}", f"I{total_row}", f"J{total_row}"):
                    ws2[addr].font = font_bb
                    ws2[addr].alignment = align_c
                    ws2[addr].number_format = num_fmt

                # ИТОГО по месяцам (оплаты Wr, выполнения Mr) без фильтра по контрагенту
                for i in range(len(months)):
                    ci = start_col_pay + i
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

                for i in range(len(months)):
                    ci = start_col_perf + i
                    col_letter = get_column_letter(ci)
                    coef = 1.12 if i < 12 else 1.16
                    ws2.cell(
                        row=total_row,
                        column=ci,
                        value=f'=SUMIFS({mr_ref}!$H:$H,{mr_ref}!$P:$P,{col_letter}$2,{mr_ref}!$G:$G,"6010")*{coef}',
                    ).number_format = num_fmt
                    ws2.cell(row=total_row, column=ci).alignment = align_c
                    ws2.cell(row=total_row, column=ci).font = font_bb

                # G/H суммы за "последние A2 месяцев" для строки ИТОГО
                ws2.cell(row=total_row, column=12, value=_rollup_pay_formula(total_row)).alignment = align_c
                ws2.cell(row=total_row, column=12).number_format = num_fmt
                ws2.cell(row=total_row, column=12).font = font_bb
                ws2.cell(row=total_row, column=13, value=_rollup_perf_formula(total_row)).alignment = align_c
                ws2.cell(row=total_row, column=13).number_format = num_fmt
                ws2.cell(row=total_row, column=13).font = font_bb

            else:
                def _sum_df(col: str) -> int:
                    return int(round(float(df_total_src[col].sum() if df_total_src is not None and not df_total_src.empty and col in df_total_src.columns else 0.0)))

                ws2[f"C{total_row}"] = _sum_df("1710_нач")
                ws2[f"D{total_row}"] = _sum_df("3310_нач")
                ws2[f"E{total_row}"] = _sum_df("нетто_нач")
                ws2[f"F{total_row}"] = _sum_df("оборот_дт")
                ws2[f"G{total_row}"] = _sum_df("оборот_кт")
                ws2[f"H{total_row}"] = _sum_df("1710")
                ws2[f"I{total_row}"] = _sum_df("3310")
                tot_net = _sum_df("сальдо поставщики")
                if df_total_src is not None and not df_total_src.empty:
                    try:
                        if float(df_total_src["сальдо поставщики"].sum()) < 0:
                            tot_net = abs(int(tot_net))
                    except Exception:
                        pass
                ws2[f"J{total_row}"] = tot_net
                for addr in (f"C{total_row}", f"D{total_row}", f"E{total_row}", f"F{total_row}", f"G{total_row}", f"H{total_row}", f"I{total_row}", f"J{total_row}"):
                    ws2[addr].font = font_bb
                    ws2[addr].alignment = align_c
                    ws2[addr].number_format = num_fmt

                for i in range(len(months)):
                    ci = start_col_pay + i
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

                ws2.cell(row=total_row, column=12, value=_rollup_pay_formula(total_row)).alignment = align_c
                ws2.cell(row=total_row, column=12).number_format = num_fmt
                ws2.cell(row=total_row, column=12).font = font_bb

            # прочее = ИТОГО - ТОП-N (для чисел и для помесячных колонок)
            other_cols = ["C", "D", "E", "F", "G", "H", "I", "J", "L"]
            if kind == "cust":
                other_cols.append("M")
            for col_letter in other_cols:
                ws2[f"{col_letter}{other_row}"] = f"={col_letter}{total_row}-{col_letter}{sum_row}"
                ws2[f"{col_letter}{other_row}"].font = font_b
                ws2[f"{col_letter}{other_row}"].alignment = align_c
                ws2[f"{col_letter}{other_row}"].number_format = num_fmt

            for i in range(len(months)):
                ci = start_col_pay + i
                col_letter = get_column_letter(ci)
                ws2.cell(row=other_row, column=ci, value=f"={col_letter}{total_row}-{col_letter}{sum_row}").number_format = num_fmt
                ws2.cell(row=other_row, column=ci).alignment = align_c
            if kind == "cust":
                for i in range(len(months)):
                    ci = start_col_perf + i
                    col_letter = get_column_letter(ci)
                    ws2.cell(row=other_row, column=ci, value=f"={col_letter}{total_row}-{col_letter}{sum_row}").number_format = num_fmt
                    ws2.cell(row=other_row, column=ci).alignment = align_c

        # === Заказчики (фиксированная раскладка + строки ТОП-N/прочее/ИТОГО)
        df_cust_pos_all = df_cust[df_cust["сальдо заказчики"] >= 0] if not df_cust.empty else df_cust
        df_cust_neg_all = df_cust[df_cust["сальдо заказчики"] < 0] if not df_cust.empty else df_cust
        cust_pos = df_cust_pos_all.nlargest(TOP_N, "сальдо заказчики") if df_cust_pos_all is not None and not df_cust_pos_all.empty else df_cust_pos_all
        cust_neg = df_cust_neg_all.nsmallest(TOP_N, "сальдо заказчики") if df_cust_neg_all is not None and not df_cust_neg_all.empty else df_cust_neg_all

        cust_pos_start = 3
        cust_pos_end = cust_pos_start + TOP_N - 1
        cust_pos_sum, cust_pos_other, cust_pos_total = cust_pos_end + 1, cust_pos_end + 2, cust_pos_end + 3
        # Отступ между блоками: 1 пустая строка, затем строка заголовков, затем данные.
        cust_gap = cust_pos_total + 1
        cust_neg_hdr = cust_gap + 1
        cust_neg_start = cust_neg_hdr + 1
        cust_neg_end = cust_neg_start + TOP_N - 1
        cust_neg_sum, cust_neg_other, cust_neg_total = cust_neg_end + 1, cust_neg_end + 2, cust_neg_end + 3
        cust_end_gap = cust_neg_total + 1

        # Для отрицательного блока заказчиков делаем отдельную строку заголовков (чтобы было явно "Нетто КЗ").
        for col in range(2, 14):  # B..M
            src = ws2.cell(row=2, column=col)
            dst = ws2.cell(row=cust_neg_hdr, column=col)
            dst.value = src.value
            dst.font = _copy(src.font)
            dst.alignment = _copy(src.alignment)
        ws2[f"J{cust_neg_hdr}"] = "Нетто КЗ"
        ws2[f"B{cust_neg_hdr}"] = "Контрагент"
        ws2[f"E{cust_neg_hdr}"] = "Нетто КЗ (нач)"
        ws2[f"J{cust_neg_hdr}"] = "Нетто КЗ (кон)"

        for i in range(TOP_N):
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
                int(round(float(r0.get("1210_нач", 0) or 0))),
                int(round(float(r0.get("3510_нач", 0) or 0))),
                int(round(float(r0.get("нетто_нач", 0) or 0))),
                int(round(float(r0.get("оборот_дт", 0) or 0))),
                int(round(float(r0.get("оборот_кт", 0) or 0))),
                int(round(float(r0.get("1210", 0) or 0))),
                int(round(float(r0.get("3510", 0) or 0))),
                int(round(float(r0.get("сальдо заказчики", 0) or 0))),
            )

        _write_summary_rows("cust", df_cust_pos_all, cust_pos_start, cust_pos_end, cust_pos_sum, cust_pos_other, cust_pos_total)

        for i in range(TOP_N):
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
                int(round(float(r0.get("1210_нач", 0) or 0))),
                int(round(float(r0.get("3510_нач", 0) or 0))),
                abs(int(round(float(r0.get("нетто_нач", 0) or 0)))),
                int(round(float(r0.get("оборот_дт", 0) or 0))),
                int(round(float(r0.get("оборот_кт", 0) or 0))),
                int(round(float(r0.get("1210", 0) or 0))),
                int(round(float(r0.get("3510", 0) or 0))),
                abs(int(round(float(r0.get("сальдо заказчики", 0) or 0)))),
            )

        _write_summary_rows("cust", df_cust_neg_all, cust_neg_start, cust_neg_end, cust_neg_sum, cust_neg_other, cust_neg_total)

        # === Поставщики (ниже блока заказчиков)
        supp_header_row = cust_end_gap + 2  # пустая строка после блока заказчиков
        # Визуальный разделитель
        for col in range(2, last_used_col + 1):  # B..(конец таблицы)
            c = ws2.cell(row=supp_header_row - 1, column=col)
            c.fill = sec_fill
            c.font = sec_font
        ws2.cell(row=supp_header_row - 1, column=2, value="ПОСТАВЩИКИ").alignment = align_l

        ws2[f"B{supp_header_row}"] = "Контрагент"
        ws2[f"C{supp_header_row}"] = "1710 нач"
        ws2[f"D{supp_header_row}"] = "3310 нач"
        ws2[f"E{supp_header_row}"] = "Нетто ДЗ (нач)"
        ws2[f"F{supp_header_row}"] = "оборот ДТ"
        ws2[f"G{supp_header_row}"] = "оборот КТ"
        ws2[f"H{supp_header_row}"] = "1710"
        ws2[f"I{supp_header_row}"] = "3310"
        # Для поставщиков в шапке по умолчанию пишем "Нетто ДЗ (кон)" (для положительного блока),
        # а для отрицательного блока сделаем отдельную шапку ниже.
        ws2[f"J{supp_header_row}"] = "Нетто ДЗ (кон)"
        ws2[f"K{supp_header_row}"] = "коммент"
        ws2[f"L{supp_header_row}"] = '="Опл L"&$A$2&"M"'
        ws2[f"M{supp_header_row}"] = ""  # выполнение не заполняем для поставщиков
        for caddr in (f"B{supp_header_row}", f"C{supp_header_row}", f"D{supp_header_row}", f"E{supp_header_row}", f"F{supp_header_row}", f"G{supp_header_row}", f"H{supp_header_row}", f"I{supp_header_row}", f"J{supp_header_row}", f"K{supp_header_row}", f"L{supp_header_row}", f"M{supp_header_row}"):
            ws2[caddr].font = font_h
            ws2[caddr].alignment = align_l if caddr.startswith("B") else align_c

        df_supp_pos_all = df_supp[df_supp["сальдо поставщики"] >= 0] if not df_supp.empty else df_supp
        df_supp_neg_all = df_supp[df_supp["сальдо поставщики"] < 0] if not df_supp.empty else df_supp
        supp_pos = df_supp_pos_all.nlargest(TOP_N, "сальдо поставщики") if df_supp_pos_all is not None and not df_supp_pos_all.empty else df_supp_pos_all
        supp_neg = df_supp_neg_all.nsmallest(TOP_N, "сальдо поставщики") if df_supp_neg_all is not None and not df_supp_neg_all.empty else df_supp_neg_all

        supp_pos_start = supp_header_row + 1
        supp_pos_end = supp_pos_start + TOP_N - 1
        supp_pos_sum, supp_pos_other, supp_pos_total = supp_pos_end + 1, supp_pos_end + 2, supp_pos_end + 3
        # Для отрицательного блока поставщиков тоже добавим отдельную строку заголовков.
        supp_neg_hdr = supp_pos_total + 2
        # Копируем заголовки поставщиков и меняем Нетто на "Нетто КЗ".
        for col in range(2, 14):  # B..M
            src = ws2.cell(row=supp_header_row, column=col)
            dst = ws2.cell(row=supp_neg_hdr, column=col)
            dst.value = src.value
            dst.font = _copy(src.font)
            dst.alignment = _copy(src.alignment)
        ws2[f"E{supp_neg_hdr}"] = "Нетто КЗ (нач)"
        ws2[f"J{supp_neg_hdr}"] = "Нетто КЗ (кон)"
        ws2[f"B{supp_neg_hdr}"] = "Контрагент"

        supp_neg_start = supp_neg_hdr + 1
        supp_neg_end = supp_neg_start + TOP_N - 1
        supp_neg_sum, supp_neg_other, supp_neg_total = supp_neg_end + 1, supp_neg_end + 2, supp_neg_end + 3

        for i in range(TOP_N):
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
                int(round(float(r0.get("1710_нач", 0) or 0))),
                int(round(float(r0.get("3310_нач", 0) or 0))),
                int(round(float(r0.get("нетто_нач", 0) or 0))),
                int(round(float(r0.get("оборот_дт", 0) or 0))),
                int(round(float(r0.get("оборот_кт", 0) or 0))),
                int(round(float(r0.get("1710", 0) or 0))),
                int(round(float(r0.get("3310", 0) or 0))),
                int(round(float(r0.get("сальдо поставщики", 0) or 0))),
            )
        _write_summary_rows("supp", df_supp_pos_all, supp_pos_start, supp_pos_end, supp_pos_sum, supp_pos_other, supp_pos_total)

        for i in range(TOP_N):
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
                int(round(float(r0.get("1710_нач", 0) or 0))),
                int(round(float(r0.get("3310_нач", 0) or 0))),
                abs(int(round(float(r0.get("нетто_нач", 0) or 0)))),
                int(round(float(r0.get("оборот_дт", 0) or 0))),
                int(round(float(r0.get("оборот_кт", 0) or 0))),
                int(round(float(r0.get("1710", 0) or 0))),
                int(round(float(r0.get("3310", 0) or 0))),
                abs(int(round(float(r0.get("сальдо поставщики", 0) or 0)))),
            )
        _write_summary_rows("supp", df_supp_neg_all, supp_neg_start, supp_neg_end, supp_neg_sum, supp_neg_other, supp_neg_total)

        # === Общее сальдо (ниже поставщиков): только значение сальдо + строки ТОП-15/прочее/ИТОГО
        total_header_row = supp_neg_total + 5  # пустая строка после блока поставщиков
        ws2[f"B{total_header_row}"] = "Контрагент"
        ws2[f"E{total_header_row}"] = "общее сальдо"
        ws2[f"B{total_header_row}"].font = font_h
        ws2[f"E{total_header_row}"].font = font_h
        ws2[f"B{total_header_row}"].alignment = align_l
        ws2[f"E{total_header_row}"].alignment = align_c

        df_total_pos_all = df_total[df_total["общее сальдо"] > 0] if not df_total.empty else df_total
        df_total_neg_all = df_total[df_total["общее сальдо"] < 0] if not df_total.empty else df_total
        total_pos = df_total_pos_all.nlargest(TOP_N, "общее сальдо") if df_total_pos_all is not None and not df_total_pos_all.empty else df_total_pos_all
        total_neg = df_total_neg_all.nsmallest(TOP_N, "общее сальдо") if df_total_neg_all is not None and not df_total_neg_all.empty else df_total_neg_all

        total_pos_start = total_header_row + 1
        total_pos_end = total_pos_start + TOP_N - 1
        total_pos_sum, total_pos_other, total_pos_total = total_pos_end + 1, total_pos_end + 2, total_pos_end + 3
        total_neg_start = total_pos_total + 2
        total_neg_end = total_neg_start + TOP_N - 1
        total_neg_sum, total_neg_other, total_neg_total = total_neg_end + 1, total_neg_end + 2, total_neg_end + 3

        def _write_total_row(r: int, contr: str, vtot: int) -> None:
            ws2.cell(row=r, column=2, value=contr).font = font_b
            ws2.cell(row=r, column=2).alignment = align_l
            # Для строк с контрагентами (топ-15) числа НЕ делаем жирными.
            ws2.cell(row=r, column=5, value=vtot).font = font_b
            ws2.cell(row=r, column=5).alignment = align_c
            ws2.cell(row=r, column=5).number_format = num_fmt

        for i in range(TOP_N):
            if total_pos is None or i >= len(total_pos.index):
                continue
            rr = total_pos_start + i
            r0 = total_pos.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_total_row(rr, contr, int(round(float(r0.get("общее сальдо", 0) or 0))))

        # Сводные строки (положительный блок)
        ws2.cell(row=total_pos_sum, column=2, value=f"ТОП-{TOP_N}").font = font_h
        ws2.cell(row=total_pos_other, column=2, value="прочее").font = font_h
        ws2.cell(row=total_pos_total, column=2, value="ИТОГО").font = font_h
        for rr in (total_pos_sum, total_pos_other, total_pos_total):
            ws2.cell(row=rr, column=2).alignment = align_l

        ws2[f"E{total_pos_sum}"] = f"=SUM(E{total_pos_start}:E{total_pos_end})"
        ws2[f"E{total_pos_sum}"].font = font_bb
        ws2[f"E{total_pos_sum}"].alignment = align_c
        ws2[f"E{total_pos_sum}"].number_format = num_fmt

        total_all = int(round(float(df_total_pos_all["общее сальдо"].sum() if df_total_pos_all is not None and not df_total_pos_all.empty else 0.0)))
        ws2[f"E{total_pos_total}"] = total_all
        ws2[f"E{total_pos_total}"].font = font_bb
        ws2[f"E{total_pos_total}"].alignment = align_c
        ws2[f"E{total_pos_total}"].number_format = num_fmt

        ws2[f"E{total_pos_other}"] = f"=E{total_pos_total}-E{total_pos_sum}"
        ws2[f"E{total_pos_other}"].font = font_b
        ws2[f"E{total_pos_other}"].alignment = align_c
        ws2[f"E{total_pos_other}"].number_format = num_fmt

        for i in range(TOP_N):
            if total_neg is None or i >= len(total_neg.index):
                continue
            rr = total_neg_start + i
            r0 = total_neg.iloc[i]
            contr = str(r0.get("Контрагент", "")).strip()
            if not contr:
                continue
            _write_total_row(rr, contr, int(round(float(r0.get("общее сальдо", 0) or 0))))

        # Сводные строки (отрицательный блок)
        ws2.cell(row=total_neg_sum, column=2, value=f"ТОП-{TOP_N}").font = font_h
        ws2.cell(row=total_neg_other, column=2, value="прочее").font = font_h
        ws2.cell(row=total_neg_total, column=2, value="ИТОГО").font = font_h
        for rr in (total_neg_sum, total_neg_other, total_neg_total):
            ws2.cell(row=rr, column=2).alignment = align_l

        ws2[f"E{total_neg_sum}"] = f"=SUM(E{total_neg_start}:E{total_neg_end})"
        ws2[f"E{total_neg_sum}"].font = font_bb
        ws2[f"E{total_neg_sum}"].alignment = align_c
        ws2[f"E{total_neg_sum}"].number_format = num_fmt

        total_all_neg = int(round(float(df_total_neg_all["общее сальдо"].sum() if df_total_neg_all is not None and not df_total_neg_all.empty else 0.0)))
        ws2[f"E{total_neg_total}"] = total_all_neg
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
        for col in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]:
            ws2.column_dimensions[col].width = 10.14

        # Пунктирные границы серого цвета вокруг блоков (для читабельности)
        dotted_side = Side(border_style="dotted", color="A0A0A0")
        dotted_border = Border(left=dotted_side, right=dotted_side, top=dotted_side, bottom=dotted_side)

        def _dotted_box(r1: int, r2: int, c1: int, c2: int) -> None:
            r1 = max(1, int(r1))
            r2 = max(1, int(r2))
            c1 = max(1, int(c1))
            c2 = max(1, int(c2))
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    ws2.cell(row=rr, column=cc).border = dotted_border

        last_used_col = start_col_perf + len(months) - 1
        # Заказчики: + и - блоки (включая строки ТОП/прочее/ИТОГО)
        _dotted_box(cust_pos_start, cust_pos_total, 2, last_used_col)
        _dotted_box(cust_neg_start, cust_neg_total, 2, last_used_col)
        # Поставщики
        _dotted_box(supp_header_row - 1, supp_pos_total, 2, last_used_col)
        _dotted_box(supp_neg_start, supp_neg_total, 2, last_used_col)

        # Чёрные разделители колонок: после "оборот КТ" и вокруг "Нетто" (для всех 4 таблиц)
        black = Side(border_style="thin", color="000000")

        def _apply_vline(row_from: int, row_to: int, col: int, which: str) -> None:
            for rr in range(int(row_from), int(row_to) + 1):
                cell = ws2.cell(row=rr, column=int(col))
                b = cell.border or Border()
                if which == "right":
                    cell.border = Border(left=b.left, right=black, top=b.top, bottom=b.bottom)
                else:
                    cell.border = Border(left=black, right=b.right, top=b.top, bottom=b.bottom)

        # Столбцы: G (оборот КТ) -> правый разделитель; J (Нетто) -> левый и правый.
        col_after_turn_kt = 7   # G
        col_net = 10            # J
        for r1, r2 in [
            (cust_pos_start - 1, cust_pos_total),
            (cust_neg_hdr, cust_neg_total),
            (supp_header_row - 1, supp_pos_total),
            (supp_neg_hdr, supp_neg_total),
        ]:
            _apply_vline(r1, r2, col_after_turn_kt, "right")
            _apply_vline(r1, r2, col_net, "left")
            _apply_vline(r1, r2, col_net, "right")

        def _apply_hline(col_from: int, col_to: int, row: int, which: str) -> None:
            for cc in range(int(col_from), int(col_to) + 1):
                cell = ws2.cell(row=int(row), column=int(cc))
                b = cell.border or Border()
                if which == "bottom":
                    cell.border = Border(left=b.left, right=b.right, top=b.top, bottom=black)
                else:
                    cell.border = Border(left=b.left, right=b.right, top=black, bottom=b.bottom)

        # Горизонтальные чёрные разделители между Нетто ДЗ и Нетто КЗ
        _apply_hline(2, last_used_col, cust_pos_total, "bottom")
        _apply_hline(2, last_used_col, cust_neg_hdr, "top")
        _apply_hline(2, last_used_col, supp_pos_total, "bottom")
        _apply_hline(2, last_used_col, supp_neg_hdr, "top")

    out = io.BytesIO()
    wb.save(out)

    # Топ-5 заказчиков по дебиторке (для проверки через egov)
    top5: list = []
    try:
        if not df_cust.empty and "сальдо заказчики" in df_cust.columns:
            top5 = (
                df_cust[df_cust["сальдо заказчики"] > 0]
                .nlargest(5, "сальдо заказчики")["Контрагент"]
                .tolist()
            )
    except Exception:
        pass

    return out.getvalue(), top5


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

        def _as_int_year(v) -> Optional[int]:
            if v is None:
                return None
            if isinstance(v, (int, float)) and float(v).is_integer():
                y = int(v)
                return y if 1900 <= y <= 2500 else None
            s = str(v).strip()
            m = re.search(r"(\d{4})", s)
            if not m:
                return None
            try:
                y = int(m.group(1))
                return y if 1900 <= y <= 2500 else None
            except Exception:
                return None

        def _scan_month_headers(sheet) -> List[Tuple[int, str]]:
            """Ищет заголовки вида YYYY_MM в 1-й строке и возвращает [(col_idx, label), ...] по порядку колонок."""
            out: List[Tuple[int, str]] = []
            max_col = int(sheet.max_column or 1)
            for c in range(1, max_col + 1):
                v = sheet.cell(row=1, column=c).value
                if v is None:
                    continue
                s = str(v).strip()
                if re.fullmatch(r"\d{4}_\d{2}", s):
                    out.append((c, s))
            return out

        years_pay = [_as_int_year(wd_ws["C1"].value), _as_int_year(wd_ws["D1"].value), _as_int_year(wd_ws["E1"].value)]
        years_perf = [_as_int_year(md_ws["C1"].value), _as_int_year(md_ws["D1"].value), _as_int_year(md_ws["E1"].value)]
        years_pay = [y for y in years_pay if y is not None]
        years_perf = [y for y in years_perf if y is not None]

        # По текущей структуре "Договоры" ожидаем 3 годовых колонки (C/D/E). Если что-то не так — не ломаемся,
        # просто подставляем "пустые" годы для сохранения структуры.
        while len(years_pay) < 3:
            years_pay.append(0)
        while len(years_perf) < 3:
            years_perf.append(0)
        years_pay = years_pay[:3]
        years_perf = years_perf[:3]

        payments_year = defaultdict(lambda: [0.0, 0.0, 0.0])
        performance_year = defaultdict(lambda: [0.0, 0.0, 0.0])

        pay_month_cols = _scan_month_headers(wd_ws)
        perf_month_cols = _scan_month_headers(md_ws)
        payments_monthly = defaultdict(lambda: [0.0] * len(pay_month_cols))
        performance_monthly = defaultdict(lambda: [0.0] * len(perf_month_cols))

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

        def collect_monthly(sheet, target_dict, month_cols: List[Tuple[int, str]]):
            for row in range(2, sheet.max_row + 1):
                n = sheet[f"A{row}"].value
                c = sheet[f"B{row}"].value
                if not n and not c:
                    continue
                key = (str(n).strip() if n else "", str(c).strip() if c else "")
                for i, (col_idx, _label) in enumerate(month_cols):
                    v = sheet.cell(row=row, column=col_idx).value
                    if v is None:
                        continue
                    try:
                        target_dict[key][i] += float(v)
                    except:
                        pass

        collect_yearly(wd_ws, payments_year)
        collect_yearly(md_ws, performance_year)
        collect_monthly(wd_ws, payments_monthly, pay_month_cols)
        collect_monthly(md_ws, performance_monthly, perf_month_cols)

        all_keys = sorted(
            set(payments_year.keys())
            | set(performance_year.keys())
            | set(payments_monthly.keys())
            | set(performance_monthly.keys()),
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
        ws["C2"] = years_pay[0] or ""
        ws["D2"] = years_pay[1] or ""
        ws["E2"] = years_pay[2] or ""
        ws["F2"] = "Total"

        ws["G1"] = "выполнения с ндс"
        ws["G2"] = years_perf[0] or ""
        ws["H2"] = years_perf[1] or ""
        ws["I2"] = years_perf[2] or ""
        ws["J2"] = "Total"

        ws["K2"] = "дз/(аванс)"

        ws["M1"] = "оплата"
        start_col_pay = column_index_from_string("M")
        pay_month_labels = [lbl for _c, lbl in pay_month_cols]
        for i, label in enumerate(pay_month_labels):
            ws[f"{get_column_letter(start_col_pay + i)}2"] = label

        start_col_perf = start_col_pay + max(1, len(pay_month_labels))
        ws[f"{get_column_letter(start_col_perf)}1"] = "выполнения с ндс"
        perf_month_labels = [lbl for _c, lbl in perf_month_cols]
        for i, label in enumerate(perf_month_labels):
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

            def _vat_coef_for_year(y: int) -> float:
                # НДС 12% до 2025 включительно, 16% начиная с 2026 (и для лет > 2026 тоже 16%).
                return 1.16 if int(y or 0) >= 2026 else 1.12

            ws[f"G{row}"] = pf[0] * _vat_coef_for_year(years_perf[0])
            ws[f"H{row}"] = pf[1] * _vat_coef_for_year(years_perf[1])
            ws[f"I{row}"] = pf[2] * _vat_coef_for_year(years_perf[2])

            ws[f"F{row}"] = f"=SUM(C{row}:E{row})"
            ws[f"J{row}"] = f"=SUM(G{row}:I{row})"
            ws[f"K{row}"] = f"=J{row}-F{row}"

            mp = payments_monthly.get(key, [0.0] * len(pay_month_labels))
            for i in range(len(pay_month_labels)):
                ws[f"{get_column_letter(start_col_pay + i)}{row}"] = mp[i]

            mf = performance_monthly.get(key, [0.0] * len(perf_month_labels))
            for i, label in enumerate(perf_month_labels):
                y = _as_int_year(label.split("_", 1)[0]) or 0
                ws[f"{get_column_letter(start_col_perf + i)}{row}"] = mf[i] * _vat_coef_for_year(y)

            row += 1

        last_row = row - 1 if row > start_row else 2

        regular = Font(name="Arial", size=10)
        bold = Font(name="Arial", size=10, bold=True)

        last_col = max(start_col_perf + len(perf_month_labels) - 1, column_index_from_string("K"))
        for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=last_col):
            for c in row:
                c.font = regular

        for col in range(1, last_col + 1):
            ws[f"{get_column_letter(col)}2"].font = bold
        for addr in ["A1", "C1", "G1", "M1", f"{get_column_letter(start_col_perf)}1", "K2"]:
            ws[addr].font = bold

        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")
        for col in range(1, last_col + 1):
            addr = f"{get_column_letter(col)}2"
            ws[addr].alignment = left if addr in ["A2", "B2"] else center

        num_format = "#,##0;[Red](#,##0)"
        numeric_cols = list("CDEFGHIJK")
        numeric_cols += [get_column_letter(c) for c in range(start_col_pay, start_col_pay + len(pay_month_labels))]
        numeric_cols += [get_column_letter(c) for c in range(start_col_perf, start_col_perf + len(perf_month_labels))]

        for col in numeric_cols:
            for r in range(3, last_row + 1):
                cell = ws[f"{col}{r}"]
                cell.alignment = center
                cell.number_format = num_format

        for r in range(1, last_row + 1):
            ws[f"A{r}"].alignment = left
            ws[f"B{r}"].alignment = left

        for addr in ["C1", "G1", "M1"]:
            try:
                ws[addr].alignment = left
            except Exception:
                pass
        try:
            ws[f"{get_column_letter(start_col_perf)}1"].alignment = left
        except Exception:
            pass

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 38
        for col in numeric_cols:
            ws.column_dimensions[col].width = 12.2 if column_index_from_string(col) >= start_col_pay else 12.6

        thin = Side(border_style="thin", color="000000")
        border_cols = ["C", "G", "K", "M", get_column_letter(start_col_perf)]
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
# Для каждого выбранного счета 13**:
# - находим лист(ы), где в названии встречается номер счета
# - в листе ищем строку в колонке A, где ячейка равна номеру счета
# - ставим I1=400 и суммарную "шапку" корзин в K6:O7 (как в эталонном файле)
# - в колонке I проставляем формулу-классификацию по строкам до первой пустой в колонке G
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

    def _apply_inventory_bucket_layout(ws) -> None:
        """
        Новая методика ликвидности запасов (как в эталонном файле):
        - I1 = 400
        - K6:O6 — корзины
        - K7:O7 — суммы по корзинам через SUMIFS (G по значениям I)
        """
        ws["I1"].value = 400
        ws["I1"].number_format = "#,##0"
        ws["I1"].font = Font(name="Arial", size=8, bold=True, color="FF4B5563")  # тёмно-серый
        ws["I1"].fill = PatternFill(fill_type="solid", fgColor="FFD9D9D9")  # светло-серый фон
        ws["I1"].alignment = Alignment(horizontal="center")

        labels = [
            "0-1 мес",
            "1-3 мес",
            "3-6 мес",
            "6-12 мес",
            "ДТ BoP = 0",
            "более 12 мес",
            "Бесконечность (ост нач)",
            "Бесконечность (обор ДТ)",
        ]
        cols = ["K", "L", "M", "N", "O", "P", "Q", "R"]

        fill_hdr = PatternFill(fill_type="solid", fgColor="E4F0DD")
        font_hdr = Font(name="Arial", size=9, bold=True, color="003F2F")
        align_hdr = Alignment(horizontal="center", vertical="top", wrap_text=True)
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col, label in zip(cols, labels):
            c = ws[f"{col}6"]
            c.value = label
            c.font = font_hdr
            c.fill = fill_hdr
            c.alignment = align_hdr
            c.border = border
            c.number_format = "#,##0.00"

        font_sum = Font(name="Arial", size=9, bold=True)
        align_sum = Alignment(horizontal="center")
        for col in cols:
            c = ws[f"{col}7"]
            c.value = f"=SUMIFS($G:$G,$I:$I,{col}6)"
            c.font = font_sum
            c.alignment = align_sum
            c.number_format = "#,##0"

    def _bucket_formula_for_row(r: int) -> str:
        return (
            f'=IF(G{r}=0,\"Нет остатка\",\n'
            f' IF(G{r}<0,\"Отрицательное\",\n'
            f' IF(C{r}=0,\"ДТ BoP = 0\",\n'
            f' IF(AND(F{r}=0,C{r}<>0,E{r}=0,G{r}=C{r}),\"Бесконечность (ост нач)\",\n'
            f' IF(AND(F{r}=0,C{r}<>0,E{r}<>0),\"Бесконечность (обор ДТ)\",\n'
            f' _xlfn.LET(_xlpm.turn,$I$1*(G{r}/F{r}),\n'
            f'     IF(_xlpm.turn<=31,\"0-1 мес\",\n'
            f'        IF(_xlpm.turn<=92,\"1-3 мес\",\n'
            f'           IF(_xlpm.turn<=183,\"3-6 мес\",\n'
            f'              IF(_xlpm.turn<=366,\"6-12 мес\",\"более 12 мес\"))))))))))'
        )

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

            _apply_inventory_bucket_layout(ws)

            # Заполняем формулы по строкам запасов: от строки ниже маркера до первой пустой в колонке G.
            start_row = found_row + 1
            last = start_row - 1
            for r in range(start_row, ws.max_row + 1):
                # Останавливаемся перед строкой "Итого" (чтобы не тянуть хвосты/сводные строки в запасы).
                a = ws.cell(row=r, column=1).value
                if a is not None and "итого" in str(a).strip().lower():
                    break
                g = ws.cell(row=r, column=7).value
                if _is_blank(g):
                    break
                last = r

            if last >= start_row:
                for rr in range(start_row, last + 1):
                    c = ws.cell(row=rr, column=9)  # I
                    c.value = _bucket_formula_for_row(rr)
                    c.font = Font(name="Arial", size=8)
                processed.append(f"{account}: {sh} (строки {start_row}-{last})")
            else:
                processed.append(f"{account}: {sh} (нет строк с данными в G ниже маркера)")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), {"processed": processed, "missing_sheets": missing_sheets, "missing_markers": missing_markers}


# =========================
# CODE 4 (Чистая прибыль по "общ*")
# Рассчитывает ЧП из движения денег (1030), исключая CFF/CFI/NWC/Амортизацию.
# Для каждого выбранного листа "общ*" добавляет справа-внизу таблицу (формулами INDEX/MATCH).
# =========================
def run_code_4_net_profit(file_bytes: bytes, obsh_sheetnames: List[str]) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))

    if not obsh_sheetnames:
        raise ValueError("Чистая прибыль: не выбраны листы «общ*».")

    for sh in obsh_sheetnames:
        if sh not in wb.sheetnames:
            raise ValueError(f"Чистая прибыль: лист не найден: {sh}")

    def _is_blank(v) -> bool:
        if v is None:
            return True
        if isinstance(v, str) and not v.strip():
            return True
        return False

    def _excel_sheet_ref(title: str) -> str:
        # Excel-референс листа в формулах: 'Лист 1'!A1. Одинарные кавычки внутри имени удваиваем.
        t = str(title).replace("'", "''")
        return f"'{t}'"

    def _find_last_used_col_in_header(ws) -> int:
        maxc = 1
        top = min(int(ws.max_row or 1), 12)
        for r in range(1, top + 1):
            for c in range(1, int(ws.max_column or 1) + 1):
                v = ws.cell(row=r, column=c).value
                if not _is_blank(v):
                    maxc = max(maxc, c)
        return maxc

    def _find_last_used_row_in_col_a(ws) -> int:
        max_scan = min(int(ws.max_row or 1), 10000)
        last = 1
        for r in range(1, max_scan + 1):
            if not _is_blank(ws.cell(row=r, column=1).value):
                last = r
        return last

    def _find_debit_credit_pairs(ws) -> Tuple[int, int, int, int, int, int]:
        """
        Возвращает (bop_deb, bop_cred, turn_deb, turn_cred, eop_deb, eop_cred) — номера колонок.
        Ищем строку заголовков, где встречаются "Дебет"/"Кредит" 3 раза подряд (6 колонок).
        """
        # Ищем по верхним строкам (обычно 4..8).
        best_row = None
        for r in range(1, 20):
            deb_cols = []
            cred_cols = []
            for c in range(1, int(ws.max_column or 1) + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                s = str(v).strip().lower()
                if "дебет" == s or s.startswith("дебет"):
                    deb_cols.append(c)
                if "кредит" == s or s.startswith("кредит"):
                    cred_cols.append(c)
            # Кандидат: минимум 3 дебета и 3 кредита, и они идут вперемешку по возрастанию.
            if len(deb_cols) >= 3 and len(cred_cols) >= 3:
                best_row = r
                break

        if best_row is None:
            raise ValueError(f"Чистая прибыль ({ws.title}): не нашёл заголовки «Дебет/Кредит» (3 пары) в верхней части листа.")

        labels = []
        for c in range(1, int(ws.max_column or 1) + 1):
            v = ws.cell(row=best_row, column=c).value
            if v is None:
                continue
            s = str(v).strip().lower()
            if s.startswith("дебет"):
                labels.append((c, "D"))
            elif s.startswith("кредит"):
                labels.append((c, "C"))
        labels.sort(key=lambda x: x[0])

        # Берём первые 6 в формате D,C,D,C,D,C
        seq = []
        for c, t in labels:
            if not seq:
                if t == "D":
                    seq.append((c, t))
            else:
                want = "C" if seq[-1][1] == "D" else "D"
                if t == want:
                    seq.append((c, t))
            if len(seq) == 6:
                break

        if len(seq) < 6 or [t for _c, t in seq] != ["D", "C", "D", "C", "D", "C"]:
            raise ValueError(f"Чистая прибыль ({ws.title}): не смог определить 3 пары «Дебет/Кредит» из заголовков.")

        bop_deb, bop_cred, turn_deb, turn_cred, eop_deb, eop_cred = [c for c, _t in seq]
        return bop_deb, bop_cred, turn_deb, turn_cred, eop_deb, eop_cred

    def _find_account_row(ws, acc: str) -> Optional[int]:
        max_scan = min(int(ws.max_row or 1), 10000)
        for r in range(3, max_scan + 1):
            v = ws.cell(row=r, column=1).value
            if v is None:
                continue
            if acc in str(v):
                return r
        return None

    def _name_or_acc(ws, row: Optional[int], acc: str) -> str:
        if row is None:
            return acc
        v = ws.cell(row=row, column=1).value
        return str(v).strip() if v is not None else acc

    def _delta_formula(sheet_ref: str, acc_row: int, cols: Tuple[int, int, int, int, int, int]) -> str:
        bop_d, bop_c, _t_d, _t_c, eop_d, eop_c = cols
        # Δ = (Кт_кон - Кт_нач) - (Дт_кон - Дт_нач)
        bd = f"INDEX({sheet_ref}!$1:$1048576,{acc_row},{bop_d})"
        bc = f"INDEX({sheet_ref}!$1:$1048576,{acc_row},{bop_c})"
        ed = f"INDEX({sheet_ref}!$1:$1048576,{acc_row},{eop_d})"
        ec = f"INDEX({sheet_ref}!$1:$1048576,{acc_row},{eop_c})"
        return f"=(({ec})-({bc}))-(({ed})-({bd}))"

    thin_black = Side(style="thin", color="000000")
    dotted_gray = Side(style="dotted", color="A0A0A0")
    header_fill = PatternFill("solid", fgColor="EDEDED")
    font_h = Font(name="Arial", size=9, bold=True)
    font_b = Font(name="Arial", size=9)
    num_fmt = "#,##0; -#,##0"

    for ws_name in obsh_sheetnames:
        ws = wb[ws_name]

        cols = _find_debit_credit_pairs(ws)
        last_row = _find_last_used_row_in_col_a(ws)
        start_row = last_row + 3
        # Таблицу ставим на 2 столбца правее ПОСЛЕДНЕГО столбца "Сальдо на конец (Кредит)"
        # (то есть правее третьей пары "Дебет/Кредит").
        start_col = int(cols[5]) + 2
        sheet_ref = _excel_sheet_ref(ws.title)

        # Заголовки таблицы
        headers = ["Счет", "Полный", "ВГО", "Без ВГО"]
        for j, h in enumerate(headers):
            cell = ws.cell(row=start_row, column=start_col + j, value=h)
            cell.font = font_h
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=dotted_gray, right=dotted_gray, top=dotted_gray, bottom=dotted_gray)

        # Строки (в таком порядке, как в ТЗ)
        # Для каждой "строки счета": если счет найден — пишем полное название, иначе — 4 цифры.
        # Для "итогов": CF/CFF/CFI/NWC/ЧП — считаем формулой по значениям выше.
        row = start_row + 1

        def _write_line(label: str, full_text: str, full_formula: Optional[str]) -> int:
            nonlocal row
            ws.cell(row=row, column=start_col + 0, value=full_text).font = font_b

            c_full = ws.cell(row=row, column=start_col + 1, value=(full_formula or 0))
            c_full.font = font_b
            c_full.number_format = num_fmt

            ws.cell(row=row, column=start_col + 2, value="").font = font_b  # ВГО пустой

            c_no = ws.cell(row=row, column=start_col + 3, value=f"=({get_column_letter(start_col+1)}{row})-({get_column_letter(start_col+2)}{row})")
            c_no.font = font_b
            c_no.number_format = num_fmt

            # Пунктирные границы вокруг строки
            for j in range(4):
                ws.cell(row=row, column=start_col + j).border = Border(left=dotted_gray, right=dotted_gray, top=dotted_gray, bottom=dotted_gray)
            cur = row
            row += 1
            return cur

        def _write_sum_line(text: str, ref_rows: List[int]) -> int:
            nonlocal row
            ws.cell(row=row, column=start_col + 0, value=text).font = font_h
            # Полный
            letters = [f"{get_column_letter(start_col+1)}{r}" for r in ref_rows]
            ws.cell(row=row, column=start_col + 1, value=("=SUM(" + ",".join(letters) + ")") if letters else 0).number_format = num_fmt
            ws.cell(row=row, column=start_col + 1).font = font_h
            # ВГО пустой
            ws.cell(row=row, column=start_col + 2, value="").font = font_h
            # Без ВГО
            letters2 = [f"{get_column_letter(start_col+3)}{r}" for r in ref_rows]
            ws.cell(row=row, column=start_col + 3, value=("=SUM(" + ",".join(letters2) + ")") if letters2 else 0).number_format = num_fmt
            ws.cell(row=row, column=start_col + 3).font = font_h
            for j in range(4):
                ws.cell(row=row, column=start_col + j).border = Border(left=dotted_gray, right=dotted_gray, top=dotted_gray, bottom=dotted_gray)
            cur = row
            row += 1
            return cur

        def _top_border(rr: int) -> None:
            for j in range(4):
                cell = ws.cell(row=rr, column=start_col + j)
                cell.border = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=thin_black,
                    bottom=cell.border.bottom,
                )

        # 1030
        r1030 = _find_account_row(ws, "1030")
        if r1030:
            base = _delta_formula(sheet_ref, r1030, cols)
            neg = f"=-({base[1:]})" if isinstance(base, str) and base.startswith("=") else f"=-({base})"
        else:
            neg = None
        r_1030_line = _write_line("1030", _name_or_acc(ws, r1030, "1030"), neg)
        r_cf = _write_sum_line("CF", [r_1030_line])
        _top_border(r_cf)

        row += 1  # пропуск строки

        # CFF: 3010 + 3030 + 4010
        r3010 = _find_account_row(ws, "3010")
        r3030 = _find_account_row(ws, "3030")
        r4010 = _find_account_row(ws, "4010")
        r_3010_line = _write_line("3010", _name_or_acc(ws, r3010, "3010"), _delta_formula(sheet_ref, r3010, cols) if r3010 else None)
        r_3030_line = _write_line("3030", _name_or_acc(ws, r3030, "3030"), _delta_formula(sheet_ref, r3030, cols) if r3030 else None)
        r_4010_line = _write_line("4010", _name_or_acc(ws, r4010, "4010"), _delta_formula(sheet_ref, r4010, cols) if r4010 else None)
        r_cff = _write_sum_line("CFF", [r_3010_line, r_3030_line, r_4010_line])
        _top_border(r_cff)

        row += 1  # пропуск строки

        # CFI: 2410
        r2410 = _find_account_row(ws, "2410")
        r_2410_line = _write_line("2410", _name_or_acc(ws, r2410, "2410"), _delta_formula(sheet_ref, r2410, cols) if r2410 else None)
        r_cfi = _write_sum_line("CFI", [r_2410_line])
        _top_border(r_cfi)

        row += 1

        # NWC: 1210,1310,1320,1330,1400,1710,1711,3100,3310,3311,3320,3350,3510
        nwc_rows = []
        for acc in ["1210", "1310", "1320", "1330", "1400", "1710", "1711", "3100", "3310", "3311", "3320", "3350", "3510"]:
            rr = _find_account_row(ws, acc)
            nwc_rows.append(_write_line(acc, _name_or_acc(ws, rr, acc), _delta_formula(sheet_ref, rr, cols) if rr else None))
        r_nwc = _write_sum_line("NWC", nwc_rows)
        _top_border(r_nwc)

        row += 1

        # Амортизация: 2420
        r2420 = _find_account_row(ws, "2420")
        r_am_line = _write_line("2420", _name_or_acc(ws, r2420, "2420"), _delta_formula(sheet_ref, r2420, cols) if r2420 else None)

        row += 1

        # ЧП = CF - CFF - CFI - NWC - Аморт
        # Делается формулами по строкам итогов.
        ws.cell(row=row, column=start_col + 0, value="ЧП").font = font_h
        ws.cell(
            row=row,
            column=start_col + 1,
            value=f"=({get_column_letter(start_col+1)}{r_cf})-({get_column_letter(start_col+1)}{r_cff})-({get_column_letter(start_col+1)}{r_cfi})-({get_column_letter(start_col+1)}{r_nwc})-({get_column_letter(start_col+1)}{r_am_line})",
        ).font = font_h
        ws.cell(row=row, column=start_col + 1).number_format = num_fmt
        ws.cell(row=row, column=start_col + 2, value="").font = font_h
        ws.cell(
            row=row,
            column=start_col + 3,
            value=f"=({get_column_letter(start_col+3)}{r_cf})-({get_column_letter(start_col+3)}{r_cff})-({get_column_letter(start_col+3)}{r_cfi})-({get_column_letter(start_col+3)}{r_nwc})-({get_column_letter(start_col+3)}{r_am_line})",
        ).font = font_h
        ws.cell(row=row, column=start_col + 3).number_format = num_fmt
        for j in range(4):
            ws.cell(row=row, column=start_col + j).border = Border(left=dotted_gray, right=dotted_gray, top=thin_black, bottom=dotted_gray)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()




# =========================
# CODE 5 (Инсайты)
# Генерирует лист(ы) «инсайты» на основе W/M/Wt/Mt (и общ/кред при наличии).
# Реализовано в отдельном модуле `insights.py`.
# =========================

def run_code_5_insights(file_bytes: bytes) -> bytes:
    """Генерация листа(ов) «инсайты» (см. модуль `insights.py`)."""
    from insights import generate_insights

    return generate_insights(file_bytes)


# =========================
# CODE 6 (Госы)
# Для каждого выбранного листа «гос*»:
#   - находит столбцы «Номер договора» и «Статус» в строке 1
#   - по каждой строке со статусом «Действует» запрашивает данные из API goszakup
#   - записывает результат правее существующей таблицы
# =========================

_GOS_BASE_OWS = "https://ows.goszakup.gov.kz"
_GOS_BASE_WEB = "https://www.goszakup.gov.kz"

_GOS_NEW_COLS = [
    "Ссылка",
    "Многолетний",
    "Год окончания",
    "Сумма 2026",
    "Сумма 2027",
    "% аванса",
    "Сумма аванса",
    "Обеспечение аванса требуется",
    "Обеспечение аванса прикреплено",
    "Факт сумма",
]


def _gos_money(x) -> float:
    if x in [None, "", "null"]:
        return 0.0
    try:
        return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except Exception:
        return 0.0


def _gos_parse_money(text: str) -> float:
    if text is None:
        return 0.0
    text = str(text).replace("\xa0", " ").replace(" ", "").replace(",", ".")
    text = re.sub(r"[^\d.]", "", text)
    return _gos_money(text) if text else 0.0


def _gos_clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def _gos_html_to_text(html: str) -> str:
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        return _gos_clean_text(soup.get_text("\n", strip=True))
    except Exception:
        text = re.sub(r"<script.*?</script>", " ", html, flags=re.S | re.I)
        text = re.sub(r"<style.*?</style>", " ", text, flags=re.S | re.I)
        text = re.sub(r"<[^>]+>", "\n", text)
        return _gos_clean_text(text)


def _gos_get_year_from_date(date_str):
    if not date_str:
        return None
    try:
        return int(str(date_str)[:4])
    except Exception:
        return None


def _gos_get_max_end_year(contract: Dict[str, Any]) -> Optional[int]:
    years = []
    for field in ["contract_end_date", "plan_exec_date", "ec_end_date"]:
        y = _gos_get_year_from_date(contract.get(field))
        if y:
            years.append(y)
    return max(years) if years else None


def _gos_get_additions_url(contract: Dict[str, Any]) -> str:
    additions_id = contract.get("root_id") or contract.get("id")
    if not additions_id:
        return ""
    return f"{_GOS_BASE_WEB}/ru/egzcontract/cpublic/additions/{additions_id}"


def _gos_get_json(url: str, token: str, params: Optional[Dict[str, Any]] = None) -> Any:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
    }
    r = _requests_lib.get(url, headers=headers, params=params, timeout=40)
    if r.status_code == 401:
        raise Exception("401: неверный или просроченный токен")
    if r.status_code == 403:
        raise Exception(f"403: нет доступа к API: {url}")
    if r.status_code == 404:
        raise Exception(f"404: не найдено: {url}")
    r.raise_for_status()
    return r.json()


def _gos_get_html_text(url: str) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    r = _requests_lib.get(url, headers=headers, timeout=40)
    r.raise_for_status()
    return _gos_html_to_text(r.text)


def _gos_unwrap_contract(data: Any) -> Dict[str, Any]:
    if isinstance(data, dict) and "data" in data:
        data = data["data"]
    if isinstance(data, list):
        if not data:
            raise Exception("Договор не найден")
        return data[0]
    if isinstance(data, dict):
        return data
    raise Exception(f"Неожиданный формат ответа: {type(data)}")


def _gos_get_contract_by_number(contract_number: str, token: str) -> Dict[str, Any]:
    if "/" in contract_number:
        url = f"{_GOS_BASE_OWS}/v3/contract/number-sys/"
    else:
        url = f"{_GOS_BASE_OWS}/v3/contract/number/"
    data = _gos_get_json(url, token, params={"number": contract_number})
    return _gos_unwrap_contract(data)


def _gos_get_contract_units(contract_id: int, token: str) -> Dict[str, Any]:
    url = f"{_GOS_BASE_OWS}/v3/contract/{contract_id}/units"
    return _gos_get_json(url, token)


def _gos_get_plan_point(pln_point_id: int, token: str) -> Optional[Dict[str, Any]]:
    url = f"{_GOS_BASE_OWS}/v3/plans/view/{pln_point_id}"
    try:
        data = _gos_get_json(url, token)
        if isinstance(data, dict) and "data" in data:
            return data["data"]
        return data
    except Exception:
        return None


def _gos_extract_current_block(additions_text: str, contract_number: str) -> str:
    start_pattern = re.escape(contract_number) + r"\s+Статус договора:"
    start_match = re.search(start_pattern, additions_text)
    if not start_match:
        return additions_text
    start = start_match.start()
    next_match = re.search(r"\n\d{12}/\d{6}/\d{2}\s+Статус договора:", additions_text[start_match.end():])
    if next_match:
        return additions_text[start: start_match.end() + next_match.start()]
    footer_patterns = ["АИИС ЭГЗ", "Техническая поддержка", "Чем я могу Вам помочь"]
    end = len(additions_text)
    for p in footer_patterns:
        m = re.search(p, additions_text[start:])
        if m:
            end = min(end, start + m.start())
    return additions_text[start:end]


def _gos_parse_year_amounts(block: str) -> Dict[str, float]:
    result: Dict[str, float] = {}
    pattern = r"Сумма\s+на\s+(\d{4})\s+год\s*[-:]\s*([\d\s]+(?:[.,]\d+)?)"
    for year, amount in re.findall(pattern, block, flags=re.I):
        result[year] = result.get(year, 0.0) + _gos_parse_money(amount)
    return {k: _gos_money(v) for k, v in result.items()}


def _gos_parse_advance_security(block: str) -> Dict[str, Any]:
    result: Dict[str, Any] = {"advance_security_required": None, "advance_security_attached": None}
    pattern = (
        r"Обеспечение исполнения договора на сумму аванса по договору\s*"
        r"Требуется:\s*(Да|Нет)\s*,\s*Прикреплен:\s*(Да|Нет)"
    )
    m = re.search(pattern, block, flags=re.I)
    if m:
        result["advance_security_required"] = 1 if m.group(1).lower() == "да" else 0
        result["advance_security_attached"] = 1 if m.group(2).lower() == "да" else 0
    return result


def _gos_parse_totals(block: str) -> Dict[str, Any]:
    result: Dict[str, Any] = {"contract_sum": None, "fact_sum": None}
    m = re.search(r"Общая итоговая сумма договора\s*:\s*([\d\s]+(?:[.,]\d+)?)", block, flags=re.I)
    if m:
        result["contract_sum"] = _gos_parse_money(m.group(1))
    m = re.search(r"Общая фактическая сумма договора\s*:\s*([\d\s]+(?:[.,]\d+)?)", block, flags=re.I)
    if m:
        result["fact_sum"] = _gos_parse_money(m.group(1))
    return result


def _gos_extract_advance_percent(contract: Dict[str, Any], token: str) -> Dict[str, Any]:
    advance_percent = 0.0
    unresolved: List[int] = []
    units = contract.get("units", {}).get("items", []) or []
    for unit in units:
        pln_point_id = unit.get("pln_point_id")
        if not pln_point_id:
            continue
        pp = _gos_get_plan_point(pln_point_id, token)
        if not pp:
            unresolved.append(pln_point_id)
            continue
        prepayment = pp.get("prepayment")
        if prepayment in [None, "", "null"]:
            prepayment = 0
        try:
            prepayment = float(prepayment)
        except Exception:
            prepayment = 0.0
        if prepayment > advance_percent:
            advance_percent = prepayment
    if advance_percent > 0:
        return {"status": "ok", "advance_percent": _gos_money(advance_percent), "message": "", "unresolved": unresolved}
    if unresolved:
        return {
            "status": "error",
            "advance_percent": None,
            "message": f"ОШИБКА: аванс не удалось проверить. Не раскрылись plans/view: {unresolved}",
            "unresolved": unresolved,
        }
    return {"status": "ok", "advance_percent": 0.0, "message": "", "unresolved": []}


def _gos_build_row(contract_number: str, token: str) -> Dict[str, Any]:
    contract = _gos_get_contract_by_number(contract_number, token)
    contract_id = contract.get("id")
    if not contract_id:
        raise Exception("В договоре нет contract_id")
    try:
        contract["units"] = _gos_get_contract_units(contract_id, token)
    except Exception:
        contract["units"] = {"items": []}

    additions_url = _gos_get_additions_url(contract)
    is_long_term = 1 if contract.get("ref_contract_year_type_id") == 2 else 0
    contract_sum = _gos_money(contract.get("contract_sum_wnds"))
    fact_sum = _gos_money(contract.get("fakt_sum_wnds"))
    amount_2026 = 0.0
    amount_2027 = 0.0
    advance_security_required = None
    advance_security_attached = None
    max_end_year = _gos_get_max_end_year(contract)
    current_block = ""
    html_status = "not_used"
    html_error = ""

    if is_long_term == 1:
        try:
            additions_text = _gos_get_html_text(additions_url) if additions_url else ""
            current_block = _gos_extract_current_block(
                additions_text,
                contract.get("contract_number_sys") or contract_number,
            )
            amounts_by_year = _gos_parse_year_amounts(current_block)
            totals = _gos_parse_totals(current_block)
            adv_sec = _gos_parse_advance_security(current_block)
            amount_2026 = _gos_money(amounts_by_year.get("2026", 0))
            amount_2027 = 0.0 if (max_end_year is not None and max_end_year <= 2026) else _gos_money(amounts_by_year.get("2027", 0))
            if totals.get("contract_sum") is not None:
                contract_sum = _gos_money(totals["contract_sum"])
            if totals.get("fact_sum") is not None:
                fact_sum = _gos_money(totals["fact_sum"])
            advance_security_required = adv_sec["advance_security_required"]
            advance_security_attached = adv_sec["advance_security_attached"]
            html_status = "ok"
        except Exception as e:
            html_error = str(e)
            html_status = "error"
            amount_2026 = f"ОШИБКА HTML: {html_error}"
            amount_2027 = f"ОШИБКА HTML: {html_error}"
    else:
        fin_year = str(contract.get("fin_year") or "")
        if fin_year == "2026":
            amount_2026 = contract_sum
        elif fin_year == "2027":
            amount_2027 = contract_sum

    advance_data = _gos_extract_advance_percent(contract, token)
    if advance_data["status"] == "error":
        advance_percent = advance_data["message"]
        advance_sum = advance_data["message"]
        advance_security_required = advance_data["message"]
        advance_security_attached = advance_data["message"]
    else:
        advance_percent = advance_data["advance_percent"]
        if advance_percent > 0:
            advance_sum = _gos_money(contract_sum * advance_percent / 100)
            if not current_block and additions_url:
                try:
                    additions_text = _gos_get_html_text(additions_url)
                    current_block = _gos_extract_current_block(
                        additions_text,
                        contract.get("contract_number_sys") or contract_number,
                    )
                    adv_sec = _gos_parse_advance_security(current_block)
                    advance_security_required = adv_sec["advance_security_required"]
                    advance_security_attached = adv_sec["advance_security_attached"]
                    html_status = "ok"
                except Exception as e:
                    html_status = "error"
                    html_error = str(e)
            if html_status != "ok":
                msg = f"ОШИБКА: HTML не получен: {html_error}"
                advance_security_required = msg
                advance_security_attached = msg
            elif advance_security_required is None:
                advance_security_required = "ОШИБКА: строка обеспечения не найдена"
            elif advance_security_attached is None:
                advance_security_attached = "ОШИБКА: строка обеспечения не найдена"
        else:
            advance_percent = 0.0
            advance_sum = 0.0
            advance_security_required = 0
            advance_security_attached = 0

    return {
        "Ссылка": additions_url,
        "Многолетний": is_long_term,
        "Год окончания": max_end_year,
        "Сумма 2026": amount_2026,
        "Сумма 2027": amount_2027,
        "% аванса": advance_percent,
        "Сумма аванса": advance_sum,
        "Обеспечение аванса требуется": advance_security_required,
        "Обеспечение аванса прикреплено": advance_security_attached,
        "Факт сумма": fact_sum,
    }


def run_code_6_gos(file_bytes: bytes, selected_gos_sheets: List[str], token: str) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))

    for sheet_name in selected_gos_sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        # Найти столбцы «Номер договора» и «Статус» в строке 1
        col_contract: Optional[int] = None
        col_status: Optional[int] = None
        for col_idx in range(1, (ws.max_column or 1) + 1):
            cell_val = ws.cell(row=1, column=col_idx).value
            if cell_val is None:
                continue
            cv = str(cell_val).strip()
            if cv == "Номер договора":
                col_contract = col_idx
            elif cv == "Статус":
                col_status = col_idx

        if col_contract is None or col_status is None:
            continue

        # Начало новых столбцов — после последнего используемого столбца
        start_col = (ws.max_column or 1) + 1

        # Заголовки новых столбцов в строке 1
        for j, col_name in enumerate(_GOS_NEW_COLS):
            ws.cell(row=1, column=start_col + j, value=col_name)

        # Обработка строк
        for row_idx in range(2, (ws.max_row or 1) + 1):
            status_val = ws.cell(row=row_idx, column=col_status).value
            if str(status_val or "").strip() != "Действует":
                continue
            contract_number = ws.cell(row=row_idx, column=col_contract).value
            if not contract_number:
                continue
            contract_number = str(contract_number).strip()
            try:
                row_data = _gos_build_row(contract_number, token)
                for j, col_name in enumerate(_GOS_NEW_COLS):
                    val = row_data.get(col_name)
                    c = ws.cell(row=row_idx, column=start_col + j)
                    if col_name == "Ссылка" and val:
                        c.value = "Открыть"
                        c.hyperlink = str(val)
                    else:
                        c.value = val
            except Exception as e:
                ws.cell(row=row_idx, column=start_col, value=f"ОШИБКА: {e}")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 7 (Объединить общие ОСВ по группе)
# Находит листы «*общ*» (с префиксами и суффиксами),
# группирует по суффиксу (году), создаёт лист GroupXX с объединёнными данными.
# =========================

_GSC_SEC_WIDTH = 8  # account (1) + data (6) + separator (1)


def _gsc_excel_sheetref(name: str) -> str:
    """Возвращает ссылку на лист в формуле Excel (с кавычками если нужно)."""
    if re.search(r"[ \t!'\[\]\*\?:~\(\)]", name) or (name and name[0].isdigit()):
        return "'" + name.replace("'", "''") + "'"
    return name


def _gsc_parse_osv_data_cols(ws) -> Dict[str, int]:
    """Определяет номера столбцов для 6 полей данных в листе ОСВ."""
    saldo_nach_col = oborot_col = saldo_kon_col = None
    header_row_idx = None

    for row_cells in ws.iter_rows(min_row=1, max_row=20):
        for cell in row_cells:
            if isinstance(cell , MergedCell):
                continue
            v = str(cell.value or "").strip().lower()
            if "сальдо на начало" in v and saldo_nach_col is None:
                saldo_nach_col = cell.column
                header_row_idx = cell.row
            if "обороты за период" in v and oborot_col is None:
                oborot_col = cell.column
            if "сальдо на конец" in v and saldo_kon_col is None:
                saldo_kon_col = cell.column

    if saldo_nach_col is None:
        saldo_nach_col = 3
    if oborot_col is None:
        oborot_col = 5
    if saldo_kon_col is None:
        saldo_kon_col = 7

    # Ищем строку с «Дебет»/«Кредит» ниже заголовков
    deb_kred_map: Dict[int, str] = {}
    search_from = header_row_idx if header_row_idx else 1
    for row_cells in ws.iter_rows(min_row=search_from, max_row=search_from + 4):
        for cell in row_cells:
            if isinstance(cell, MergedCell):
                continue
            v = str(cell.value or "").strip().lower()
            if v == "дебет":
                deb_kred_map[cell.column] = "дебет"
            elif v == "кредит":
                deb_kred_map[cell.column] = "кредит"
        if deb_kred_map:
            break

    def find_pair(g_col: int):
        cols_at = sorted(c for c in deb_kred_map if c >= g_col)
        deb = next((c for c in cols_at if deb_kred_map[c] == "дебет"), g_col)
        kred = next((c for c in cols_at if deb_kred_map[c] == "кредит" and c > deb), deb + 1)
        return deb, kred

    if deb_kred_map:
        d1, k1 = find_pair(saldo_nach_col)
        d2, k2 = find_pair(oborot_col)
        d3, k3 = find_pair(saldo_kon_col)
    else:
        d1, k1 = saldo_nach_col, saldo_nach_col + 1
        d2, k2 = oborot_col, oborot_col + 1
        d3, k3 = saldo_kon_col, saldo_kon_col + 1

    return {
        "saldo_nach_deb": d1, "saldo_nach_kred": k1,
        "oborot_deb": d2, "oborot_kred": k2,
        "saldo_kon_deb": d3, "saldo_kon_kred": k3,
    }


_GSC_FIELD_KEYS = [
    "saldo_nach_deb", "saldo_nach_kred",
    "oborot_deb", "oborot_kred",
    "saldo_kon_deb", "saldo_kon_kred",
]


def run_code_7_group_osv(file_bytes: bytes, selected_groups: Dict[str, List[str]]) -> bytes:
    """
    selected_groups: {suffix: [sheet_names]} — например {"25": ["Uобщ25", "Тобщ25"]}
    """
    wb = load_workbook(io.BytesIO(file_bytes))

    # ── стили ──────────────────────────────────────────────────────────
    # #,##0 → Excel применяет разделитель тысяч для ВСЕХ групп по локали (пробел в RU/KZ)
    # третья секция пустая → ноль отображается как пустая ячейка, значение остаётся 0
    NUM_FMT = "#,##0_ ;[Red]-#,##0 ;"

    # шрифты
    F       = Font(name="Arial", size=8)
    F_BOLD  = Font(name="Arial", size=8, bold=True)
    F_HDR   = Font(name="Arial", size=8, bold=True, color="FFFFFF")

    # заливки
    FILL_H1     = PatternFill("solid", fgColor="17375E")  # строка 1 (тёмный синий)
    FILL_H2     = PatternFill("solid", fgColor="2E74B5")  # строка 2 (средний синий)
    FILL_GRP    = PatternFill("solid", fgColor="DEEAF1")  # группа (светло-голубой)
    FILL_CO_ODD = PatternFill("solid", fgColor="F2F7FC")  # нечётные компании
    FILL_ITOGO  = PatternFill("solid", fgColor="D9E1F2")  # итого
    FILL_SEP    = PatternFill("solid", fgColor="D4DCE8")  # разделитель
    FILL_NONE   = PatternFill(fill_type=None)

    # стороны границ
    _T = Side(style="thin",   color="D9D9D9")
    _M = Side(style="medium", color="2E74B5")
    _N = Side(style=None)

    def _border(left=_T, right=_T, top=_T, bottom=_T):
        return Border(left=left, right=right, top=top, bottom=bottom)

    # выравнивание
    AL_C  = Alignment(horizontal="center",  vertical="center", wrap_text=True)
    AL_L  = Alignment(horizontal="left",    vertical="center", wrap_text=True)
    AL_R  = Alignment(horizontal="right",   vertical="center")
    AL_SEP = Alignment()

    insert_pos = 0
    for suffix in sorted(selected_groups.keys()):
        source_sheets = selected_groups[suffix]
        valid_sheets = [s for s in source_sheets if s in wb.sheetnames]
        if len(valid_sheets) < 2:
            continue

        # ── читаем данные исходных ОСВ ─────────────────────────────────
        companies: List[Dict[str, Any]] = []
        for sh in valid_sheets:
            ws_src = wb[sh]
            company_name = str(ws_src.cell(1, 1).value or sh).strip()
            period = str(ws_src.cell(2, 1).value or "").strip()
            data_cols = _gsc_parse_osv_data_cols(ws_src)
            accounts: List[str] = []
            for row_cells in ws_src.iter_rows(min_col=1, max_col=1):
                val = str(row_cells[0].value or "").strip()
                if re.match(r"^\d{4}", val):
                    accounts.append(val)
            companies.append({
                "sheet_name": sh, "company_name": company_name,
                "period": period, "data_cols": data_cols, "accounts": accounts,
            })

        if not companies:
            continue

        # ── объединяем счета ────────────────────────────────────────────
        seen: Set[str] = set()
        all_accounts: List[str] = []
        for cd in companies:
            for acc in cd["accounts"]:
                if acc not in seen:
                    seen.add(acc)
                    all_accounts.append(acc)

        def _sort_key(s: str):
            m = re.match(r"^(\d{4})", s)
            return (int(m.group(1)) if m else 99999, s)

        all_accounts.sort(key=_sort_key)

        # XX00 счета → жирные (код оканчивается на 00)
        round_set: Set[str] = set()
        for acc in all_accounts:
            m = re.match(r"^(\d{4})", acc)
            if m and m.group(1)[2:] == "00":
                round_set.add(acc)

        # ── создаём лист ────────────────────────────────────────────────
        group_title = f"Group{suffix}" if suffix else "Group"
        group_title = make_unique_sheet_title(wb, group_title)
        ws_g = wb.create_sheet(group_title, insert_pos)
        insert_pos += 1
        n = len(companies)

        # вспомогательные функции адресации
        def sec_start(k: int) -> int:        return 1 + k * _GSC_SEC_WIDTH
        def sec_acc_col(k: int) -> int:      return sec_start(k)
        def sec_data_col(k: int, o: int) -> int: return sec_start(k) + 1 + o

        def sumifs(src_sh: str, src_col: int, lkp_col: int, row: int) -> str:
            sr  = _gsc_excel_sheetref(src_sh)
            gr  = _gsc_excel_sheetref(group_title)
            scl = get_column_letter(src_col)
            lcl = get_column_letter(lkp_col)
            return f"=SUMIFS({sr}!{scl}:{scl},{sr}!$A:$A,{gr}!${lcl}{row})"

        DATA_ROW_START = 3
        itogo_row      = DATA_ROW_START + len(all_accounts)
        # количество колонок: (n+1) секций, последняя без разделителя
        total_cols     = n * _GSC_SEC_WIDTH + 7

        # ── пишем контент ───────────────────────────────────────────────

        # строка 1
        group_header = " + ".join(cd["company_name"] for cd in companies)
        ws_g.cell(1, sec_acc_col(0), group_header)
        ws_g.cell(1, sec_data_col(0, 0), "Сальдо на начало")
        ws_g.cell(1, sec_data_col(0, 2), "Обороты за период")
        ws_g.cell(1, sec_data_col(0, 4), "Сальдо на конец")
        for k, cd in enumerate(companies, 1):
            ws_g.cell(1, sec_acc_col(k), cd["sheet_name"])
            ws_g.cell(1, sec_data_col(k, 0), "Сальдо на начало")
            ws_g.cell(1, sec_data_col(k, 2), "Обороты за период")
            ws_g.cell(1, sec_data_col(k, 4), "Сальдо на конец")

        # строка 2
        ws_g.cell(2, sec_acc_col(0), companies[0]["period"])
        for k in range(n + 1):
            for o, lbl in enumerate(["Дебет", "Кредит", "Дебет", "Кредит", "Дебет", "Кредит"]):
                ws_g.cell(2, sec_data_col(k, o), lbl)

        # строки счетов
        for i, acc in enumerate(all_accounts):
            r = DATA_ROW_START + i
            ws_g.cell(r, sec_acc_col(0), acc)
            for o in range(6):
                parts = [get_column_letter(sec_data_col(k, o)) + str(r) for k in range(1, n + 1)]
                ws_g.cell(r, sec_data_col(0, o), "=" + "+".join(parts))
            for k, cd in enumerate(companies, 1):
                ws_g.cell(r, sec_acc_col(k), acc)
                for o, fkey in enumerate(_GSC_FIELD_KEYS):
                    src_col = cd["data_cols"].get(fkey)
                    if src_col:
                        ws_g.cell(r, sec_data_col(k, o), sumifs(cd["sheet_name"], src_col, sec_acc_col(k), r))

        # итого
        ws_g.cell(itogo_row, sec_acc_col(0), "Итого")
        for o in range(6):
            parts = [get_column_letter(sec_data_col(k, o)) + str(itogo_row) for k in range(1, n + 1)]
            ws_g.cell(itogo_row, sec_data_col(0, o), "=" + "+".join(parts))
        for k, cd in enumerate(companies, 1):
            ws_g.cell(itogo_row, sec_acc_col(k), "Итого")
            for o, fkey in enumerate(_GSC_FIELD_KEYS):
                src_col = cd["data_cols"].get(fkey)
                if src_col:
                    ws_g.cell(itogo_row, sec_data_col(k, o), sumifs(cd["sheet_name"], src_col, sec_acc_col(k), itogo_row))

        # ── объединение ячеек заголовка ─────────────────────────────────
        for k in range(n + 1):
            # аккаунт-ячейка: строки 1-2 не объединяем, только ставим wrap
            ws_g.merge_cells(
                start_row=1, start_column=sec_data_col(k, 0),
                end_row=1,   end_column=sec_data_col(k, 1))
            ws_g.merge_cells(
                start_row=1, start_column=sec_data_col(k, 2),
                end_row=1,   end_column=sec_data_col(k, 3))
            ws_g.merge_cells(
                start_row=1, start_column=sec_data_col(k, 4),
                end_row=1,   end_column=sec_data_col(k, 5))

        # ── применяем дизайн ко всем ячейкам ───────────────────────────
        for r in range(1, itogo_row + 1):
            for c in range(1, total_cols + 1):
                cell = ws_g.cell(r, c)
                sec_idx = (c - 1) // _GSC_SEC_WIDTH
                offset  = (c - 1) %  _GSC_SEC_WIDTH
                is_sep  = offset == 7
                is_acc  = offset == 0
                is_data = not is_sep and not is_acc

                # шрифт
                if is_sep:
                    cell.font = F
                elif r <= 2:
                    cell.font = F_HDR
                elif r == itogo_row or (is_acc and str(cell.value or "") in round_set):
                    cell.font = F_BOLD
                else:
                    cell.font = F

                # заливка
                if r == 1:
                    cell.fill = FILL_H1
                elif r == 2:
                    cell.fill = FILL_H2
                elif is_sep:
                    cell.fill = FILL_SEP
                elif r == itogo_row:
                    cell.fill = FILL_ITOGO
                elif sec_idx == 0:
                    cell.fill = FILL_GRP
                elif sec_idx % 2 == 0:
                    cell.fill = FILL_CO_ODD
                else:
                    cell.fill = FILL_NONE

                # формат чисел
                if r >= DATA_ROW_START and is_data:
                    cell.number_format = NUM_FMT

                # выравнивание
                if is_sep:
                    cell.alignment = AL_SEP
                elif r <= 2:
                    cell.alignment = AL_C
                elif is_acc:
                    cell.alignment = AL_L
                else:
                    cell.alignment = AL_R

                # границы
                if is_sep:
                    cell.border = Border()
                else:
                    right_side = _M if offset == 6 else _T
                    top_side   = _M if r == itogo_row or r == DATA_ROW_START else _T
                    bot_side   = _M if r == 2 else _T
                    cell.border = _border(right=right_side, top=top_side, bottom=bot_side)

        # ── ширины колонок и высоты строк ──────────────────────────────
        for c in range(1, total_cols + 1):
            offset = (c - 1) % _GSC_SEC_WIDTH
            ltr    = get_column_letter(c)
            if offset == 7:
                ws_g.column_dimensions[ltr].width = 0.6
            elif offset == 0:
                ws_g.column_dimensions[ltr].width = 30
            else:
                ws_g.column_dimensions[ltr].width = 11.5

        ws_g.row_dimensions[1].height = 30
        ws_g.row_dimensions[2].height = 22
        for r in range(DATA_ROW_START, itogo_row):
            ws_g.row_dimensions[r].height = 13
        ws_g.row_dimensions[itogo_row].height = 16

        # ── заморозка ───────────────────────────────────────────────────
        ws_g.freeze_panes = "B3"


    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 8 (Объединение ОСВ по счёту)
# Берёт ≥2 листа с одинаковым 4-значным суффиксом счёта (например 1210),
# определяет их хронологический порядок по датам в шапке,
# сливает строки (дубликаты убираются), суммирует обороты,
# BOP берёт из самого старого листа, EOP считает сам,
# создаёт новый лист с идентичным форматированием и удаляет источники.
# =========================

_YEAR_RE = re.compile(r'\d{4}')
_OSV_DASH_RE = re.compile(r'\s*[–—-]\s*')


def _get_osv_header_text(ws, account: str) -> Optional[str]:
    """Возвращает полный текст ячейки с номером счёта из первых 5 строк."""
    for r in range(1, min(6, (ws.max_row or 1) + 1)):
        for c in range(1, min((ws.max_column or 1) + 1, 25)):
            v = ws.cell(row=r, column=c).value
            if v is not None and account in str(v):
                return str(v).strip()
    return None


def _get_osv_end_year(ws, account: str) -> Optional[int]:
    """
    Сканирует первые 5 строк листа, ищет ячейку с номером счёта,
    берёт текст ПОСЛЕ него и возвращает максимальный год (1900-2100).
    Это год конца периода — по нему определяем, какой лист старее.

    Работает для любого формата:
    "...по счету 1210 за 01.01.2025 - 10.07.2025"  -> 2025
    "...по счету 1710 за Январь 2025 г. - Июнь 2026" -> 2026
    """
    for r in range(1, min(6, (ws.max_row or 1) + 1)):
        for c in range(1, min((ws.max_column or 1) + 1, 25)):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            s = str(v)
            idx = s.find(account)
            if idx == -1:
                continue
            tail = s[idx + len(account):]
            years = [int(y) for y in _YEAR_RE.findall(tail) if 1900 <= int(y) <= 2100]
            if years:
                return max(years)
    return None


def _read_osv_data_for_merge(ws, account: str) -> Dict[str, Dict[str, float]]:
    """
    Читает строки между строкой счёта и «Итого», возвращает:
    {name: {bop_dt, bop_kt, turn_dt, turn_kt, eop_dt, eop_kt}}
    Колонки: A=1 (name), C=3 (bop_dt), D=4 (bop_kt),
            E=5 (turn_dt), F=6 (turn_kt), G=7 (eop_dt), H=8 (eop_kt).
    """
    start_row = find_first_row_with_value(ws, account, col=1)
    end_row = find_first_row_with_value(ws, "Итого", col=1)
    if end_row is None:
        for r in range(1, (ws.max_row or 1) + 1):
            v = ws.cell(row=r, column=1).value
            if v and str(v).strip().lower().startswith("итого"):
                end_row = r
                break
    if not start_row or not end_row or end_row <= start_row:
        return {}

    result: Dict[str, Dict[str, float]] = {}
    for r in range(start_row + 1, end_row):
        raw = ws.cell(row=r, column=1).value
        if raw is None:
            continue
        name = str(raw).strip()
        if not name or name.lower().startswith("итого") or name == account:
            continue
        vals = {
            "bop_dt":  to_number(ws.cell(row=r, column=3).value) or 0.0,
            "bop_kt":  to_number(ws.cell(row=r, column=4).value) or 0.0,
            "turn_dt": to_number(ws.cell(row=r, column=5).value) or 0.0,
            "turn_kt": to_number(ws.cell(row=r, column=6).value) or 0.0,
            "eop_dt":  to_number(ws.cell(row=r, column=7).value) or 0.0,
            "eop_kt":  to_number(ws.cell(row=r, column=8).value) or 0.0,
        }
        if any(abs(v) > 1e-9 for v in vals.values()):
            result[name] = vals
    return result


def run_code_8_merge_osv(file_bytes: bytes, merge_groups: Dict[str, List[str]]) -> bytes:
    """
    merge_groups: {account: [sheet_name, ...]}  — листы для объединения.
    Для каждой группы:
    • определяет хронологический порядок по датам шапки,
    • BOP берёт из самого старого листа,
    • обороты суммирует по всем,
    • EOP рассчитывает по правилу дебет/кредит счёта,
    • создаёт новый лист с форматированием как у clean_osv_sheet_inplace,
    • удаляет исходные листы.
    """
    wb = load_workbook(io.BytesIO(file_bytes))
    num_fmt_osv = "#,##0; -#,##0"

    for account, sheet_names in merge_groups.items():
        valid = [s for s in sheet_names if s in wb.sheetnames]
        if len(valid) < 2:
            continue

        # ── 1. Определяем порядок: по году конца периода из строки «Оборотно…» ──
        # _get_osv_end_year ищет максимальный год (1900-2100) в тексте ПОСЛЕ
        # номера счёта в первых 5 строках листа. Меньший год = старее.
        year_by_sheet: Dict[str, int] = {}
        for sn in valid:
            y = _get_osv_end_year(wb[sn], account)
            year_by_sheet[sn] = y if y is not None else 9999

        valid_sorted = sorted(valid, key=lambda sn: year_by_sheet[sn])

        oldest_header = _get_osv_header_text(wb[valid_sorted[0]], account)
        newest_header = _get_osv_header_text(wb[valid_sorted[-1]], account)

        period_text = ""
        if oldest_header and newest_header and valid_sorted[0] != valid_sorted[-1]:
            # Начало периода — из старого: всё до первого даша
            m_old = _OSV_DASH_RE.search(oldest_header)
            start_part = oldest_header[:m_old.start()].rstrip() if m_old else oldest_header
            # Конец периода — из нового: всё после последнего даша
            m_new_iter = list(_OSV_DASH_RE.finditer(newest_header))
            end_part = newest_header[m_new_iter[-1].end():].strip() if m_new_iter else newest_header
            period_text = f"{start_part} – {end_part}" if end_part else oldest_header
        elif oldest_header:
            period_text = oldest_header
        else:
            years_found = [year_by_sheet[sn] for sn in valid_sorted if year_by_sheet[sn] != 9999]
            if len(years_found) >= 2:
                period_text = f"{years_found[0]} – {years_found[-1]}"
            elif years_found:
                period_text = str(years_found[0])

        oldest_sn = valid_sorted[0]
        ws_oldest = wb[oldest_sn]
        company_label = extract_company_label_from_a1(ws_oldest)

        # ── 2. Читаем данные ────────────────────────────────────────────
        # BOP — только из самого старого листа
        data_old = _read_osv_data_for_merge(ws_oldest, account)

        merged: Dict[str, Dict[str, float]] = {}
        for name, vals in data_old.items():
            merged[name] = {
                "bop_dt":  vals["bop_dt"],
                "bop_kt":  vals["bop_kt"],
                "turn_dt": 0.0,
                "turn_kt": 0.0,
            }

        # Обороты суммируем из ВСЕХ листов (в хронологическом порядке)
        for sn in valid_sorted:
            ws_src = wb[sn]
            data_src = _read_osv_data_for_merge(ws_src, account)
            for name, vals in data_src.items():
                if name not in merged:
                    merged[name] = {"bop_dt": 0.0, "bop_kt": 0.0, "turn_dt": 0.0, "turn_kt": 0.0}
                merged[name]["turn_dt"] += vals["turn_dt"]
                merged[name]["turn_kt"] += vals["turn_kt"]

        # ── 3. Рассчитываем EOP ─────────────────────────────────────────
        first_digit = int(account[0]) if account and account[0].isdigit() else 1
        is_debit_account = first_digit <= 2  # 1xxx/2xxx — дебетовые; 3xxx+ — кредитовые

        for name, vals in merged.items():
            if is_debit_account:
                net = vals["bop_dt"] - vals["bop_kt"] + vals["turn_dt"] - vals["turn_kt"]
                vals["eop_dt"] = max(net, 0.0)
                vals["eop_kt"] = max(-net, 0.0)
            else:
                net = vals["bop_kt"] - vals["bop_dt"] + vals["turn_kt"] - vals["turn_dt"]
                vals["eop_kt"] = max(net, 0.0)
                vals["eop_dt"] = max(-net, 0.0)

        # Убираем строки где всё нули
        merged = {
            name: vals for name, vals in merged.items()
            if any(abs(v) > 1e-9 for v in vals.values())
        }

        # ── 4. Сортировка: убываем по рабочей стороне EOP ──────────────
        eop_key = "eop_dt" if is_debit_account else "eop_kt"
        sorted_rows = sorted(merged.items(), key=lambda x: x[1][eop_key], reverse=True)

        # ── 5. Итого ────────────────────────────────────────────────────
        itogo: Dict[str, float] = {"bop_dt": 0.0, "bop_kt": 0.0, "turn_dt": 0.0,
                                    "turn_kt": 0.0, "eop_dt": 0.0, "eop_kt": 0.0}
        for _, vals in sorted_rows:
            for k in itogo:
                itogo[k] += vals[k]

        # ── 6. Создаём новый лист ───────────────────────────────────────
        # Вставляем на место первого исходного листа
        first_idx = next((i for i, sn in enumerate(wb.sheetnames) if sn == valid[0]), None)

        # Имя: если у всех листов один и тот же prefix — сохраняем его
        prefixes = [split_prefix_suffix4(sn)[0].strip() for sn in valid]
        common_prefix = prefixes[0] if len(set(prefixes)) == 1 else ""
        new_title = make_unique_with_fixed_suffix(wb, common_prefix, account)

        ws_new = (wb.create_sheet(new_title, first_idx)
                if first_idx is not None else wb.create_sheet(new_title))

        # Шапка
        ws_new.cell(row=1, column=1, value=company_label or f"ОСВ по счёту {account}")
        ws_new.cell(row=2, column=1, value=period_text)

        ACCOUNT_ROW = 3
        # Строка счёта (заголовки колонок)
        hdrs = {
            1: account,
            3: "Сальдо нач Дт",
            4: "Сальдо нач Кт",
            5: "Обороты Дт",
            6: "Обороты Кт",
            7: "Сальдо кон Дт",
            8: "Сальдо кон Кт",
        }
        font_hdr = Font(name="Arial", size=9, bold=True)
        align_c  = Alignment(horizontal="center", vertical="center")
        align_l  = Alignment(horizontal="left",   vertical="center")
        align_r  = Alignment(horizontal="right",  vertical="center")

        for col, val in hdrs.items():
            cell = ws_new.cell(row=ACCOUNT_ROW, column=col, value=val)
            cell.font = font_hdr
            cell.alignment = align_l if col == 1 else align_c
            cell.number_format = num_fmt_osv if col >= 3 else "@"

        # Строки данных
        DATA_START = ACCOUNT_ROW + 1
        font_body = Font(name="Arial", size=9)
        font_neg  = Font(name="Arial", size=9, color="FFCC0000")
        total_col = 7 if is_debit_account else 8

        for i, (name, vals) in enumerate(sorted_rows):
            r = DATA_START + i
            ws_new.cell(row=r, column=1, value=name).alignment = align_l
            ws_new.cell(row=r, column=1).font = font_body

            for col, key in [(3, "bop_dt"), (4, "bop_kt"), (5, "turn_dt"),
                            (6, "turn_kt"), (7, "eop_dt"), (8, "eop_kt")]:
                v = vals[key]
                cell = ws_new.cell(row=r, column=col, value=v if abs(v) > 1e-9 else None)
                cell.number_format = num_fmt_osv
                cell.alignment = align_r
                cell.font = font_neg if (col == total_col and v < 0) else font_body

        # Строка Итого
        ITOGO_ROW = DATA_START + len(sorted_rows)
        ws_new.cell(row=ITOGO_ROW, column=1, value="Итого").font = font_hdr
        ws_new.cell(row=ITOGO_ROW, column=1).alignment = align_l
        for col, key in [(3, "bop_dt"), (4, "bop_kt"), (5, "turn_dt"),
                        (6, "turn_kt"), (7, "eop_dt"), (8, "eop_kt")]:
            v = itogo[key]
            cell = ws_new.cell(row=ITOGO_ROW, column=col, value=v if abs(v) > 1e-9 else None)
            cell.number_format = num_fmt_osv
            cell.alignment = align_r
            cell.font = font_hdr

        # ── 7. Форматирование (идентично clean_osv_sheet_inplace) ───────
        ws_new.column_dimensions["A"].width = 50
        for col_letter in ["B", "C", "D", "E", "F", "G", "H"]:
            ws_new.column_dimensions[col_letter].width = 18
        set_all_rows_height(ws_new, 12)

        # Статус-ячейки (как в clean_osv_sheet_inplace): проверяем сходимость итого
        try:
            expected = itogo[eop_key]
            actual = sum(vals[eop_key] for _, vals in sorted_rows)
            diff_pct = 0.0 if abs(expected) < 1e-9 else abs(actual - expected) / abs(expected)
            pct_col = total_col
            status_col = total_col - 1
            out_cell = ws_new.cell(row=1, column=pct_col)
            out_cell.value = float(diff_pct)
            out_cell.number_format = "0.0%"
            out_cell.font = Font(name="Arial", size=9, bold=True, color="FF4B5563")
            out_cell.alignment = align_c

            status_cell = ws_new.cell(row=1, column=status_col)
            status_cell.font = Font(name="Arial", size=9, bold=True, color="FF111827")
            status_cell.alignment = align_c
            if diff_pct <= 1e-12:
                status_cell.value = "MATCH"
                status_cell.fill = PatternFill(fill_type="solid", fgColor="FFE4F0DD")
            elif diff_pct < 0.10:
                status_cell.value = "MISMATCH"
                status_cell.fill = PatternFill(fill_type="solid", fgColor="FFFFF2CC")
            else:
                status_cell.value = "CHECK"
                status_cell.fill = PatternFill(fill_type="solid", fgColor="FFFCE4EC")
        except Exception:
            pass

        # ── 8. Удаляем исходные листы ───────────────────────────────────
        for sn in valid_sorted:
            if sn in wb.sheetnames:
                del wb[sn]

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# EGOV — вспомогательные функции
# =========================
def find_wt_sheets(file_bytes: bytes) -> List[str]:
    """Возвращает имена листов, которые заканчиваются на 'wt' (регистр не важен)."""
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
        result = [s for s in wb.sheetnames if s.strip().lower().endswith("wt")]
        wb.close()
        return result
    except Exception:
        return []


# ─────────────────────────────────────────────────────────────────────────────
# ADATA — синхронный клиент (токен → поллинг)
# ─────────────────────────────────────────────────────────────────────────────
import time as _time_lib

_ADATA_KEY   = "vBpEhjCB7WmjKoAA5LfbIfM20ZnLYq1"
_ADATA_BASE  = "https://api.adata.kz/api/company"
_ADATA_ROOT  = "https://api.adata.kz/api"
_ADATA_RATE  = 1.0 / 1.5          # мин. интервал между новыми токенами (1.5 req/s)
_adata_session   = _requests_lib.Session()
_adata_cache: Dict[str, dict] = {}
_adata_last_req  = 0.0


def _adata_get_token(module: str, bin_: str, root: bool = False) -> Tuple[Optional[str], Optional[str]]:
    global _adata_last_req
    base = _ADATA_ROOT if root else _ADATA_BASE
    url  = f"{base}/{module}/{_ADATA_KEY}?iinBin={bin_}"
    gap  = _adata_last_req + _ADATA_RATE - _time_lib.monotonic()
    if gap > 0:
        _time_lib.sleep(gap)
    _adata_last_req = _time_lib.monotonic()
    try:
        data = _adata_session.get(url, timeout=30,
                                  headers={"Accept": "application/json"}).json()
    except Exception as e:
        return None, f"error:{e}"
    msg = str(data.get("message") or "")
    if "not found" in msg.lower():
        return None, "not_found"
    token = data.get("token")
    if not token:
        return None, "no_token"
    return token, None


def _adata_poll(token: str, check_url: str = "") -> Tuple[str, Optional[dict]]:
    url = f"{check_url or (_ADATA_BASE + '/info/check')}/{_ADATA_KEY}?token={token}"
    for i in range(20):
        if i:
            _time_lib.sleep(min(i, 6))
        try:
            data = _adata_session.get(url, timeout=30,
                                      headers={"Accept": "application/json"}).json()
        except Exception:
            continue
        msg = str(data.get("message") or "")
        if msg == "ready" and data.get("data") is not None:
            return "ready", data["data"]
        if "not found" in msg.lower():
            return "not_found", None
    return "wait", None


def _adata_fetch(module: str, bin_: str, root: bool = False,
                 check_url: str = "") -> Tuple[str, Optional[dict]]:
    token, err = _adata_get_token(module, bin_, root=root)
    if err:
        return err, None
    return _adata_poll(token, check_url=check_url)


def _adata_fetch_taxes(bin_: str) -> Tuple[str, list]:
    """Помесячные налоги за текущий и прошлый год."""
    from datetime import date as _date
    token, err = _adata_get_token("tax-deduction/extended", bin_)
    if err:
        return err, []
    today  = _date.today()
    years  = [today.year - 1, today.year]
    rows   = []
    for year in years:
        page, last_page = 1, 1
        while page <= last_page:
            url = f"{_ADATA_BASE}/info/check/{_ADATA_KEY}?token={token}&page={page}&year={year}"
            got = False
            for i in range(20):
                if i:
                    _time_lib.sleep(min(i, 6))
                try:
                    data = _adata_session.get(url, timeout=30,
                                              headers={"Accept": "application/json"}).json()
                except Exception:
                    continue
                msg = str(data.get("message") or "")
                if msg == "ready" and data.get("data") is not None:
                    d = data["data"]
                    rows.extend(d.get("details") or [])
                    last_page = d.get("last_page", last_page)
                    got = True
                    break
                if "not found" in msg.lower():
                    return "not_found", []
            if not got:
                return "wait", rows
            page += 1
    return "ready", rows


def _adata_tax_windows(rows: list) -> Tuple[float, float, float]:
    """Считает L12M / L6M / L3M из помесячных налоговых платежей."""
    from datetime import date as _date
    today  = _date.today()
    monthly: Dict[Tuple[int,int], float] = {}
    for it in rows:
        if it.get("pay_code") != "tax" or it.get("entry_code") != "payment":
            continue
        rd  = it.get("receipt_date")
        amt = it.get("amount")
        if not rd or amt in (None, ""):
            continue
        parts = str(rd).split("-")
        if len(parts) < 2:
            continue
        try:
            y, m = int(parts[0]), int(parts[1])
        except ValueError:
            continue
        monthly[(y, m)] = monthly.get((y, m), 0.0) + float(amt)

    end_y, end_m = today.year, today.month - 1
    if end_m == 0:
        end_m, end_y = 12, today.year - 1

    def _win(months: int) -> float:
        sm, sy = end_m - months + 1, end_y
        while sm <= 0:
            sm += 12; sy -= 1
        total = 0.0
        y, m = sy, sm
        while (y, m) <= (end_y, end_m):
            total += monthly.get((y, m), 0.0)
            m += 1
            if m > 12:
                m, y = 1, y + 1
        return total

    return _win(12), _win(6), _win(3)


def _adata_risk_flags(info: dict, tax_rows: list) -> dict:
    """Извлекает флаги риска и вердикт из данных ADATA."""
    from datetime import date as _date

    def gp(path: str):
        cur = info
        for k in path.split("."):
            if isinstance(cur, dict) and k in cur:
                cur = cur[k]
            else:
                return None
        return cur

    def num(v) -> float:
        try: return float(v or 0)
        except: return 0.0

    def b(v) -> bool:
        if isinstance(v, bool): return v
        if isinstance(v, (int, float)): return bool(v)
        if isinstance(v, str): return v.strip().lower() == "true"
        return False

    today = _date.today()
    flags = []   # (цвет, текст)

    # ── Критичные (красные) ──────────────────────────────────────────────────
    if b(gp("riskFactor.company.bankrupt")):
        flags.append(("R", "Банкрот"))
    if b(gp("riskFactor.company.liquidating_taxpayer")):
        flags.append(("R", "В ликвидации"))
    if b(gp("trustworthyExtended.finished")):
        flags.append(("R", "Банкротство завершено"))
    if not b(gp("status.company_status")):
        flags.append(("R", "Недействующее"))

    tax_debt = num(gp("status.tax_debt"))
    if tax_debt > 50_000_000:
        flags.append(("R", f"Налог.долг {tax_debt/1e6:.1f} млн"))
    elif tax_debt > 500_000:
        flags.append(("Y", f"Налог.долг {tax_debt/1e6:.1f} млн"))

    enf = num(gp("riskFactor.company.enforcement_debt_sum"))
    if enf > 25_000_000:
        flags.append(("R", f"Исполы {enf/1e6:.1f} млн"))
    elif enf > 250_000:
        flags.append(("Y", f"Исполы {enf/1e6:.1f} млн"))
    elif b(gp("riskFactor.company.enforcement_debt")):
        flags.append(("Y", "Есть исполы"))

    # ── Мягкие (жёлтые) ─────────────────────────────────────────────────────
    if b(gp("riskFactor.company.seized_bank_account")):
        flags.append(("Y", "Арест счетов"))
    if b(gp("riskFactor.company.seized_property")):
        flags.append(("Y", "Арест имущества"))

    dir_enf = num(gp("riskFactor.head.enforcement_debt_sum"))
    if dir_enf > 25_000_000:
        flags.append(("R", f"Долг директора {dir_enf/1e6:.1f} млн"))
    elif dir_enf > 250_000:
        flags.append(("Y", f"Долг директора {dir_enf/1e6:.1f} млн"))

    reg_str = gp("basic.date_registration")
    if reg_str:
        try:
            p = str(reg_str).split("-")
            reg = _date(int(p[0]), int(p[1]), int(p[2]))
            yr_ago = today.replace(year=today.year - 1) if today.month > 2 else \
                     today.replace(year=today.year - 1, day=28)
            if reg > yr_ago:
                flags.append(("Y", "Моложе года"))
        except Exception:
            pass

    l12, l6, l3 = _adata_tax_windows(tax_rows)
    if 0 < l12 < 10_000_000:
        flags.append(("Y", f"LTM {l12/1e6:.1f} млн"))
    if l12 > 0 and l6 < l12 * 0.10:
        flags.append(("R", "LSM < 10% LTM"))
    elif l12 > 0 and l6 < l12 * 0.25:
        flags.append(("Y", "LSM < 25% LTM"))

    rd = gp("riskFactor.company.tax_risk_degree") or ""
    if rd and str(rd).lower() not in ("низкий", "low", ""):
        flags.append(("Y", f"Налог.риск: {rd}"))

    has_r = any(f[0] == "R" for f in flags)
    has_y = any(f[0] == "Y" for f in flags)
    verdict = "🔴 Критично" if has_r else ("🟡 Риски" if has_y else "🟢 Норма")
    flag_str = "; ".join(("🔴 " if f[0]=="R" else "🟡 ") + f[1] for f in flags)

    return {
        "verdict":    verdict,
        "flags":      flag_str,
        "tax_l12m":   round(l12 / 1e6, 2),
        "tax_l6m":    round(l6  / 1e6, 2),
        "tax_l3m":    round(l3  / 1e6, 2),
        "tax_debt":   tax_debt,
        "enf_sum":    enf,
        "name":       gp("basic.short_name") or gp("basic.name_ru") or "",
        "director":   gp("basic.fullname_director") or "",
        "employee_n": int(num(gp("basic.employee_count"))),
        "oked":       gp("basic.oked") or "",
        "active":     b(gp("status.company_status")),
        "is_ip":      gp("basic.counterparty_type") in ("enterprise", "individual"),
    }


def _adata_company_summary(bin_: str) -> dict:
    """Полная оценка компании через ADATA. Результат кэшируется на сессию."""
    bin_ = str(bin_ or "").strip()
    if len(bin_) < 10:
        return {"error": "Нет БИН"}
    if bin_ in _adata_cache:
        return _adata_cache[bin_]

    status, info = _adata_fetch("info", bin_)
    if status != "ready" or not info:
        result = {"error": f"ADATA: {status}"}
        _adata_cache[bin_] = result
        return result

    is_ip = info.get("basic", {}).get("counterparty_type") in ("enterprise", "individual")
    tax_rows: list = []
    if not is_ip:
        _, tax_rows = _adata_fetch_taxes(bin_)

    risk = _adata_risk_flags(info, tax_rows)
    result = {"found": True, "bin": bin_, **risk}
    _adata_cache[bin_] = result
    return result


def _adata_affiliations_by_type(bin_: str) -> Tuple[str, Optional[dict]]:
    """Аффилированные лица по типу (director/founder/old_director/old_founder/phone).

    Эндпоинт: relation-scheme/affiliation-by-type
    Check URL: relation-scheme/check  (свой, не общий info/check)
    """
    cache_key = f"affil_{bin_}"
    if cache_key in _adata_cache:
        return "ready", _adata_cache[cache_key]
    check_url = f"{_ADATA_BASE}/relation-scheme/check"
    status, data = _adata_fetch("relation-scheme/affiliation-by-type", bin_,
                                check_url=check_url)
    if status == "ready" and data:
        _adata_cache[cache_key] = data
    return status, data


def _collect_affil_bins(data: dict) -> Dict[str, Tuple[str, str]]:
    """Из ответа affiliation-by-type собирает {biin: (тип_связи, имя)}."""
    TYPE_LABELS = {
        "director":     "Директор",
        "founder":      "Учредитель",
        "old_director": "Бывший директор",
        "old_founder":  "Бывший учредитель",
        "phone":        "Общий телефон",
    }
    bins: Dict[str, Tuple[str, str]] = {}
    for key, label in TYPE_LABELS.items():
        for item in (data.get(key) or []):
            biin = str(item.get("biin") or "").strip()
            if not biin or len(biin) < 10:
                continue
            name = str(item.get("name") or "").strip()
            if biin in bins:
                prev_label = bins[biin][0]
                if label not in prev_label:
                    bins[biin] = (prev_label + " / " + label, name)
            else:
                bins[biin] = (label, name)
    return bins


def _adata_director_companies(iinbin: str) -> Tuple[str, List[dict]]:
    """Компании, в которых iinbin является директором/учредителем.

    Эндпоинт: company/director-in-companies   Check: info/check (стандартный).
    Поддерживает пагинацию через ?page=N.
    """
    cache_key = f"dir_co_{iinbin}"
    if cache_key in _adata_cache:
        return "ready", _adata_cache[cache_key]
    token, err = _adata_get_token("director-in-companies", iinbin)
    if err:
        return err, []
    all_details: List[dict] = []
    page = 1
    last_page = 1
    check_base = f"{_ADATA_BASE}/info/check/{_ADATA_KEY}"
    while page <= last_page:
        url = f"{check_base}?token={token}&page={page}"
        got = False
        for i in range(20):
            if i:
                _time_lib.sleep(min(i, 6))
            try:
                resp = _adata_session.get(url, timeout=30,
                                          headers={"Accept": "application/json"}).json()
            except Exception:
                continue
            msg = str(resp.get("message") or "")
            if msg == "ready" and resp.get("data") is not None:
                d = resp["data"]
                all_details.extend(d.get("details") or [])
                last_page = int(d.get("last_page") or last_page)
                got = True
                break
            if "not found" in msg.lower():
                return "not_found", []
        if not got:
            break
        page += 1
    if all_details:
        _adata_cache[cache_key] = all_details
    return "ready", all_details


def _extract_client_bin(filename: str) -> Optional[str]:
    """Извлекает 12-значный БИН из имени файла (_анализ_123456789012.xlsx)."""
    m = re.search(r'\d{12}', filename or "")
    return m.group(0) if m else None


def read_wt_counterparties_with_amounts(
        file_bytes: bytes, sheet_name: str) -> List[Tuple[str, str, float]]:
    """Читает wt-лист: возвращает (таблица, контрагент, сумма_AP).

    Колонка B — название контрагента.
    Колонка AP (42) — оборот за весь период.
    """
    _SUMMARY  = {'топ', 'всего', 'доля'}
    _SKIP_TBL = re.compile(r'inflow|outflow', re.IGNORECASE)
    _SKIP_VAL = re.compile(
        r'^[\s\-–—\.,:;/\\0-9]+$'
        r'|^(?:н/?д|нет\s+данных?|не\s+определ|прочи[ех]?|разные?|итого)$',
        re.IGNORECASE,
    )
    results: List[Tuple[str, str, float]] = []
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb[sheet_name]
        current_table: Optional[str] = None
        in_table = False
        for row in ws.iter_rows(min_col=2, max_col=42, values_only=True):
            name_val = row[0]    # col B
            ap_val   = row[40]   # col AP (offset 40 from col B)
            if name_val is None:
                continue
            s = str(name_val).strip()
            if not s:
                continue
            s_low = s.lower()
            if s_low in _SUMMARY:
                if s_low == 'топ':
                    in_table = False
                    current_table = None
                continue
            if not in_table:
                current_table = s
                in_table = True
            else:
                if current_table and _SKIP_TBL.search(current_table):
                    continue
                if _SKIP_VAL.match(s):
                    continue
                if current_table:
                    try:
                        amount = float(ap_val) if ap_val is not None else 0.0
                    except (TypeError, ValueError):
                        amount = 0.0
                    results.append((current_table, s, amount))
        wb.close()
    except Exception:
        pass
    return results


def run_related_companies_sheet(
        file_bytes: bytes, wt_sheets: List[str], filename: str) -> Tuple[bytes, str]:
    """Модуль "Связанные компании".

    1. Извлекает БИН клиента из имени файла.
    2. Параллельно: ищет БИНы контрагентов (ГБД ЮЛ) + запрашивает аффилированных (ADATA).
    3. Создаёт лист "связанные" всегда — со всеми контрагентами и их статусом.
    Возвращает (file_bytes, статус_сообщение).
    """
    import concurrent.futures
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    client_bin = _extract_client_bin(filename)

    # Собираем контрагентов со всех wt-листов (поддержка нескольких wt с префиксами)
    all_cps: List[Tuple[str, str, str, float]] = []  # (wt_sh, table, name, amount)
    for wt_sh in wt_sheets:
        for tbl, name, amt in read_wt_counterparties_with_amounts(file_bytes, wt_sh):
            all_cps.append((wt_sh, tbl, name, amt))

    # ── Параллельные задачи ──────────────────────────────────────────────────
    def _lookup_bins() -> Dict[str, dict]:
        out: Dict[str, dict] = {}
        seen: set = set()
        for _, _, name, _ in all_cps:
            if name in seen:
                continue
            seen.add(name)
            eg = _egov_company_summary(name)
            out[name] = eg
        return out

    def _lookup_affil() -> Tuple[Dict[str, Tuple[str, str]], Optional[dict]]:
        if not client_bin:
            return {}, None
        status, data = _adata_affiliations_by_type(client_bin)
        if status == "ready" and data:
            return _collect_affil_bins(data), data
        return {}, None

    def _lookup_dir_cos() -> List[dict]:
        if not client_bin:
            return []
        _, details = _adata_director_companies(client_bin)
        return details or []

    if all_cps and client_bin:
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as ex:
            f_bins    = ex.submit(_lookup_bins)
            f_affil   = ex.submit(_lookup_affil)
            f_dir_cos = ex.submit(_lookup_dir_cos)
            name_to_eg            = f_bins.result()
            affil_bins, affil_raw = f_affil.result()
            dir_cos_raw: List[dict] = f_dir_cos.result()
    elif all_cps:
        name_to_eg = _lookup_bins()
        affil_bins, affil_raw = {}, None
        dir_cos_raw = []
    else:
        name_to_eg  = {}
        affil_bins, affil_raw = {}, None
        dir_cos_raw = []

    # Мержим director-in-companies в affil_bins (дедупликация по BIN)
    DIR_CO_LABEL = "Директор/учредитель в компании"
    for item in dir_cos_raw:
        b = str(item.get("bin") or "").strip()
        if not b:
            continue
        n = str(item.get("short_name") or "").strip()
        if b in affil_bins:
            prev_lbl, prev_nm = affil_bins[b]
            if DIR_CO_LABEL not in prev_lbl:
                affil_bins[b] = (prev_lbl + " / " + DIR_CO_LABEL, prev_nm or n)
        else:
            affil_bins[b] = (DIR_CO_LABEL, n)

    # ── Строим полную таблицу ────────────────────────────────────────────────
    # дедупликация: один контрагент может встречаться в нескольких wt
    rows_affiliated = []
    rows_other      = []
    seen_key: set   = set()

    for wt_sh, tbl, name, amount in all_cps:
        key = (wt_sh, tbl, name)
        if key in seen_key:
            continue
        seen_key.add(key)

        eg        = name_to_eg.get(name, {})
        # БИН берём только при confirmed match (found=True), иначе ES мог вернуть ложный результат
        cp_bin    = str(eg.get("bin", "") or "").strip() if eg.get("found") else ""
        bin_found = bool(cp_bin)
        is_affil  = bin_found and cp_bin in affil_bins
        rel_type  = affil_bins[cp_bin][0] if is_affil else ""
        affil_nm  = affil_bins[cp_bin][1] if is_affil else ""
        egov_name = str(eg.get("nameru", "") or "").strip() if eg.get("found") else ""

        rec = {
            "Лист wt":         wt_sh,
            "Таблица":         tbl,
            "Контрагент (wt)": name,
            "Оборот (AP)":     amount,
            "БИН":             cp_bin,
            "Найдено в ГБД":   egov_name,
            "Связанный?":      "ДА" if is_affil else "",
            "Тип связи":       rel_type,
            "Имя в ADATA":     affil_nm,
        }
        if is_affil:
            rows_affiliated.append(rec)
        else:
            rows_other.append(rec)

    all_rows = rows_affiliated + rows_other
    n_affil  = len(rows_affiliated)

    # ── Создаём лист "связанные" ─────────────────────────────────────────────
    wb = load_workbook(io.BytesIO(file_bytes))
    if "связанные" in wb.sheetnames:
        del wb["связанные"]
    ws = wb.create_sheet("связанные")

    # Стили
    hdr_font   = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    hdr_fill   = PatternFill(fill_type="solid", fgColor="FF0D0D0D")
    body_font  = Font(name="Calibri", size=10)
    bold_font  = Font(name="Calibri", size=10, bold=True)
    acc_fill   = PatternFill(fill_type="solid", fgColor="FFFF6600")  # orange — связанный
    dim_fill   = PatternFill(fill_type="solid", fgColor="FFF5F5F5")  # светлый — прочие
    meta_font  = Font(name="Calibri", size=10, italic=True, color="FF666666")
    ok_font    = Font(name="Calibri", size=10, bold=True, color="FF000000")
    thin       = Side(style="thin", color="FFCCCCCC")
    border     = Border(bottom=thin)
    align_c    = Alignment(horizontal="center", vertical="center")
    align_l    = Alignment(horizontal="left",   vertical="center")

    # Строка 1: сводка
    ws.merge_cells("A1:I1")
    summary_parts = [f"БИН клиента: {client_bin or 'не определён'}"]
    summary_parts.append(f"wt-листов: {len(wt_sheets)}")
    summary_parts.append(f"контрагентов: {len(all_rows)}")
    summary_parts.append(f"связанных: {n_affil}")
    summary_parts.append(f"директор в компаниях: {len(dir_cos_raw)}")
    if not client_bin:
        summary_parts.append("⚠ БИН не найден в имени файла")
    c = ws["A1"]
    c.value = "   |   ".join(summary_parts)
    c.font  = meta_font
    c.alignment = align_l
    ws.row_dimensions[1].height = 18

    # Строка 2: заголовки
    HEADERS = ["Лист wt", "Таблица", "Контрагент (wt)", "Оборот (AP)",
               "БИН", "Найдено в ГБД", "Связанный?", "Тип связи", "Имя в ADATA"]
    for ci, h in enumerate(HEADERS, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = align_c

    # Данные начиная с row 3
    for ri, rec in enumerate(all_rows, 3):
        is_affil_row = bool(rec["Связанный?"])
        row_fill = acc_fill if is_affil_row else dim_fill
        for ci, h in enumerate(HEADERS, 1):
            val = rec[h]
            c   = ws.cell(row=ri, column=ci, value=val)
            c.border = border
            c.alignment = align_c if h in ("Лист wt", "Оборот (AP)", "БИН", "Связанный?", "Тип связи") else align_l
            if is_affil_row:
                c.fill = row_fill
                c.font = ok_font
            else:
                c.font = body_font
            if h == "Оборот (AP)" and isinstance(val, (int, float)):
                c.number_format = "#,##0"
        # Контрагент col всегда bold
        ws.cell(row=ri, column=3).font = bold_font if not is_affil_row else ok_font

    if not all_rows:
        ws.cell(row=3, column=1, value="Контрагенты не найдены на wt-листах").font = meta_font

    COL_WIDTHS = [14, 32, 46, 16, 16, 46, 14, 22, 46]
    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[ws.cell(row=2, column=ci).column_letter].width = w
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    # ── Вторая секция: все аффилированные из ADATA (raw JSON) ────────────────
    TYPE_ORDER = [
        ("director",     "Директор"),
        ("founder",      "Учредитель"),
        ("old_director", "Бывший директор"),
        ("old_founder",  "Бывший учредитель"),
        ("phone",        "Общий телефон"),
    ]
    sec2_start = len(all_rows) + 5  # через 2 пустые строки после основной таблицы

    sec2_hdr_fill = PatternFill(fill_type="solid", fgColor="FF1A1A1A")
    sec2_grp_fill = PatternFill(fill_type="solid", fgColor="FF2D2D2D")
    sec2_grp_font = Font(name="Calibri", size=10, bold=True, color="FFFF6600")
    sec2_body_font = Font(name="Calibri", size=10, color="FF111111")

    # Заголовок секции
    ws.merge_cells(f"A{sec2_start}:D{sec2_start}")
    c = ws.cell(row=sec2_start, column=1,
                value=f"Аффилированные лица клиента (ADATA) — БИН {client_bin or '?'}")
    c.font      = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    c.fill      = sec2_hdr_fill
    c.alignment = align_l
    ws.row_dimensions[sec2_start].height = 16

    # Заголовки колонок
    SEC2_HDRS = ["Тип связи", "БИН / ИИН", "Имя", "Доп. инфо"]
    hdr_row = sec2_start + 1
    for ci, h in enumerate(SEC2_HDRS, 1):
        c = ws.cell(row=hdr_row, column=ci, value=h)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = align_c
    ws.row_dimensions[hdr_row].height = 16

    cur_row = hdr_row + 1
    if affil_raw:
        for key, label in TYPE_ORDER:
            items = affil_raw.get(key) or []
            if not items:
                continue
            # Строка-группа с названием типа
            ws.merge_cells(f"A{cur_row}:D{cur_row}")
            gc = ws.cell(row=cur_row, column=1, value=f"▸  {label}  ({len(items)})")
            gc.font      = sec2_grp_font
            gc.fill      = sec2_grp_fill
            gc.alignment = align_l
            cur_row += 1
            for item in items:
                biin  = str(item.get("biin") or item.get("bin") or "").strip()
                name  = str(item.get("name") or item.get("nameru") or "").strip()
                # Собираем доп. поля (все кроме biin/bin/name/nameru)
                extra_parts = []
                for fk, fv in item.items():
                    if fk.lower() in ("biin", "bin", "name", "nameru"):
                        continue
                    if fv not in (None, "", [], {}):
                        extra_parts.append(f"{fk}: {fv}")
                extra = "; ".join(extra_parts)

                row_vals = [label, biin, name, extra]
                for ci, val in enumerate(row_vals, 1):
                    c = ws.cell(row=cur_row, column=ci, value=val)
                    c.font      = sec2_body_font
                    c.border    = border
                    c.alignment = align_c if ci in (1, 2) else align_l
                cur_row += 1
    else:
        ws.cell(row=cur_row, column=1,
                value="Данные от ADATA не получены").font = meta_font

    # Ширины колонок для секции 2 (A-D)
    for ci, w in zip(range(1, 5), [22, 18, 50, 60]):
        ws.column_dimensions[ws.cell(row=hdr_row, column=ci).column_letter].width = max(
            ws.column_dimensions[ws.cell(row=hdr_row, column=ci).column_letter].width or 0, w)

    # ── Секция 3: Компании, в которых клиент является директором/учредителем ─
    sec3_start = cur_row + 2
    ws.merge_cells(f"A{sec3_start}:E{sec3_start}")
    c3h = ws.cell(row=sec3_start, column=1,
                  value=f"Компании, где клиент — директор/учредитель (ADATA director-in-companies) — БИН {client_bin or '?'}")
    c3h.font      = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    c3h.fill      = PatternFill(fill_type="solid", fgColor="FF1A1A1A")
    c3h.alignment = align_l
    ws.row_dimensions[sec3_start].height = 16

    SEC3_HDRS = ["БИН", "Название", "Активна?", "Надёжная?", "Дата окончания"]
    s3_hdr_row = sec3_start + 1
    for ci, h in enumerate(SEC3_HDRS, 1):
        c = ws.cell(row=s3_hdr_row, column=ci, value=h)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = align_c
    ws.row_dimensions[s3_hdr_row].height = 16

    s3_cur = s3_hdr_row + 1
    if dir_cos_raw:
        for item in dir_cos_raw:
            b    = str(item.get("bin") or "").strip()
            nm   = str(item.get("short_name") or "").strip()
            inactive = item.get("is_inactive")
            trust    = item.get("trustworthy")
            end_dt   = str(item.get("end_date") or "").strip()
            active_str = "нет" if inactive else "да"
            trust_str  = "да" if trust else ("нет" if trust is False else "")
            row_vals = [b, nm, active_str, trust_str, end_dt]
            is_inactive_row = bool(inactive)
            for ci, val in enumerate(row_vals, 1):
                c = ws.cell(row=s3_cur, column=ci, value=val)
                c.font      = Font(name="Calibri", size=10,
                                   color="FF888888" if is_inactive_row else "FF111111")
                c.border    = border
                c.alignment = align_c if ci in (1, 3, 4, 5) else align_l
            s3_cur += 1
    else:
        ws.cell(row=s3_cur, column=1,
                value="Данные от ADATA не получены или клиент не зарегистрирован как директор").font = meta_font

    for ci, w in zip(range(1, 6), [18, 52, 12, 14, 18]):
        col_letter = ws.cell(row=s3_hdr_row, column=ci).column_letter
        ws.column_dimensions[col_letter].width = max(
            ws.column_dimensions[col_letter].width or 0, w)

    out = io.BytesIO()
    wb.save(out)
    msg = (f"Связанные: {n_affil} из {len(all_rows)} контрагентов, "
           f"директор в {len(dir_cos_raw)} компаниях "
           f"(БИН клиента: {client_bin or 'не определён'})")
    return out.getvalue(), msg


# ─────────────────────────────────────────────────────────────────────────────

def read_wt_all_counterparties(file_bytes: bytes, sheet_name: str) -> List[Tuple[str, str]]:
    """
    Читает все таблицы с листа wt.
    Возвращает список (имя_таблицы, название_контрагента).

    Структура листа wt:
      - Строка-заголовок таблицы (col B)
      - Строки с контрагентами (col B)
      - "Топ" — маркер конца, не включаем
      - "Всего", "Доля" — служебные строки, пропускаем
      - Пустые строки — разделители

    Пропускаем таблицы, содержащие "Inflow" или "Outflow" в названии
    (они дубликаты Прочая_ДЗ/КЗ).
    """
    _SUMMARY = {'топ', 'всего', 'доля'}
    _SKIP_TBL = re.compile(r'inflow|outflow', re.IGNORECASE)
    _SKIP_VAL = re.compile(
        r'^[\s\-–—\.,:;/\\0-9]+$'
        r'|^(?:н/?д|нет\s+данных?|не\s+определ|прочи[ех]?|разные?|итого)$',
        re.IGNORECASE,
    )
    results: List[Tuple[str, str]] = []
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        ws = wb[sheet_name]
        current_table: Optional[str] = None
        in_table: bool = False
        for row in ws.iter_rows(min_col=2, max_col=2, values_only=True):
            raw = row[0]
            if raw is None:
                continue
            s = str(raw).strip()
            if not s:
                continue
            s_low = s.lower()
            if s_low in _SUMMARY:
                if s_low == 'топ':
                    in_table = False
                    current_table = None
                continue
            if not in_table:
                current_table = s
                in_table = True
            else:
                if current_table and _SKIP_TBL.search(current_table):
                    continue
                if _SKIP_VAL.match(s):
                    continue
                if current_table:
                    results.append((current_table, s))
        wb.close()
    except Exception:
        pass
    return results


def run_egov_clients_sheet(file_bytes: bytes, wt_sheets: List[str]) -> bytes:
    """
    Для каждого листа из wt_sheets читает все таблицы контрагентов,
    делает lookup через egov, создаёт/обновляет лист 'контрагенты'.
    Если листов >1 — добавляет столбец 'Компания' (из w-листа, ячейка B2).
    Всегда добавляет столбец 'Таблица' (откуда взят контрагент на листе wt).
    """
    wb = load_workbook(io.BytesIO(file_bytes))
    multi = len(wt_sheets) > 1

    def _company_name(wt_sh: str) -> str:
        prefix, _ = split_prefix_suffix2(wt_sh)
        target = (prefix + "w").lower()
        for sh in wb.sheetnames:
            if sh.lower() == target:
                try:
                    val = wb[sh].cell(row=2, column=2).value
                    return str(val).strip() if val else wt_sh
                except Exception:
                    pass
        return wt_sh

    # Собираем все строки: (company, table_name, counterparty_name)
    all_rows: List[Tuple[str, str, str]] = []
    for wt_sh in wt_sheets:
        company = _company_name(wt_sh) if multi else ""
        pairs = read_wt_all_counterparties(file_bytes, wt_sh)
        for tbl_name, cp_name in pairs:
            all_rows.append((company, tbl_name, cp_name))

    if not all_rows:
        return file_bytes

    for _dead in ("контрагенты", "клиенты"):
        if _dead in wb.sheetnames:
            del wb[_dead]
    ws_k = wb.create_sheet("контрагенты")

    # --- Стили ---
    hdr_font  = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    hdr_fill  = PatternFill("solid", fgColor="1A1A1A")
    cell_font = Font(name="Calibri", size=10)
    muted_font = Font(name="Calibri", size=10, color="888888", italic=True)
    thin      = Side(border_style="thin", color="333333")
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_c   = Alignment(horizontal="center", vertical="center", wrap_text=False)
    align_l   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    if multi:
        headers = [
            ("Компания",                26),
            ("Таблица",                 18),
            ("Контрагент (wt)",         30),
            ("Полное название (ГБД ЮЛ)", 44),
            ("БИН",                     14),
            ("Дата регистрации",        16),
            ("Директор",                34),
        ]
        COL_COMPANY  = 1
        COL_TABLE    = 2
        COL_CONTR    = 3
        COL_FULLNAME = 4
        COL_BIN      = 5
        COL_DATE     = 6
        COL_DIRECTOR = 7
    else:
        headers = [
            ("Таблица",                 18),
            ("Контрагент (wt)",         30),
            ("Полное название (ГБД ЮЛ)", 44),
            ("БИН",                     14),
            ("Дата регистрации",        16),
            ("Директор",                34),
        ]
        COL_COMPANY  = 0   # нет
        COL_TABLE    = 1
        COL_CONTR    = 2
        COL_FULLNAME = 3
        COL_BIN      = 4
        COL_DATE     = 5
        COL_DIRECTOR = 6

    for ci, (h, w) in enumerate(headers, 1):
        cell = ws_k.cell(row=1, column=ci, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = border
        cell.alignment = align_c
        ws_k.column_dimensions[get_column_letter(ci)].width = w
    ws_k.row_dimensions[1].height = 18

    left_cols = {c for c in (COL_COMPANY, COL_TABLE, COL_CONTR, COL_FULLNAME, COL_DIRECTOR) if c}

    for ri, (company, tbl_name, name) in enumerate(all_rows, 2):
        info = _egov_company_summary(name)
        ws_k.row_dimensions[ri].height = 16

        def _c(col: int, val, fnt=None, _ri=ri):
            cell = ws_k.cell(row=_ri, column=col, value=val)
            cell.font = fnt or cell_font
            cell.border = border
            cell.alignment = align_l if col in left_cols else align_c
            return cell

        if COL_COMPANY:
            _c(COL_COMPANY, company)
        _c(COL_TABLE,  tbl_name)
        _c(COL_CONTR,  name)

        if info.get("found"):
            _c(COL_FULLNAME, info.get("nameru", ""))
            _c(COL_BIN,      info.get("bin", ""))
            _c(COL_DATE,     info.get("datereg", ""))
            _c(COL_DIRECTOR, info.get("director", ""))
        else:
            note = info.get("note") or "Не найдено в ГБД ЮЛ"
            _c(COL_FULLNAME, note, muted_font)

    ws_k.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="CF app", page_icon=None, layout="wide", initial_sidebar_state="collapsed")

# Bloomberg Terminal theme
BG     = "#000000"
CARD   = "#0A0A0A"
TEXT   = "#E8E8E8"
MUTED  = "#888888"
ACCENT = "#FF6600"       # Bloomberg orange
ACCENT_2 = "#FFB347"     # amber highlight
BORDER = "#2A2A2A"
BTN_BG    = "#1A1A1A"
BTN_TEXT  = ACCENT
BTN_BORDER = "#FF6600"
BTN_HOVER  = "#2A2A2A"

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;700&display=swap');

    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header {{visibility: hidden;}}

    .stApp {{
        background: {BG};
        color: {TEXT};
    }}

    html, body, [class*="css"], .stApp {{
        font-size: 13px !important;
        font-family: "IBM Plex Mono", "Consolas", monospace !important;
    }}

    button, input, textarea, select, option,
    label,
    [data-baseweb],
    [data-baseweb] *,
    [data-testid] * {{
        font-family: "IBM Plex Mono", "Consolas", monospace !important;
    }}

    .block-container {{
        max-width: 1160px;
        padding-top: 1rem;
        padding-bottom: 1rem;
    }}

    /* Убираем все скругления — Bloomberg всё квадратное */
    *, *::before, *::after {{
        border-radius: 0 !important;
    }}

    /* Заголовки секций — оранжевая левая полоса */
    .stMarkdown h4 {{
        font-size: 11px !important;
        font-weight: 700 !important;
        letter-spacing: 1.5px !important;
        text-transform: uppercase !important;
        color: {ACCENT} !important;
        font-family: "IBM Plex Mono", "Consolas", monospace !important;
        border-left: 3px solid {ACCENT};
        padding-left: 8px;
        margin: 0.8rem 0 0.5rem 0 !important;
    }}

    /* File uploader */
    [data-testid="stFileUploader"] section,
    [data-testid="stForm"],
    .stExpander {{
        border-radius: 0 !important;
        border: 1px solid {BORDER} !important;
    }}

    [data-testid="stFileUploader"] section {{
        padding: 12px;
        background: {CARD};
        box-shadow: none;
    }}
    [data-testid="stFileUploader"] * {{
        color: {TEXT} !important;
    }}
    [data-testid="stFileUploader"] small {{
        color: {MUTED} !important;
    }}

    /* Кнопки */
    div.stButton > button,
    div.stDownloadButton > button,
    button[kind="primary"],
    button[kind="secondary"],
    div[data-baseweb="button"] > button,
    div[data-baseweb="button"] button {{
        background: {BTN_BG} !important;
        color: {BTN_TEXT} !important;
        border: 1px solid {BTN_BORDER} !important;
        border-radius: 0 !important;
        padding: 0.55rem 0.90rem !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
        text-transform: uppercase !important;
        box-shadow: none !important;
        transition: background 100ms ease !important;
    }}
    div.stButton > button:hover,
    div.stDownloadButton > button:hover,
    button[kind="primary"]:hover,
    button[kind="secondary"]:hover,
    div[data-baseweb="button"] > button:hover,
    div[data-baseweb="button"] button:hover {{
        background: {ACCENT} !important;
        color: #000000 !important;
    }}
    div.stButton > button:focus-visible,
    div.stDownloadButton > button:focus-visible {{
        outline: none !important;
        box-shadow: 0 0 0 2px {ACCENT} !important;
    }}

    [data-testid="stFileUploader"] button,
    [data-testid="stFileUploaderDropzone"] button {{
        background: {BTN_BG} !important;
        color: {BTN_TEXT} !important;
        border: 1px solid {BTN_BORDER} !important;
        border-radius: 0 !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
        text-transform: uppercase !important;
    }}
    [data-testid="stFileUploader"] button:hover,
    [data-testid="stFileUploaderDropzone"] button:hover {{
        background: {ACCENT} !important;
        color: #000 !important;
    }}

    [data-testid="stFileUploaderDeleteBtn"] button {{
        background: {CARD} !important;
        color: {MUTED} !important;
        border: 1px solid {BORDER} !important;
        border-radius: 0 !important;
    }}
    [data-testid="stFileUploaderDeleteBtn"] svg {{
        fill: {MUTED} !important;
        color: {MUTED} !important;
    }}

    input, textarea, [data-baseweb="input"] input {{
        background: {CARD} !important;
        color: {TEXT} !important;
        border-color: {BORDER} !important;
        border-radius: 0 !important;
    }}
    input::placeholder, textarea::placeholder {{
        color: {MUTED} !important;
    }}
    input:focus, textarea:focus, [data-baseweb="input"] input:focus {{
        outline: none !important;
        box-shadow: 0 0 0 2px {ACCENT} !important;
        border-color: {ACCENT} !important;
    }}

    div.stButton > button, div.stDownloadButton > button {{
        width: 100%;
        font-size: 13px !important;
    }}

    .stMarkdown, .stMarkdown p, label {{
        color: {TEXT} !important;
    }}
    .stCaption, [data-testid="stCaptionContainer"] p {{
        color: {MUTED} !important;
        font-size: 11px !important;
    }}

    [data-testid="stCheckbox"] * {{
        color: {TEXT} !important;
    }}
    input[type="checkbox"], input[type="radio"] {{
        accent-color: {ACCENT} !important;
    }}

    /* Горизонтальная линия — тонкая оранжевая */
    hr {{
        border: none !important;
        border-top: 1px solid {ACCENT} !important;
        margin: 0.6rem 0 !important;
        opacity: 0.4 !important;
    }}

    /* Алерты */
    [data-testid="stAlert"] {{
        border-radius: 0 !important;
        border: 1px solid {BORDER} !important;
        border-left: 3px solid {ACCENT} !important;
        box-shadow: none !important;
        background: {CARD} !important;
    }}

    /* Spinner */
    [data-testid="stSpinner"] * {{
        color: {ACCENT} !important;
    }}

    /* Progress bar */
    [data-testid="stProgressBar"] > div > div {{
        background: {ACCENT} !important;
    }}
    [data-testid="stProgressBar"] > div {{
        background: #1A1A1A !important;
        border-radius: 0 !important;
        height: 6px !important;
    }}

    /* Select / Multiselect */
    [data-baseweb="select"] > div,
    [data-baseweb="select"] div[class*="ValueContainer"],
    [data-baseweb="popover"] ul,
    [data-baseweb="menu"] {{
        background: {CARD} !important;
        border-color: {BORDER} !important;
        border-radius: 0 !important;
        color: {TEXT} !important;
    }}
    [data-baseweb="select"] span,
    [data-baseweb="select"] div {{
        color: {TEXT} !important;
    }}
    [data-baseweb="option"]:hover,
    [data-baseweb="option"][aria-selected="true"] {{
        background: #1A0800 !important;
        color: {ACCENT} !important;
    }}
    [data-baseweb="tag"] {{
        background: #1A0800 !important;
        border: 1px solid {ACCENT} !important;
        border-radius: 0 !important;
    }}
    [data-baseweb="tag"] span {{
        color: {ACCENT} !important;
    }}

    /* Expander */
    .stExpander > details {{
        background: {CARD} !important;
        border: 1px solid {BORDER} !important;
    }}
    .stExpander > details > summary {{
        background: #0F0F0F !important;
        color: {ACCENT} !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
        text-transform: uppercase !important;
        border-bottom: 1px solid {BORDER} !important;
        padding: 8px 12px !important;
    }}
    .stExpander > details > summary:hover {{
        background: #1A0800 !important;
    }}
    .stExpander > details[open] > summary {{
        border-left: 3px solid {ACCENT} !important;
    }}

    /* Info / success / warning / error */
    [data-testid="stAlert"][data-baseweb="notification"] {{
        border-left: 3px solid {ACCENT} !important;
        background: {CARD} !important;
        color: {TEXT} !important;
    }}
    div[class*="stInfo"] {{ border-left-color: #4488FF !important; }}
    div[class*="stSuccess"] {{ border-left-color: #00CC44 !important; }}
    div[class*="stWarning"] {{ border-left-color: #FFCC00 !important; }}
    div[class*="stError"] {{ border-left-color: #FF2222 !important; }}

    /* Metric виджеты */
    [data-testid="stMetric"] {{
        background: {CARD} !important;
        border: 1px solid {BORDER} !important;
        padding: 8px 12px !important;
    }}
    [data-testid="stMetricValue"] {{
        color: {ACCENT} !important;
        font-size: 20px !important;
        font-weight: 700 !important;
    }}
    [data-testid="stMetricLabel"] {{
        color: {MUTED} !important;
        font-size: 10px !important;
        letter-spacing: 1px !important;
        text-transform: uppercase !important;
    }}

    /* Columns separator */
    [data-testid="column"] {{
        padding: 0 6px !important;
    }}

    /* Scrollbar */
    ::-webkit-scrollbar {{ width: 5px; height: 5px; }}
    ::-webkit-scrollbar-track {{ background: {BG}; }}
    ::-webkit-scrollbar-thumb {{ background: #333; }}
    ::-webkit-scrollbar-thumb:hover {{ background: {ACCENT}; }}

    /* Topbar */
    .bb-topbar {{
        position: fixed;
        top: 0; left: 0; right: 0;
        height: 26px;
        background: {ACCENT};
        color: #000;
        display: flex;
        align-items: center;
        padding: 0 16px;
        font-size: 10px;
        font-weight: 700;
        letter-spacing: .1em;
        z-index: 9999;
        font-family: "IBM Plex Mono", monospace;
        gap: 0;
        overflow: hidden;
        white-space: nowrap;
    }}
    .bb-topbar-sep {{
        width: 1px; height: 100%;
        background: rgba(0,0,0,.30);
        margin: 0 10px; flex-shrink: 0;
    }}
    .bb-topbar-right {{
        margin-left: auto;
        opacity: .55;
        font-weight: 400;
        flex-shrink: 0;
        padding-left: 10px;
    }}
    .bb-topbar-dim {{
        opacity: .70;
        font-weight: 500;
    }}
    .bb-topbar-mods {{
        opacity: .80;
        font-weight: 400;
        letter-spacing: .06em;
    }}
    /* сдвигаем контент под топбар */
    .appview-container .main .block-container {{
        padding-top: 2.2rem !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

topbar_ph = st.empty()


def _render_topbar(company: str = "", bin_: str = "", modules: list = None) -> str:
    co  = company.upper()[:35] if company else ""
    bn  = bin_ or ""
    co_part = (
        f'<span class="bb-topbar-sep"></span>'
        f'<span class="bb-topbar-dim">{co}</span>'
    ) if co else ""
    bn_part = (
        f'<span class="bb-topbar-sep"></span>'
        f'<span class="bb-topbar-dim">БИН&nbsp;{bn}</span>'
    ) if bn else ""
    mods = modules or []
    mods_part = (
        f'<span class="bb-topbar-sep"></span>'
        f'<span class="bb-topbar-mods">{" · ".join(mods)}</span>'
    ) if mods else ""
    return (
        f'<div class="bb-topbar">'
        f'<span>CREDITFLOW</span>'
        f'{co_part}{bn_part}{mods_part}'
        f'<div class="bb-topbar-right">КЗ</div>'
        f'</div>'
    )


topbar_ph.markdown(_render_topbar(), unsafe_allow_html=True)


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
    for k in ["prepared_bytes", "prepared_name", "availability", "prep_report", "processed_bytes", "obsh_meta"]:
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

    existing_osv_info = list_existing_osv_sheets_with_a1(analysis_wb_tmp)
    if existing_osv_info:
        st.markdown("#### _Анализ: найденные ОСВ")
        for sh, suf, a1 in existing_osv_info[:12]:
            tail = f" | {_short(a1)}" if a1 else ""
            st.caption(f"{sh} → {suf}{tail}")
        if len(existing_osv_info) > 12:
            st.caption(f"... и еще {len(existing_osv_info) - 12}")

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
    inv_accounts_found = availability.get("inventory_accounts_found") or []
    saldo_ok = bool(availability.get("saldo_ok"))
    contracts_ok = bool(availability.get("contracts_ok"))
    profit_ok = bool(availability.get("profit_ok"))
    obsh_sheets = availability.get("obsh_sheets") or []
    insights_ok = bool(availability.get("insights_ok"))
    insights_existing = availability.get("insights_existing_titles") or []
    insights_missing = availability.get("insights_missing_titles") or []
    insights_legacy = availability.get("insights_legacy_titles") or []
    gos_ok = bool(availability.get("gos_ok"))
    gos_sheets = availability.get("gos_sheets") or []
    group_osv_ok = bool(availability.get("group_osv_ok"))
    valid_group_osv: Dict[str, List[str]] = availability.get("valid_group_osv") or {}
    clients_ok = bool(availability.get("clients_ok"))
    wt_sheets_avail: List[str] = availability.get("wt_sheets") or []

    opt_profit = st.checkbox("Чистая прибыль", value=False, disabled=(not profit_ok))
    if not profit_ok:
        st.caption("Чистая прибыль недоступна: не найдены листы, содержащие «общ» в названии (например: «общ26», «Gобщ26»).")

    selected_obsh: List[str] = []
    if opt_profit and obsh_sheets:
        # Показываем выбор листов "общ*" + подпись компании из A1
        if "obsh_meta" not in st.session_state:
            meta = {}
            try:
                wb_tmp = load_workbook(io.BytesIO(st.session_state["prepared_bytes"]), read_only=True, data_only=True)
                for sh in obsh_sheets:
                    try:
                        a1 = wb_tmp[sh].cell(row=1, column=1).value
                        meta[sh] = _short(a1) if a1 is not None else ""
                    except Exception:
                        meta[sh] = ""
            except Exception:
                meta = {sh: "" for sh in obsh_sheets}
            st.session_state["obsh_meta"] = meta
        meta = st.session_state.get("obsh_meta") or {}

        st.caption("Выбери листы «общ*», по которым строить расчёт:")
        for sh in obsh_sheets:
            pad, col1, col2 = st.columns([1, 2, 4])
            with pad:
                st.write("")
            with col1:
                on = st.checkbox(sh, value=False, key=f"obsh_pick::{sh}")
            with col2:
                st.caption(meta.get(sh, ""))
            if on:
                selected_obsh.append(sh)

    opt_contracts = st.checkbox("Контракты", value=False, disabled=(not contracts_ok))
    if not contracts_ok:
        st.caption("Контракты недоступны: не найдены пары листов *Wd/*Md.")

    opt_saldo = st.checkbox("Сальдо", value=False, disabled=(not saldo_ok))
    if not saldo_ok:
        st.caption("Сальдо недоступно: не найдены листы, заканчивающиеся на 1210/1710/3310/3510.")

    inv_available_any = bool(inv_accounts_found)
    opt_inventory = st.checkbox("Запасы", value=False, disabled=(not inv_available_any))
    if not inv_available_any:
        st.caption("Запасы недоступны: не найдены ОСВ-листы по счетам 13**.")

    inventory_accounts: List[str] = inv_accounts_found if opt_inventory else []

    opt_gos = st.checkbox("Госы", value=False, disabled=(not gos_ok))
    if not gos_ok:
        st.caption("Госы недоступны: не найдены листы, содержащие «гос» в названии.")

    selected_gos_sheets: List[str] = []
    gos_token = "711d774c72600d3a14b6df6ec7642533"
    if opt_gos and gos_sheets:
        st.caption("Выбери листы «гос*» для обработки:")
        for sh in gos_sheets:
            pad, col1 = st.columns([1, 6])
            with pad:
                st.write("")
            with col1:
                on = st.checkbox(sh, value=False, key=f"gos_pick::{sh}")
            if on:
                selected_gos_sheets.append(sh)

    opt_clients = st.checkbox("Контрагенты", value=False, disabled=(not clients_ok))
    if not clients_ok:
        st.caption("Контрагенты недоступны: не найден лист wt.")

    # Связанные компании: нужен wt-лист + БИН в имени файла
    _has_client_bin = any(_extract_client_bin(u.name or "") for u in uploads if hasattr(u, "name"))
    related_ok = clients_ok and _has_client_bin
    opt_related = st.checkbox("Связанные компании (ADATA)", value=False, disabled=(not related_ok))
    if not related_ok:
        if not clients_ok:
            st.caption("Связанные: не найден лист wt.")
        else:
            st.caption("Связанные: БИН клиента не найден в имени файла (_анализ_БИНКЛИЕНТА.xlsx).")
    selected_wt_sheets: List[str] = []
    if (opt_clients or opt_related) and wt_sheets_avail:
        if len(wt_sheets_avail) == 1:
            selected_wt_sheets = wt_sheets_avail[:]
            st.caption(f"Источник: лист **{wt_sheets_avail[0]}**")
        else:
            st.caption("Выберите листы wt для обработки:")
            for _wsh in wt_sheets_avail:
                _pad, _col1 = st.columns([1, 6])
                with _pad:
                    st.write("")
                with _col1:
                    if st.checkbox(_wsh, value=False, key=f"wt_pick::{_wsh}"):
                        selected_wt_sheets.append(_wsh)

    opt_group_osv = st.checkbox("Объединить общие ОСВ по группе", value=False, disabled=(not group_osv_ok))
    if not group_osv_ok:
        st.caption("Недоступно: не найдены листы «*общ*» с одинаковым суффиксом и ≥2 разными префиксами.")

    selected_groups: Dict[str, List[str]] = {}
    if opt_group_osv and valid_group_osv:
        st.caption("Выбери группы для объединения:")
        for suf in sorted(valid_group_osv.keys()):
            sheets_in_group = valid_group_osv[suf]
            label_name = f"Group{suf}" if suf else "Group"
            label = f"{label_name} — {', '.join(sheets_in_group)}"
            pad, col1 = st.columns([1, 6])
            with pad:
                st.write("")
            with col1:
                if st.checkbox(label, value=False, key=f"grp_pick::{suf}"):
                    selected_groups[suf] = sheets_in_group

    # ── Объединение ОСВ по счёту ─────────────────────────────────────────
    try:
        _tmp_wb8 = load_workbook(io.BytesIO(st.session_state["prepared_bytes"]), read_only=True)
        mergeable_osv: Dict[str, List[str]] = compute_mergeable_osv_from_wb(_tmp_wb8)
        _tmp_wb8.close()
    except Exception:
        mergeable_osv = {}
    merge_osv_ok = bool(mergeable_osv)

    opt_merge_osv = st.checkbox("Объединить ОСВ по счёту", value=False, disabled=(not merge_osv_ok))
    if not merge_osv_ok:
        st.caption("Недоступно: нет счетов, у которых ≥2 листа с одинаковым 4-значным суффиксом.")

    selected_merge_osv: Dict[str, List[str]] = {}
    if opt_merge_osv and mergeable_osv:
        st.caption(
            "Порядок листов определяется автоматически по году в строке «Оборотно-сальдовая ведомость…»."
        )

        _year_cache: Dict[str, Optional[int]] = {}
        try:
            _tmp_wb8b = load_workbook(io.BytesIO(st.session_state["prepared_bytes"]), read_only=True)
            for _sn in _tmp_wb8b.sheetnames:
                _, _suf = split_prefix_suffix4(_sn)
                if _suf.isdigit() and len(_suf) == 4:
                    try:
                        _year_cache[_sn] = _get_osv_end_year(_tmp_wb8b[_sn], _suf)
                    except Exception:
                        _year_cache[_sn] = None
            _tmp_wb8b.close()
        except Exception:
            pass

        for acc in sorted(mergeable_osv.keys()):
            sheets = mergeable_osv[acc]
            sheets_by_year = sorted(sheets, key=lambda s: _year_cache.get(s) or 9999)

            def _lbl(sn: str) -> str:
                y = _year_cache.get(sn)
                return f"{sn} [{y}]" if y else f"{sn} [год ?]"

            cb_label = f"Счёт {acc}:  " + "  +  ".join(_lbl(s) for s in sheets_by_year)
            pad, col1 = st.columns([1, 6])
            with pad:
                st.write("")
            with col1:
                if st.checkbox(cb_label, value=False, key=f"merge_osv::{acc}"):
                    if any(_year_cache.get(s) is None for s in sheets):
                        st.caption(
                            f"  ⚠ Год не найден у части листов — порядок может быть неверным. "
                            f"Убедись, что в строке «Оборотно-сальдовая ведомость» есть год."
                        )
                    selected_merge_osv[acc] = sheets_by_year

    opt_insights = st.checkbox("Инсайты", value=False, disabled=(not insights_ok))
    if insights_existing:
        st.warning("Инс уже есть в файле: " + ", ".join(insights_existing) + ". Повторно не генерирую.")
    if insights_legacy:
        st.caption("Найдены старые листы «инсайты» — при запуске будут переименованы в «инс» и перенесены в конец файла.")
    if not insights_ok:
        if insights_existing:
            st.caption("Инс недоступен: для всех найденных префиксов лист «инс» уже существует.")
        else:
            st.caption("Инсайты недоступны: нужны листы W, M, Mt, Wt (лист «общ» — только для блока 6).")
    elif insights_missing and insights_existing:
        st.caption("Будут добавлены только отсутствующие: " + ", ".join(insights_missing) + ".")

    selected_modes: List[str] = []
    if opt_saldo:
        selected_modes.append("Сальдо")
    if opt_contracts:
        selected_modes.append("Контракты")
    if opt_clients and selected_wt_sheets:
        selected_modes.append("Контрагенты")

    # ── Динамический Bloomberg-топбар ──────────────────────────────────────
    _tb_raw = ""
    if analysis_u:
        _tb_raw = re.sub(r'\.xlsx?$', '', analysis_u.name, flags=re.I)
        _tb_raw = re.sub(r'^_Анализ[_\s]*', '', _tb_raw, flags=re.I).strip()
        _tb_raw = re.sub(r'\b\d{12}\b', '', _tb_raw).strip()
    elif analysis_title:
        _tb_raw = re.sub(r'\b\d{12}\b', '', analysis_title).strip()
    _tb_bin = next(
        (_extract_client_bin(u.name or "") for u in uploads
         if hasattr(u, "name") and _extract_client_bin(u.name or "")),
        "",
    )
    _tb_mods: List[str] = []
    if opt_profit:       _tb_mods.append("ЧП")
    if opt_contracts:    _tb_mods.append("КОНТР")
    if opt_saldo:        _tb_mods.append("САЛЬД")
    if opt_inventory:    _tb_mods.append("ЗАП")
    if opt_gos:          _tb_mods.append("ГОС")
    if opt_clients:      _tb_mods.append("WT")
    if opt_related:      _tb_mods.append("СВЯЗ")
    if opt_group_osv:    _tb_mods.append("ГРУППА")
    if opt_merge_osv:    _tb_mods.append("ОБЪЕД")
    if opt_insights:     _tb_mods.append("ИНС")
    topbar_ph.markdown(_render_topbar(_tb_raw, _tb_bin, _tb_mods or None), unsafe_allow_html=True)
    # ───────────────────────────────────────────────────────────────────────

    st.write("")
    st.markdown("#### Запуск")

    clients_sel_ok = (not opt_clients) or bool(selected_wt_sheets)
    has_any_mode = (bool(selected_modes) or opt_inventory or opt_profit or opt_insights
                    or opt_gos or (opt_group_osv and bool(selected_groups))
                    or (opt_merge_osv and bool(selected_merge_osv))
                    or opt_related)
    inventory_ok = (not opt_inventory) or bool(inventory_accounts)
    profit_sel_ok = (not opt_profit) or bool(selected_obsh)
    gos_sel_ok = (not opt_gos) or bool(selected_gos_sheets)
    group_osv_sel_ok = (not opt_group_osv) or bool(selected_groups)
    merge_osv_sel_ok = (not opt_merge_osv) or bool(selected_merge_osv)
    run_btn = st.button("Обработать", disabled=(
        (not has_any_mode) or (not inventory_ok) or (not profit_sel_ok)
        or (not gos_sel_ok) or (not group_osv_sel_ok) or (not merge_osv_sel_ok)
        or (not clients_sel_ok)
    ))

    status_box = st.empty()
    progress = st.progress(0)

    if run_btn:
        try:
            status_box.info("Подготовка…")
            progress.progress(10)
            time.sleep(0.05)

            out_bytes = st.session_state["prepared_bytes"]

            if opt_profit:
                status_box.info("Обработка: Чистая прибыль…")
                progress.progress(30)
                out_bytes = run_code_4_net_profit(out_bytes, selected_obsh)
                progress.progress(40)

            if "Контракты" in selected_modes:
                status_box.info("Обработка: Контракты…")
                progress.progress(45)
                out_bytes = run_code_2(out_bytes)
                progress.progress(60)

            if "Сальдо" in selected_modes:
                status_box.info("Обработка: Сальдо…")
                progress.progress(65)
                out_bytes, _top5 = run_code_1(out_bytes)
                st.session_state["top5_customers"] = _top5
                progress.progress(78)

            if opt_inventory:
                status_box.info("Обработка: Запасы…")
                progress.progress(82)
                out_bytes, inv_report = run_code_3_inventory(out_bytes, inventory_accounts)
                progress.progress(94)
                if inv_report.get("missing_sheets") or inv_report.get("missing_markers"):
                    parts = []
                    if inv_report.get("missing_sheets"):
                        parts.append("не найдены листы: " + ", ".join(inv_report["missing_sheets"]))
                    if inv_report.get("missing_markers"):
                        parts.append("в листах не найден счет в колонке A")
                    st.warning("Запасы: " + "; ".join(parts))

            if opt_insights:
                status_box.info("Обработка: Инсайты…")
                progress.progress(97)
                out_bytes = run_code_5_insights(out_bytes)
                progress.progress(99)

            if opt_gos and selected_gos_sheets:
                status_box.info("Обработка: Госы (запросы к API, может занять время)…")
                progress.progress(99)
                out_bytes = run_code_6_gos(out_bytes, selected_gos_sheets, gos_token)

            if opt_group_osv and selected_groups:
                status_box.info("Обработка: объединение ОСВ по группе…")
                progress.progress(99)
                out_bytes = run_code_7_group_osv(out_bytes, selected_groups)

            if opt_merge_osv and selected_merge_osv:
                status_box.info("Обработка: объединение ОСВ по счёту…")
                progress.progress(99)
                out_bytes = run_code_8_merge_osv(out_bytes, selected_merge_osv)

            if "Контрагенты" in selected_modes and selected_wt_sheets:
                status_box.info("Обработка: Контрагенты (запросы к ГБД ЮЛ, может занять время)…")
                progress.progress(99)
                out_bytes = run_egov_clients_sheet(out_bytes, selected_wt_sheets)

            if opt_related:
                # Берём имя файла анализа (первый файл с "_анализ_" или 12 цифрами в имени)
                _analysis_filename = ""
                for _u in uploads:
                    if "_анализ_" in (_u.name or "").lower() or re.search(r'\d{12}', _u.name or ""):
                        _analysis_filename = _u.name or ""
                        break
                # Листы wt: выбранные пользователем, иначе все доступные
                _rel_wt = selected_wt_sheets if selected_wt_sheets else wt_sheets_avail
                status_box.info("Обработка: Связанные компании (ADATA + ГБД ЮЛ параллельно)…")
                progress.progress(99)
                out_bytes, _related_msg = run_related_companies_sheet(
                    out_bytes, _rel_wt, _analysis_filename)
                st.caption(f"Связанные: {_related_msg}")

            st.session_state["processed_bytes"] = out_bytes
            st.session_state["processed_name"] = st.session_state.get("prepared_name") or "output.xlsx"
            status_box.success("Готово.")
            progress.progress(100)
        except Exception as e:
            progress.progress(0)
            status_box.error(f"Ошибка: {e}")

    st.write("")
    st.markdown("#### Скачать")
    download_bytes = st.session_state.get("processed_bytes") or st.session_state.get("prepared_bytes")
    download_name = st.session_state.get("processed_name") or st.session_state.get("prepared_name") or "output.xlsx"
    # Если внутри есть проект VBA, Excel ожидает .xlsm; иначе сохраняем как .xlsx.
    want_xlsm = bool(download_bytes and has_vba_project(download_bytes))
    if want_xlsm:
        if not download_name.lower().endswith(".xlsm"):
            download_name = re.sub(r"\.(xlsx|xlsm)$", ".xlsm", download_name, flags=re.IGNORECASE)
            if not download_name.lower().endswith(".xlsm"):
                download_name += ".xlsm"
        download_mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
    else:
        if not download_name.lower().endswith(".xlsx"):
            download_name = re.sub(r"\.(xlsx|xlsm)$", ".xlsx", download_name, flags=re.IGNORECASE)
            if not download_name.lower().endswith(".xlsx"):
                download_name += ".xlsx"
        download_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    if not has_any_mode:
        confirm_merge = st.checkbox("Только объединить (без обработок)", value=False)
        st.download_button(
            label="Скачать",
            data=download_bytes,
            file_name=download_name,
            mime=download_mime,
            use_container_width=True,
            disabled=(not confirm_merge),
        )
    else:
        st.download_button(
            label="Скачать",
            data=download_bytes,
            file_name=download_name,
            mime=download_mime,
            use_container_width=True,
        )


