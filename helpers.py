# -*- coding: utf-8 -*-
"""
helpers.py — набор вспомогательных функций без бизнес-логики.

Содержит:
- нормализацию строк/валют/названий
- парсинг чисел и дат (в т.ч. Excel serial date)
- поиск строк-заголовков разделов/таблиц
- сопоставление биржи по тексту площадки
- извлечение ISIN из произвольного текста
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple


# Предкомпилированный регэксп для ISIN РФ (ускоряет многократные вызовы)
_ISIN_RE = re.compile(r"\bRU[A-Z0-9]{10}\b", re.IGNORECASE)


def norm_text(value: Any) -> str:
    """
    Безопасная нормализация ячейки в строку: str(..).strip()
    None -> "".
    """
    if value is None:
        return ""
    return str(value).strip()


def parse_decimal(value: Any, locale_comma: bool = True) -> Optional[float]:
    """
    Парсинг числовых значений из Excel:
    - поддержка строк вида "1 234,56" (locale_comma=True) или "1,234.56"
    - None/""/"-" -> None
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "-"):
        return None

    # Убираем пробелы/неразрывные пробелы (разделители тысяч)
    s = s.replace("\u00A0", " ").replace(" ", "")
    if locale_comma:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(value: Any) -> Optional[str]:
    """
    Универсальный парсер дат:
    - поддерживает Excel datetime, Excel serial number (с учётом базы 1899-12-30)
    - строки формата dd.mm.yyyy, dd-mm-yyyy, yyyy-mm-dd, yyyy.mm.dd
    Возвращает строку в формате "YYYY-MM-DD" либо None.
    """
    if value is None or str(value).strip() == "":
        return None

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    # Популярные текстовые форматы
    s = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y.%m.%d"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Excel serial date (число дней от 1899-12-30 с "ошибкой 1900 года")
    try:
        num = float(s)
        base = datetime(1899, 12, 30)
        dt = base + timedelta(days=num)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def normalize_currency(ccy: Optional[str]) -> Optional[str]:
    """
    Нормализация кода валюты к верхнему регистру.
    Специальный случай: RUR -> RUB.
    """
    if not ccy:
        return ccy
    c = ccy.strip().upper()
    if c in ("RUR", "RUB"):
        return "RUB"
    return c


def match_header_indices(values: List[Any], mapping: Dict[str, List[str]]) -> Dict[str, int]:
    """
    По гибкому словарю mapping {"логическое_имя": ["подстрока1", ...]}
    возвращает индексы столбцов в строке заголовков.
    Совпадение по подстроке в нижнем регистре.
    """
    idx_map: Dict[str, int] = {}
    for j, cell in enumerate(values):
        text = norm_text(cell).lower()
        if not text:
            continue
        for key, variants in mapping.items():
            for v in variants:
                if v.lower() in text:
                    # Первое совпадение фиксируем, последующие игнорируем
                    idx_map.setdefault(key, j)
    return idx_map


def find_section_title(ws, pattern: re.Pattern) -> Optional[int]:
    """
    Поиск строки (1-based) с заголовком раздела по regex pattern.
    Возвращает индекс строки или None.
    """
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        row_text = " | ".join(norm_text(c) for c in row)
        if pattern.search(row_text):
            return i
    return None


def find_header_below(
    ws,
    start_row: int,
    mapping: Dict[str, List[str]],
    needed: List[str],
    lookahead: int = 25,
) -> Tuple[int, Dict[str, int]]:
    """
    В окне строк [start_row+1 .. start_row+lookahead] ищет заголовок таблицы.
    Критерий: найденные индексы по mapping содержат все ключи из needed.
    Возвращает (header_row_index, idx_map), где idx_map — 0-based индексы столбцов.
    """
    max_r = min(ws.max_row, start_row + lookahead)
    for i in range(start_row + 1, max_r + 1):
        values = [ws.cell(i, j).value for j in range(1, ws.max_column + 1)]
        idx_map = match_header_indices(values, mapping)
        if set(needed).issubset(idx_map.keys()):
            return i, idx_map
    raise RuntimeError("Не найден заголовок таблицы после строки раздела.")


def to_exchange(place: Optional[str], mapping: Optional[Dict[str, str]] = None) -> Optional[str]:
    """
    Грубое сопоставление биржи по тексту площадки.
    По умолчанию использует внутренний словарь:
      'москов'/'moex'/'мосбир' -> MCX
      'спб'/'spb'               -> SPB
      'nasdaq'                  -> NASDAQ
      'nyse'                    -> NYSE
      'lse'                     -> LSE
      'hk'                      -> HK
    Можно передать свой mapping.
    """
    if not place:
        return None
    s = place.strip().lower()
    if mapping is None:
        mapping = {
            "москов": "MCX",
            "moex": "MCX",
            "мосбир": "MCX",
            "спб": "SPB",
            "spb": "SPB",
            "nasdaq": "NASDAQ",
            "nyse": "NYSE",
            "lse": "LSE",
            "hk": "HK",
        }
    for key, val in mapping.items():
        if key in s:
            return val
    return None


def norm_key(s: str) -> str:
    """
    Нормализатор ключей для нестрогого текстового сопоставления (подстрока):
    - нижний регистр
    - 'ё' -> 'е'
    - удаление кавычек/служебных знаков
    - удаление всех не-буквенно-цифровых символов (склейка)
    """
    t = (s or "").lower()
    t = t.replace("ё", "е")
    t = re.sub(r"[\"'`«»]", "", t)
    t = re.sub(r"[^a-z0-9а-я]+", "", t)
    return t


def isin_from_text(text: str) -> Optional[str]:
    """
    Извлекает ISIN (формата RUXXXXXXXXXX) из произвольного текста.
    """
    m = _ISIN_RE.search(text or "")
    return m.group(0).upper() if m else None