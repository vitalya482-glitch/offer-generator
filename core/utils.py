from __future__ import annotations

import json
import re
from pathlib import Path

MONTHS_RU = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля", 5: "мая", 6: "июня",
    7: "июля", 8: "августа", 9: "сентября", 10: "октября", 11: "ноября", 12: "декабря",
}


def money(value: float, decimals: int = 2) -> str:
    if abs(value - round(value)) < 0.005:
        return f"{round(value):,}".replace(",", " ")
    return f"{value:,.{decimals}f}".replace(",", " ").replace(".", ",")


def qty_text(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else str(value).replace(".", ",")


def sanitize_filename(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value).strip()
    return value or "client"


def load_json(path: Path, default: dict | None = None) -> dict:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default or {}


def first_not_empty(*values):
    for value in values:
        if value not in (None, ""):
            return value
    return ""


def as_float(value, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default


_UNITS = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_UNITS_F = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_TEENS = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
_TENS = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"]
_HUNDREDS = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"]


def _triad_to_words(n: int, female: bool = False) -> list[str]:
    words: list[str] = []
    words.append(_HUNDREDS[n // 100])
    rest = n % 100
    if 10 <= rest <= 19:
        words.append(_TEENS[rest - 10])
    else:
        words.append(_TENS[rest // 10])
        words.append((_UNITS_F if female else _UNITS)[rest % 10])
    return [w for w in words if w]


def _plural(n: int, forms: tuple[str, str, str]) -> str:
    n = abs(n) % 100
    n1 = n % 10
    if 10 < n < 20:
        return forms[2]
    if 1 < n1 < 5:
        return forms[1]
    if n1 == 1:
        return forms[0]
    return forms[2]


def number_to_words_ru(value: float) -> str:
    n = int(round(value))
    if n == 0:
        return "ноль"
    parts: list[str] = []
    billions = n // 1_000_000_000
    millions = (n // 1_000_000) % 1000
    thousands = (n // 1000) % 1000
    rest = n % 1000
    if billions:
        parts += _triad_to_words(billions) + [_plural(billions, ("миллиард", "миллиарда", "миллиардов"))]
    if millions:
        parts += _triad_to_words(millions) + [_plural(millions, ("миллион", "миллиона", "миллионов"))]
    if thousands:
        parts += _triad_to_words(thousands, female=True) + [_plural(thousands, ("тысяча", "тысячи", "тысяч"))]
    if rest:
        parts += _triad_to_words(rest)
    return " ".join(parts)


def currency_label(currency: str) -> str:
    return {"KZT": "тенге", "EUR": "EUR", "USD": "USD", "RUB": "RUB"}.get(currency.upper(), currency)


def currency_suffix(currency: str) -> str:
    return {"KZT": "тенге", "EUR": "EUR", "USD": "USD", "RUB": "RUB"}.get(currency.upper(), currency)
