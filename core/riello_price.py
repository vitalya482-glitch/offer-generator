from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from core.runtime_paths import resource_path


DEFAULT_PRICE_PDF = Path("prices") / "PT Price List 2025 - SAM.pdf"


@dataclass(frozen=True)
class RielloPriceItem:
    model: str
    code: str
    dimensions: str
    weight_kg: float
    price: float
    currency: str = "EUR"
    section: str = ""
    description: str = ""
    power: str = ""

    @property
    def display_name(self) -> str:
        if self.code:
            return f"{self.model} — {self.code}"
        return self.model


_FALLBACK_ITEMS: tuple[RielloPriceItem, ...] = (
    RielloPriceItem(
        model="SRT 20 PM P",
        code="DSRTK20ANBP0RUA",
        dimensions='445(19") x 664 x 397(9U)',
        weight_kg=41,
        price=3000,
        section="SRT/SRM Power Modules",
        description="Силовой модуль для установки в шкафы SRT/SRM Rack Cabinets, параллельная конфигурация, без IN/OUT/BATT кабелей.",
        power="20 kVA / 20 kW",
    ),
    RielloPriceItem(
        model="SRT 20 PM SP",
        code="DSRTK20ANBSPRUA",
        dimensions='445(19") x 664 x 397(9U)',
        weight_kg=41,
        price=3540,
        section="SRT/SRM Power Modules",
        description="Силовой модуль для установки в коммерческие 19-дюймовые шкафы, с параллельной конфигурацией и комплектом кабелей.",
        power="20 kVA / 20 kW",
    ),
    RielloPriceItem(
        model="SRT 60 PWC",
        code="GSRTK60UNB00RUA",
        dimensions="700 x 750 x 2060",
        weight_kg=165,
        price=5100,
        section="SRT/SRM Rack Cabinets",
        description="Шкаф SRT Rack Cabinet для установки 3 силовых модулей SRT 20 PM.",
        power="60 kVA / 60 kW",
    ),
    RielloPriceItem(
        model="SRM 60 PWC",
        code="GSRMK60UNB00RUA",
        dimensions="700 x 750 x 2060",
        weight_kg=165,
        price=5700,
        section="SRT/SRM Rack Cabinets",
        description="Шкаф SRM Rack Cabinet для установки 3 силовых модулей SRM 20 PM.",
        power="60 kVA / 60 kW",
    ),
    RielloPriceItem(
        model="S3T 100",
        code="ES3TM10ANB00RUA",
        dimensions="500x830x1600",
        weight_kg=180,
        price=10725,
        section="Series Sentryum S3T",
        description="ИБП Riello Sentryum S3T.",
        power="100 kVA / 100 kW",
    ),
    RielloPriceItem(
        model="MHT 100",
        code="EMHTM10ANB00RUB",
        dimensions="800 x 850 x 1900",
        weight_kg=700,
        price=10455,
        section="Series Master HP MHT",
        description="ИБП Riello Master HP MHT.",
        power="100 kVA / 90 kW",
    ),
    RielloPriceItem(
        model="MHE 100",
        code="EMHEM10ANB00RUB",
        dimensions="800 x 850 x 1900",
        weight_kg=850,
        price=13070,
        section="Series Master HE MHE",
        description="ИБП Riello Master HE MHE.",
        power="100 kVA / 100 kW",
    ),
)


_MODEL_PREFIXES = ("S3T", "S3M", "SRT", "SRM", "MHT", "MHE", "NXE")
_MODEL_RE = re.compile(r"^(?:" + "|".join(_MODEL_PREFIXES) + r")\b[ A-Z0-9+*/().,\-\"]*$", re.IGNORECASE)
_CODE_RE = re.compile(r"^[A-Z]{1,6}[A-Z0-9]{6,}$")
_NUMBER_RE = re.compile(r"^-?\d+(?:[\s\d]*\d)?(?:[,.]\d+)?$")
_DIM_RE = re.compile(r"\d+\s*(?:\([^)]*\))?\s*x\s*\d+", re.IGNORECASE)
_POWER_PAIR_RE = re.compile(r"^(\d+(?:[,.]\d+)?)\s*/\s*(\d+(?:[,.]\d+)?)$")
_POWER_KVA_RE = re.compile(r"(\d+(?:[,.]\d+)?)\s*kVA\b", re.IGNORECASE)
_POWER_KW_RE = re.compile(r"(\d+(?:[,.]\d+)?)\s*kW\b", re.IGNORECASE)
_MODEL_POWER_RE = re.compile(r"\b(\d+(?:[,.]\d+)?)\b")
_SECTION_RE = re.compile(r"^(Series|Section)\b", re.IGNORECASE)


def default_price_path() -> Path:
    return resource_path(DEFAULT_PRICE_PDF)


def _to_float(value: str, default: float = 0.0) -> float:
    text = str(value or "").strip().replace("\xa0", " ")
    text = text.replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except Exception:
        return default


def _money_to_float(value: str) -> float:
    text = str(value or "").strip().replace("\xa0", " ")
    text = re.sub(r"[^0-9,\.\s-]", "", text)
    return _to_float(text, 0.0)


def _clean_line(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def _is_number(value: str) -> bool:
    return bool(_NUMBER_RE.match(_clean_line(value)))


def _is_code(value: str) -> bool:
    return bool(_CODE_RE.match(_clean_line(value)))


def _is_model(value: str) -> bool:
    text = _clean_line(value)
    return bool(_MODEL_RE.match(text))


def _format_number(value: float) -> str:
    return f"{value:g}"


def _power_from_pair(value: str) -> str:
    match = _POWER_PAIR_RE.match(_clean_line(value))
    if not match:
        return ""
    kva = _to_float(match.group(1), 0.0)
    kw = _to_float(match.group(2), 0.0)
    if kva <= 0 and kw <= 0:
        return ""
    return f"{_format_number(kva)} kVA / {_format_number(kw)} kW"


def _power_from_model(model: str) -> str:
    match = _MODEL_POWER_RE.search(model or "")
    if not match:
        return ""
    power = _to_float(match.group(1), 0.0)
    if power <= 0:
        return ""
    return f"{_format_number(power)} kVA / {_format_number(power)} kW"


def _description_for_item(model: str, section: str) -> str:
    section_lower = (section or "").lower()
    model_upper = (model or "").upper()
    if " 20 PM" in model_upper or "power modules" in section_lower:
        return "Силовой модуль Riello Sentryum Rack."
    if model_upper.endswith("60 PWC") or "rack cabinets" in section_lower:
        return "Шкаф Riello SRT/SRM Rack Cabinet для установки силовых модулей."
    if model_upper.startswith("S3T "):
        return "ИБП Riello Sentryum S3T."
    if model_upper.startswith("S3M "):
        return "ИБП Riello Sentryum S3M."
    if model_upper.startswith("MHT "):
        return "ИБП Riello Master HP MHT."
    if model_upper.startswith("MHE "):
        return "ИБП Riello Master HE MHE."
    if model_upper.startswith("NXE "):
        return "ИБП Riello NextEnergy NXE."
    return section or "Позиция Riello из PDF-прайса."


def _extract_pdf_lines(pdf_path: Path) -> list[tuple[int, str]]:
    try:
        import fitz  # PyMuPDF
    except Exception as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("Для чтения PDF-прайса Riello нужен PyMuPDF из requirements.txt") from exc

    doc = fitz.open(str(pdf_path))
    lines: list[tuple[int, str]] = []
    for page_number, page in enumerate(doc, start=1):
        page_lines = [_clean_line(line) for line in page.get_text("text").splitlines()]
        page_lines = [line for line in page_lines if line]
        lines.extend((page_number, line) for line in page_lines)
    return lines


def _parse_pdf_lines(lines: list[tuple[int, str]]) -> list[RielloPriceItem]:
    items: list[RielloPriceItem] = []
    seen: set[tuple[str, str]] = set()
    section = "Riello PDF price list"

    raw_lines = [line for _page, line in lines]
    pages = [page for page, _line in lines]

    for i, line in enumerate(raw_lines):
        if _SECTION_RE.match(line) or line.lower().startswith("power modules"):
            section = line
            continue

        if not _is_model(line):
            continue
        if i + 1 >= len(raw_lines) or not _is_code(raw_lines[i + 1]):
            continue

        model = line
        code = raw_lines[i + 1]
        block: list[str] = []
        for k in range(i + 2, min(len(raw_lines), i + 14)):
            candidate = raw_lines[k]
            if k > i + 2 and _is_model(candidate) and k + 1 < len(raw_lines):
                next_value = raw_lines[k + 1]
                if _is_code(next_value) or _POWER_PAIR_RE.match(next_value):
                    break
            if candidate in {"MODEL", "Code"} or "PRICE LIST" in candidate.upper():
                break
            block.append(candidate)

        dim_index = next((idx for idx, value in enumerate(block) if _DIM_RE.search(value)), None)
        if dim_index is None:
            continue
        if dim_index + 2 >= len(block):
            continue

        dimensions = block[dim_index]
        weight_text = block[dim_index + 1]
        price_text = block[dim_index + 2]
        if not (_is_number(weight_text) and _is_number(price_text)):
            found = False
            for idx in range(dim_index + 1, min(len(block) - 1, dim_index + 5)):
                if _is_number(block[idx]) and _is_number(block[idx + 1]):
                    weight_text = block[idx]
                    price_text = block[idx + 1]
                    found = True
                    break
            if not found:
                continue

        power = ""
        for value in block[:dim_index]:
            power = _power_from_pair(value)
            if power:
                break
        if not power:
            power = _power_from_model(model)

        model_upper = model.upper()
        # NXE TCE and similar rows are accessories / battery extension cabinets,
        # not UPS models for the main power selection list.
        if " TCE " in f" {model_upper} " or model_upper.startswith("NXE TCE"):
            continue

        price = _money_to_float(price_text)
        weight = _to_float(weight_text)
        if price <= 0:
            continue

        key = (model.upper(), code.upper())
        if key in seen:
            continue
        seen.add(key)

        items.append(
            RielloPriceItem(
                model=model,
                code=code,
                dimensions=dimensions,
                weight_kg=weight,
                price=price,
                currency="EUR",
                section=section,
                description=_description_for_item(model, section),
                power=power,
            )
        )

    return items


def parse_price_pdf(pdf_path: str | Path | None = None) -> list[RielloPriceItem]:
    path = Path(pdf_path) if pdf_path else default_price_path()
    if not path.exists():
        return list(_FALLBACK_ITEMS)

    lines = _extract_pdf_lines(path)
    items = _parse_pdf_lines(lines)
    return items or list(_FALLBACK_ITEMS)


def load_price_items(pdf_path: str | Path | None = None) -> list[RielloPriceItem]:
    return parse_price_pdf(pdf_path)


def find_item(items: list[RielloPriceItem], model_or_code: str, *, contains: bool = False) -> RielloPriceItem | None:
    needle = _clean_line(model_or_code).upper()
    if not needle:
        return None
    for item in items:
        haystacks = (item.model.upper(), item.code.upper(), item.display_name.upper())
        if any(needle == h for h in haystacks):
            return item
    if contains:
        for item in items:
            if needle in item.model.upper() or needle in item.code.upper() or needle in item.display_name.upper():
                return item
    return None


def item_power_kva(item: RielloPriceItem) -> float:
    """Возвращает мощность позиции в кВА для подбора по модели из прайса."""
    match = _POWER_KVA_RE.search(str(item.power or ""))
    if match:
        return _to_float(match.group(1), 0.0)
    match = _MODEL_POWER_RE.search(item.model or "")
    if match:
        return _to_float(match.group(1), 0.0)
    return 0.0


def item_power_kw(item: RielloPriceItem) -> float:
    """Возвращает мощность позиции в кВт, если она есть в прайсе."""
    match = _POWER_KW_RE.search(str(item.power or ""))
    if match:
        return _to_float(match.group(1), 0.0)
    match = _MODEL_POWER_RE.search(item.model or "")
    if match:
        return _to_float(match.group(1), 0.0)
    return 0.0


def item_power_label(item: RielloPriceItem) -> str:
    if item.power:
        return item.power.replace("kVA", "кВА").replace("kW", "кВт")
    kva = item_power_kva(item)
    return f"{kva:g} кВА" if kva else "—"


def format_price(value: float) -> str:
    try:
        return f"{float(value):,.0f}".replace(",", " ")
    except Exception:
        return str(value)


def item_display_with_power(item: RielloPriceItem) -> str:
    return f"{item.display_name} — {item_power_label(item)}"


def item_display_with_price(item: RielloPriceItem) -> str:
    return f"{item.model} — {format_price(item.price)} {item.currency} — {item_power_label(item)} — {item.code}"


def rack_cabinets(items: list[RielloPriceItem]) -> list[RielloPriceItem]:
    return [
        item
        for item in items
        if item.model.upper().endswith("60 PWC")
        or "RACK CABINETS" in item.section.strip().upper()
    ]


def nearest_power_items(items: list[RielloPriceItem], required_power_kw: float) -> list[RielloPriceItem]:
    """
    Возвращает все позиции ближайшей подходящей мощности.

    Для Riello подбираем по кВА: пользователь обычно вводит модельную мощность
    20/60/100, а в прайсе часть серий имеет PF < 1, например MHT 100 = 100 kVA / 90 kW.
    """
    candidates = [(item_power_kva(item), item) for item in items]
    candidates = [(power, item) for power, item in candidates if power > 0]
    if not candidates:
        return list(items)

    required = max(float(required_power_kw or 0), 0.0)
    powers = sorted({power for power, _item in candidates})
    if required <= 0:
        target_power = powers[0]
    else:
        target_power = next((power for power in powers if power >= required), powers[-1])

    result = [item for power, item in candidates if power == target_power]

    def sort_key(item: RielloPriceItem) -> tuple[int, float, str, str]:
        model_upper = item.model.upper()
        prefix_order = {"SRT": 0, "SRM": 1, "S3T": 2, "S3M": 3, "MHT": 4, "MHE": 5, "NXE": 6}
        prefix = model_upper.split(" ", 1)[0]
        prefix_priority = prefix_order.get(prefix, 99)
        return prefix_priority, item.price, model_upper, item.code.upper()

    return sorted(result, key=sort_key)


def power_modules(items: list[RielloPriceItem], prefix: str = "SRT") -> list[RielloPriceItem]:
    prefix = (prefix or "SRT").upper()
    result = []
    for item in items:
        model = item.model.upper()
        if not model.startswith(prefix + " "):
            continue
        if " 20 PM" in model:
            result.append(item)
    return result
