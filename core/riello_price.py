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
        power="до 60 kVA / 60 kW",
    ),
    RielloPriceItem(
        model="SRM 60 PWC",
        code="GSRMK60UNB00RUA",
        dimensions="700 x 750 x 2060",
        weight_kg=165,
        price=5700,
        section="SRT/SRM Rack Cabinets",
        description="Шкаф SRM Rack Cabinet для установки 3 силовых модулей SRM 20 PM.",
        power="до 60 kVA / 60 kW",
    ),
)


_MODEL_RE = re.compile(r"^(SRT|SRM)\s+.+", re.IGNORECASE)
_CODE_RE = re.compile(r"^[A-Z0-9]{8,}$")
_NUMBER_RE = re.compile(r"^-?\d+(?:[\s\d]*\d)?(?:[,.]\d+)?$")
_DIM_RE = re.compile(r"\d+.*x.*\d+", re.IGNORECASE)
_POWER_KW_RE = re.compile(r"(\d+(?:[,.]\d+)?)\s*kW\b", re.IGNORECASE)
_MODEL_POWER_RE = re.compile(r"\b(\d+(?:[,.]\d+)?)\b")


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


def _item_description(model: str, section: str) -> tuple[str, str]:
    section_lower = section.lower()
    model_upper = model.upper()
    if " 20 PM" in model_upper or "power modules" in section_lower:
        return "20 kVA / 20 kW", "Силовой модуль Riello Sentryum Rack."
    if model_upper.endswith("60 PWC") or section.strip().lower() == "srt/srm rack cabinets":
        return "до 60 kVA / 60 kW", "Шкаф SRT/SRM Rack Cabinet для установки 3 силовых модулей."
    return "", section


def _extract_sentryum_rack_text(pdf_path: Path) -> list[str]:
    try:
        import fitz  # PyMuPDF
    except Exception as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("Для чтения PDF-прайса Riello нужен PyMuPDF из requirements.txt") from exc

    doc = fitz.open(str(pdf_path))
    lines: list[str] = []
    in_section = False
    for page in doc:
        page_lines = [_clean_line(line) for line in page.get_text("text").splitlines()]
        page_lines = [line for line in page_lines if line]
        page_text = "\n".join(page_lines).upper()
        if "SERIES SENTRYUM RACK" in page_text:
            in_section = True
        if not in_section:
            continue
        for line in page_lines:
            lines.append(line)
        # The Rack page ends after the accessories block. Stop before the next product series.
        if "SRT/SRM accessories" in page_text:
            break
    return lines


def parse_price_pdf(pdf_path: str | Path | None = None) -> list[RielloPriceItem]:
    path = Path(pdf_path) if pdf_path else default_price_path()
    if not path.exists():
        return list(_FALLBACK_ITEMS)

    lines = _extract_sentryum_rack_text(path)
    items: list[RielloPriceItem] = []
    section = "SERIES SENTRYUM RACK (SRT/SRM)"
    seen: set[tuple[str, str]] = set()

    heading_markers = (
        "SRT/SRM Power Modules",
        "Power Modules for installation",
        "SA - without parallel",
        "SP - including parallel",
        "SP - with parallel",
        "SRT/SRM Rack Cabinets",
        "SRT/SRM accessories",
    )

    i = 0
    while i < len(lines):
        line = lines[i]
        if any(marker.lower() in line.lower() for marker in heading_markers):
            section = line
            i += 1
            continue

        if _MODEL_RE.match(line) and i + 4 < len(lines) and _CODE_RE.match(lines[i + 1]):
            model = line
            code = lines[i + 1]
            dimensions = lines[i + 2]
            weight_text = lines[i + 3]
            price_text = lines[i + 4]

            if _DIM_RE.search(dimensions) and _NUMBER_RE.match(weight_text) and _NUMBER_RE.match(price_text):
                key = (model.upper(), code.upper())
                if key not in seen:
                    power, description = _item_description(model, section)
                    items.append(
                        RielloPriceItem(
                            model=model,
                            code=code,
                            dimensions=dimensions,
                            weight_kg=_to_float(weight_text),
                            price=_money_to_float(price_text),
                            currency="EUR",
                            section=section,
                            description=description,
                            power=power,
                        )
                    )
                    seen.add(key)
                i += 5
                continue
        i += 1

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


def item_power_kw(item: RielloPriceItem) -> float:
    """Возвращает мощность позиции в кВт, насколько это можно понять из прайса/модели."""
    for text in (item.power, item.model, item.description, item.section):
        match = _POWER_KW_RE.search(str(text or ""))
        if match:
            return _to_float(match.group(1), 0.0)
    # В строках Riello мощность часто есть в модели: SRT 60 PWC, SRT 20 PM P и т.п.
    match = _MODEL_POWER_RE.search(item.model)
    if match:
        return _to_float(match.group(1), 0.0)
    return 0.0


def format_price(value: float) -> str:
    try:
        return f"{float(value):,.0f}".replace(",", " ")
    except Exception:
        return str(value)


def item_display_with_power(item: RielloPriceItem) -> str:
    power = item_power_kw(item)
    power_text = f"{power:g} кВт" if power else "мощность не указана"
    return f"{item.display_name} — {power_text}"


def item_display_with_price(item: RielloPriceItem) -> str:
    power = item_power_kw(item)
    power_text = f"{power:g} кВт" if power else "мощность не указана"
    return f"{item.model} — {format_price(item.price)} {item.currency} — {power_text} — {item.code}"


def rack_cabinets(items: list[RielloPriceItem]) -> list[RielloPriceItem]:
    return [
        item
        for item in items
        if item.model.upper().endswith("60 PWC")
        or item.section.strip().upper() == "SRT/SRM RACK CABINETS"
    ]


def nearest_power_items(items: list[RielloPriceItem], required_power_kw: float) -> list[RielloPriceItem]:
    """
    Возвращает все позиции ближайшей подходящей мощности.

    Логика для страницы Riello: сначала ищем минимальную мощность >= требуемой.
    Если в прайсе нет модели выше/равной требуемой, показываем самую мощную доступную.
    """
    candidates = [(item_power_kw(item), item) for item in items]
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
        prefix_priority = 0 if model_upper.startswith("SRT ") else 1 if model_upper.startswith("SRM ") else 2
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
