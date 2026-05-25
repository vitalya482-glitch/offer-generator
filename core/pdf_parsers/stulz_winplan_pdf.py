from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pypdf import PdfReader

from core.stulz_reference import load_stulz_winplan


@dataclass
class StulzTechRow:
    name: str
    value: str = ""
    is_section: bool = False


def extract_pdf_text(path: str | Path) -> str:
    reader = PdfReader(str(path))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    # Fix common WinPlan PDF text-layer inversions.
    text = re.sub(r"([0-9][0-9\s,.]*)Power consumption:\s*kW", r"Power consumption: \1 kW", text)
    text = re.sub(r"V([0-9][0-9\s,.]*)Control voltage:", r"Control voltage: \1 V", text)
    text = re.sub(r"Heat rejection:\s*kW([0-9][0-9\s,.]*)", r"Heat rejection: \1 kW", text)
    text = re.sub(r"kW/kW([0-9][0-9\s,.]*)COP:", r"COP: \1 kW/kW", text)
    text = re.sub(r"(?<![A-Za-z])([0-9][0-9\s,.]*)Number:", r"Number: \1", text)
    return text


def _clean_value(value: str) -> str:
    value = value or ""
    value = value.replace("mі/h", "м³/ч").replace("m³/h", "м³/ч").replace("m3/h", "м³/ч")
    value = value.replace("rpm", "Об/мин").replace("rel.%", "%")
    value = value.replace("dB(A)", "дБ(A)").replace("dB(А)", "дБ(А)")
    value = re.sub(r"\s+", " ", value)
    value = value.replace(" kW/kW", " кВт/кВт")
    value = value.replace(" kW", " кВт")
    value = value.replace(" mm", " мм")
    value = value.replace(" Pa", " Па")
    value = value.replace(" kg", " кг")
    return value.strip(" ;")


def _segment(text: str, start_marker: str, end_marker: str | None = None) -> str:
    start = text.find(start_marker)
    if start < 0:
        return ""
    start += len(start_marker)
    if end_marker:
        end = text.find(end_marker, start)
        if end >= 0:
            return text[start:end]
    return text[start:]


def _extract_after_label(section_text: str, label: str, labels: list[str]) -> str:
    pos = section_text.find(label)
    if pos < 0:
        return ""
    start = pos + len(label)
    end = len(section_text)
    for other in labels:
        if other == label:
            continue
        other_pos = section_text.find(other, start)
        if other_pos >= 0:
            end = min(end, other_pos)
    raw = section_text[start:end]
    raw = raw.split("\n", 1)[0] if "\n" in raw and len(raw.split("\n", 1)[0].strip()) > 0 else raw
    cleaned = _clean_value(raw)

    # Some PDF text layers put the value immediately before the label, e.g. "V9,1Control voltage:".
    if not cleaned:
        before = section_text[:pos][-30:]
        match = re.search(r"([0-9][0-9\s,.]*\s*(?:V|A|kW|Pa|rpm|°C|%|dB\(A\))?)\s*$", before)
        if match:
            cleaned = _clean_value(match.group(1))
    return cleaned


def _dedupe_rows(rows: list[StulzTechRow]) -> list[StulzTechRow]:
    result: list[StulzTechRow] = []
    seen: set[tuple[str, str, bool]] = set()
    for row in rows:
        key = (row.name, row.value, row.is_section)
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def parse_stulz_winplan_specs(path: str | Path) -> list[StulzTechRow]:
    text = extract_pdf_text(path)
    config = load_stulz_winplan()

    unit_text = _segment(text, "Unit\n", "Fan (Data per unit)") or _segment(text, "Unit", "Fan (Data per unit)")
    fan_text = _segment(text, "Fan (Data per unit)", "Compressor (Data per compressor)")
    compressor_text = _segment(text, "Compressor (Data per compressor)", "Condenser")
    condenser_text = _segment(text, "Condenser", "Noise data")

    unit_labels = ["Unit type:", "Cooling capacity (total):", "Cooling capacity (sensible):", "Net total cooling capacity:", "Net sensible cooling capacity:", "Condensing temperature:", "EER:", "AER:", "Sound power level:", "LpA (2m freefield):", "Number of refrigerant circuits:", "Number of compressors:", "Total power consumption:", "Airflow:", "Air velocity:", "Return air temperature:", "Return air humidity:", "Supply air temperature:", "Altitude above sea level:", "Height:", "Width:", "Depth:", "Weight:", "Belt type:", "Refrigerant:", "Power supply:"]
    fan_labels = ["Fan type:", "Number:", "Max. revolutions:", "Revolutions:", "Nominal power:", "Power consumption:", "ESP external static pressure:", "Total pressure drop:", "Control voltage:"]
    compressor_labels = ["Electrical power consumption:", "Heat rejection:", "COP:", "Number:", "Evaporating temperature:"]
    condenser_labels = ["Unit type:", "Ambient temperature:", "Sound pressure group:", "LpA (5m freefield):", "Required condenser capacity:", "Available condenser capacity:", "Difference:", "Number of fans:", "Number of condensers:", "Weight:", "Current consumption (per fan):", "Power consumption (per fan):", "Height:", "Width:", "Depth:", "Airflow:"]

    sections = [
        ("", unit_text, unit_labels, ["Unit type:", "Cooling capacity (total):", "Net sensible cooling capacity:", "Condensing temperature:", "EER:", "Sound power level:", "LpA (2m freefield):", "Airflow:", "Return air temperature:", "Return air humidity:", "Supply air temperature:", "Height:", "Width:", "Depth:", "Weight:", "Refrigerant:", "Power supply:"]),
        ("Вентилятор:", fan_text, fan_labels, ["Fan type:", "Number:", "Max. revolutions:", "Revolutions:", "Nominal power:", "Power consumption:", "ESP external static pressure:", "Total pressure drop:", "Control voltage:"]),
        ("Компрессор:", compressor_text, compressor_labels, ["Electrical power consumption:", "COP:", "Number:", "Heat rejection:", "Evaporating temperature:"]),
        ("Выносной блок (Конденсор):", condenser_text, condenser_labels, ["Unit type:", "Ambient temperature:", "LpA (5m freefield):", "Required condenser capacity:", "Available condenser capacity:", "Difference:", "Number of fans:", "Number of condensers:", "Weight:", "Current consumption (per fan):", "Power consumption (per fan):", "Height:", "Width:", "Depth:", "Airflow:"]),
    ]

    ru_by_source: dict[str, str] = {}
    for row in config:
        src = (row.get("source_name") or "").strip()
        ru = (row.get("ru_name") or src).strip()
        if src and src not in ru_by_source:
            ru_by_source[src] = ru

    rows: list[StulzTechRow] = []
    for section_title, section_text, all_labels, output_labels in sections:
        if section_title:
            rows.append(StulzTechRow(section_title, "", True))
        for label in output_labels:
            value = _extract_after_label(section_text, label, all_labels)
            if not value:
                continue
            rows.append(StulzTechRow(ru_by_source.get(label, label), value, False))

    return _dedupe_rows(rows)
