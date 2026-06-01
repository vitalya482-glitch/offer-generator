from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from core.models import CalcData
from core.pdf_parsers.stulz_calc_pdf import StulzCalcTotals, StulzOptionRow, parse_stulz_calc_options, parse_stulz_calc_totals
from core.pdf_parsers.stulz_winplan_pdf import StulzTechRow, parse_stulz_winplan_specs


@dataclass
class StulzSpecificationData:
    calc_pdf: Path | None = None
    winplan_pdf: Path | None = None
    drawing_pdf: Path | None = None
    totals: StulzCalcTotals | None = None
    options: list[StulzOptionRow] = field(default_factory=list)
    technical_specs: list[StulzTechRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _normalize(value: str) -> str:
    value = (value or "").lower().replace("ё", "е")
    value = re.sub(r"[^a-zа-я0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _model_tokens(calc: CalcData) -> list[str]:
    tokens: list[str] = []
    for item in calc.items:
        text = item.name or ""
        # Examples: CCD 171 A, CCD171A, ASD 632 AS.
        matches = re.findall(r"\b([A-ZА-Я]{2,5})\s*([0-9]{2,4})\s*([A-ZА-Я]{0,3})\b", text, re.IGNORECASE)
        for prefix, number, suffix in matches:
            joined = f"{prefix}{number}{suffix}".upper()
            spaced = " ".join(part for part in (prefix.upper(), number, suffix.upper()) if part)
            tokens.extend([joined, spaced])
    if calc.model:
        tokens.append(calc.model)
    return [token for token in dict.fromkeys(tokens) if token]




def _compact_model(value: str) -> str:
    return _normalize(value).replace(" ", "").upper()


def _find_calc_pdf_by_parsed_model(pdfs: list[Path], tokens: list[str]) -> Path | None:
    token_set = {_compact_model(token) for token in tokens if _compact_model(token)}
    if not token_set:
        return None

    for path in sorted(pdfs):
        if "calc" not in path.stem.lower():
            continue
        try:
            totals = parse_stulz_calc_totals(path)
        except Exception:
            continue
        parsed_model = _compact_model(getattr(totals, "model", ""))
        if parsed_model and parsed_model in token_set:
            return path

    return None

def _score_file(path: Path, tokens: list[str], kind: str) -> int:
    name = _normalize(path.stem)
    score = 0

    if kind == "winplan" and "winplan" in name:
        score += 100
    if kind == "calc" and "calc" in name:
        score += 100

    # Drawing PDFs usually have the model name only, without Calc/WinPlan markers.
    if kind == "drawing":
        if "calc" in name or "winplan" in name:
            score -= 100

    for token in tokens:
        token_norm = _normalize(token)
        if token_norm and token_norm in name:
            score += 20
        compact = token_norm.replace(" ", "")
        if compact and compact in name.replace(" ", ""):
            score += 15

    return score


def find_stulz_pdf_pair(pdf_dir: str | Path | None, calc: CalcData) -> tuple[Path | None, Path | None]:
    if not pdf_dir:
        return None, None
    root = Path(pdf_dir)
    if not root.exists():
        return None, None

    pdfs = [p for p in root.rglob("*.pdf") if p.is_file()]
    if not pdfs:
        return None, None

    tokens = _model_tokens(calc)
    winplan_candidates = sorted(pdfs, key=lambda p: _score_file(p, tokens, "winplan"), reverse=True)

    winplan_pdf = next((p for p in winplan_candidates if _score_file(p, tokens, "winplan") >= 100), None)

    # Calc.pdf must be matched by the equipment row inside the PDF, not by file/project name.
    # This prevents false matches such as project title "ASD221A" when the real unit row is "ASD 211 A".
    calc_pdf = _find_calc_pdf_by_parsed_model(pdfs, tokens)

    return calc_pdf, winplan_pdf


def find_stulz_drawing_pdf(pdf_dir: str | Path | None, calc: CalcData) -> Path | None:
    """Find model drawing PDF, for example ASD261AS.PDF.

    Calc and WinPlan PDFs are intentionally ignored here.
    """
    if not pdf_dir:
        return None
    root = Path(pdf_dir)
    if not root.exists():
        return None

    pdfs = [p for p in root.rglob("*.pdf") if p.is_file()]
    if not pdfs:
        return None

    tokens = _model_tokens(calc)
    candidates = sorted(pdfs, key=lambda p: _score_file(p, tokens, "drawing"), reverse=True)
    return next((p for p in candidates if _score_file(p, tokens, "drawing") >= 20), None)


def build_stulz_specification(pdf_dir: str | Path | None, calc: CalcData) -> StulzSpecificationData:
    data = StulzSpecificationData()
    calc_pdf, winplan_pdf = find_stulz_pdf_pair(pdf_dir, calc)
    data.calc_pdf = calc_pdf
    data.winplan_pdf = winplan_pdf
    data.drawing_pdf = find_stulz_drawing_pdf(pdf_dir, calc)

    if calc_pdf:
        try:
            data.totals = parse_stulz_calc_totals(calc_pdf)
        except Exception as exc:
            data.warnings.append(f"Не удалось прочитать итоговые цены из {calc_pdf.name}: {exc}")
        try:
            data.options = parse_stulz_calc_options(calc_pdf, calc.quantity)
        except Exception as exc:
            data.warnings.append(f"Не удалось прочитать опции из {calc_pdf.name}: {exc}")
    else:
        data.warnings.append("Calc PDF для опций не найден.")

    if winplan_pdf:
        try:
            data.technical_specs = parse_stulz_winplan_specs(winplan_pdf)
        except Exception as exc:
            data.warnings.append(f"Не удалось прочитать WinPlan PDF {winplan_pdf.name}: {exc}")
    else:
        data.warnings.append("WinPlan PDF для технических характеристик не найден.")

    return data
