from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from core.models import CalcData
from core.pdf_parsers.stulz_calc_pdf import StulzOptionRow, parse_stulz_calc_options
from core.pdf_parsers.stulz_winplan_pdf import StulzTechRow, parse_stulz_winplan_specs


@dataclass
class StulzSpecificationData:
    calc_pdf: Path | None = None
    winplan_pdf: Path | None = None
    drawing_pdf: Path | None = None
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



def _canonical_model(prefix: str, number: str, suffix: str = "") -> str:
    return f"{prefix}{number}{suffix}".upper().replace(" ", "").replace("-", "")


def extract_stulz_models_from_text(text: str) -> list[str]:
    """Extract STULZ model codes from a file/folder name.

    The specification folder is the source of truth for Word specification
    blocks. Folder names often look like ``...@ASU-211A-...@ASU 211 A``;
    PDF names may be ``ASD712A WinPlan.pdf``. This helper accepts both
    compact and spaced/dashed variants and keeps only known STULZ prefixes.
    """
    prefixes = {"ASU", "ASD", "CRS", "SXL", "CCD", "CSD", "CWS", "CWD"}
    result: list[str] = []
    seen: set[str] = set()
    source = text or ""
    pattern = re.compile(r"\b([A-ZА-Я]{2,5})[\s_-]*([0-9]{2,4})[\s_-]*([A-ZА-Я]{0,3})\b", re.IGNORECASE)
    for match in pattern.finditer(source):
        prefix, number, suffix = match.groups()
        prefix = prefix.upper()
        suffix = suffix.upper()
        if prefix not in prefixes:
            continue
        model = _canonical_model(prefix, number, suffix)
        if model not in seen:
            result.append(model)
            seen.add(model)
    return result


def list_stulz_specification_models(pdf_dir: str | Path | None) -> list[tuple[str, int]]:
    """Return models found in the selected specification folder.

    The GUI must not fill the specification table from the commercial offer
    Excel rows. It must use the actual contents of the selected
    specifications folder: first model-specific subfolders, then PDFs directly
    in the selected folder if there are no subfolders with model codes.
    Quantity is the number of matching model folders/files and can still be
    edited by the user in the table.
    """
    if not pdf_dir:
        return []
    root = Path(pdf_dir)
    if not root.exists():
        return []

    counts: dict[str, int] = {}
    order: list[str] = []

    def add(model: str) -> None:
        if model not in counts:
            counts[model] = 0
            order.append(model)
        counts[model] += 1

    # Prefer immediate subfolders: in the customer's structure each supplier
    # folder represents one specification set for a concrete model.
    for child in sorted((p for p in root.iterdir() if p.is_dir()), key=lambda p: p.name.lower()):
        models = extract_stulz_models_from_text(child.name)
        if not models:
            # If the model is not in the folder name, inspect PDF names inside.
            for pdf in sorted(child.rglob("*.pdf"), key=lambda p: str(p).lower()):
                models.extend(extract_stulz_models_from_text(pdf.stem))
        for model in dict.fromkeys(models):
            add(model)

    if counts:
        return [(model, counts[model]) for model in order]

    # Fallback for a flat folder with PDFs.
    for pdf in sorted(root.rglob("*.pdf"), key=lambda p: str(p).lower()):
        for model in extract_stulz_models_from_text(pdf.stem):
            add(model)

    return [(model, counts[model]) for model in order]

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
    calc_candidates = sorted(pdfs, key=lambda p: _score_file(p, tokens, "calc"), reverse=True)

    winplan_pdf = next((p for p in winplan_candidates if _score_file(p, tokens, "winplan") >= 100), None)
    calc_pdf = next((p for p in calc_candidates if _score_file(p, tokens, "calc") >= 100), None)

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
