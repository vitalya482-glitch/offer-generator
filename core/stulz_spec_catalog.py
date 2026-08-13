from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from core.pdf_parsers.stulz_calc_pdf import parse_stulz_calc_totals


@dataclass(frozen=True)
class StulzSpecEntry:
    """One physical STULZ specification set discovered from a Calc PDF.

    Important: entries are intentionally NOT grouped by model. The same model can
    legitimately appear several times in one project (for example for different
    cities or different configurations). The parent directory of Calc.pdf is the
    boundary of one specification set.
    """

    key: str
    model: str
    quantity: float
    calc_pdf: Path
    source_dir: Path
    source_label: str


def _format_model(value: str) -> str:
    return (value or "").replace(" ", "").strip()


def _source_label(root: Path, source_dir: Path) -> str:
    try:
        relative = source_dir.relative_to(root)
        text = str(relative)
        if text and text != ".":
            return text
    except Exception:
        pass
    return source_dir.name or str(source_dir)


def discover_stulz_spec_entries(spec_dir: str | Path | None) -> list[StulzSpecEntry]:
    """Discover STULZ specification sets below *spec_dir*.

    Every successfully parsed Calc PDF becomes a separate entry. This deliberately
    preserves duplicate model names instead of summing their quantities.
    """

    if not spec_dir:
        return []

    root = Path(spec_dir)
    if not root.exists():
        return []

    entries: list[StulzSpecEntry] = []
    for calc_pdf in sorted(root.rglob("*.pdf")):
        if not calc_pdf.is_file() or "calc" not in calc_pdf.stem.lower():
            continue

        try:
            totals = parse_stulz_calc_totals(calc_pdf)
        except Exception:
            continue

        model = _format_model(getattr(totals, "model", ""))
        if not model:
            continue

        raw_qty = getattr(totals, "quantity", None)
        try:
            quantity = float(raw_qty) if raw_qty not in (None, "") else 1.0
        except Exception:
            quantity = 1.0
        if quantity <= 0:
            quantity = 1.0

        source_dir = calc_pdf.parent
        try:
            key = str(calc_pdf.resolve()).lower()
        except Exception:
            key = str(calc_pdf).lower()

        entries.append(
            StulzSpecEntry(
                key=key,
                model=model,
                quantity=quantity,
                calc_pdf=calc_pdf,
                source_dir=source_dir,
                source_label=_source_label(root, source_dir),
            )
        )

    return entries
