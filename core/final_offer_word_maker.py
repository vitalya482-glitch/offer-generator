from __future__ import annotations

"""Universal DOCX commercial-offer builder.

The module deliberately contains no brand, Excel or GUI logic.  A caller gives
it a Word template, ordinary tag values and zero or more repeating table rows.
The same builder can therefore be used by HVAC, Stulz, Eltek, Riello and future
brands as their templates are migrated to common {{tag_name}} placeholders.
"""

from copy import deepcopy
from dataclasses import dataclass, field
import os
from pathlib import Path
import re
import tempfile
from typing import Any, Iterable, Mapping, Sequence

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Row


_TAG_RE = re.compile(r"\{\{\s*([A-Za-zА-Яа-я0-9_\-.]+)\s*\}\}")


@dataclass(frozen=True)
class RepeatingTable:
    """Description of one repeatable Word table row.

    ``row_marker`` is any tag contained in the template row, for example
    ``item_name`` or ``{{item_name}}``.  Each item in ``rows`` is a mapping of
    tag names to values.  The template row is cloned for every item and then
    removed.
    """

    row_marker: str
    rows: Sequence[Mapping[str, Any]]
    required: bool = True
    keep_rows_together: bool = True
    repeat_header: bool = True


@dataclass
class OfferBuildResult:
    output_path: Path
    replaced_tags: set[str] = field(default_factory=set)
    unresolved_tags: set[str] = field(default_factory=set)
    created_table_rows: int = 0
    warnings: list[str] = field(default_factory=list)


class OfferTemplateError(ValueError):
    """Raised when a required structural marker is absent from a template."""


def make_final_offer(
    template_path: str | Path,
    output_path: str | Path,
    tags: Mapping[str, Any] | None = None,
    repeating_tables: Sequence[RepeatingTable] | None = None,
    *,
    clear_unresolved: bool = True,
    atomic_save: bool = True,
    overwrite: bool = False,
) -> OfferBuildResult:
    """Fill a DOCX template and save the final commercial offer.

    Tag keys may be passed either as ``offer_date`` or ``{{offer_date}}``.
    Replacements are case-insensitive.  Placeholders split by Word into several
    runs are replaced without flattening the whole paragraph, which preserves
    the formatting of neighbouring text.
    """

    template = Path(template_path)
    output = Path(output_path)
    if not template.is_file():
        raise FileNotFoundError(f"Шаблон Word не найден: {template}")
    if template.suffix.casefold() != ".docx":
        raise OfferTemplateError(f"Ожидается DOCX-шаблон: {template}")
    if output.suffix.casefold() != ".docx":
        output = output.with_suffix(".docx")
    if output.exists() and not overwrite:
        raise FileExistsError(f"Файл уже существует: {output}")

    document = Document(str(template))
    result = OfferBuildResult(output_path=output)
    normalized_tags = _normalize_mapping(tags or {})

    for block in repeating_tables or ():
        created = _fill_repeating_table(document, block, result)
        result.created_table_rows += created

    replaced = _replace_mapping_everywhere(document, normalized_tags)
    result.replaced_tags.update(replaced)

    unresolved = collect_unresolved_tags(document)
    result.unresolved_tags = set(unresolved)
    if unresolved:
        result.warnings.append(
            "В шаблоне остались незаполненные теги: "
            + ", ".join(sorted(f"{{{{{name}}}}}" for name in unresolved))
        )
        if clear_unresolved:
            empty_values = {name.casefold(): "" for name in unresolved}
            result.replaced_tags.update(_replace_mapping_everywhere(document, empty_values))

    output.parent.mkdir(parents=True, exist_ok=True)
    if atomic_save:
        _save_atomic(document, output)
    else:
        document.save(str(output))
    _verify_saved_file(output)
    return result


def collect_unresolved_tags(document: DocumentObject) -> set[str]:
    """Return placeholder names still present anywhere in the DOCX."""

    result: set[str] = set()
    for paragraph in _iter_all_paragraphs(document):
        for match in _TAG_RE.finditer(_paragraph_text(paragraph)):
            result.add(match.group(1).strip())
    return result


def normalize_tag_name(value: Any) -> str:
    text = str(value or "").strip()
    match = _TAG_RE.fullmatch(text)
    if match:
        text = match.group(1)
    return text.strip().casefold()


def _normalize_mapping(values: Mapping[str, Any]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for key, value in values.items():
        name = normalize_tag_name(key)
        if name:
            normalized[name] = "" if value is None else str(value)
    return normalized


def _fill_repeating_table(
    document: DocumentObject,
    block: RepeatingTable,
    result: OfferBuildResult,
) -> int:
    marker = normalize_tag_name(block.row_marker)
    if not marker:
        raise OfferTemplateError("Для повторяемой таблицы не указан row_marker.")

    found = _find_template_row(document, marker)
    if found is None:
        message = f"В Word-шаблоне не найдена строка с тегом {{{{{marker}}}}}."
        if block.required:
            raise OfferTemplateError(message)
        result.warnings.append(message)
        return 0

    table, template_row, template_row_index = found
    if block.repeat_header and template_row_index > 0:
        _set_repeat_table_header(table.rows[template_row_index - 1]._tr)
    source_xml = deepcopy(template_row._tr)
    insert_after = template_row._tr
    created = 0

    for raw_values in block.rows:
        row_values = _normalize_mapping(raw_values)
        new_xml = deepcopy(source_xml)
        if block.keep_rows_together:
            _set_row_cant_split(new_xml)
        insert_after.addnext(new_xml)
        insert_after = new_xml
        new_row = _Row(new_xml, table)
        result.replaced_tags.update(_replace_mapping_in_row(new_row, row_values))
        created += 1

    table._tbl.remove(template_row._tr)
    return created


def _find_template_row(document: DocumentObject, marker: str):
    for table in _iter_all_tables(document):
        for row_index, row in enumerate(table.rows):
            row_text = "\n".join(cell.text for cell in row.cells)
            names = {match.group(1).strip().casefold() for match in _TAG_RE.finditer(row_text)}
            if marker in names:
                return table, row, row_index
    return None



def _set_row_cant_split(row_xml) -> None:
    tr_pr = row_xml.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def _set_repeat_table_header(row_xml) -> None:
    tr_pr = row_xml.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))

def _replace_mapping_in_row(row, values: Mapping[str, str]) -> set[str]:
    replaced: set[str] = set()
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replaced.update(_replace_mapping_in_paragraph(paragraph, values))
        for nested_table in cell.tables:
            for nested_row in nested_table.rows:
                replaced.update(_replace_mapping_in_row(nested_row, values))
    return replaced


def _replace_mapping_everywhere(
    document: DocumentObject,
    values: Mapping[str, str],
) -> set[str]:
    replaced: set[str] = set()
    if not values:
        return replaced
    for paragraph in _iter_all_paragraphs(document):
        replaced.update(_replace_mapping_in_paragraph(paragraph, values))
    return replaced


def _replace_mapping_in_paragraph(paragraph, values: Mapping[str, str]) -> set[str]:
    full_text = _paragraph_text(paragraph)
    if "{{" not in full_text:
        return set()

    matches = list(_TAG_RE.finditer(full_text))
    replaced: set[str] = set()
    # Replace from right to left so positions of earlier placeholders stay valid.
    for match in reversed(matches):
        raw_name = match.group(1).strip()
        key = raw_name.casefold()
        if key not in values:
            continue
        if _replace_range_across_runs(paragraph, match.start(), match.end(), values[key]):
            replaced.add(raw_name)
    return replaced


def _replace_range_across_runs(paragraph, start: int, end: int, replacement: str) -> bool:
    if start < 0 or end < start:
        return False
    if not paragraph.runs:
        text = paragraph.text
        paragraph.text = text[:start] + replacement + text[end:]
        return True

    spans: list[tuple[int, int, Any]] = []
    position = 0
    for run in paragraph.runs:
        run_start = position
        run_end = position + len(run.text)
        spans.append((run_start, run_end, run))
        position = run_end

    involved = [item for item in spans if item[1] > start and item[0] < end]
    if not involved:
        return False

    first_start, _first_end, first_run = involved[0]
    last_start, _last_end, last_run = involved[-1]
    prefix = first_run.text[: max(0, start - first_start)]
    suffix = last_run.text[max(0, end - last_start) :]

    if first_run is last_run:
        first_run.text = prefix + replacement + suffix
    else:
        first_run.text = prefix + replacement
        for _run_start, _run_end, run in involved[1:-1]:
            run.text = ""
        last_run.text = suffix
    return True


def _paragraph_text(paragraph) -> str:
    if paragraph.runs:
        return "".join(run.text for run in paragraph.runs)
    return paragraph.text or ""


def _iter_all_tables(document: DocumentObject):
    seen: set[Any] = set()

    def emit(tables):
        for table in tables:
            identity = table._tbl
            if identity in seen:
                continue
            seen.add(identity)
            yield table
            for row in table.rows:
                for cell in row.cells:
                    yield from emit(cell.tables)

    yield from emit(document.tables)
    for section in document.sections:
        for container in (section.header, section.footer):
            yield from emit(container.tables)


def _iter_all_paragraphs(document: DocumentObject):
    seen: set[Any] = set()

    def emit_paragraphs(paragraphs):
        for paragraph in paragraphs:
            identity = paragraph._p
            if identity not in seen:
                seen.add(identity)
                yield paragraph

    def emit_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from emit_paragraphs(cell.paragraphs)
                    yield from emit_tables(cell.tables)

    yield from emit_paragraphs(document.paragraphs)
    yield from emit_tables(document.tables)
    for section in document.sections:
        for container in (section.header, section.footer):
            yield from emit_paragraphs(container.paragraphs)
            yield from emit_tables(container.tables)


def _save_atomic(document: DocumentObject, output: Path) -> None:
    temp_path: Path | None = None
    try:
        fd, raw_temp_path = tempfile.mkstemp(
            suffix=".docx",
            prefix=f".{output.stem}_",
            dir=str(output.parent),
        )
        os.close(fd)
        temp_path = Path(raw_temp_path)
        document.save(str(temp_path))
        _verify_saved_file(temp_path)
        os.replace(temp_path, output)
    finally:
        if temp_path is not None and temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass


def _verify_saved_file(path: Path) -> None:
    try:
        stat = path.stat()
    except OSError as exc:
        raise OSError(f"Не удалось подтвердить сохранение Word-файла: {path}") from exc
    if not path.is_file() or stat.st_size <= 0:
        raise OSError(f"Word-файл не создан или имеет нулевой размер: {path}")


__all__ = [
    "OfferBuildResult",
    "OfferTemplateError",
    "RepeatingTable",
    "collect_unresolved_tags",
    "make_final_offer",
    "normalize_tag_name",
]
