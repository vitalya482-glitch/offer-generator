from __future__ import annotations

"""Generate a browsable architecture map for SAM Offer Generator.

The report is intentionally generated from the source tree instead of being a
second hand-maintained inventory.  The small amount of business-flow text is
kept here because static imports alone cannot explain runtime orchestration.
"""

import ast
from dataclasses import dataclass, asdict
from datetime import datetime
import json
from pathlib import Path
import re
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "docs" / "architecture.html"
SKIP_PARTS = {
    ".git",
    ".venv",
    "__pycache__",
    "node_modules",
    "dist",
    "build",
    "tmp",
    ".codex-riello-kp",
}
SOURCE_DIRS = ("core", "brands", "gui", "config", "mcp", "tools", "tests")
ROOT_PY = ("app.py",)
ROOT_MJS = (
    "build_calculation.mjs",
    "build_template_calculation.mjs",
    "help_unmerge.mjs",
    "inspect_workbook.mjs",
    "render_template.mjs",
    "verify_saved.mjs",
)


LAYER_META = {
    "entrypoint": ("Точка входа", "Запуск GUI, self-check и переключение на editable source modules.", "orange"),
    "gui": ("Desktop GUI", "PySide6-окна, страницы направлений, диалоги и выбор файлов.", "blue"),
    "brands": ("Brand logic", "Сценарии Stulz, Riello, DC Eltek, HVAC и Generator.", "purple"),
    "core": ("Core engine", "Общие модели, чтение Excel/PDF, расчёты и генерация DOCX.", "green"),
    "config": ("Configuration", "Редактируемые JSON-справочники и настройки.", "yellow"),
    "mcp": ("MCP assistant", "Локальный MCP-сервер для анализа файлов проекта.", "pink"),
    "tools": ("Build / tools", "Сборка, упаковка, runtime signature и release helpers.", "orange"),
    "tests": ("Tests", "Автоматические проверки отдельных парсеров и генераторов.", "gray"),
    "scripts": ("Node helpers", "Вспомогательные скрипты для workbook/artifact tooling.", "gray"),
    "root": ("Root scripts", "Разовые CLI-скрипты и compatibility helpers.", "gray"),
}


REQUIREMENT_ALIASES = {
    "PySide6": "PySide6",
    "PySide6_Addons": "PySide6 Addons",
    "PySide6_Essentials": "PySide6 Essentials",
    "shiboken6": "shiboken6",
    "python-docx": "python-docx",
    "openpyxl": "openpyxl",
    "num2words": "num2words",
    "pypdf": "pypdf",
    "PyMuPDF": "PyMuPDF",
    "pyinstaller": "PyInstaller",
}
IMPORT_TO_REQUIREMENT = {
    "PySide6": "PySide6",
    "docx": "python-docx",
    "openpyxl": "openpyxl",
    "num2words": "num2words",
    "pypdf": "pypdf",
    "fitz": "PyMuPDF",
    "mcp": "mcp",
}


@dataclass
class SourceRecord:
    id: str
    path: str
    module: str
    layer: str
    kind: str
    doc: str = ""
    lines: int = 0
    classes: list[dict[str, Any]] | None = None
    functions: list[dict[str, Any]] | None = None
    imports: list[str] | None = None
    external_imports: list[str] | None = None
    dependencies: list[str] | None = None
    manifest_module: str | None = None

    def finalize(self) -> None:
        self.classes = self.classes or []
        self.functions = self.functions or []
        self.imports = sorted(set(self.imports or []))
        self.external_imports = sorted(set(self.external_imports or []))
        self.dependencies = sorted(set(self.dependencies or []))


def iter_files() -> Iterable[Path]:
    for name in ROOT_PY:
        path = ROOT / name
        if path.is_file():
            yield path
    for directory in SOURCE_DIRS:
        base = ROOT / directory
        if not base.exists():
            continue
        for path in sorted(base.rglob("*.py")):
            if not any(part in SKIP_PARTS for part in path.parts):
                yield path
    for name in ROOT_MJS:
        path = ROOT / name
        if path.is_file():
            yield path


def rel_path(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def layer_for(path: Path) -> str:
    rel = rel_path(path)
    if path.suffix == ".mjs":
        return "scripts"
    if "/" not in rel:
        return "entrypoint" if rel == "app.py" else "root"
    return rel.split("/", 1)[0] if rel.split("/", 1)[0] in LAYER_META else "root"


def module_name(path: Path) -> str:
    rel = rel_path(path)
    if path.suffix == ".mjs":
        return rel.removesuffix(".mjs").replace("/", ".")
    parts = rel.removesuffix(".py").split("/")
    if parts[-1] == "__init__":
        parts.pop()
    return ".".join(parts) or path.stem


def node_text(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        parent = node_text(node.value)
        return f"{parent}.{node.attr}" if parent else node.attr
    if isinstance(node, ast.Subscript):
        return node_text(node.value)
    if isinstance(node, ast.Constant):
        return repr(node.value)
    return ast.unparse(node) if hasattr(ast, "unparse") else "?"


def short_doc(node: ast.AST) -> str:
    text = ast.get_docstring(node) or ""
    text = re.sub(r"\s+", " ", text).strip()
    return text[:240]


def function_info(node: ast.FunctionDef | ast.AsyncFunctionDef) -> dict[str, Any]:
    args = [a.arg for a in (*node.args.posonlyargs, *node.args.args, *node.args.kwonlyargs)]
    if node.args.vararg:
        args.append("*" + node.args.vararg.arg)
    return {
        "name": node.name,
        "line": node.lineno,
        "args": args[:10],
        "doc": short_doc(node),
        "async": isinstance(node, ast.AsyncFunctionDef),
    }


def class_info(node: ast.ClassDef) -> dict[str, Any]:
    methods = [
        function_info(item)
        for item in node.body
        if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef))
    ]
    return {
        "name": node.name,
        "line": node.lineno,
        "bases": [node_text(base) for base in node.bases],
        "doc": short_doc(node),
        "methods": methods,
    }


def parse_requirements() -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    path = ROOT / "requirements.txt"
    if not path.exists():
        return result
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        match = re.match(r"([^<>=!~\s]+)\s*(.*)", line)
        if not match:
            continue
        package, constraint = match.groups()
        result.append({
            "name": REQUIREMENT_ALIASES.get(package, package),
            "package": package,
            "version": constraint or "не зафиксирована",
            "used_by": [],
        })
    # The local MCP server has its own requirements file.
    mcp_req = ROOT / "mcp" / "sam_assistant" / "requirements.txt"
    if mcp_req.exists():
        for raw in mcp_req.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            match = re.match(r"([^<>=!~\s]+)\s*(.*)", line)
            if match and not any(item["package"].casefold() == match.group(1).casefold() for item in result):
                result.append({
                    "name": match.group(1),
                    "package": match.group(1),
                    "version": match.group(2) or "локальная зависимость MCP",
                    "used_by": [],
                })
    return result


def parse_python(path: Path, module_index: dict[str, str]) -> SourceRecord:
    rel = rel_path(path)
    record = SourceRecord(
        id=f"file:{rel}",
        path=rel,
        module=module_name(path),
        layer=layer_for(path),
        kind="python",
        lines=len(path.read_text(encoding="utf-8", errors="replace").splitlines()),
        classes=[],
        functions=[],
        imports=[],
        external_imports=[],
        dependencies=[],
    )
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except SyntaxError as exc:
        record.doc = f"Не удалось разобрать AST: {exc}"
        record.finalize()
        return record
    record.doc = short_doc(tree)
    internal_candidates: set[str] = set()
    external: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                root = alias.name.split(".", 1)[0]
                if root in {"brands", "core", "gui", "config", "mcp", "tools", "tests"}:
                    internal_candidates.add(alias.name)
                elif root != "__future__":
                    external.add(root)
        elif isinstance(node, ast.ImportFrom):
            if node.module is None:
                continue
            root = node.module.split(".", 1)[0]
            if root in {"brands", "core", "gui", "config", "mcp", "tools", "tests"}:
                internal_candidates.add(node.module)
            elif root != "__future__":
                external.add(root)
    for candidate in internal_candidates:
        target = module_index.get(candidate) or module_index.get(candidate.rsplit(".", 1)[0])
        if target and target != record.id:
            record.dependencies.append(target)
            record.imports.append(candidate)
    record.external_imports = sorted(external)
    # Classes may intentionally be nested inside a factory/function.  The
    # main Qt window is one such case, so collect every class declaration.
    record.classes = [class_info(node) for node in ast.walk(tree) if isinstance(node, ast.ClassDef)]
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            record.functions.append(function_info(node))
    record.finalize()
    return record


def parse_mjs(path: Path) -> SourceRecord:
    text = path.read_text(encoding="utf-8", errors="replace")
    imports = sorted(set(re.findall(r"from\s+['\"]([^'\"]+)['\"]|import\s+['\"]([^'\"]+)['\"]", text)))
    flat_imports = sorted({a or b for a, b in imports})
    return SourceRecord(
        id=f"file:{rel_path(path)}",
        path=rel_path(path),
        module=module_name(path),
        layer="scripts",
        kind="javascript",
        lines=len(text.splitlines()),
        classes=[],
        functions=[],
        imports=flat_imports,
        external_imports=sorted({item.split("/", 1)[0] for item in flat_imports}),
        dependencies=[],
    )


def load_manifest() -> list[dict[str, Any]]:
    path = ROOT / "MODULES_MANIFEST.json"
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return data.get("modules", [])


def assign_manifest(records: list[SourceRecord], manifest: list[dict[str, Any]]) -> None:
    for record in records:
        for item in manifest:
            paths = item.get("paths", [])
            if any(record.path == p or record.path.startswith(p.rstrip("/") + "/") for p in paths):
                record.manifest_module = item.get("id")
                break


def package_used_by(records: list[SourceRecord], requirements: list[dict[str, str]]) -> None:
    package_by_import = {"mcp": "mcp"}
    package_by_import.update(IMPORT_TO_REQUIREMENT)
    for record in records:
        for imported in record.external_imports or []:
            package = package_by_import.get(imported, imported)
            for requirement in requirements:
                if requirement["package"].casefold() == package.casefold():
                    if record.path not in requirement["used_by"]:
                        requirement["used_by"].append(record.path)


def flow_data() -> list[dict[str, Any]]:
    return [
        {
            "id": "main",
            "title": "Основной сценарий генерации КП",
            "description": "Один и тот же общий контур используется несколькими направлениями: GUI собирает контекст, brand-модуль разбирает исходные файлы и передаёт данные в Word-рендерер.",
            "steps": [
                {"label": "1 · Пользователь", "detail": "Выбирает направление, папку проекта, Excel и Word-шаблон.", "refs": ["app.py", "gui/main_window.py", "gui/project_folder_selector.py"]},
                {"label": "2 · Сканирование", "detail": "Поиск Excel/DOCX/PDF с debounce и кешем по папке проекта.", "refs": ["core/project_scanner.py", "gui/path_helpers.py"]},
                {"label": "3 · Контекст", "detail": "Страница направления формирует OfferContext и валидирует входные пути.", "refs": ["core/models.py", "gui/pages/stulz_page.py", "gui/pages/riello_page.py", "gui/pages/dc_eltek_page.py"]},
                {"label": "4 · Разбор данных", "detail": "Excel/PDF читаются адаптером конкретного направления; позиции, суммы, НДС и валюты нормализуются.", "refs": ["core/excel_reader.py", "core/excel_calc_parser.py", "core/pdf_parsers/stulz_calc_pdf.py", "brands/stulz.py", "brands/dc_eltek.py"]},
                {"label": "5 · Формирование", "detail": "Brand-модуль строит текст, позиции, спецификации и замены тегов.", "refs": ["brands/stulz.py", "brands/riello.py", "brands/dc_eltek.py", "brands/hvac/offer_builder.py"]},
                {"label": "6 · Результат", "detail": "Общий DOCX-рендерер заполняет шаблон, повторяющиеся таблицы и сохраняет готовое коммерческое предложение.", "refs": ["core/docx_renderer.py", "core/final_offer_word_maker.py", "core/riello_offer_word.py"]},
            ],
        },
        {
            "id": "startup",
            "title": "Запуск приложения и ленивые вкладки",
            "description": "Главное окно не создаёт все страницы сразу: выбранная вкладка импортируется только при открытии.",
            "steps": [
                {"label": "app.py", "detail": "Парсит аргументы, чистит legacy updater-файлы и запускает run_gui().", "refs": ["app.py"]},
                {"label": "MainWindow", "detail": "Создаёт QMainWindow, настройки QSettings и вкладки направлений.", "refs": ["gui/main_window.py"]},
                {"label": "Registry", "detail": "BRANDS сопоставляет имя направления с Python-модулем бренда.", "refs": ["brands/registry.py"]},
                {"label": "Brand page", "detail": "После выбора вкладки import_module загружает конкретный QWidget-представитель.", "refs": ["gui/pages/stulz_page.py", "gui/pages/riello_workspace.py", "gui/pages/dc_eltek_page.py"]},
            ],
        },
        {
            "id": "stulz",
            "title": "Ветка Stulz",
            "description": "Составная runtime-цепочка поверх базовой логики Stulz: выбор позиций, легенда, валюта, опции и спецификации.",
            "steps": [
                {"label": "StulzPage", "detail": "Собирает OfferContext и проектные настройки.", "refs": ["gui/pages/stulz_page.py", "gui/pages/stulz_page_runtime.py"]},
                {"label": "Runtime layers", "detail": "full_content → position_selection → legend/compressor/cable helpers.", "refs": ["brands/stulz_full_content_runtime.py", "brands/stulz_position_selection_runtime.py", "brands/stulz_legend_runtime.py", "brands/stulz_compressor_runtime.py"]},
                {"label": "Calc + PDF", "detail": "Excel даёт расчёт и позиции, PDF — опции/итоги и технические характеристики.", "refs": ["core/excel_reader.py", "core/pdf_parsers/stulz_calc_pdf.py", "core/pdf_parsers/stulz_winplan_pdf.py"]},
                {"label": "DOCX", "detail": "Вставляются спецификация, опции, чертежи и итоговые суммы.", "refs": ["core/stulz_specification.py", "core/docx_renderer.py"]},
            ],
        },
        {
            "id": "riello",
            "title": "Ветка Riello",
            "description": "Двухэтапная рабочая область: подбор/создание Calc, затем чтение Calc и генерация Word КП.",
            "steps": [
                {"label": "Источник цен", "detail": "ExcelParserToCalc читает supplier workbook и нормализует строки.", "refs": ["gui/pages/riello_workspace.py", "core/excel_parser_to_calc.py"]},
                {"label": "Calc", "detail": "ExcelCalcMaker создаёт рабочую Excel-калькуляцию с формулами и стилями.", "refs": ["core/excel_calc_maker.py", "core/riello_excel_exporter.py"]},
                {"label": "Offer page", "detail": "RielloOfferPage читает готовый Calc, выбирает шаблон и запускает make_offer.", "refs": ["gui/pages/riello_offer_page.py", "core/riello_offer_word.py"]},
                {"label": "Word", "detail": "Общие операции DOCX выполняются через docx_renderer.", "refs": ["core/docx_renderer.py"]},
            ],
        },
        {
            "id": "generic",
            "title": "HVAC / DC Eltek / Generator",
            "description": "Направления используют общий парсер расчётов и общий Word builder там, где это уже подключено.",
            "steps": [
                {"label": "DC Eltek", "detail": "Собственный layout detector, чтение XLSX XML и таблица оборудования.", "refs": ["gui/pages/dc_eltek_page.py", "brands/dc_eltek.py", "brands/dc_eltek_en.py"]},
                {"label": "HVAC", "detail": "Парсит CalcData и строит предложение через универсальный финальный builder.", "refs": ["gui/pages/hvac_page.py", "brands/hvac/offer_builder.py", "core/final_offer_word_maker.py"]},
                {"label": "Generator", "detail": "Текущая страница является каркасом, зарегистрированным в общей GUI-архитектуре.", "refs": ["gui/pages/genset_page.py", "brands/generator.py"]},
            ],
        },
    ]


def stats(records: list[SourceRecord], requirements: list[dict[str, str]]) -> dict[str, int]:
    return {
        "files": len(records),
        "python": sum(r.kind == "python" for r in records),
        "classes": sum(len(r.classes or []) for r in records),
        "functions": sum(len(r.functions or []) for r in records),
        "internal_edges": sum(len(r.dependencies or []) for r in records),
        "libraries": len(requirements),
    }


def build_data() -> dict[str, Any]:
    paths = list(iter_files())
    module_index: dict[str, str] = {}
    for path in paths:
        if path.suffix == ".py":
            module_index.setdefault(module_name(path), f"file:{rel_path(path)}")
    records = [parse_python(path, module_index) if path.suffix == ".py" else parse_mjs(path) for path in paths]
    manifest = load_manifest()
    assign_manifest(records, manifest)
    requirements = parse_requirements()
    # Node helper scripts are part of the project too.  They do not use
    # requirements.txt, so add their real package dependencies explicitly.
    for record in records:
        for imported in record.external_imports or []:
            if imported.startswith("@oai/") and not any(item["package"] == imported for item in requirements):
                requirements.append({
                    "name": imported,
                    "package": imported,
                    "version": "Node.js helper",
                    "used_by": [],
                })
    package_used_by(records, requirements)
    for requirement in requirements:
        requirement["used_by"].sort()
    for record in records:
        record.finalize()
    return {
        "generated_at": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z"),
        "project": "SAM Offer Generator",
        "stats": stats(records, requirements),
        "layers": [
            {"id": key, "title": value[0], "description": value[1], "color": value[2]}
            for key, value in LAYER_META.items()
        ],
        "records": [asdict(record) for record in records],
        "requirements": requirements,
        "manifest": manifest,
        "flows": flow_data(),
        "read_first": [
            {"path": "app.py", "why": "Точка входа и режим self-check."},
            {"path": "gui/main_window.py", "why": "Главное окно, вкладки и ленивое подключение страниц."},
            {"path": "brands/registry.py", "why": "Регистрация направлений и выбор brand-модуля."},
            {"path": "core/models.py", "why": "OfferContext, CalcData и предметные модели."},
            {"path": "core/project_scanner.py", "why": "Поиск файлов проекта и кеширование."},
            {"path": "core/excel_reader.py", "why": "Чтение и индексация Excel для Stulz."},
            {"path": "core/docx_renderer.py", "why": "Общий механизм заполнения Word-шаблонов."},
        ],
    }


HTML_TEMPLATE = r'''<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SAM Offer Generator · Architecture</title>
  <style>
    :root { color-scheme: light; --bg:#f4f7fb; --panel:#fff; --ink:#172033; --muted:#68748a; --line:#dfe6ef; --accent:#dc3b52; --accent-soft:#fff0f2; --blue:#2c73d2; --green:#138a64; --purple:#7950c9; --yellow:#b07800; --orange:#ce6a1b; --pink:#b33f88; --gray:#738096; --shadow:0 12px 34px rgba(21,38,63,.08); }
    * { box-sizing:border-box; }
    body { margin:0; background:var(--bg); color:var(--ink); font:14px/1.45 Segoe UI, Inter, system-ui, -apple-system, sans-serif; }
    button,input,select { font:inherit; }
    button { cursor:pointer; }
    .shell { max-width:1480px; margin:0 auto; padding:28px 28px 52px; }
    header { display:flex; justify-content:space-between; gap:24px; align-items:flex-end; margin-bottom:22px; }
    h1,h2,h3,p { margin:0; }
    h1 { font-size:30px; letter-spacing:-.035em; }
    h2 { font-size:20px; letter-spacing:-.02em; }
    h3 { font-size:15px; }
    .eyebrow { color:var(--accent); font-weight:700; text-transform:uppercase; letter-spacing:.12em; font-size:11px; margin-bottom:5px; }
    .subtitle,.muted { color:var(--muted); }
    .subtitle { max-width:800px; margin-top:7px; }
    .generated { color:var(--muted); font-size:12px; white-space:nowrap; }
    .tabs { display:flex; flex-wrap:wrap; gap:7px; padding:6px; background:var(--panel); border:1px solid var(--line); border-radius:13px; box-shadow:var(--shadow); margin-bottom:18px; }
    .tab { border:0; background:transparent; color:var(--muted); border-radius:9px; padding:9px 14px; font-weight:600; }
    .tab:hover,.tab.active { color:var(--ink); background:var(--accent-soft); }
    .tab.active { color:var(--accent); }
    .toolbar { display:flex; flex-wrap:wrap; gap:9px; align-items:center; margin-bottom:18px; }
    .search { flex:1 1 300px; min-width:230px; border:1px solid var(--line); background:var(--panel); border-radius:9px; padding:10px 12px; color:var(--ink); outline:none; }
    .search:focus,select:focus { border-color:var(--accent); box-shadow:0 0 0 3px var(--accent-soft); }
    select { border:1px solid var(--line); border-radius:9px; background:var(--panel); padding:10px 12px; color:var(--ink); }
    .stats { display:grid; grid-template-columns:repeat(6,minmax(120px,1fr)); gap:10px; margin-bottom:22px; }
    .stat { background:var(--panel); border:1px solid var(--line); border-radius:12px; padding:14px; box-shadow:var(--shadow); }
    .stat strong { display:block; font-size:25px; letter-spacing:-.04em; }
    .stat span { color:var(--muted); font-size:12px; }
    .section { margin-top:24px; }
    .section-head { display:flex; justify-content:space-between; gap:14px; align-items:baseline; margin-bottom:11px; }
    .grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(290px,1fr)); gap:12px; }
    .card { background:var(--panel); border:1px solid var(--line); border-radius:13px; padding:15px; box-shadow:var(--shadow); }
    .card.clickable { transition:transform .15s ease, box-shadow .15s ease, border-color .15s ease; cursor:pointer; }
    .card.clickable:hover { transform:translateY(-2px); border-color:#c4cfdd; box-shadow:0 14px 34px rgba(21,38,63,.12); }
    .card-head { display:flex; justify-content:space-between; gap:10px; align-items:flex-start; }
    .path { color:var(--accent); font:12px/1.35 Consolas, monospace; overflow-wrap:anywhere; }
    .desc { color:var(--muted); margin-top:8px; }
    .badge { display:inline-flex; align-items:center; gap:5px; border-radius:999px; padding:3px 8px; color:#fff; font-size:11px; font-weight:700; white-space:nowrap; }
    .badge.blue { background:var(--blue); } .badge.green { background:var(--green); } .badge.purple { background:var(--purple); } .badge.yellow { background:var(--yellow); } .badge.orange { background:var(--orange); } .badge.pink { background:var(--pink); } .badge.gray { background:var(--gray); }
    .meta { display:flex; flex-wrap:wrap; gap:6px; margin-top:12px; color:var(--muted); font-size:12px; }
    .meta span { background:#f2f5f9; border-radius:6px; padding:4px 7px; }
    .manifest-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:10px; }
    .manifest { border-left:4px solid var(--accent); }
    .manifest .deps { margin-top:10px; display:flex; flex-wrap:wrap; gap:5px; }
    .deps span { color:var(--muted); background:#f2f5f9; border-radius:5px; padding:3px 6px; font:11px Consolas,monospace; }
    .arch-flow { display:flex; align-items:stretch; gap:8px; overflow-x:auto; padding:5px 0 12px; }
    .flow-node { flex:1 0 145px; min-height:114px; border:1px solid var(--line); border-top:4px solid var(--accent); border-radius:11px; background:var(--panel); padding:12px; box-shadow:var(--shadow); }
    .flow-node b { display:block; margin-bottom:6px; }
    .flow-node small { display:block; color:var(--muted); }
    .arrow { flex:0 0 20px; align-self:center; text-align:center; color:var(--accent); font-size:20px; }
    .flow-list { display:grid; gap:12px; }
    .flow-block { background:var(--panel); border:1px solid var(--line); border-radius:13px; padding:16px; box-shadow:var(--shadow); }
    .flow-block > p { color:var(--muted); margin-top:5px; }
    .flow-steps { display:flex; gap:8px; overflow-x:auto; padding:15px 0 2px; }
    .flow-step { min-width:210px; flex:1; border-left:3px solid var(--blue); padding:0 8px 0 11px; }
    .flow-step b { display:block; margin-bottom:4px; }
    .flow-step p { color:var(--muted); font-size:13px; }
    .refs { display:flex; flex-wrap:wrap; gap:4px; margin-top:8px; }
    .ref { border:0; background:var(--accent-soft); color:var(--accent); border-radius:5px; padding:3px 6px; font:11px Consolas,monospace; }
    .table-wrap { overflow:auto; background:var(--panel); border:1px solid var(--line); border-radius:13px; box-shadow:var(--shadow); }
    table { width:100%; border-collapse:collapse; min-width:760px; }
    th,td { text-align:left; vertical-align:top; padding:11px 13px; border-bottom:1px solid var(--line); }
    th { color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.07em; background:#f8fafc; position:sticky; top:0; }
    tr:last-child td { border-bottom:0; }
    code { font:12px Consolas,monospace; color:#334155; }
    .class-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(310px,1fr)); gap:12px; }
    .class-card { border-left:3px solid var(--purple); }
    .class-card .base { color:var(--muted); font:12px Consolas,monospace; margin-top:3px; }
    .methods { margin-top:10px; display:flex; flex-wrap:wrap; gap:5px; }
    .method { background:#f2f5f9; border-radius:5px; padding:4px 6px; font:11px Consolas,monospace; }
    .empty { padding:30px; text-align:center; color:var(--muted); background:var(--panel); border:1px dashed var(--line); border-radius:12px; }
    .detail { position:fixed; inset:0; background:rgba(23,32,51,.35); display:flex; justify-content:flex-end; z-index:10; }
    .detail[hidden] { display:none; }
    .drawer { width:min(620px,100%); height:100%; overflow:auto; background:var(--panel); padding:25px; box-shadow:-20px 0 50px rgba(21,38,63,.2); }
    .drawer-close { float:right; border:0; background:#f2f5f9; border-radius:7px; padding:6px 9px; color:var(--muted); }
    .drawer h2 { padding-right:40px; }
    .drawer-section { margin-top:22px; }
    .drawer-list { margin:8px 0 0; padding-left:19px; }
    .drawer-list li { margin:5px 0; }
    .footer { color:var(--muted); font-size:12px; margin-top:28px; }
    @media (max-width:900px) { .stats { grid-template-columns:repeat(3,1fr); } header { display:block; } .generated { margin-top:10px; } }
    @media (max-width:540px) { .shell { padding:18px 14px 38px; } h1 { font-size:24px; } .stats { grid-template-columns:repeat(2,1fr); } .tabs { border-radius:10px; } .tab { flex:1 0 30%; } }
  </style>
</head>
<body>
  <main class="shell">
    <header>
      <div><div class="eyebrow">Architecture map</div><h1>SAM Offer Generator</h1><p class="subtitle">Интерактивная карта логики проекта: от запуска GUI и выбора направления до разбора Excel/PDF и формирования Word-коммерческого предложения.</p></div>
      <div class="generated" id="generated"></div>
    </header>
    <nav class="tabs" aria-label="Разделы карты">
      <button class="tab active" data-view="overview">Обзор</button>
      <button class="tab" data-view="flows">Потоки</button>
      <button class="tab" data-view="modules">Модули</button>
      <button class="tab" data-view="classes">Классы</button>
      <button class="tab" data-view="dependencies">Зависимости</button>
      <button class="tab" data-view="libraries">Библиотеки</button>
    </nav>
    <div class="toolbar">
      <input class="search" id="search" type="search" placeholder="Поиск по модулям, классам, функциям или файлам…" aria-label="Поиск">
      <select id="layer" aria-label="Слой"><option value="">Все слои</option></select>
      <select id="kind" aria-label="Тип"><option value="">Все типы</option><option value="python">Python</option><option value="javascript">JavaScript</option></select>
    </div>
    <section class="stats" id="stats"></section>
    <div id="content"></div>
    <p class="footer">Сгенерировано статическим AST-анализом исходников. Для обновления после изменений: <code>python tools/generate_architecture_html.py</code></p>
  </main>
  <aside class="detail" id="detail" hidden><div class="drawer"><button class="drawer-close" id="close-detail" type="button">Закрыть</button><div id="detail-body"></div></div></aside>
  <script>
    const DATA = __DATA__;
    const state = { view: 'overview', query: '', layer: '', kind: '' };
    const $ = (selector, root = document) => root.querySelector(selector);
    const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
    const esc = (value) => String(value ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
    const layerById = Object.fromEntries(DATA.layers.map(item => [item.id, item]));
    const recordByPath = Object.fromEntries(DATA.records.map(item => [item.path, item]));
    const recordById = Object.fromEntries(DATA.records.map(item => [item.id, item]));
    const getRecord = (ref) => recordByPath[ref] || recordById[ref];
    const badge = (layer) => { const item = layerById[layer] || {title:layer, color:'gray'}; return `<span class="badge ${item.color}">${esc(item.title)}</span>`; };
    const matches = (record) => {
      const query = state.query.trim().toLowerCase();
      if (state.layer && record.layer !== state.layer) return false;
      if (state.kind && record.kind !== state.kind) return false;
      if (!query) return true;
      const haystack = JSON.stringify(record).toLowerCase();
      return haystack.includes(query);
    };
    function setView(view) { state.view = view; $$('.tab').forEach(btn => btn.classList.toggle('active', btn.dataset.view === view)); render(); }
    function renderStats() {
      const s = DATA.stats;
      const items = [['files','Файлов в карте'],['python','Python-модулей'],['classes','Классов'],['functions','Функций'],['internal_edges','Внутренних связей'],['libraries','Зависимостей']];
      $('#stats').innerHTML = items.map(([key,label]) => `<div class="stat"><strong>${s[key]}</strong><span>${label}</span></div>`).join('');
    }
    function renderOverview() {
      const manifest = DATA.manifest;
      const layerNodes = ['entrypoint','gui','brands','core','config'].map(id => layerById[id]).filter(Boolean);
      return `<section class="section"><div class="section-head"><h2>Как проходит запрос пользователя</h2><span class="muted">основной runtime-контур</span></div>
        <div class="arch-flow">${layerNodes.map((layer, index) => `${index ? '<div class="arrow" aria-hidden="true">→</div>' : ''}<div class="flow-node" style="border-top-color:var(--${layer.color})"><b>${esc(layer.title)}</b><small>${esc(layer.description)}</small></div>`).join('')}</div>
      </section>
      <section class="section"><div class="section-head"><h2>Модульные блоки релиза</h2><span class="muted">из MODULES_MANIFEST.json</span></div><div class="manifest-grid">${manifest.map(item => `<article class="card manifest clickable" data-open-manifest="${esc(item.id)}"><div class="card-head"><h3>${esc(item.title)}</h3><span class="badge gray">${esc(item.id)}</span></div><p class="desc">${esc(item.description || '')}</p><div class="deps">${(item.depends_on || []).map(dep => `<span>← ${esc(dep)}</span>`).join('') || '<span>без зависимостей</span>'}</div></article>`).join('')}</div></section>
      <section class="section"><div class="section-head"><h2>С чего читать код</h2><span class="muted">рекомендуемый маршрут</span></div><div class="table-wrap"><table><thead><tr><th>Файл</th><th>Зачем смотреть</th></tr></thead><tbody>${DATA.read_first.map(item => `<tr><td><code>${esc(item.path)}</code></td><td>${esc(item.why)}</td></tr>`).join('')}</tbody></table></div></section>`;
    }
    function renderFlows() {
      return `<section class="section"><div class="flow-list">${DATA.flows.map(flow => `<article class="flow-block"><h2>${esc(flow.title)}</h2><p>${esc(flow.description)}</p><div class="flow-steps">${flow.steps.map((step, index) => `<div class="flow-step"><b>${esc(step.label)}</b><p>${esc(step.detail)}</p><div class="refs">${(step.refs || []).map(ref => `<button type="button" class="ref" data-open-path="${esc(ref)}">${esc(ref)}</button>`).join('')}</div></div>`).join('')}</div></article>`).join('')}</div></section>`;
    }
    function renderModules() {
      const records = DATA.records.filter(matches).sort((a,b) => a.path.localeCompare(b.path));
      if (!records.length) return '<div class="empty">Ничего не найдено. Измените поиск или фильтр.</div>';
      return `<section class="section"><div class="section-head"><h2>Исходные файлы и модули</h2><span class="muted">${records.length} показано</span></div><div class="grid">${records.map(record => `<article class="card clickable" data-open-path="${esc(record.path)}"><div class="card-head"><div><h3>${esc(record.module)}</h3><div class="path">${esc(record.path)}</div></div>${badge(record.layer)}</div>${record.doc ? `<p class="desc">${esc(record.doc)}</p>` : ''}<div class="meta"><span>${record.lines} строк</span><span>${(record.classes || []).length} классов</span><span>${(record.functions || []).length} функций</span><span>${(record.dependencies || []).length} связей</span></div></article>`).join('')}</div></section>`;
    }
    function renderClasses() {
      const rows = [];
      DATA.records.filter(matches).forEach(record => (record.classes || []).forEach(cls => rows.push({ ...cls, path:record.path, module:record.module, layer:record.layer })));
      rows.sort((a,b) => `${a.name}${a.path}`.localeCompare(`${b.name}${b.path}`));
      if (!rows.length) return '<div class="empty">В выбранном наборе нет классов.</div>';
      return `<section class="section"><div class="section-head"><h2>Классы и методы</h2><span class="muted">${rows.length} найдено</span></div><div class="class-grid">${rows.map(cls => `<article class="card class-card clickable" data-open-path="${esc(cls.path)}"><div class="card-head"><h3>${esc(cls.name)}</h3>${badge(cls.layer)}</div><div class="base">${cls.bases?.length ? `наследует: ${esc(cls.bases.join(', '))}` : 'самостоятельный класс'} · строка ${cls.line}</div><div class="path" style="margin-top:8px">${esc(cls.module)}</div>${cls.doc ? `<p class="desc">${esc(cls.doc)}</p>` : ''}<div class="methods">${(cls.methods || []).slice(0,18).map(method => `<span class="method">${esc(method.name)}()</span>`).join('')}${(cls.methods || []).length > 18 ? `<span class="method">+${cls.methods.length - 18}</span>` : ''}</div></article>`).join('')}</div></section>`;
    }
    function renderDependencies() {
      const rows = DATA.records.filter(matches).flatMap(record => (record.dependencies || []).map(target => ({source:record, target:getRecord(target)}))).filter(row => row.target);
      return `<section class="section"><div class="section-head"><h2>Внутренние связи модулей</h2><span class="muted">импорты между пакетами проекта</span></div><div class="table-wrap"><table><thead><tr><th>Источник</th><th>→</th><th>Цель</th><th>Импортирует</th></tr></thead><tbody>${rows.map(row => `<tr><td><button class="ref" type="button" data-open-path="${esc(row.source.path)}">${esc(row.source.path)}</button></td><td style="color:var(--accent);font-size:18px">→</td><td><button class="ref" type="button" data-open-path="${esc(row.target.path)}">${esc(row.target.path)}</button></td><td><code>${esc((row.source.imports || []).filter(item => item === row.target.module || item.startsWith(row.target.module + '.')).join(', ') || row.target.module)}</code></td></tr>`).join('') || '<tr><td colspan="4">Связи не найдены.</td></tr>'}</tbody></table></div></section>`;
    }
    function renderLibraries() {
      const libs = [...DATA.requirements].sort((a,b) => a.name.localeCompare(b.name));
      return `<section class="section"><div class="section-head"><h2>Внешние библиотеки</h2><span class="muted">requirements.txt и MCP requirements</span></div><div class="grid">${libs.map(lib => `<article class="card"><div class="card-head"><h3>${esc(lib.name)}</h3><span class="badge gray">${esc(lib.version)}</span></div><p class="path" style="margin-top:7px">${esc(lib.package)}</p><p class="desc">Используется в ${(lib.used_by || []).length} файлах.</p><div class="refs">${(lib.used_by || []).slice(0,12).map(path => `<button class="ref" type="button" data-open-path="${esc(path)}">${esc(path)}</button>`).join('')}${(lib.used_by || []).length > 12 ? `<span class="muted">и ещё ${lib.used_by.length - 12}</span>` : ''}</div></article>`).join('')}</div></section>`;
    }
    function render() {
      renderStats();
      const renderers = {overview:renderOverview, flows:renderFlows, modules:renderModules, classes:renderClasses, dependencies:renderDependencies, libraries:renderLibraries};
      $('#content').innerHTML = renderers[state.view]();
      $$('[data-open-path]').forEach(button => button.addEventListener('click', () => openRecord(button.dataset.openPath)));
      $$('[data-open-manifest]').forEach(card => card.addEventListener('click', () => openManifest(card.dataset.openManifest)));
    }
    function openRecord(path) {
      const record = getRecord(path);
      if (!record) return;
      $('#detail-body').innerHTML = `<div class="eyebrow">${esc(record.kind)}</div><h2>${esc(record.module)}</h2><div class="path" style="margin-top:8px">${esc(record.path)}</div><div class="meta">${badge(record.layer)}<span>${record.lines} строк</span><span>${(record.dependencies || []).length} внутренних связей</span></div>${record.doc ? `<div class="drawer-section"><h3>Назначение</h3><p class="desc">${esc(record.doc)}</p></div>` : ''}<div class="drawer-section"><h3>Классы</h3>${(record.classes || []).length ? `<ul class="drawer-list">${record.classes.map(cls => `<li><b>${esc(cls.name)}</b>${cls.bases?.length ? ` <span class="muted">(${esc(cls.bases.join(', '))})</span>` : ''} · строка ${cls.line}<div class="muted">${(cls.methods || []).map(method => esc(method.name) + '()').join(', ') || 'методов нет'}</div></li>`).join('')}</ul>` : '<p class="muted">Классов нет.</p>'}</div><div class="drawer-section"><h3>Функции</h3>${(record.functions || []).length ? `<ul class="drawer-list">${record.functions.map(fn => `<li><code>${esc(fn.name)}(${esc((fn.args || []).join(', '))})</code> · строка ${fn.line}${fn.doc ? `<div class="muted">${esc(fn.doc)}</div>` : ''}</li>`).join('')}</ul>` : '<p class="muted">Публичных функций верхнего уровня нет.</p>'}</div><div class="drawer-section"><h3>Зависимости</h3>${(record.dependencies || []).length ? `<ul class="drawer-list">${record.dependencies.map(dep => { const target = getRecord(dep); return `<li><button class="ref" type="button" data-open-path="${esc(target?.path || dep)}">${esc(target?.path || dep)}</button></li>`; }).join('')}</ul>` : '<p class="muted">Внутренних импортов нет.</p>'}</div>`;
      $('#detail').hidden = false;
      $$('#detail-body [data-open-path]').forEach(button => button.addEventListener('click', () => openRecord(button.dataset.openPath)));
    }
    function openManifest(id) {
      const item = DATA.manifest.find(entry => entry.id === id);
      if (!item) return;
      $('#detail-body').innerHTML = `<div class="eyebrow">release module</div><h2>${esc(item.title)}</h2><div class="path" style="margin-top:8px">${esc(item.id)}</div><div class="drawer-section"><p class="desc">${esc(item.description || '')}</p></div><div class="drawer-section"><h3>Зависит от</h3><p class="desc">${(item.depends_on || []).join(', ') || 'нет'}</p></div><div class="drawer-section"><h3>Состав</h3><ul class="drawer-list">${(item.paths || []).map(path => `<li><code>${esc(path)}</code></li>`).join('')}</ul></div>`;
      $('#detail').hidden = false;
    }
    $$('.tab').forEach(btn => btn.addEventListener('click', () => setView(btn.dataset.view)));
    $('#search').addEventListener('input', event => { state.query = event.target.value; if (state.view === 'overview' || state.view === 'flows') setView('modules'); else render(); });
    $('#layer').innerHTML += DATA.layers.map(layer => `<option value="${esc(layer.id)}">${esc(layer.title)}</option>`).join('');
    $('#layer').addEventListener('change', event => { state.layer = event.target.value; if (state.view === 'overview' || state.view === 'flows') setView('modules'); else render(); });
    $('#kind').addEventListener('change', event => { state.kind = event.target.value; if (state.view === 'overview' || state.view === 'flows') setView('modules'); else render(); });
    $('#close-detail').addEventListener('click', () => $('#detail').hidden = true);
    $('#detail').addEventListener('click', event => { if (event.target.id === 'detail') $('#detail').hidden = true; });
    document.addEventListener('keydown', event => { if (event.key === 'Escape') $('#detail').hidden = true; });
    $('#generated').textContent = `обновлено ${DATA.generated_at}`;
    render();
  </script>
</body>
</html>
'''


def main() -> int:
    data = build_data()
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = HTML_TEMPLATE.replace("__DATA__", payload)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"Generated {OUTPUT}")
    print(json.dumps(data["stats"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
