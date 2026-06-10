from __future__ import annotations


def split_project_path(path_text: str) -> list[str]:
    raw = path_text.strip().strip('"')
    if not raw:
        return []
    normalized = raw.replace('\\', '/')
    return [part.strip() for part in normalized.split('/') if part.strip()]


def extract_client_from_project_dir(path_text: str) -> str:
    """Return client name from standard SAM server project path.

    Example:
    //Diskstationnew/Exchange/01_Work/01_STULZ/02_Projects/Client/2206/Project
    -> Client
    """
    parts = split_project_path(path_text)
    if not parts:
        return ''

    lowered = [part.lower() for part in parts]
    project_markers = {'02_projects', '2_projects', 'projects', 'проекты'}

    for index, part in enumerate(lowered):
        if part in project_markers and index + 1 < len(parts):
            return parts[index + 1]

    # Fallback for path variants without the 02_Projects marker.
    for index, part in enumerate(lowered):
        if 'stulz' in part and index + 2 < len(parts):
            candidate = parts[index + 2]
            if candidate and not candidate.lower().endswith('projects'):
                return candidate

    return ''


def extract_brand_from_project_dir(path_text: str, available_brands: list[str] | tuple[str, ...]) -> str:
    """Return brand/direction from project path."""
    parts = split_project_path(path_text)
    if not parts:
        return ''

    available = set(available_brands)
    brand_rules = (
        ('stulz', 'Stulz'),
        ('01_stulz', 'Stulz'),
        ('riello', 'Riello'),
        ('ups', 'Riello'),
        ('02_ups', 'Riello'),
        ('battery', 'DC Eltek'),
        ('batteries', 'DC Eltek'),
        ('05_batteries', 'DC Eltek'),
        ('dc_eltek', 'DC Eltek'),
        ('dc eltek', 'DC Eltek'),
        ('eltek', 'DC Eltek'),
        ('genset', 'Generator'),
        ('gen_set', 'Generator'),
        ('03_genset', 'Generator'),
        ('generator', 'Generator'),
        ('generators', 'Generator'),
    )

    for part in parts:
        clean = part.lower().replace('-', '_').replace(' ', '_')
        for marker, brand in brand_rules:
            marker_clean = marker.replace(' ', '_')
            if marker_clean in clean and brand in available:
                return brand

    return ''


def _is_suppliers_dir_name(name: str) -> bool:
    clean = name.lower().replace('-', '_').replace(' ', '_')
    return 'supplier' in clean or 'suppliers' in clean


def find_suppliers_dir(project_dir_text: str) -> str:
    """Find the project suppliers folder.

    Standard SAM folders are usually named like 02_Suppliers, 04_Suppliers,
    Suppliers, etc. The number may change, but the word suppliers usually stays.
    """
    from pathlib import Path

    project_text = str(project_dir_text or '').strip().strip('"')
    if not project_text:
        return ''

    project_dir = Path(project_text)
    if not project_dir.exists() or not project_dir.is_dir():
        return ''

    direct_matches: list[Path] = []
    nested_matches: list[Path] = []

    try:
        for item in project_dir.iterdir():
            if item.is_dir() and _is_suppliers_dir_name(item.name):
                direct_matches.append(item)
    except OSError:
        return ''

    if direct_matches:
        return str(sorted(direct_matches, key=lambda p: p.name.lower())[0])

    # Fallback: sometimes the project folder selected by the user is one level
    # above or below the usual project root. Keep this scan shallow and safe.
    try:
        for item in project_dir.rglob('*'):
            try:
                if item.is_dir() and _is_suppliers_dir_name(item.name):
                    nested_matches.append(item)
            except OSError:
                continue
    except OSError:
        nested_matches = []

    if not nested_matches:
        return ''

    return str(sorted(nested_matches, key=lambda p: (len(p.parts), p.name.lower()))[0])


def infer_specifications_dir(project_dir_text: str, pdf_dirs: list | tuple | None = None) -> str:
    """Guess the folder that contains supplier specification PDFs.

    Priority:
    1. Find a *Suppliers* folder inside the selected project folder.
    2. Inside it, prefer a folder with WinPlan and Calc PDF files.
    3. If such PDFs are not found, use the Suppliers folder itself.
    4. If no Suppliers folder exists, fall back to the old PDF-based logic.

    The returned folder is only a default. User can still change it manually.
    """
    from pathlib import Path
    import os

    project_text = str(project_dir_text or '').strip().strip('"')
    if not project_text:
        return ''

    project_dir = Path(project_text)
    if not project_dir.exists():
        return project_text

    suppliers_text = find_suppliers_dir(str(project_dir))
    suppliers_dir = Path(suppliers_text) if suppliers_text else None

    candidates = [Path(p) for p in (pdf_dirs or []) if str(p).strip()]

    def pdf_features(folder: Path) -> tuple[bool, bool, bool]:
        has_winplan = False
        has_calc = False
        has_any_spec_pdf = False
        try:
            for item in folder.iterdir():
                if not item.is_file() or item.suffix.lower() != '.pdf':
                    continue
                name = item.name.lower()
                is_spec = 'winplan' in name or 'win_plan' in name or 'calc' in name or 'option' in name
                has_any_spec_pdf = has_any_spec_pdf or is_spec
                has_winplan = has_winplan or 'winplan' in name or 'win_plan' in name
                has_calc = has_calc or 'calc' in name
        except OSError:
            return False, False, False
        return has_winplan, has_calc, has_any_spec_pdf

    def has_spec_pdf(folder: Path) -> bool:
        return pdf_features(folder)[2]

    def is_relative_to_safe(path: Path, parent: Path) -> bool:
        try:
            path.resolve().relative_to(parent.resolve())
            return True
        except Exception:
            return False

    def score_pdf_dir(folder: Path) -> tuple[int, int, float, str]:
        has_winplan, has_calc, has_any = pdf_features(folder)
        score = 0
        if suppliers_dir and is_relative_to_safe(folder, suppliers_dir):
            score += 100
        if has_winplan and has_calc:
            score += 60
        elif has_winplan or has_calc:
            score += 35
        elif has_any:
            score += 15
        try:
            mtime = folder.stat().st_mtime
        except OSError:
            mtime = 0.0
        # Prefer more specific nested folders when score is equal.
        return score, len(folder.parts), mtime, str(folder).lower()

    # First try to locate WinPlan/Calc PDFs specifically inside Suppliers.
    if suppliers_dir and suppliers_dir.exists():
        supplier_pdf_dirs: list[Path] = []

        # Reuse already scanned PDF directories when possible.
        for folder in candidates:
            if folder.exists() and is_relative_to_safe(folder, suppliers_dir) and has_spec_pdf(folder):
                supplier_pdf_dirs.append(folder)

        if not supplier_pdf_dirs:
            try:
                for folder in suppliers_dir.rglob('*'):
                    if folder.is_dir() and has_spec_pdf(folder):
                        supplier_pdf_dirs.append(folder)
            except OSError:
                supplier_pdf_dirs = []

        if supplier_pdf_dirs:
            best_score = max(score_pdf_dir(folder)[0] for folder in supplier_pdf_dirs)
            best_dirs = [folder for folder in supplier_pdf_dirs if score_pdf_dir(folder)[0] == best_score]

            # If there are several model folders with spec PDFs, use their common
            # parent so the Stulz scanner can see all models below it.
            try:
                common = Path(os.path.commonpath([str(p) for p in best_dirs]))
            except Exception:
                common = sorted(best_dirs, key=score_pdf_dir, reverse=True)[0]

            if common.is_file():
                common = common.parent
            return str(common)

        return str(suppliers_dir)

    # Old behavior below: use PDF directories from the selected project.
    if has_spec_pdf(project_dir):
        return str(project_dir)

    spec_dirs = [p for p in candidates if p.exists() and has_spec_pdf(p)]
    if not spec_dirs:
        # Lightweight fallback: inspect one level below the project folder.
        try:
            spec_dirs = [p for p in project_dir.iterdir() if p.is_dir() and has_spec_pdf(p)]
        except OSError:
            spec_dirs = []

    if not spec_dirs:
        return str(project_dir)

    # If PDFs are in model folders, use their common parent as the default.
    try:
        common = Path(os.path.commonpath([str(p) for p in spec_dirs]))
    except Exception:
        common = spec_dirs[0]

    if common.is_file():
        common = common.parent
    return str(common)
