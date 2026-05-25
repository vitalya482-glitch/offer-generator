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
