from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

# These folders are bundled by PyInstaller into _internal, but in our portable
# release they are also copied next to the EXE and are updated by the small
# App-No-Runtime module. They must not affect the runtime signature, otherwise
# every change in config/update.json, prices, templates or icons forces the
# 60+ MB runtime ZIP to be downloaded again.
IGNORED_TOP_LEVEL_DIRS = {"assets", "config", "prices", "templates"}
IGNORED_FILES = {"release_info.json"}


def is_ignored_runtime_path(relative_path: Path) -> bool:
    parts = relative_path.parts
    if not parts:
        return False
    if parts[0].lower() in IGNORED_TOP_LEVEL_DIRS:
        return True
    if relative_path.name.lower() in IGNORED_FILES:
        return True
    return False


def directory_content_sha256(root: Path) -> str:
    """Stable hash of true runtime contents: relative paths + file bytes, no timestamps.

    Application data/config folders are intentionally ignored. They are shipped in
    the app module and resolved from the writable app root by core.runtime_paths.
    """
    root = root.resolve()
    if not root.exists() or not root.is_dir():
        raise FileNotFoundError(f"Runtime folder was not found: {root}")

    digest = hashlib.sha256()
    files = sorted(
        (
            path
            for path in root.rglob("*")
            if path.is_file() and not is_ignored_runtime_path(path.relative_to(root))
        ),
        key=lambda path: path.relative_to(root).as_posix().lower(),
    )
    for file_path in files:
        rel = file_path.relative_to(root).as_posix().encode("utf-8")
        digest.update(rel)
        digest.update(b"\0")
        with file_path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        digest.update(b"\0")
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description="Print stable content SHA256 for a runtime directory.")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    print(directory_content_sha256(args.path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
