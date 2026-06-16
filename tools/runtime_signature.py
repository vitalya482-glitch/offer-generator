from __future__ import annotations

import argparse
import hashlib
from pathlib import Path


def directory_content_sha256(root: Path) -> str:
    """Stable hash of directory contents: relative paths + file bytes, no timestamps."""
    root = root.resolve()
    if not root.exists() or not root.is_dir():
        raise FileNotFoundError(f"Runtime folder was not found: {root}")

    digest = hashlib.sha256()
    files = sorted(
        (path for path in root.rglob("*") if path.is_file()),
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
