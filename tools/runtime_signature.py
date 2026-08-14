from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path

# The updater must decide whether the heavy _internal runtime package really
# changed. Hashing the built _internal directory is NOT suitable for that:
# PyInstaller output is not byte-for-byte deterministic between builds, so the
# hash may change even when Python and all runtime libraries are identical.
#
# We therefore keep an explicit, stable runtime identity in
# config/runtime_version.txt. It is bumped only when the actual runtime stack
# changes (Python/PySide/dependencies/build layout). The runtime ZIP itself is
# still protected by its own SHA256 in offer-generator.json.
RUNTIME_VERSION_FILE = Path(__file__).resolve().parents[1] / "config" / "runtime_version.txt"
RUNTIME_VERSION_RE = re.compile(r"^[0-9a-f]{64}$")

# These folders are bundled by PyInstaller into _internal, but in our portable
# release they are also copied next to the EXE and are updated by the small
# App-No-Runtime module. They are ignored only for the diagnostic build hash.
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
    """Diagnostic hash of built runtime contents.

    This value is useful in CI logs, but it must NOT be used as the updater's
    runtime version because PyInstaller builds are not guaranteed to be
    reproducible byte-for-byte.
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


def declared_runtime_version() -> str:
    if not RUNTIME_VERSION_FILE.exists():
        raise FileNotFoundError(f"Runtime version file was not found: {RUNTIME_VERSION_FILE}")

    value = RUNTIME_VERSION_FILE.read_text(encoding="utf-8").strip().lower()
    if not RUNTIME_VERSION_RE.fullmatch(value):
        raise ValueError(
            "config/runtime_version.txt must contain exactly one 64-character lowercase hex runtime id"
        )
    return value


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Print stable runtime version and log the actual build hash for diagnostics."
    )
    parser.add_argument("path", type=Path)
    args = parser.parse_args()

    # Keep validating the built runtime and log its real content hash. This lets
    # us investigate build drift without making the updater download 70+ MB on
    # every harmless rebuild.
    actual_build_hash = directory_content_sha256(args.path)
    print(f"Runtime build content SHA256 (diagnostic): {actual_build_hash}", file=sys.stderr)

    # stdout is consumed by GitHub Actions as versions.runtime.
    print(declared_runtime_version())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
