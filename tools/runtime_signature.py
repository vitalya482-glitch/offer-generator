from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import platform
import re
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools.update_layout_policy import (  # noqa: E402
    RUNTIME_FINGERPRINT_SCHEMA,
    RUNTIME_IGNORED_ROOT_FILES,
    RUNTIME_IGNORED_TOP_LEVEL_DIRS,
    RUNTIME_PACKAGE_VERSION_NAMES,
    RUNTIME_VERSION_FILE as POLICY_RUNTIME_VERSION_FILE,
    validate_update_layout_policy,
)


RUNTIME_VERSION_FILE = REPO_ROOT / POLICY_RUNTIME_VERSION_FILE
REQUIREMENTS_FILE = REPO_ROOT / "requirements.txt"
RUNTIME_VERSION_RE = re.compile(r"^[0-9a-f]{64}$")


def is_ignored_runtime_path(relative_path: Path) -> bool:
    parts = relative_path.parts
    if not parts:
        return False
    if parts[0].lower() in RUNTIME_IGNORED_TOP_LEVEL_DIRS:
        return True
    if relative_path.name.lower() in RUNTIME_IGNORED_ROOT_FILES:
        return True
    return False


def directory_content_sha256(root: Path) -> str:
    """Return a diagnostic hash of the built _internal directory.

    PyInstaller output is not guaranteed to be byte-for-byte reproducible and it
    can include compiled copies of application modules.  This hash is logged for
    diagnostics only; it is deliberately not used as versions.runtime.
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


def _read_manual_runtime_epoch() -> str:
    """Read the optional manual runtime epoch used for exceptional ABI changes."""

    if not RUNTIME_VERSION_FILE.exists():
        raise FileNotFoundError(f"Runtime version file was not found: {RUNTIME_VERSION_FILE}")

    value = RUNTIME_VERSION_FILE.read_text(encoding="utf-8").strip().lower()
    if not RUNTIME_VERSION_RE.fullmatch(value):
        raise ValueError(
            "config/runtime_version.txt must contain exactly one 64-character lowercase hex runtime id"
        )
    return value


def _file_sha256(path: Path) -> str:
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"Runtime fingerprint input was not found: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _package_version(name: str) -> str:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return "not-installed"


def effective_runtime_version() -> str:
    """Return a stable runtime id derived from heavy runtime inputs only.

    This value intentionally ignores application source files.  Python/UI/brand
    fixes are shipped through App-No-Runtime.zip.  Bump config/runtime_version.txt
    manually when a runtime-affecting change is made that is not represented by
    requirements.txt, Python version, architecture or bundled LVKUpdater version.
    """

    validate_update_layout_policy()
    components = [
        f"schema={RUNTIME_FINGERPRINT_SCHEMA}",
        f"epoch={_read_manual_runtime_epoch()}",
        f"requirements={_file_sha256(REQUIREMENTS_FILE)}",
        f"python={sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        f"implementation={platform.python_implementation()}",
        f"machine={platform.machine().lower()}",
    ]
    components.extend(
        f"{package_name}={_package_version(package_name)}"
        for package_name in RUNTIME_PACKAGE_VERSION_NAMES
    )
    payload = "\n".join(components).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Print the stable effective runtime version and log the actual built "
            "_internal content hash for diagnostics."
        )
    )
    parser.add_argument("path", type=Path)
    args = parser.parse_args()

    validate_update_layout_policy()
    actual_build_hash = directory_content_sha256(args.path)
    runtime_version = effective_runtime_version()

    print(f"Runtime build content SHA256 (diagnostic): {actual_build_hash}", file=sys.stderr)
    print(f"Effective runtime version: {runtime_version}", file=sys.stderr)

    # stdout is consumed by GitHub Actions as versions.runtime.
    print(runtime_version)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
