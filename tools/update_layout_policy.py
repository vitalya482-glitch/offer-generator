from __future__ import annotations

"""Single source of truth for SAM Offer Generator update packaging.

This file intentionally keeps the updater split rules in one place so future
HVAC/Stulz/GUI/parser fixes cannot accidentally move app source files back into
``versions.runtime`` and force users to download the heavy runtime package.

Rules:
- App-No-Runtime.zip contains the executable, editable Python source packages,
  templates, config and other project assets.
- Runtime-Win64.zip contains the heavy PyInstaller runtime and third-party
  dependencies.
- Normal changes in app.py, brands/, core/, gui/ or tools/ must change only the
  app package version, not the runtime version.
- Runtime changes only when Python/dependencies/runtime epoch changes.
"""

from pathlib import Path


APP_ROOT_FILES: tuple[str, ...] = (
    "app.py",
    "README.md",
    "MODULES_MANIFEST.json",
    "config.example.json",
    "requirements.txt",
)

APP_SOURCE_MODULE_DIRS: tuple[str, ...] = (
    "brands",
    "core",
    "gui",
    "tools",
)

APP_CONFIG_DIR = "config"

APP_OPTIONAL_ROOT_DIRS: tuple[str, ...] = (
    "assets",
    "prices",
    "templates",
)

# Keep this schema value stable unless the runtime/dependency split itself is
# intentionally changed.  Do NOT bump it for normal app source fixes.
RUNTIME_FINGERPRINT_SCHEMA = "4-editable-app-sources"
RUNTIME_VERSION_FILE = Path("config/runtime_version.txt")

# These packages represent the heavy runtime layer. If requirements or installed
# dependency versions change, Runtime-Win64.zip must be refreshed.
RUNTIME_PACKAGE_VERSION_NAMES: tuple[str, ...] = (
    "pyinstaller",
    "PySide6",
    "openpyxl",
    "python-docx",
    "pypdf",
)

# Anything listed here is app/update content and must be ignored by the runtime
# fingerprint.  This is the guardrail that prevents source-only fixes from
# downloading Runtime-Win64.zip again.
RUNTIME_IGNORED_TOP_LEVEL_DIRS: frozenset[str] = frozenset(
    APP_OPTIONAL_ROOT_DIRS
    + APP_SOURCE_MODULE_DIRS
    + (APP_CONFIG_DIR,)
)

RUNTIME_IGNORED_ROOT_FILES: frozenset[str] = frozenset({
    "app.py",
    "release_info.json",
})

EXCLUDED_DIR_NAMES: frozenset[str] = frozenset({
    "__pycache__",
    ".git",
    ".pytest_cache",
    ".mypy_cache",
})

EXCLUDED_SUFFIXES: frozenset[str] = frozenset({
    ".pyc",
    ".pyo",
})

RELEASE_LAYOUT_NAME = "pyinstaller-onedir-portable-with-editable-app-sources"
RELEASE_LAYOUT_NOTE = (
    "Keep the full folder together; app Python sources live next to the EXE "
    "and override frozen copies."
)


def validate_update_layout_policy() -> None:
    """Fail fast if the app/runtime split is accidentally made unsafe."""

    required_app_runtime_ignored = set(APP_SOURCE_MODULE_DIRS) | set(APP_OPTIONAL_ROOT_DIRS) | {APP_CONFIG_DIR}
    missing_ignored = required_app_runtime_ignored - set(RUNTIME_IGNORED_TOP_LEVEL_DIRS)
    if missing_ignored:
        raise RuntimeError(
            "Update layout policy regression: app folders are not ignored by runtime fingerprint: "
            + ", ".join(sorted(missing_ignored))
        )

    missing_app_file_ignored = {"app.py"} - set(RUNTIME_IGNORED_ROOT_FILES)
    if missing_app_file_ignored:
        raise RuntimeError(
            "Update layout policy regression: app.py must not affect runtime fingerprint."
        )


def policy_summary() -> str:
    return "\n".join(
        [
            "SAM Offer Generator update layout policy:",
            "  App-No-Runtime files: " + ", ".join(APP_ROOT_FILES),
            "  App-No-Runtime dirs: "
            + ", ".join((APP_CONFIG_DIR,) + APP_SOURCE_MODULE_DIRS + APP_OPTIONAL_ROOT_DIRS),
            "  Runtime ignores dirs: " + ", ".join(sorted(RUNTIME_IGNORED_TOP_LEVEL_DIRS)),
            "  Runtime ignores files: " + ", ".join(sorted(RUNTIME_IGNORED_ROOT_FILES)),
            "  Runtime packages: " + ", ".join(RUNTIME_PACKAGE_VERSION_NAMES),
            "  Runtime schema: " + RUNTIME_FINGERPRINT_SCHEMA,
        ]
    )


if __name__ == "__main__":
    validate_update_layout_policy()
    print(policy_summary())
