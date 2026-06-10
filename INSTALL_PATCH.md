# Offer Generator modular release patch

This archive contains only the files that must be added or replaced in the existing `offer-generator-main` project.

## How to apply

1. Make a backup of your current project folder.
2. Extract this archive.
3. Copy everything from `offer-generator-patch/` into the root of your project folder.
4. Allow replacement of existing files.
5. Commit the result to GitHub.

Example target structure after copying:

```text
offer-generator-main/
  .github/workflows/main.yml
  sam_offer_generator.spec
  MODULES_MANIFEST.json
  core/runtime_paths.py
  tools/package_modules.py
  tools/prepare_portable_release.py
  scripts/build-windows-portable.ps1
  scripts/build-source-modules.ps1
```

## Files replaced

```text
.github/workflows/main.yml
.gitignore
CHATGPT_CONTEXT.md
FILE_NOTES.md
README.md
core/stulz_reference.py
sam_offer_generator.spec
```

## Files added

```text
GITHUB_RELEASES.md
MODULES_MANIFEST.json
PROJECT_ANALYSIS.md
core/runtime_paths.py
scripts/build-source-modules.ps1
scripts/build-windows-portable.ps1
tools/package_modules.py
tools/prepare_portable_release.py
INSTALL_PATCH.md
```

## What this changes

- PyInstaller build changes from one-file style to portable one-folder release.
- Runtime DLLs, `.pyd`, Qt/PySide6 files and Python dependencies go to `_internal/` in the built release.
- Source modules can be packaged separately: `core`, `brands`, `gui`, `config`, `prices`, `mcp-sam-assistant`, and GitHub build files.
- GitHub Actions builds both a portable Windows folder/zip and separate module ZIP artifacts.

## Local build commands on Windows

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build-windows-portable.ps1
powershell -ExecutionPolicy Bypass -File scripts/build-source-modules.ps1
```

## GitHub Release

Push to `main` to build artifacts. Push a version tag to create a release:

```bash
git tag v0.2.0
git push origin v0.2.0
```
