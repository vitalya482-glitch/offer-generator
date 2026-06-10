from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_MANIFEST = ROOT_DIR / "MODULES_MANIFEST.json"

EXCLUDED_DIRS = {
    ".git",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".venv",
    "venv",
    "env",
    "build",
    "dist",
}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}


def should_skip(path: Path) -> bool:
    rel_parts = path.relative_to(ROOT_DIR).parts
    if any(part in EXCLUDED_DIRS for part in rel_parts):
        return True
    if path.name.startswith("~$"):
        return True
    if path.suffix.lower() in EXCLUDED_SUFFIXES:
        return True
    return False


def iter_module_files(module: dict) -> list[Path]:
    files: list[Path] = []
    seen: set[Path] = set()
    for raw_path in module.get("paths", []):
        path = (ROOT_DIR / raw_path).resolve()
        if not path.exists():
            continue
        if path.is_file():
            candidates = [path]
        else:
            candidates = [p for p in path.rglob("*") if p.is_file()]
        for candidate in candidates:
            if should_skip(candidate):
                continue
            if candidate not in seen:
                seen.add(candidate)
                files.append(candidate)
    return sorted(files, key=lambda p: p.relative_to(ROOT_DIR).as_posix().lower())


def zip_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_module_zip(module: dict, files: list[Path], output_dir: Path, generated_at: str) -> Path:
    module_id = module["id"]
    zip_path = output_dir / f"offer-generator-{module_id}.zip"
    module_info = {
        "project": "SAM Offer Generator",
        "module": module,
        "generated_at_utc": generated_at,
        "file_count": len(files),
    }

    with ZipFile(zip_path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("MODULE_INFO.json", json.dumps(module_info, ensure_ascii=False, indent=2))
        for file_path in files:
            arcname = file_path.relative_to(ROOT_DIR).as_posix()
            zf.write(file_path, arcname)

    return zip_path


def build_index(manifest: dict, module_results: list[dict], generated_at: str) -> str:
    lines = [
        "# SAM Offer Generator modular downloads",
        "",
        f"Generated at UTC: `{generated_at}`",
        "",
        "Each ZIP below is a standalone source module package. Some modules depend on other modules; see the dependency column before replacing files in a repository.",
        "",
        "| Module | ZIP | Depends on | Files | SHA256 |",
        "|---|---|---:|---:|---|",
    ]
    for item in module_results:
        module = item["module"]
        depends = ", ".join(module.get("depends_on") or []) or "-"
        lines.append(
            "| {title} (`{id}`) | `{zip_name}` | {depends} | {file_count} | `{sha256}` |".format(
                title=module.get("title", module["id"]),
                id=module["id"],
                zip_name=item["zip_path"].name,
                depends=depends,
                file_count=item["file_count"],
                sha256=item["sha256"],
            )
        )
    lines.extend([
        "",
        "Recommended release asset for normal users: `SAM-Offer-Generator-windows-portable.zip`.",
        "Recommended release assets for development: download only the module ZIPs you need or download `offer-generator-source-modules.zip`.",
        "",
    ])
    return "\n".join(lines)


def write_all_modules_zip(output_dir: Path, module_results: list[dict], index_path: Path) -> Path:
    bundle_path = output_dir / "offer-generator-source-modules.zip"
    with ZipFile(bundle_path, "w", ZIP_DEFLATED) as zf:
        zf.write(index_path, index_path.name)
        for item in module_results:
            zf.write(item["zip_path"], item["zip_path"].name)
    return bundle_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Create modular source ZIPs for SAM Offer Generator.")
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument("--output", type=Path, default=ROOT_DIR / "dist" / "source-modules")
    args = parser.parse_args()

    manifest_path = args.manifest.resolve()
    output_dir = args.output.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    generated_at = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat()

    module_results: list[dict] = []
    for module in manifest.get("modules", []):
        files = iter_module_files(module)
        zip_path = write_module_zip(module, files, output_dir, generated_at)
        module_results.append({
            "module": module,
            "zip_path": zip_path,
            "file_count": len(files),
            "sha256": zip_sha256(zip_path),
        })

    index_text = build_index(manifest, module_results, generated_at)
    index_path = output_dir / "MODULES_INDEX.md"
    index_path.write_text(index_text, encoding="utf-8")

    bundle_path = write_all_modules_zip(output_dir, module_results, index_path)
    hashes_path = output_dir / "SHA256SUMS.txt"
    lines = []
    for item in module_results:
        lines.append(f"{item['sha256']}  {item['zip_path'].name}")
    lines.append(f"{zip_sha256(bundle_path)}  {bundle_path.name}")
    hashes_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"Wrote module packages to {output_dir}")
    for item in module_results:
        print(f"- {item['zip_path'].name}: {item['file_count']} files")
    print(f"- {bundle_path.name}: all module zips")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
