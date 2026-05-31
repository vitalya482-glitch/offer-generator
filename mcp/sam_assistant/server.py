from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path
from typing import Any

import fitz
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("sam-assistant")

EXCLUDED_DIRS = {".git", ".venv", "build", "dist", "__pycache__", ".continue"}
TEXT_EXTENSIONS = {".py", ".md", ".json", ".txt", ".yaml", ".yml", ".spec"}


def _resolve_project_path(project_path: str) -> Path:
    root = Path(project_path).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        raise ValueError(f"Project path does not exist or is not a directory: {root}")
    return root


def _resolve_inside_project(project_path: str, path: str) -> Path:
    root = _resolve_project_path(project_path)

    candidate = Path(path).expanduser()
    if not candidate.is_absolute():
        candidate = root / candidate

    candidate = candidate.resolve()

    if root != candidate and root not in candidate.parents:
        raise ValueError(f"Path is outside project: {candidate}")

    return candidate


def _iter_project_files(root: Path, extensions: set[str] | None = None):
    for current_root, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in EXCLUDED_DIRS]

        for filename in files:
            path = Path(current_root) / filename
            if extensions is None or path.suffix.lower() in extensions:
                yield path


def _read_text_file(path: Path, max_chars: int | None = None) -> str:
    text = path.read_text(encoding="utf-8", errors="ignore")
    if max_chars is not None and len(text) > max_chars:
        return text[:max_chars] + "\n\n...[truncated]..."
    return text


@mcp.tool()
def scan_project(project_path: str = ".") -> dict[str, Any]:
    """Return a compact project tree excluding cache/build folders."""
    root = _resolve_project_path(project_path)

    items: list[str] = []
    for path in _iter_project_files(root):
        items.append(str(path.relative_to(root)))

    return {
        "project_path": str(root),
        "file_count": len(items),
        "files": sorted(items),
    }


@mcp.tool()
def search_code(project_path: str = ".", query: str = "", max_results: int = 50) -> dict[str, Any]:
    """Search text inside project files."""
    root = _resolve_project_path(project_path)
    if not query:
        raise ValueError("query is required")

    pattern = re.compile(re.escape(query), re.IGNORECASE)
    results: list[dict[str, Any]] = []

    for path in _iter_project_files(root, TEXT_EXTENSIONS):
        text = _read_text_file(path)

        for line_number, line in enumerate(text.splitlines(), start=1):
            if pattern.search(line):
                results.append({
                    "file": str(path.relative_to(root)),
                    "line": line_number,
                    "text": line.strip(),
                })
                break

        if len(results) >= max_results:
            break

    return {
        "query": query,
        "count": len(results),
        "results": results,
    }


@mcp.tool()
def read_project_file(project_path: str = ".", path: str = "", max_chars: int = 20000) -> dict[str, Any]:
    """Read a text file inside the project."""
    if not path:
        raise ValueError("path is required")

    file_path = _resolve_inside_project(project_path, path)
    if not file_path.exists() or not file_path.is_file():
        raise FileNotFoundError(str(file_path))

    return {
        "path": str(file_path),
        "content": _read_text_file(file_path, max_chars=max_chars),
    }


@mcp.tool()
def list_new_files(project_path: str = ".") -> dict[str, Any]:
    """Return git status --short for the project."""
    root = _resolve_project_path(project_path)

    result = subprocess.run(
        ["git", "-C", str(root), "status", "--short"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )

    return {
        "project_path": str(root),
        "returncode": result.returncode,
        "stdout": result.stdout.strip(),
        "stderr": result.stderr.strip(),
    }


@mcp.tool()
def find_specification_logic(project_path: str = ".", max_results: int = 100) -> dict[str, Any]:
    """Find files and lines related to specification logic."""
    root = _resolve_project_path(project_path)

    terms = [
        "specification",
        "specifications",
        "spec",
        "спецификац",
        "spec_preview",
        "spec_block",
        "build_specification",
    ]

    results: list[dict[str, Any]] = []

    for path in _iter_project_files(root, TEXT_EXTENSIONS):
        text = _read_text_file(path)

        for line_number, line in enumerate(text.splitlines(), start=1):
            lower = line.lower()
            if any(term.lower() in lower for term in terms):
                results.append({
                    "file": str(path.relative_to(root)),
                    "line": line_number,
                    "text": line.strip(),
                })
                break

        if len(results) >= max_results:
            break

    return {
        "terms": terms,
        "count": len(results),
        "results": results,
    }


@mcp.tool()
def read_pdf_text(project_path: str = ".", path: str = "", max_chars: int = 20000) -> dict[str, Any]:
    """Extract text from a PDF inside the project."""
    if not path:
        raise ValueError("path is required")

    pdf_path = _resolve_inside_project(project_path, path)
    if not pdf_path.exists() or not pdf_path.is_file():
        raise FileNotFoundError(str(pdf_path))

    doc = fitz.open(pdf_path)
    pages: list[str] = []

    for index, page in enumerate(doc, start=1):
        pages.append(f"\n--- Page {index} ---\n{page.get_text()}")

    text = "\n".join(pages)
    if len(text) > max_chars:
        text = text[:max_chars] + "\n\n...[truncated]..."

    return {
        "path": str(pdf_path),
        "pages": len(doc),
        "text": text,
    }


@mcp.tool()
def analyze_price_pdf(project_path: str = ".", path: str = "") -> dict[str, Any]:
    """Analyze whether a PDF looks like a price list."""
    if not path:
        raise ValueError("path is required")

    pdf_path = _resolve_inside_project(project_path, path)
    if not pdf_path.exists() or not pdf_path.is_file():
        raise FileNotFoundError(str(pdf_path))

    doc = fitz.open(pdf_path)
    text = "\n".join(page.get_text() for page in doc)

    keywords = ["price", "model", "eur", "usd", "cooling", "capacity", "kw", "прайс", "цена", "модель"]
    found = [keyword for keyword in keywords if keyword in text.lower()]

    return {
        "path": str(pdf_path),
        "pages": len(doc),
        "first_3000_chars": text[:3000],
        "found_keywords": found,
        "looks_like_price_list": any(k in found for k in ["price", "eur", "usd", "прайс", "цена"]),
    }


if __name__ == "__main__":
    mcp.run()