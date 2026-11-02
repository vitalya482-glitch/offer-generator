from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, TypeVar

from PySide6.QtCore import QObject, Signal
from PySide6.QtWidgets import QFileDialog, QGridLayout, QLineEdit

from core.project_scanner import clear_scan_cache, scan_project_files
from gui.path_helpers import (
    extract_brand_from_project_dir,
    extract_client_from_project_dir,
    infer_output_dir,
    infer_specifications_dir,
)

T = TypeVar("T")


@dataclass(frozen=True, slots=True)
class ProjectFolderInfo:
    """Common information derived from a selected project folder."""

    path: str
    client_name: str
    brand_name: str
    output_dir: str
    specifications_dir: str


class ProjectFolderSelector(QObject):
    """Reusable project-folder card and shared project-folder logic.

    A page supplies its settings key and available brands. The selector owns
    the folder row, remembers the full path, derives common project metadata,
    and can pass a temporary scan result to a page callback. The callback must
    consume results synchronously; their lists are cleared immediately after it
    returns so pages keep only data they actually display or need later.
    """

    projectSelected = Signal(str, str)  # old path, new path
    pathEdited = Signal(str)

    def __init__(
        self,
        owner,
        settings,
        *,
        settings_key: str,
        available_brands: tuple[str, ...] | list[str],
        initial_path: str = "",
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self.owner = owner
        self.settings = settings
        self.settings_key = settings_key
        self.available_brands = tuple(available_brands)
        self._updating_display = False
        self._path = str(initial_path or "").strip()

        self.widget = owner._card("Папка проекта")
        grid = QGridLayout()
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)
        self.widget.layout().addLayout(grid)

        self.path_edit = QLineEdit(owner._display_dir(self._path))
        self.path_edit.setToolTip(self._path)
        owner._add_row(grid, 0, "Папка проекта", self.path_edit, "Выбрать", self.browse)
        self.path_edit.textChanged.connect(self._on_text_changed)

    def project_path_text(self) -> str:
        return self._path or self.path_edit.toolTip() or self.path_edit.text().strip()

    def client_name(self) -> str:
        return extract_client_from_project_dir(self.project_path_text())

    def brand_name(self) -> str:
        return extract_brand_from_project_dir(self.project_path_text(), self.available_brands)

    def project_info(self, pdf_dirs=()) -> ProjectFolderInfo:
        path = self.project_path_text().strip()
        return ProjectFolderInfo(
            path=path,
            client_name=extract_client_from_project_dir(path),
            brand_name=extract_brand_from_project_dir(path, self.available_brands),
            output_dir=infer_output_dir(path) if path else "",
            specifications_dir=infer_specifications_dir(path, pdf_dirs) if path else "",
        )

    def set_project_path(self, path: str) -> None:
        self._path = str(path or "").strip()
        self._updating_display = True
        try:
            self.owner._set_line_path(self.path_edit, self._path, is_file=False)
        finally:
            self._updating_display = False
        self.settings.setValue(self.settings_key, self._path)

    def browse(self) -> None:
        old_path = self.project_path_text().strip()
        start = old_path or str(Path.home())
        path = QFileDialog.getExistingDirectory(self.path_edit, "Выберите папку проекта", start)
        if not path:
            return
        self.set_project_path(path)
        self.projectSelected.emit(old_path, path)

    def scan(self, consumer: Callable[[dict[str, list[Path]]], T], *, force: bool = False) -> T:
        """Scan the selected project and release temporary result containers."""
        path = self.project_path_text().strip()
        if not path:
            raise ValueError("Папка проекта не выбрана")
        if force:
            clear_scan_cache()

        result = scan_project_files(Path(path), use_cache=not force)
        try:
            return consumer(result)
        finally:
            for paths in result.values():
                paths.clear()
            result.clear()

    def scan_async(self, consumer: Callable[[dict[str, list[Path]]], T], *, force: bool = False) -> None:
        """Scan a project in a worker thread, then deliver its result on the GUI thread."""
        path = self.project_path_text().strip()
        if not path:
            raise ValueError("Папка проекта не выбрана")

        def scan_snapshot() -> dict[str, list[Path]]:
            if force:
                clear_scan_cache()
            return scan_project_files(Path(path), use_cache=not force)

        def deliver(result: dict[str, list[Path]]) -> None:
            try:
                consumer(result)
            finally:
                for paths in result.values():
                    paths.clear()
                result.clear()

        self.owner.run_background(scan_snapshot, deliver)

    def _on_text_changed(self, text: str) -> None:
        if self._updating_display:
            return
        self._path = str(text or "").strip()
        self.path_edit.setToolTip(self._path)
        self.settings.setValue(self.settings_key, self._path)
        self.pathEdited.emit(self._path)
