from __future__ import annotations

from PySide6.QtWidgets import QTabWidget, QVBoxLayout, QWidget

from gui.pages.riello_offer_page import RielloOfferPage
from gui.pages.riello_page import RielloPage as RielloSelectionPage


class RielloPage(QWidget):
    """Riello workspace with two explicit stages: Calc and Word commercial offer."""

    brand_name = "Riello"

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.selection_page = RielloSelectionPage(owner)
        self.offer_page = RielloOfferPage(owner)

        self.tabs.addTab(self.selection_page, "1. Подбор оборудования / Calc")
        self.tabs.addTab(self.offer_page, "2. Формирование КП")
        self.tabs.currentChanged.connect(self._on_stage_changed)
        layout.addWidget(self.tabs)

        saved_stage = str(owner.settings.value("riello/workspace_tab", "0") or "0")
        try:
            self.tabs.setCurrentIndex(max(0, min(1, int(saved_stage))))
        except Exception:
            self.tabs.setCurrentIndex(0)

    def _on_stage_changed(self, index: int) -> None:
        self.owner.settings.setValue("riello/workspace_tab", index)
        self.owner.settings.sync()
        if index == 1:
            try:
                self.offer_page.adopt_project(self.selection_page.project_path_text())
            except Exception:
                pass
            try:
                self.offer_page.refresh_preview()
            except Exception:
                pass

    def project_path_text(self) -> str:
        current = self.tabs.currentWidget()
        if current is not None and hasattr(current, "project_path_text"):
            try:
                return str(current.project_path_text())
            except Exception:
                pass
        try:
            return str(self.selection_page.project_path_text())
        except Exception:
            return ""

    def output_path_text(self) -> str:
        current = self.tabs.currentWidget()
        if current is not None and hasattr(current, "output_path_text"):
            try:
                return str(current.output_path_text())
            except Exception:
                pass
        return ""

    def clear_cache(self) -> None:
        for page in (self.selection_page, self.offer_page):
            if hasattr(page, "clear_cache"):
                try:
                    page.clear_cache()
                except Exception:
                    pass

    def on_settings_changed(self) -> None:
        for page in (self.selection_page, self.offer_page):
            if hasattr(page, "on_settings_changed"):
                try:
                    page.on_settings_changed()
                except Exception:
                    pass
