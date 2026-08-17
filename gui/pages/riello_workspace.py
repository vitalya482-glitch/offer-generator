from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import QLabel, QTabWidget, QVBoxLayout, QWidget

from core.runtime_paths import resource_path
from gui.pages.riello_offer_page import RielloOfferPage
from gui.pages.riello_page import RielloPage as RielloSelectionPage


_DEFAULT_RIELLO_WORD_TEMPLATE = Path("templates") / "riello" / "Offer_Company_17-08-26_TAGGED_UPS.docx"


class RielloOfferStagePage(RielloOfferPage):
    """Riello Word stage with the bundled tagged template as the default."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.ensure_default_word_template()

    def ensure_default_word_template(self) -> None:
        template_path = resource_path(_DEFAULT_RIELLO_WORD_TEMPLATE)
        if not template_path.exists():
            return

        combo = self.template_combo
        current = self._path_from_combo(combo).strip()
        index = self.owner._find_combo_path(combo, str(template_path))
        if index < 0:
            self.owner._add_path_item(combo, str(template_path), is_file=True)
            index = combo.count() - 1

        # Prefer a manually/project-selected Word file. When there is no valid
        # current choice, use the bundled Riello template automatically.
        if not current or not Path(current).exists():
            combo.setCurrentIndex(index)
            self._set_saved("riello_offer/template_path", str(template_path))
            self.settings.sync()

    def scan_project(self, force: bool = False) -> None:
        super().scan_project(force=force)
        self.ensure_default_word_template()


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
        self.offer_page = RielloOfferStagePage(owner)

        self._polish_selection_stage()

        self.tabs.addTab(self.selection_page, "1. Подбор оборудования / Calc")
        self.tabs.addTab(self.offer_page, "2. Формирование КП")
        self.tabs.currentChanged.connect(self._on_stage_changed)
        layout.addWidget(self.tabs)

        saved_stage = str(owner.settings.value("riello/workspace_tab", "0") or "0")
        try:
            self.tabs.setCurrentIndex(max(0, min(1, int(saved_stage))))
        except Exception:
            self.tabs.setCurrentIndex(0)

    def _polish_selection_stage(self) -> None:
        """Use specification wording without rewriting the mature picker code."""

        page = self.selection_page
        try:
            page.add_ups_btn.setText("+ Добавить в спецификацию")
        except Exception:
            pass
        try:
            page.add_option_btn.setText("+ Добавить опцию")
        except Exception:
            pass

        # The existing quote table already behaves as the Calc specification.
        # Rename only its visible card title/status wording.
        for label in page.findChildren(QLabel):
            text = label.text().strip()
            if text == "Лист подбора оборудования":
                label.setText("Спецификация Riello")
            elif text == "Выберите папку проекта и добавьте оборудование в лист подбора":
                label.setText("Выберите папку проекта и добавьте оборудование в спецификацию")

    def _on_stage_changed(self, index: int) -> None:
        self.owner.settings.setValue("riello/workspace_tab", index)
        self.owner.settings.sync()
        if index == 1:
            try:
                self.offer_page.adopt_project(self.selection_page.project_path_text())
            except Exception:
                pass
            try:
                self.offer_page.ensure_default_word_template()
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
