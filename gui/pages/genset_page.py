from __future__ import annotations

from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget

from gui.project_folder_selector import ProjectFolderSelector


class GensetPage(QWidget):
    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        self.project_selector = ProjectFolderSelector(
            owner,
            owner.settings,
            settings_key="genset/project_dir",
            available_brands=("Stulz", "Riello", "DC Eltek", "Generator", "HVAC", "Battery"),
            initial_path=str(owner.settings.value("genset/project_dir", "") or ""),
            parent=self,
        )
        layout.addWidget(self.project_selector.widget)
        card = owner._card("Genset")
        label = QLabel("Логика Genset будет подключена следующим этапом.")
        label.setWordWrap(True)
        card.layout().addWidget(label)
        layout.addWidget(card)
        layout.addStretch(1)

    def project_path_text(self) -> str:
        return self.project_selector.project_path_text()

    def clear_cache(self) -> None:
        self.project_selector.set_project_path("")
        self.owner.settings.remove("genset/project_dir")
