from __future__ import annotations

from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget


class GensetPage(QWidget):
    def __init__(self, owner) -> None:
        super().__init__(owner)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        card = owner._card("Genset")
        label = QLabel("Логика Genset будет подключена следующим этапом.")
        label.setWordWrap(True)
        card.layout().addWidget(label)
        layout.addWidget(card)
        layout.addStretch(1)
