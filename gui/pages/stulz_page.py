from __future__ import annotations

from PySide6.QtWidgets import QGridLayout, QVBoxLayout, QWidget


class StulzPage(QWidget):
    """Страница Stulz. Содержит текущую рабочую логику формы КП."""

    def __init__(self, owner) -> None:
        super().__init__(owner)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        stulz_card = owner._card("Stulz: файлы и параметры КП")
        grid = QGridLayout()
        stulz_card.layout().addLayout(grid)
        grid.setColumnStretch(1, 1)
        grid.setVerticalSpacing(12)
        grid.setHorizontalSpacing(10)

        owner._add_row(grid, 0, "Клиент", owner.client_edit, None, None)
        owner._add_row(grid, 1, "Excel-расчет", owner.calc_combo, "Обновить", lambda: owner._scan_project(force=True))
        owner._add_row(grid, 2, "Лист Excel", owner.sheet_combo, "Листы", owner._load_sheets)
        owner._add_row(grid, 3, "Папка спецификаций", owner.spec_edit, "Выбрать", owner._browse_spec_dir)
        owner._add_row(grid, 4, "Папка результата", owner.output_edit, "Выбрать", owner._browse_output_dir)

        layout.addWidget(stulz_card)
