from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSize, Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.excel_parser_to_calc import ExcelParserToCalc
from core.excel_calc_maker import ExcelCalcMaker
from core.runtime_paths import resource_path
from gui.pages.riello_offer_page import RielloOfferPage
from gui.pages.riello_page import RielloPage as RielloSelectionPage
from gui.project_folder_selector import ProjectFolderSelector


class _ContentSizedTableWidget(QTableWidget):
    """A table whose layout height follows its header and visible data rows."""

    def sizeHint(self) -> QSize:  # noqa: N802 - Qt API name
        hint = super().sizeHint()
        header_height = self.horizontalHeader().sizeHint().height()
        rows_height = self.verticalHeader().length()
        frame = self.frameWidth() * 2
        return QSize(hint.width(), header_height + rows_height + frame)


class RielloPage(QWidget):
    """Riello workspace split into Calc, offer and price-list stages."""

    brand_name = "Riello"

    def __init__(self, owner) -> None:
        super().__init__(owner)
        self.owner = owner

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        # Keep the existing picker intact in its own third stage. This is the
        # current, already working Riello flow and remains the source for the
        # future Calc assembly stage.
        self.selection_page = RielloSelectionPage(owner)
        self.offer_page = RielloOfferPage(owner)
        self.calc_stage = self._build_calc_stage()

        self.tabs.addTab(self.calc_stage, "1. Собрать расчёт Calc")
        self.tabs.addTab(self.offer_page, "2. Собрать КП")
        self.tabs.addTab(self.selection_page, "3. Позиции по прайсу")
        self.tabs.currentChanged.connect(self._on_stage_changed)
        layout.addWidget(self.tabs)

        saved_stage = str(owner.settings.value("riello/workspace_tab", "0") or "0")
        try:
            self.tabs.setCurrentIndex(max(0, min(2, int(saved_stage))))
        except Exception:
            self.tabs.setCurrentIndex(0)

    def _on_stage_changed(self, index: int) -> None:
        self.owner.settings.setValue("riello/workspace_tab", index)
        self.owner.settings.sync()
        if index == 1:
            try:
                self.offer_page.refresh_preview()
            except Exception:
                pass

    def _build_calc_stage(self) -> QWidget:
        """Build the Calc source selection and position review stage."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(14)

        title = QLabel("Riello: сборка расчёта Calc")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            "Выберите проект, Excel-источник цен и лист. Проверьте найденные позиции "
            "и снимите отметки у тех, которые не нужно включать в расчёт."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        self.calc_selector = ProjectFolderSelector(
            self.owner,
            self.owner.settings,
            settings_key="riello_calc/project_dir",
            available_brands=("Stulz", "Riello", "DC Eltek", "Generator", "HVAC", "Battery"),
            initial_path=str(self.owner.settings.value("riello_calc/project_dir", "") or ""),
            parent=page,
        )
        self.calc_selector.projectSelected.connect(self._on_calc_project_changed)
        self.calc_selector.pathEdited.connect(self._on_calc_project_edited)
        layout.addWidget(self.calc_selector.widget)

        source_card = self.owner._card("Источник цен")
        source_grid = QGridLayout()
        source_grid.setColumnStretch(1, 1)
        source_grid.setVerticalSpacing(12)
        source_grid.setHorizontalSpacing(10)
        source_card.layout().addLayout(source_grid)
        self.price_file_edit = QLineEdit()
        self.price_file_edit.setReadOnly(True)
        self.owner._add_row(source_grid, 0, "Excel-файл с ценами", self.price_file_edit, "Выбрать", self._browse_price_file)
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItem("Сначала выберите Excel-файл", "")
        self.owner._add_row(source_grid, 1, "Лист источника", self.sheet_combo, None, None)
        self.sheet_combo.currentIndexChanged.connect(self._on_sheet_changed)
        layout.addWidget(source_card)

        self._parser = ExcelParserToCalc()
        self._source_path = ""
        saved_source = str(self.owner.settings.value("riello_calc/price_file", "") or "")
        saved_sheet = str(self.owner.settings.value("riello_calc/price_sheet", "") or "")
        if saved_source:
            self._set_price_source(saved_source, preferred_sheet=saved_sheet, silent=True)

        card = self.owner._card("Состав расчёта")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        card_layout = card.layout()
        card_layout.setAlignment(Qt.AlignTop)
        self.parse_button = QPushButton("Собрать данные из прайсов")
        self.parse_button.setObjectName("PrimaryButton")
        self.parse_button.clicked.connect(self._parse_supplier_prices)
        actions = QHBoxLayout()
        actions.setSpacing(10)
        actions.addWidget(self.parse_button)
        self.generate_calc_button = QPushButton("Сформировать Calc из выбранных позиций")
        self.generate_calc_button.setObjectName("PrimaryButton")
        self.generate_calc_button.setEnabled(False)
        self.generate_calc_button.clicked.connect(self._generate_calc)
        actions.addWidget(self.generate_calc_button)
        actions.addStretch(1)
        card_layout.addLayout(actions)

        self.calc_status = QLabel("Выберите Excel-файл с ценами и лист источника.")
        self.calc_status.setObjectName("Hint")
        self.calc_status.setWordWrap(True)
        card_layout.addWidget(self.calc_status)

        self.calc_results = _ContentSizedTableWidget(0, 7)
        self.calc_results.setObjectName("CalcParserResults")
        self.calc_results.setHorizontalHeaderLabels((
            "В расчёт", "Наименование", "Бренд", "Кол-во",
            "Стоимость за единицу", "Сумма", "Условия поставки",
        ))
        self.calc_results.setEditTriggers(QTableWidget.NoEditTriggers)
        self.calc_results.setSelectionBehavior(QTableWidget.SelectRows)
        self.calc_results.setAlternatingRowColors(True)
        header = self.calc_results.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        for col in range(2, 7):
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        self.calc_results.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        card_layout.addWidget(self.calc_results)
        self.calc_results.itemChanged.connect(self._on_calc_item_changed)
        self._calc_rows: list[dict] = []

        summary_title = QLabel("Итоги по включённым позициям")
        summary_title.setObjectName("FormLabel")
        card_layout.addWidget(summary_title)
        self.calc_summary = QPlainTextEdit()
        self.calc_summary.setObjectName("CalcSummary")
        self.calc_summary.setReadOnly(True)
        self.calc_summary.setMaximumHeight(112)
        self.calc_summary.setPlaceholderText("Здесь появятся количество позиций и итоговая стоимость с валютой.")
        card_layout.addWidget(self.calc_summary)
        layout.addWidget(card)
        # The stage lives inside a resizable scroll area. Without a terminal
        # stretch, Qt distributes the viewport's spare height among preferred
        # widgets (labels and cards), making the page look vertically broken.
        # Keep the content at its natural height and let the scroll area handle
        # any overflow instead.
        layout.addStretch(1)
        return page

    def _on_calc_project_changed(self, _old_path: str, _new_path: str) -> None:
        self._clear_price_source()

    def _on_calc_project_edited(self, _path: str) -> None:
        if self._source_path:
            self._clear_price_source()

    def _clear_price_source(self) -> None:
        self._source_path = ""
        self.price_file_edit.clear()
        self.price_file_edit.setToolTip("")
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItem("Сначала выберите Excel-файл", "")
        self.sheet_combo.blockSignals(False)
        self.owner.settings.remove("riello_calc/price_file")
        self.owner.settings.remove("riello_calc/price_sheet")
        self._calc_rows = []
        self.calc_results.setRowCount(0)
        self.calc_summary.clear()
        self.generate_calc_button.setEnabled(False)
        self.calc_status.setText("Выберите Excel-файл с ценами и лист источника.")

    def _browse_price_file(self) -> None:
        project_path = self.calc_selector.project_path_text().strip()
        if not project_path:
            self.calc_status.setText("Сначала выберите папку проекта.")
            return
        try:
            source_paths = self._parser.list_price_sources(project_path)
            start_dir = str(source_paths[0].parent) if source_paths else project_path
        except Exception as exc:
            self.calc_status.setText(f"Не удалось открыть список прайсов: {exc}")
            return
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Excel-файл с ценами",
            start_dir,
            "Excel (*.xlsx *.xlsm)",
        )
        if selected:
            self._set_price_source(selected)

    def _set_price_source(self, file_path: str, *, preferred_sheet: str = "", silent: bool = False) -> None:
        try:
            sources = self._parser.list_price_sources(self.calc_selector.project_path_text())
            resolved = str(Path(file_path).resolve()).casefold()
            if resolved not in {str(path.resolve()).casefold() for path in sources}:
                raise ValueError("Выберите Excel-файл из папки «02_Suppliers» текущего проекта.")
            sheets = self._parser.sheet_names(file_path)
            if not sheets:
                raise ValueError("В выбранном Excel-файле нет листов.")
        except Exception as exc:
            self._source_path = ""
            self.price_file_edit.clear()
            self.sheet_combo.clear()
            self.sheet_combo.addItem("Сначала выберите Excel-файл", "")
            self.owner.settings.remove("riello_calc/price_file")
            self.owner.settings.remove("riello_calc/price_sheet")
            if not silent:
                self.calc_status.setText(f"Не удалось загрузить источник: {exc}")
            return

        self._source_path = str(file_path)
        self.owner._set_line_path(self.price_file_edit, self._source_path, is_file=True)
        self.price_file_edit.setToolTip(self._source_path)
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItem("Выберите лист источника", "")
        for sheet_name in sheets:
            self.sheet_combo.addItem(sheet_name, sheet_name)
        index = self.sheet_combo.findData(preferred_sheet) if preferred_sheet else -1
        self.sheet_combo.setCurrentIndex(index if index >= 0 else 0)
        self.sheet_combo.blockSignals(False)
        self.owner.settings.setValue("riello_calc/price_file", self._source_path)
        if index < 0:
            self.owner.settings.remove("riello_calc/price_sheet")
        if not silent:
            self.calc_status.setText(f"Выберите лист в файле «{Path(file_path).name}».")

    def _on_sheet_changed(self, index: int) -> None:
        sheet_name = str(self.sheet_combo.itemData(index) or "")
        if sheet_name:
            self.owner.settings.setValue("riello_calc/price_sheet", sheet_name)

    def selected_calc_rows(self) -> list[dict]:
        """Return only rows checked for inclusion in the future Calc build."""
        included = []
        for row_no in range(self.calc_results.rowCount()):
            checkbox = self.calc_results.item(row_no, 0)
            if checkbox and checkbox.checkState() == Qt.Checked:
                included.append(dict(self._calc_rows[row_no]))
        return included

    def _parse_supplier_prices(self) -> None:
        project_path = self.calc_selector.project_path_text().strip()
        if not project_path:
            self.calc_status.setText("Ошибка: сначала выберите папку проекта.")
            return
        source_file = self._source_path
        if not source_file:
            self.calc_status.setText("Ошибка: выберите Excel-файл с источником цен.")
            return
        sheet_name = str(self.sheet_combo.currentData() or "")
        if not sheet_name:
            self.calc_status.setText("Ошибка: выберите лист источника цен.")
            return

        self.parse_button.setEnabled(False)
        self.parse_button.setText("Читаю Excel…")
        self.calc_status.setText("Читаю выбранный источник цен…")
        try:
            result = self._parser.parse(project_path, source_file=source_file, sheet_name=sheet_name)
            if not result["ok"]:
                self.calc_status.setText(f"Ошибка: {result['error']}")
                self.calc_results.setRowCount(0)
                self._calc_rows = []
                self.calc_summary.clear()
                self.generate_calc_button.setEnabled(False)
                return
            self._populate_calc_table(result["rows"])
            self.calc_status.setText(f"Источник: «{Path(source_file).name}», лист «{sheet_name}». Снимите отметку, чтобы исключить позицию.")
        except Exception as exc:
            self.calc_status.setText(f"Ошибка разбора прайсов: {exc}")
            self.calc_results.setRowCount(0)
            self._calc_rows = []
            self.calc_summary.clear()
            self.generate_calc_button.setEnabled(False)
        finally:
            self.parse_button.setEnabled(True)
            self.parse_button.setText("Собрать данные из прайсов")

    def _generate_calc(self) -> None:
        rows = self.selected_calc_rows()
        if not rows:
            self.calc_status.setText("Выберите хотя бы одну позицию для расчёта.")
            return
        project = self.calc_selector.project_info()
        if not project.path:
            self.calc_status.setText("Сначала выберите папку проекта.")
            return

        template = resource_path(Path("templates") / "riello" / "calc_08-04-26 UPS.xlsx")
        self.generate_calc_button.setEnabled(False)
        self.generate_calc_button.setText("Формирую Calc…")
        try:
            output = ExcelCalcMaker().make_calc(
                rows,
                brand=self.brand_name,
                client_name=project.client_name or Path(project.path).name,
                template_path=template,
                output_dir=project.output_dir or project.path,
            )
            self.calc_status.setText(f"Calc сформирован: {output}")
            self.calc_status.setToolTip(str(output))
            message = QMessageBox(self)
            message.setWindowTitle("SAM Offer Generator")
            message.setIcon(QMessageBox.Information)
            message.setText("Расчёт собран")
            message.setInformativeText(str(output))
            open_file_button = message.addButton("Открыть расчёт", QMessageBox.ActionRole)
            open_folder_button = message.addButton("Перейти в папку с расчётом", QMessageBox.ActionRole)
            ok_button = message.addButton("ОК", QMessageBox.AcceptRole)
            message.setDefaultButton(ok_button)
            message.exec()
            clicked = message.clickedButton()
            if clicked == open_file_button:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(output)))
            elif clicked == open_folder_button:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(output.parent)))
        except Exception as exc:
            self.calc_status.setText(f"Не удалось сформировать Calc: {exc}")
        finally:
            self.generate_calc_button.setText("Сформировать Calc из выбранных позиций")
            self.generate_calc_button.setEnabled(bool(self.selected_calc_rows()))

    def _populate_calc_table(self, rows: list[dict]) -> None:
        self._calc_rows = [dict(row) for row in rows]
        self.calc_results.blockSignals(True)
        self.calc_results.setRowCount(len(rows))
        for row_no, row in enumerate(rows):
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            checkbox.setCheckState(Qt.Checked)
            checkbox.setTextAlignment(Qt.AlignCenter)
            self.calc_results.setItem(row_no, 0, checkbox)
            values = (
                row.get("name", ""), row.get("brand", ""), self._format_number(row.get("quantity")),
                self._format_number(row.get("unit_price")), self._format_number(row.get("total")),
                row.get("delivery_terms", ""),
            )
            for col, value in enumerate(values, start=1):
                item = QTableWidgetItem(str(value))
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if col in (3, 4, 5):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.calc_results.setItem(row_no, col, item)
        self.calc_results.blockSignals(False)
        self.calc_results.resizeRowsToContents()
        self.calc_results.updateGeometry()
        self.generate_calc_button.setEnabled(bool(self.selected_calc_rows()))
        self._update_calc_summary()

    def _on_calc_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 0:
            return
        included = len(self.selected_calc_rows())
        total = len(getattr(self, "_calc_rows", []))
        self.calc_status.setText(f"Позиций в таблице: {total}; включено: {included}; исключено: {total - included}.")
        self.generate_calc_button.setEnabled(included > 0)
        self._update_calc_summary()

    def _update_calc_summary(self) -> None:
        rows = self.selected_calc_rows()
        quantities = [self._to_float(row.get("quantity")) for row in rows]
        total_quantity = sum(value for value in quantities if value is not None)
        totals_by_currency: dict[str, float] = {}
        for row in rows:
            total = self._to_float(row.get("total"))
            if total is None:
                continue
            currency = str(row.get("currency") or "Валюта не определена")
            totals_by_currency[currency] = totals_by_currency.get(currency, 0.0) + total

        lines = [
            f"Позиций включено: {len(rows)} из {len(self._calc_rows)}",
            f"Общее количество: {self._format_number(total_quantity)}",
            "Итоговая стоимость:",
        ]
        if totals_by_currency:
            lines.extend(f"  {currency}: {self._format_number(total)}" for currency, total in totals_by_currency.items())
        else:
            lines.append("  —")
        self.calc_summary.setPlainText("\n".join(lines))

    @staticmethod
    def _to_float(value) -> float | None:
        if value is None or value == "":
            return None
        try:
            return float(str(value).replace(" ", "").replace(",", "."))
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _format_number(value) -> str:
        if value is None or value == "":
            return ""
        if isinstance(value, (int, float)):
            return f"{value:,.2f}".replace(",", " ").replace(".", ",")
        return str(value)

    def project_path_text(self) -> str:
        current = self.tabs.currentWidget()
        if current is self.calc_stage:
            return self.calc_selector.project_path_text()
        if current is not None and hasattr(current, "project_path_text"):
            try:
                return str(current.project_path_text())
            except Exception:
                pass
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
        self.calc_selector.set_project_path("")
        self.owner.settings.remove("riello_calc/project_dir")
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
