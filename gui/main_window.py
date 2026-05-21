from __future__ import annotations

APP_FOOTER = """
Направления:
Stulz · Riello · DC Eltek · Generator

Разработчик:
Литвинов Виталий Константинович
"""

import sys
from pathlib import Path

from openpyxl import load_workbook

from brands.registry import BRANDS, get_brand_module
from core.excel_reader import list_sheets
from core.models import OfferContext
from core.project_scanner import scan_project_files


def run_gui() -> None:
    try:
        from PySide6.QtCore import Qt, QSettings, QUrl
        from PySide6.QtGui import QDesktopServices, QFont, QIcon
        from PySide6.QtWidgets import (
            QApplication,
            QComboBox,
            QFileDialog,
            QFrame,
            QGridLayout,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QScrollArea,
            QSizePolicy,
            QSpacerItem,
            QTextEdit,
            QVBoxLayout,
            QWidget,
        )
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("Для запуска GUI установите PySide6: pip install PySide6") from exc

    class OfferGeneratorWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.settings = QSettings("SAM Group", "SAM Offer Generator")
            self.setWindowTitle("SAM Offer Generator")
            self.setMinimumSize(900, 620)
            self.setWindowIcon(QIcon())

            self.project_edit = QLineEdit(self._saved("project_dir", ""))
            self.client_edit = QLineEdit(self._saved("client", "ТОО Example"))
            self.brand_combo = QComboBox()
            self.brand_combo.addItems(BRANDS.keys())
            self.brand_combo.setCurrentText(self._saved("brand", "Stulz"))
            self.calc_combo = QComboBox()
            self.calc_combo.setEditable(True)
            self.template_combo = QComboBox()
            self.template_combo.setEditable(True)
            self.sheet_combo = QComboBox()
            self.sheet_combo.setEditable(True)
            self.output_edit = QLineEdit(self._saved("output_dir", ""))
            self.preview = QTextEdit()
            self.preview.setReadOnly(True)
            self.status_label = QLabel("Выберите папку проекта")
            self._auto_client_value = ""

            self._base_font_size = 10
            self._last_scale = 0.0
            self._responsive_widgets: list[QWidget] = []

            self._build_ui()
            self._apply_style()
            self._autofill_client_from_project_dir()
            self._autofill_brand_from_project_dir()
            self._scan_project()
            self._refresh_preview()

        def _saved(self, key: str, default: str) -> str:
            value = self.settings.value(key, default)
            return str(value) if value is not None else default

        def _build_ui(self) -> None:
            central = QWidget()
            root = QHBoxLayout(central)
            root.setContentsMargins(0, 0, 0, 0)
            root.setSpacing(0)

            self.sidebar = QFrame()
            sidebar = self.sidebar
            sidebar.setObjectName("Sidebar")
            sidebar.setMinimumWidth(220)
            side = QVBoxLayout(sidebar)
            side.setContentsMargins(28, 34, 28, 28)
            side.setSpacing(18)

            brand = QLabel("SAM\nGROUP")
            brand.setObjectName("Brand")
            title = QLabel("Offer Generator")
            title.setObjectName("SideTitle")
            subtitle = QLabel("Папка проекта → расчет Excel → шаблон Word → готовое КП")
            subtitle.setObjectName("SideSubtitle")
            subtitle.setWordWrap(True)
            badge = QLabel("Project folder workflow")
            badge.setObjectName("Badge")
            badge.setAlignment(Qt.AlignCenter)

            side.addWidget(brand)
            side.addWidget(title)
            side.addWidget(subtitle)
            side.addSpacing(12)
            side.addWidget(badge)
            side.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding))
            side.addWidget(QLabel(APP_FOOTER))

            self.content = QFrame()
            content = self.content
            content.setObjectName("Content")
            content.setMinimumWidth(560)
            content_layout = QVBoxLayout(content)
            content_layout.setContentsMargins(34, 28, 34, 28)
            content_layout.setSpacing(18)

            header = QHBoxLayout()
            h_text = QVBoxLayout()
            page_title = QLabel("Новое коммерческое предложение")
            page_title.setObjectName("PageTitle")
            page_subtitle = QLabel("Сначала выберите папку проекта на сервере. Программа найдет Excel и Word внутри папки.")
            page_subtitle.setObjectName("PageSubtitle")
            h_text.addWidget(page_title)
            h_text.addWidget(page_subtitle)
            self.generate_btn = QPushButton("Сформировать КП")
            self.generate_btn.setObjectName("PrimaryButton")
            self.generate_btn.clicked.connect(self._generate)
            header.addLayout(h_text, stretch=1)
            header.addWidget(self.generate_btn, stretch=0, alignment=Qt.AlignTop)
            content_layout.addLayout(header)

            form_card = self._card("Папка проекта и файлы")
            grid = QGridLayout()
            form_card.layout().addLayout(grid)
            grid.setColumnStretch(1, 1)
            grid.setVerticalSpacing(12)
            grid.setHorizontalSpacing(10)

            self._add_row(grid, 0, "Папка проекта", self.project_edit, "Выбрать", self._browse_project_dir)
            self._add_row(grid, 1, "Направление", self.brand_combo, None, None)
            self._add_row(grid, 2, "Клиент", self.client_edit, None, None)
            self._add_row(grid, 3, "Excel-расчет", self.calc_combo, "Обновить", self._scan_project)
            self._add_row(grid, 4, "Word-шаблон", self.template_combo, "Обзор", self._browse_template_file)
            self._add_row(grid, 5, "Лист Excel", self.sheet_combo, "Листы", self._load_sheets)
            self._add_row(grid, 6, "Папка результата", self.output_edit, "Выбрать", self._browse_output_dir)
            content_layout.addWidget(form_card)

            bottom = QHBoxLayout()
            preview_card = self._card("Проверка данных")
            preview_card.layout().addWidget(self.preview)
            bottom.addWidget(preview_card, stretch=2)

            status_card = self._card("Статус")
            status_text = QLabel(
                "1. Выберите папку проекта\n"
                "2. Выберите направление\n"
                "3. Выберите Excel и Word\n"
                "4. Нажмите \"Сформировать КП\"\n\n"
                "Результат сохраняется в выбранную папку результата."
            )
            status_text.setWordWrap(True)
            status_card.layout().addWidget(status_text)
            status_card.layout().addWidget(self.status_label)
            bottom.addWidget(status_card, stretch=1)
            content_layout.addLayout(bottom)

            scroll = QScrollArea()
            scroll.setObjectName("ContentScroll")
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setWidget(content)

            root.addWidget(sidebar)
            root.addWidget(scroll, stretch=1)
            self.setCentralWidget(central)
            self._apply_responsive_metrics(force=True)

            self.project_edit.textChanged.connect(self._on_project_dir_changed)
            self.calc_combo.currentTextChanged.connect(self._load_sheets)
            self.calc_combo.currentTextChanged.connect(self._refresh_preview)
            self.template_combo.currentTextChanged.connect(self._refresh_preview)
            self.sheet_combo.currentTextChanged.connect(self._refresh_preview)
            self.brand_combo.currentTextChanged.connect(self._refresh_preview)
            self.client_edit.textChanged.connect(self._refresh_preview)
            self.output_edit.textChanged.connect(self._refresh_preview)

        def _add_row(self, grid, row: int, label: str, widget, button_text: str | None, command) -> None:
            lab = QLabel(label)
            lab.setObjectName("FormLabel")
            lab.setMinimumWidth(110)
            lab.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

            widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            widget.setMinimumWidth(0)
            if isinstance(widget, QComboBox):
                widget.setMinimumContentsLength(1)
                widget.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
            self._responsive_widgets.append(widget)

            grid.addWidget(lab, row, 0)
            grid.addWidget(widget, row, 1)
            if button_text:
                btn = QPushButton(button_text)
                btn.setObjectName("GhostButton")
                btn.setMinimumWidth(92)
                btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                if command:
                    btn.clicked.connect(command)
                self._responsive_widgets.append(btn)
                grid.addWidget(btn, row, 2)
            else:
                grid.setColumnMinimumWidth(2, 92)

        def _card(self, title: str) -> QFrame:
            frame = QFrame()
            frame.setObjectName("Card")
            layout = QVBoxLayout(frame)
            layout.setContentsMargins(20, 18, 20, 20)
            layout.setSpacing(12)
            label = QLabel(title)
            label.setObjectName("CardTitle")
            layout.addWidget(label)
            return frame

        def _ui_scale(self) -> float:
            width = max(self.width(), 900)
            height = max(self.height(), 620)
            return max(0.86, min(1.25, min(width / 1440, height / 900)))

        def _apply_responsive_metrics(self, force: bool = False) -> None:
            scale = self._ui_scale()
            if not force and abs(scale - self._last_scale) < 0.03:
                return
            self._last_scale = scale

            sidebar_width = int(max(220, min(320, self.width() * 0.22)))
            self.sidebar.setFixedWidth(sidebar_width)

            content_margin_x = int(22 * scale)
            content_margin_y = int(20 * scale)
            self.content.layout().setContentsMargins(content_margin_x, content_margin_y, content_margin_x, content_margin_y)
            self.content.layout().setSpacing(int(16 * scale))

            min_h = int(38 * scale)
            for widget in self._responsive_widgets:
                widget.setMinimumHeight(min_h)

            self.generate_btn.setMinimumWidth(int(220 * scale))
            self.generate_btn.setMinimumHeight(int(42 * scale))
            self._apply_style(scale)

        def resizeEvent(self, event) -> None:  # noqa: N802 - Qt API name
            super().resizeEvent(event)
            self._apply_responsive_metrics()

        def _apply_style(self, scale: float | None = None) -> None:
            scale = scale if scale is not None else self._ui_scale()
            app = QApplication.instance()
            if app:
                app.setFont(QFont("Segoe UI", max(9, int(10 * scale))))

            def px(value: int) -> int:
                return max(1, int(value * scale))

            self.setStyleSheet(f"""
                QMainWindow {{ background: #F4F6F8; }}
                #ContentScroll {{ background: #F4F6F8; border: none; }}
                #Sidebar {{ background: #15171B; color: #FFFFFF; }}
                #Content {{ background: #F4F6F8; }}
                #Brand {{ color: #FFFFFF; font-size: {px(34)}px; font-weight: 900; letter-spacing: 2px; }}
                #SideTitle {{ color: #FFFFFF; font-size: {px(23)}px; font-weight: 700; }}
                #SideSubtitle {{ color: #B8C0CC; font-size: {px(13)}px; line-height: 1.4; }}
                #Badge {{ background: #D71920; color: white; border-radius: {px(16)}px; padding: {px(8)}px {px(12)}px; font-weight: 700; }}
                #PageTitle {{ color: #171A1F; font-size: {px(28)}px; font-weight: 800; }}
                #PageSubtitle {{ color: #667085; font-size: {px(13)}px; }}
                #Card {{ background: #FFFFFF; border: 1px solid #E7EAF0; border-radius: {px(18)}px; }}
                #CardTitle {{ color: #171A1F; font-size: {px(15)}px; font-weight: 800; }}
                #FormLabel {{ color: #344054; font-weight: 700; font-size: {px(12)}px; }}
                QLineEdit, QComboBox, QTextEdit {{
                    background-color: #FFFFFF;
                    color: #101828;
                    border: 2px solid #D0D5DD;
                    border-radius: {px(12)}px;
                    padding: 0 {px(12)}px;
                    min-height: {px(38)}px;
                    font-size: {px(13)}px;
                    selection-background-color: #D71920;
                }}
                QLineEdit:focus, QComboBox:focus, QTextEdit:focus {{ border: 2px solid #D71920; }}
                QTextEdit {{ min-height: {px(150)}px; padding: {px(12)}px {px(14)}px;}}
                QPushButton {{ border-radius: {px(11)}px; padding: {px(9)}px {px(14)}px; font-weight: 800; font-size: {px(12)}px; }}
                #PrimaryButton {{ background: #D71920; color: white; border: 1px solid #D71920; }}
                #PrimaryButton:hover {{ background: #B9151B; }}
                #GhostButton {{ background: #FFFFFF; color: #1D2939; border: 1px solid #D0D5DD; }}
                #GhostButton:hover {{ border: 1px solid #D71920; color: #D71920; }}
                QLabel {{ color: #475467; font-size: {px(12)}px; }}
            """)

        def _split_project_path(self, path_text: str) -> list[str]:
            raw = path_text.strip().strip('"')
            if not raw:
                return []

            normalized = raw.replace("\\", "/")
            return [part.strip() for part in normalized.split("/") if part.strip()]

        def _extract_client_from_project_dir(self, path_text: str) -> str:
            """Return client name from server project path.

            Expected server structure:
            //Diskstationnew/Exchange/01_Work/01_STULZ/02_Projects/Client/2206/Project
            or:
            /diskstationnew/exchange/01_Work/01_STULZ/02_Projects/Client/...
            """
            parts = self._split_project_path(path_text)
            if not parts:
                return ""

            lowered = [part.lower() for part in parts]

            project_markers = {"02_projects", "2_projects", "projects", "проекты"}
            for index, part in enumerate(lowered):
                if part in project_markers and index + 1 < len(parts):
                    return parts[index + 1]

            # Fallback for the standard SAM structure if the marker is absent:
            # ... / 01_STULZ / Client / project-code / project-name
            for index, part in enumerate(lowered):
                if "stulz" in part and index + 2 < len(parts):
                    candidate = parts[index + 2]
                    if candidate and not candidate.lower().endswith("projects"):
                        return candidate

            return ""

        def _extract_brand_from_project_dir(self, path_text: str) -> str:
            """Return brand/direction from the standard SAM folder path."""
            parts = self._split_project_path(path_text)
            if not parts:
                return ""

            brand_rules = (
                ("stulz", "Stulz"),
                ("riello", "Riello"),
                ("dc_eltek", "DC Eltek"),
                ("dc eltek", "DC Eltek"),
                ("eltek", "DC Eltek"),
                ("generator", "Generator"),
                ("generators", "Generator"),
            )

            for part in parts:
                clean = part.lower().replace("-", "_").replace(" ", "_")
                for marker, brand in brand_rules:
                    marker_clean = marker.replace(" ", "_")
                    if marker_clean in clean and self.brand_combo.findText(brand) >= 0:
                        return brand

            return ""

        def _autofill_brand_from_project_dir(self) -> None:
            brand = self._extract_brand_from_project_dir(self.project_edit.text())
            if not brand or self.brand_combo.currentText() == brand:
                return

            self.brand_combo.blockSignals(True)
            self.brand_combo.setCurrentText(brand)
            self.brand_combo.blockSignals(False)
            self._refresh_preview()

        def _on_project_dir_changed(self) -> None:
            self._autofill_client_from_project_dir()
            self._autofill_brand_from_project_dir()
            self._scan_project()

        def _autofill_client_from_project_dir(self, force: bool = False) -> None:
            client = self._extract_client_from_project_dir(self.project_edit.text())
            if not client:
                return

            current = self.client_edit.text().strip()
            should_update = (
                force
                or not current
                or current == "ТОО Example"
                or current == self._auto_client_value
            )
            if not should_update:
                return

            self.client_edit.blockSignals(True)
            self.client_edit.setText(client)
            self.client_edit.blockSignals(False)
            self._auto_client_value = client
            self._refresh_preview()

        def _browse_project_dir(self) -> None:
            path = QFileDialog.getExistingDirectory(self, "Выберите папку проекта", self.project_edit.text())
            if path:
                self.project_edit.setText(path)
                self.output_edit.setText(path)
                self._autofill_client_from_project_dir(force=True)
                self._autofill_brand_from_project_dir()
                self._scan_project()

        def _browse_output_dir(self) -> None:
            path = QFileDialog.getExistingDirectory(self, "Выберите папку результата", self.output_edit.text() or self.project_edit.text())
            if path:
                self.output_edit.setText(path)

        def _browse_template_file(self) -> None:
            path, _ = QFileDialog.getOpenFileName(self, "Выберите Word-шаблон", self.project_edit.text(), "Word (*.docx)")
            if path:
                if self.template_combo.findText(path) < 0:
                    self.template_combo.addItem(path)
                self.template_combo.setCurrentText(path)

        def _scan_project(self) -> None:
            project_dir = Path(self.project_edit.text().strip()) if self.project_edit.text().strip() else None
            if not project_dir or not project_dir.exists():
                self.status_label.setText("Папка проекта не выбрана")
                return

            found = scan_project_files(project_dir)
            old_calc = self.calc_combo.currentText()
            old_template = self.template_combo.currentText()

            self.calc_combo.blockSignals(True)
            self.template_combo.blockSignals(True)
            self.calc_combo.clear()
            self.template_combo.clear()
            self.calc_combo.addItems([str(p) for p in found["excel"]])
            self.template_combo.addItems([str(p) for p in found["word"]])
            if old_calc and self.calc_combo.findText(old_calc) >= 0:
                self.calc_combo.setCurrentText(old_calc)
            if old_template and self.template_combo.findText(old_template) >= 0:
                self.template_combo.setCurrentText(old_template)
            self.calc_combo.blockSignals(False)
            self.template_combo.blockSignals(False)

            if not self.output_edit.text().strip():
                self.output_edit.setText(str(project_dir))

            self.status_label.setText(f"Найдено Excel: {len(found['excel'])}, Word: {len(found['word'])}")
            self._load_sheets()
            self._refresh_preview()

        def _load_sheets(self) -> None:
            current = self.sheet_combo.currentText().strip()
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            try:
                calc_path = Path(self.calc_combo.currentText().strip())
                if calc_path.exists():
                    sheets = list_sheets(calc_path)
                    self.sheet_combo.addItems(sheets)
                    if current and current in sheets:
                        self.sheet_combo.setCurrentText(current)
            except Exception:
                self.sheet_combo.addItem("")
            finally:
                self.sheet_combo.blockSignals(False)
            self._refresh_preview()

        def _make_context(self) -> OfferContext:
            project_dir = Path(self.project_edit.text().strip())
            output_dir = Path(self.output_edit.text().strip() or project_dir)
            pdf_dir = project_dir if project_dir.exists() else None
            return OfferContext(
                brand=self.brand_combo.currentText(),
                project_dir=project_dir,
                template_path=Path(self.template_combo.currentText().strip()),
                calc_path=Path(self.calc_combo.currentText().strip()),
                output_dir=output_dir,
                client_name=self.client_edit.text().strip() or "Client",
                sheet_name=self.sheet_combo.currentText().strip() or None,
                pdf_dir=pdf_dir,
            )

        def _validate_context(self, context: OfferContext) -> None:
            if not context.project_dir.exists():
                raise FileNotFoundError("Выберите существующую папку проекта.")
            if not context.calc_path.exists():
                raise FileNotFoundError("Выберите существующий Excel-файл калькуляции.")
            if not context.template_path.exists():
                raise FileNotFoundError("Выберите существующий Word-шаблон.")
            if context.template_path.suffix.lower() != ".docx":
                raise ValueError("Word-шаблон должен быть файлом .docx")

        def _refresh_preview(self) -> None:
            try:
                context = self._make_context()
                if not context.calc_path.exists():
                    self.preview.setPlainText("Excel-файл пока не выбран или не найден.")
                    return
                module = get_brand_module(context.brand)
                self.preview.setPlainText(module.preview(context))
            except Exception as exc:
                self.preview.setPlainText(f"Не удалось прочитать данные: {exc}")

        def _remember_values(self) -> None:
            self.settings.setValue("project_dir", self.project_edit.text())
            self.settings.setValue("client", self.client_edit.text())
            self.settings.setValue("brand", self.brand_combo.currentText())
            self.settings.setValue("output_dir", self.output_edit.text())

        def _generate(self) -> None:
            try:
                self.generate_btn.setEnabled(False)
                self.status_label.setText("Формирую документ...")
                QApplication.processEvents()

                context = self._make_context()
                self._validate_context(context)
                module = get_brand_module(context.brand)
                out = module.make_offer(context)

                self._remember_values()
                self.status_label.setText(f"Готово: {out.name}")

                msg = QMessageBox(self)
                msg.setWindowTitle("SAM Offer Generator")
                msg.setIcon(QMessageBox.Question)
                msg.setText("Коммерческое предложение успешно сформировано.")
                msg.setInformativeText(str(out))
                open_folder_btn = msg.addButton("Открыть папку", QMessageBox.ActionRole)
                open_file_btn = msg.addButton("Открыть КП", QMessageBox.ActionRole)
                msg.addButton("Закрыть", QMessageBox.RejectRole)
                msg.exec()
                clicked = msg.clickedButton()
                if clicked == open_folder_btn:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(out.parent)))
                elif clicked == open_file_btn:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(str(out)))

            except Exception as exc:
                self.status_label.setText("Ошибка формирования")
                QMessageBox.critical(self, "Ошибка", str(exc))
            finally:
                self.generate_btn.setEnabled(True)
                self._refresh_preview()

    app = QApplication.instance() or QApplication(sys.argv)
    window = OfferGeneratorWindow()
    window.show()
    app.exec()
