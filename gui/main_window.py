from __future__ import annotations

import sys
from pathlib import Path

from brands.registry import BRANDS
from core.manager_profile import ManagerProfile
from core.project_scanner import clear_scan_cache
from core.runtime_paths import app_icon_path
from gui.ui_style import stylesheet, ui_scale


APP_FOOTER = """
Направления: Stulz · Riello · DC Eltek · Generator
Разработчик: Литвинов Виталий Константинович
"""

SIGNERS = {
    "saniya": {
        "name": "Сания Санаткызы",
        "position": "Коммерческий директор",
    },
    "alisher": {
        "name": "Анаркулов Алишер",
        "position": "Исполнительный директор",
    },
}


def run_gui() -> None:
    try:
        from PySide6.QtCore import Qt, QSettings
        from PySide6.QtGui import QFont, QIcon
        from PySide6.QtWidgets import (
            QApplication,
            QButtonGroup,
            QComboBox,
            QDialog,
            QFrame,
            QGridLayout,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QRadioButton,
            QScrollArea,
            QSizePolicy,
            QSpacerItem,
            QTabWidget,
            QVBoxLayout,
            QWidget,
        )
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("Для запуска GUI установите PySide6: pip install PySide6") from exc

    from gui.settings_dialog import SettingsDialog
    from gui.pages.stulz_page import StulzPage
    from gui.pages.riello_page import RielloPage
    from gui.pages.dc_eltek_page import DcEltekPage
    from gui.pages.battery_page import BatteryPage
    from gui.pages.hvac_page import HVACPage
    from gui.pages.genset_page import GensetPage

    class OfferGeneratorWindow(QMainWindow):
        """Главное окно: только каркас приложения и общие сервисы.

        Вся рабочая область справа принадлежит страницам брендов. MainWindow не хранит
        поля проекта, расчета, шаблонов и не формирует КП — страницы сами собирают
        свои формы, контексты и кнопки действий.
        """

        def __init__(self) -> None:
            super().__init__()
            self.settings = QSettings("SAM Group", "SAM Offer Generator")
            self.setWindowTitle("SAM Offer Generator")
            self.setMinimumSize(900, 620)

            icon_path = app_icon_path()
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))

            self._updating_path_display = False
            self._base_font_size = 10
            self._last_scale = 0.0
            self._responsive_widgets: list[QWidget] = []

            # Данные исполнителя — общие для всех страниц и редактируются в настройках.
            self.manager_name_edit = QLineEdit(self._saved("manager_name", ""))
            self.manager_position_edit = QLineEdit(self._saved("manager_position", ""))
            self.manager_email_edit = QLineEdit(self._saved("manager_email", ""))
            self.manager_phone_edit = QLineEdit(self._saved("manager_phone", ""))

            self.signer_group = QButtonGroup(self)
            self.signer_saniya_radio = QRadioButton("Сания Санаткызы\nКоммерческий директор")
            self.signer_alisher_radio = QRadioButton("Анаркулов Алишер\nИсполнительный директор")
            for radio in (self.signer_saniya_radio, self.signer_alisher_radio):
                radio.setObjectName("SidebarRadio")
                self.signer_group.addButton(radio)
                self._responsive_widgets.append(radio)

            if self._saved("signer_key", "saniya") == "alisher":
                self.signer_alisher_radio.setChecked(True)
            else:
                self.signer_saniya_radio.setChecked(True)

            self.signer_saniya_radio.toggled.connect(self._on_signer_changed)
            self.signer_alisher_radio.toggled.connect(self._on_signer_changed)

            self._build_ui()
            self._apply_style()

        # ------------------------- общие настройки -------------------------
        def _saved(self, key: str, default: str) -> str:
            value = self.settings.value(key, default)
            return str(value) if value is not None else default

        def _clear_cache(self) -> None:
            # Главный экран не знает, какие поля хранит бренд. Он сбрасывает общий кэш
            # сканирования и просит каждую страницу очистить свои данные.
            clear_scan_cache()
            for key in (
                "project_dir",
                "output_dir",
                "calc_path",
                "template_path",
                "sheet_name",
                "spec_dir",
                "template_dir",
                "calc_template_path",
                "calc_template_dir",
            ):
                self.settings.remove(key)

            for page in self._all_brand_pages():
                if hasattr(page, "clear_cache"):
                    try:
                        page.clear_cache()
                    except Exception:
                        pass
            self.settings.sync()

        def _open_settings_dialog(self) -> None:
            dialog = SettingsDialog(self)
            if dialog.exec() == QDialog.Accepted:
                dialog.apply_to_owner()

        def _check_updates(self) -> None:
            try:
                from core.update_client import (
                    UpdateError,
                    build_update_plan,
                    download_update_packages,
                    start_updater,
                )

                plan = build_update_plan()
                release = plan.release
                current = plan.current_version
                latest = plan.latest_version or "неизвестно"

                if not plan.has_update:
                    QMessageBox.information(self, "Обновления", f"Установлена актуальная версия: {current}")
                    return

                if release.app_asset is None:
                    QMessageBox.warning(
                        self,
                        "Обновления",
                        "Новая версия найдена, но в GitHub Release нет App-модуля "
                        "SAM-Offer-Generator-App-No-Runtime.zip",
                    )
                    return

                size_mb = plan.total_size / 1024 / 1024 if plan.total_size else 0
                module_lines = "\n".join(
                    f"- {package.asset.name}: {package.asset.size / 1024 / 1024:.1f} MB"
                    for package in plan.packages
                )
                runtime_note = (
                    "Runtime-модуль тоже будет обновлен, потому что изменились настоящие runtime-файлы "
                    "в папке _internal. Изменения config/prices/templates/assets теперь не считаются "
                    "причиной для скачивания runtime.\n\n"
                    if plan.runtime_required
                    else ""
                )
                question = (
                    f"Доступна новая версия: {latest}\n"
                    f"Текущая версия: {current}\n\n"
                    f"Будет скачано: {size_mb:.1f} MB.\n"
                    f"{module_lines}\n\n"
                    f"{runtime_note}"
                    "После скачивания программа закроется, updater заменит файлы "
                    "и запустит программу снова. Продолжить?"
                )
                answer = QMessageBox.question(
                    self,
                    "Обновление программы",
                    question,
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return

                packages = download_update_packages(plan)
                QMessageBox.information(
                    self,
                    "Обновление",
                    "Обновление скачано. Сейчас программа закроется, updater применит обновление "
                    "и запустит программу снова.",
                )
                start_updater(packages=packages, restart=True)
                QApplication.instance().quit()
            except UpdateError as exc:
                QMessageBox.warning(self, "Обновления", str(exc))
            except Exception as exc:
                QMessageBox.critical(self, "Обновления", f"Ошибка обновления: {exc}")

        # ------------------------- общие данные -------------------------
        def _manager_profile(self) -> ManagerProfile:
            return ManagerProfile(
                name=self.manager_name_edit.text().strip(),
                position=self.manager_position_edit.text().strip(),
                email=self.manager_email_edit.text().strip(),
                phone=self.manager_phone_edit.text().strip(),
            )

        def _set_manager_profile(self, profile: ManagerProfile) -> None:
            self.manager_name_edit.setText(profile.name)
            self.manager_position_edit.setText(profile.position)
            self.manager_email_edit.setText(profile.email)
            self.manager_phone_edit.setText(profile.phone)

        def _save_manager_profile(self) -> None:
            profile = self._manager_profile()
            self.settings.setValue("manager_name", profile.name)
            self.settings.setValue("manager_position", profile.position)
            self.settings.setValue("manager_email", profile.email)
            self.settings.setValue("manager_phone", profile.phone)
            self.settings.setValue("manager_profile_locked", "1")
            self.settings.sync()
            self._notify_pages_settings_changed()

        def _has_saved_manager_profile(self) -> bool:
            locked = self._saved("manager_profile_locked", "") == "1"
            saved = any(
                self._saved(key, "").strip()
                for key in ("manager_name", "manager_position", "manager_email", "manager_phone")
            )
            return locked or saved

        def _selected_signer_key(self) -> str:
            return "alisher" if self.signer_alisher_radio.isChecked() else "saniya"

        def _selected_signer(self) -> dict[str, str]:
            return SIGNERS[self._selected_signer_key()]

        def _on_signer_changed(self) -> None:
            self.settings.setValue("signer_key", self._selected_signer_key())
            self.settings.sync()
            self._notify_pages_settings_changed()

        # ------------------------- точки подключения страниц -------------------------
        def _all_brand_pages(self) -> list[QWidget]:
            return [
                page
                for page in (
                    getattr(self, "stulz_page", None),
                    getattr(self, "riello_page", None),
                    getattr(self, "dc_eltek_page", None),
                    getattr(self, "battery_page", None),
                    getattr(self, "hvac_page", None),
                    getattr(self, "genset_page", None),
                )
                if page is not None
            ]

        def _active_brand_page(self):
            if not hasattr(self, "brand_tabs"):
                return None
            return self.brand_tabs.currentWidget()

        def _notify_pages_settings_changed(self) -> None:
            for page in self._all_brand_pages():
                if hasattr(page, "on_settings_changed"):
                    try:
                        page.on_settings_changed()
                    except Exception:
                        pass

        def current_project_dir(self) -> str:
            page = self._active_brand_page()
            if page is not None and hasattr(page, "project_path_text"):
                try:
                    return str(page.project_path_text())
                except Exception:
                    pass
            return self._saved("project_dir", "")

        def _brand_for_tab_index(self, index: int) -> str:
            if hasattr(self, "brand_tabs") and 0 <= index < self.brand_tabs.count():
                return self.brand_tabs.tabText(index)
            return "Stulz"

        def _tab_index_for_brand(self, brand: str) -> int:
            if hasattr(self, "brand_tabs"):
                for i in range(self.brand_tabs.count()):
                    if self.brand_tabs.tabText(i) == brand:
                        return i
            names = list(BRANDS.keys())
            return names.index(brand) if brand in names else 0

        def _select_tab_for_brand(self, brand: str) -> None:
            if not hasattr(self, "brand_tabs"):
                return
            self.brand_tabs.setCurrentIndex(self._tab_index_for_brand(brand))

        def _on_brand_tab_changed(self, index: int) -> None:
            self.settings.setValue("brand", self._brand_for_tab_index(index))
            self.settings.sync()

        # ------------------------- UI helpers for pages -------------------------
        def _display_file(self, path_text: str) -> str:
            return Path(path_text).name if path_text else ""

        def _display_dir(self, path_text: str) -> str:
            if not path_text:
                return ""
            p = Path(path_text)
            return p.name or path_text

        def _path_from_combo(self, combo: QComboBox) -> str:
            data = combo.currentData()
            return str(data) if data else combo.currentText().strip()

        def _add_path_item(self, combo: QComboBox, path_text: str, is_file: bool = True) -> None:
            full = str(path_text)
            display = self._display_file(full) if is_file else self._display_dir(full)
            combo.addItem(display, full)
            index = combo.count() - 1
            combo.setItemData(index, full, Qt.ToolTipRole)

        def _find_combo_path(self, combo: QComboBox, path_text: str) -> int:
            full = str(path_text)
            for i in range(combo.count()):
                if str(combo.itemData(i) or combo.itemText(i)) == full:
                    return i
            return -1

        def _set_line_path(self, line_edit: QLineEdit, path_text: str, is_file: bool = False) -> None:
            full = str(path_text)
            display = self._display_file(full) if is_file else self._display_dir(full)
            self._updating_path_display = True
            line_edit.setText(display)
            self._updating_path_display = False
            line_edit.setToolTip(full)

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

        # ------------------------- layout/styling -------------------------
        def _build_ui(self) -> None:
            central = QWidget()
            root = QHBoxLayout(central)
            root.setContentsMargins(0, 0, 0, 0)
            root.setSpacing(0)

            self.sidebar = QFrame()
            self.sidebar.setObjectName("Sidebar")
            self.sidebar.setMinimumWidth(220)
            side = QVBoxLayout(self.sidebar)
            side.setContentsMargins(18, 16, 18, 14)
            side.setSpacing(7)

            brand = QLabel("SAM\nGROUP")
            brand.setObjectName("Brand")
            title = QLabel("Offer Generator")
            title.setObjectName("SideTitle")
            subtitle = QLabel("Страницы брендов сами управляют своими расчетами и КП")
            subtitle.setObjectName("SideSubtitle")
            subtitle.setWordWrap(True)

            self.settings_btn = QPushButton("Настройки")
            self.settings_btn.setObjectName("Badge")
            self.settings_btn.clicked.connect(self._open_settings_dialog)
            self._responsive_widgets.append(self.settings_btn)

            self.update_btn = QPushButton("Обновления")
            self.update_btn.setObjectName("Badge")
            self.update_btn.clicked.connect(self._check_updates)
            self._responsive_widgets.append(self.update_btn)

            side.addWidget(brand)
            side.addWidget(title)
            side.addWidget(subtitle)
            side.addSpacing(6)
            side.addWidget(self.settings_btn)
            side.addWidget(self.update_btn)
            side.addSpacing(6)

            signer_title = QLabel("Подписант")
            signer_title.setObjectName("SidebarSectionTitle")
            side.addWidget(signer_title)
            side.addWidget(self.signer_saniya_radio)
            side.addWidget(self.signer_alisher_radio)
            side.addSpacerItem(QSpacerItem(20, 12, QSizePolicy.Minimum, QSizePolicy.Expanding))

            footer = QLabel(APP_FOOTER)
            footer.setObjectName("SidebarFooter")
            footer.setWordWrap(True)
            side.addWidget(footer)

            self.content = QFrame()
            self.content.setObjectName("Content")
            self.content.setMinimumWidth(560)
            content_layout = QVBoxLayout(self.content)
            content_layout.setContentsMargins(34, 28, 34, 28)
            content_layout.setSpacing(18)

            self.brand_tabs = QTabWidget()
            self.brand_tabs.setObjectName("BrandTabs")
            self.brand_tabs.setDocumentMode(True)

            self.stulz_page = StulzPage(self)
            self.riello_page = RielloPage(self)
            self.dc_eltek_page = DcEltekPage(self)
            self.battery_page = BatteryPage(self)
            self.hvac_page = HVACPage(self)
            self.genset_page = GensetPage(self)

            self.brand_tabs.addTab(self.stulz_page, "Stulz")
            self.brand_tabs.addTab(self.riello_page, "Riello")
            self.brand_tabs.addTab(self.dc_eltek_page, "DC Eltek")
            self.brand_tabs.addTab(self.battery_page, "Battery")
            self.brand_tabs.addTab(self.hvac_page, "HVAC")
            self.brand_tabs.addTab(self.genset_page, "Genset")
            self.brand_tabs.currentChanged.connect(self._on_brand_tab_changed)
            self.brand_tabs.setCurrentIndex(self._tab_index_for_brand(self._saved("brand", "Stulz")))
            content_layout.addWidget(self.brand_tabs)

            scroll = QScrollArea()
            scroll.setObjectName("ContentScroll")
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setWidget(self.content)

            self.sidebar_scroll = QScrollArea()
            self.sidebar_scroll.setObjectName("SidebarScroll")
            self.sidebar_scroll.setWidgetResizable(True)
            self.sidebar_scroll.setFrameShape(QFrame.NoFrame)
            self.sidebar_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.sidebar_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.sidebar_scroll.setWidget(self.sidebar)

            root.addWidget(self.sidebar_scroll)
            root.addWidget(scroll, stretch=1)
            self.setCentralWidget(central)
            self._apply_responsive_metrics(force=True)

        def _ui_scale(self) -> float:
            return ui_scale(self.width(), self.height())

        def _apply_responsive_metrics(self, force: bool = False) -> None:
            scale = self._ui_scale()
            if not force and abs(scale - self._last_scale) < 0.03:
                return
            self._last_scale = scale

            sidebar_width = int(max(220, min(240, self.width() * 0.20)))
            self.sidebar.setFixedWidth(sidebar_width)
            if hasattr(self, "sidebar_scroll"):
                self.sidebar_scroll.setFixedWidth(sidebar_width)

            content_margin_x = int(22 * scale)
            content_margin_y = int(20 * scale)
            self.content.layout().setContentsMargins(
                content_margin_x,
                content_margin_y,
                content_margin_x,
                content_margin_y,
            )
            self.content.layout().setSpacing(int(16 * scale))

            widget_h = int(28 * scale)
            for widget in self._responsive_widgets:
                widget.setMinimumHeight(widget_h)

            self._apply_style(scale)

            for page in self._all_brand_pages():
                if hasattr(page, "apply_responsive_metrics"):
                    try:
                        page.apply_responsive_metrics(scale)
                    except Exception:
                        pass

        def resizeEvent(self, event) -> None:  # noqa: N802 - Qt API name
            super().resizeEvent(event)
            self._apply_responsive_metrics()

        def _apply_style(self, scale: float | None = None) -> None:
            scale = scale if scale is not None else self._ui_scale()
            app = QApplication.instance()
            if app:
                app.setFont(QFont("Segoe UI", max(9, int(10 * scale))))
            self.setStyleSheet(stylesheet(scale))

        def closeEvent(self, event) -> None:  # noqa: N802 - Qt API name
            for page in self._all_brand_pages():
                if hasattr(page, "remember_values"):
                    try:
                        page.remember_values()
                    except Exception:
                        pass
            self.settings.setValue("signer_key", self._selected_signer_key())
            self.settings.sync()
            super().closeEvent(event)

    app = QApplication.instance() or QApplication(sys.argv)
    icon_path = app_icon_path()
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = OfferGeneratorWindow()
    window.show()
    app.exec()
