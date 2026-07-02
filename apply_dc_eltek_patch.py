from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent
MAIN_WINDOW = ROOT / "gui" / "main_window.py"

if not MAIN_WINDOW.exists():
    raise SystemExit("Не найден gui/main_window.py. Запусти скрипт из корня проекта offer-generator.")

text = MAIN_WINDOW.read_text(encoding="utf-8")
original = text

# 1. Импорт страницы
if "from gui.pages.dc_eltek_page import DCEltekPage" not in text:
    anchor = "from gui.pages.battery_page import BatteryPage"
    if anchor in text:
        text = text.replace(anchor, anchor + "\nfrom gui.pages.dc_eltek_page import DCEltekPage", 1)
    else:
        raise SystemExit("Не нашёл импорт BatteryPage — main_window.py отличается от ожидаемого.")

# 2. Добавить страницу в _all_brand_pages
old = 'getattr(self, "battery_page", None),'
new = 'getattr(self, "battery_page", None),\n                getattr(self, "dc_eltek_page", None),'
if 'getattr(self, "dc_eltek_page", None)' not in text:
    if old in text:
        text = text.replace(old, new, 1)
    else:
        raise SystemExit("Не нашёл список страниц _all_brand_pages.")

# 3. Создать объект страницы
old = "self.battery_page = BatteryPage(self)"
new = "self.battery_page = BatteryPage(self)\n            self.dc_eltek_page = DCEltekPage(self)"
if "self.dc_eltek_page = DCEltekPage(self)" not in text:
    if old in text:
        text = text.replace(old, new, 1)
    else:
        raise SystemExit("Не нашёл создание BatteryPage.")

# 4. Добавить вкладку после Battery
old = 'self.brand_tabs.addTab(self.battery_page, "Battery")'
new = 'self.brand_tabs.addTab(self.battery_page, "Battery")\n            self.brand_tabs.addTab(self.dc_eltek_page, "DC Eltek")'
if 'self.brand_tabs.addTab(self.dc_eltek_page, "DC Eltek")' not in text:
    if old in text:
        text = text.replace(old, new, 1)
    else:
        raise SystemExit("Не нашёл вкладку Battery.")

if text == original:
    print("main_window.py уже был пропатчен, изменений нет.")
else:
    MAIN_WINDOW.write_text(text, encoding="utf-8")
    print("main_window.py обновлён: вкладка DC Eltek подключена.")
