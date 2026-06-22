# Добавить в gui/main_window.py

from gui.pages.hvac_page import HVACPage

# В месте, где добавляются вкладки Stulz/Riello:
# self.tabs.addTab(StulzPage(self), "Stulz")
# self.tabs.addTab(RielloPage(self), "Riello")
self.tabs.addTab(HVACPage(self), "HVAC")
