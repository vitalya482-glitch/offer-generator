from __future__ import annotations


def run_gui() -> None:
    """Run the existing main window with the corrected STULZ page class.

    The main window imports StulzPage inside run_gui().  Replacing the class on
    the already-loaded module immediately before that import keeps the main-window
    implementation untouched and isolates STULZ-specific behavior in its page.
    """

    from gui.pages import stulz_page as stulz_page_module
    from gui.pages.stulz_page_runtime import StulzPage

    stulz_page_module.StulzPage = StulzPage

    from gui.main_window import run_gui as _run_gui

    _run_gui()
