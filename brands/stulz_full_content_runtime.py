from __future__ import annotations

import brands.stulz_position_selection_runtime as _previous


# Re-export the complete mature STULZ runtime: physical specification mapping,
# legends, cable quantities, compressor sections and selectable Calc positions.
#
# IMPORTANT: build a stable snapshot before writing anything into globals().
# The previous runtime itself contains a private attribute named ``_previous``.
# The old loop overwrote this module's own ``_previous`` variable midway through
# the export and the very next ``getattr(_previous, ...)`` was executed against
# brands.stulz_compressor_runtime instead of stulz_position_selection_runtime.
# That produced:
#   module 'brands.stulz_compressor_runtime' has no attribute '_selection_ui'
# during commercial-offer generation/import.
_runtime_exports = tuple(
    (name, getattr(_previous, name))
    for name in dir(_previous)
    if not name.startswith("__")
)
for _name, _value in _runtime_exports:
    globals()[_name] = _value


# Install the final GUI-only layout layer last. It removes nested scrolling from
# main STULZ blocks and expands them to their full text/row content.
import gui.pages.stulz_full_content_runtime as _full_content_ui  # noqa: E402,F401
