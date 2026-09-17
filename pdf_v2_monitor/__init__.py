"""PDF v2 monitor — PySide6 GUI for real-time pipeline progress observation.

Launch as::

    python -m pdf_v2_monitor "C:\\path\\to\\PDF" [--project AGCC_287]

Or from main.py via the monitor button.
"""

from pdf_v2_monitor.monitor_window import MonitorWindow

__all__ = ["MonitorWindow"]
