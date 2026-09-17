"""Legacy wrapper that opens the PDF v2 control center on the Settings tab."""

from __future__ import annotations

import os
import subprocess
import sys
from typing import Any, Optional


def show_pdf_v2_settings(parent: Optional[Any] = None) -> None:
    del parent
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    subprocess.Popen(
        [sys.executable, "-m", "pdf_v2_monitor", "--tab", "settings"],
        cwd=root,
    )
