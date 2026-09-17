"""Compatibility shim: old path ``tmp/analyze_rfp_parts.py``.

Prefer::

    python -X utf8 -m RFQ.rfp_parts --manifest --strict
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import main

if __name__ == "__main__":
    print(
        "NOTE: script moved to RFQ/rfp_parts/; prefer: python -X utf8 -m RFQ.rfp_parts ...",
        file=sys.stderr,
    )
    main()
