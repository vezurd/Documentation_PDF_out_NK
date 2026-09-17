"""
Read-only path smoke tests: local repo paths vs UNC. Does not modify any file.
Run from repo root: python tmp/test_paths_readonly.py

On Windows, prefer UTF-8 console for paths with Cyrillic / special punctuation:
  chcp 65001
  set PYTHONUTF8=1
  python tmp/test_paths_readonly.py
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def _safe_print(s: str) -> None:
    """Avoid UnicodeEncodeError on cp1251 consoles when path contains U+2010 etc."""
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.buffer.write((s + "\n").encode(enc, errors="backslashreplace"))


def try_read(label: str, path: str | Path) -> None:
    p = Path(path) if not isinstance(path, Path) else path
    _safe_print(f"\n--- {label} ---")
    _safe_print(f"repr: {ascii(str(p))}")
    _safe_print(f"exists: {p.exists()}")
    _safe_print(f"is_file: {p.is_file()}")
    try:
        raw = p.read_bytes()
        _safe_print(f"bytes: {len(raw)}")
        text = raw.decode("utf-8")
        obj = json.loads(text)
        _safe_print(f"json top-level keys (sample): {list(obj)[:8]}")
    except OSError as e:
        _safe_print(f"OSError: {e!r}")
    except UnicodeDecodeError as e:
        _safe_print(f"UnicodeDecodeError: {e!r}")
    except json.JSONDecodeError as e:
        _safe_print(f"JSONDecodeError: {e!r}")


def main() -> int:
    here = Path(__file__).resolve()
    root = here.parents[1]
    os.chdir(root)
    _safe_print("cwd: " + os.getcwd())
    _safe_print("sys.argv[0]: " + sys.argv[0])

    # Local: relative from repo root (forward slashes OK on Windows)
    try_read("local relative forward-slash", "pdf_parsing_v2_engine/templates/agcc_287/od_page1.json")
    try_read("local pathlib from root", root / "pdf_parsing_v2_engine" / "templates" / "agcc_287" / "od_page1.json")

    # UNC as user might paste (single backslash after \\ is wrong in Python string — we document both)
    unc_user_style = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\4130\06_4130-KSB1\Для передачи\05_рев.0_AN-02_AGCC.287‐4130‐KSB1_as-build\PDF\adapt_debug\adapt_debug_2026_04_20__214628__AGCC.287-4130-KSB1.CJ-0010_0-AN01_RU__cj_page1.json"
    try_read("UNC raw string r'\\\\bcc\\...' (correct Python UNC)", unc_user_style)

    # Non-raw string without doubling backslashes: invalid escape \e (SyntaxWarning / wrong path).
    _safe_print("\n--- note ---")
    _safe_print("Use r'\\\\server\\share\\...' or '\\\\\\\\server\\\\share\\\\...' for UNC in Python source.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
