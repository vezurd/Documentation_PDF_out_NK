from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(r"\\bcc\root\CurProjects\di_manegers")
NESTED = ROOT / "Пятницкий_П"
TARGET = Path(
    r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
    r"\Амурский ГХК\Поставки\Отчет систем ПИ"
    r"\Отчет для заказчика АГХК 07.09.2026.xlsb"
)


def dump(label: str, p: Path) -> None:
    print(f"=== {label} ===")
    print("repr", ascii(str(p)))
    print("os.exists", os.path.exists(p))
    print("path.exists", p.exists())
    try:
        names = os.listdir(p)
        print("listdir n=", len(names))
        for name in names[:40]:
            print(" ", ascii(name))
    except Exception as exc:
        print("listdir", type(exc).__name__, exc)
    print()


dump("root", ROOT)
dump("pyatnitskiy", NESTED)

print("=== target ===")
print("repr", ascii(str(TARGET)))
print("os.exists", os.path.exists(TARGET))
print("path.exists", TARGET.exists())
if TARGET.exists():
    st = TARGET.stat()
    print("size_mb", round(st.st_size / 1024 / 1024, 2))
    print("mtime", st.st_mtime)
