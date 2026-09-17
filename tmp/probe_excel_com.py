from __future__ import annotations

import sys

try:
    import win32com.client
except ImportError:
    print("NO_WIN32COM")
    sys.exit(1)


def main() -> None:
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception as exc:
        print("NO_EXCEL", type(exc).__name__, exc)
        sys.exit(2)
    print("excel visible", bool(excel.Visible))
    print("workbooks", excel.Workbooks.Count)
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(i)
        print("--- workbook", i, "---")
        print("name", wb.Name)
        print("full", wb.FullName)
        print("path", wb.Path)
        print("sheets", wb.Worksheets.Count)
        for j in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(j)
            print("  sheet", j, ascii(ws.Name), "used", ws.UsedRange.Rows.Count, "x", ws.UsedRange.Columns.Count)


if __name__ == "__main__":
    main()
