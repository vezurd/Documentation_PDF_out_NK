"""python -m pdf_template_editor — запуск GUI редактора шаблонов."""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from . import launch_template_editor


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF template editor")
    parser.add_argument("pdf_path", nargs="?", default=None, help="Optional PDF path to open")
    args = parser.parse_args()
    launch_template_editor(args.pdf_path)
