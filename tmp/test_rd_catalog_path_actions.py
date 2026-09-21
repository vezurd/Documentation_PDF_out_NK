"""Tests for UNC/local path open helpers."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.path_actions import open_containing_folder, open_directory, open_path


def test_open_path_falls_back_to_explorer_on_startfile_error() -> None:
    with (
        patch("rd_catalog.path_actions.os.startfile", side_effect=OSError("nope")),
        patch("rd_catalog.path_actions.subprocess.Popen") as popen,
        patch("rd_catalog.path_actions.os.name", "nt"),
    ):
        ok, message = open_path(r"\\bcc\eng\PrDoc\folder")
    assert ok is True
    assert message.endswith("folder")
    popen.assert_called_once()
    args = popen.call_args[0][0]
    assert args[0] == "explorer"
    assert args[1].startswith("/root,")


def test_open_path_reports_failure_when_explorer_also_fails() -> None:
    with (
        patch("rd_catalog.path_actions.os.startfile", side_effect=OSError("nope")),
        patch(
            "rd_catalog.path_actions.subprocess.Popen",
            side_effect=OSError("explorer"),
        ),
        patch("rd_catalog.path_actions.os.name", "nt"),
    ):
        ok, message = open_path(r"C:\missing")
    assert ok is False
    assert "Не удалось открыть" in message


def test_open_directory_uses_explorer_first() -> None:
    with (
        patch("rd_catalog.path_actions.subprocess.Popen") as popen,
        patch("rd_catalog.path_actions.os.name", "nt"),
        patch("rd_catalog.path_actions.os.startfile") as startfile,
    ):
        ok, message = open_directory(r"\\bcc\eng\PrDoc\377_x\05_рев.03_AGCC.287‐6160‐SOS")
    assert ok is True
    startfile.assert_not_called()
    popen.assert_called_once()
    args = popen.call_args[0][0]
    assert args[0] == "explorer"
    assert args[1].startswith("/root,")
    assert "\u2010" in args[1]


def test_open_containing_folder_is_lexical() -> None:
    folder = r"\\bcc\eng\PrDoc\377_x\PDF"
    file_path = folder + r"\AGCC.287-1000-KSB.OD-0001_0_RU.pdf"
    with patch(
        "rd_catalog.path_actions.open_directory", return_value=(True, folder)
    ) as opener:
        ok, message = open_containing_folder(file_path)
    assert ok is True
    opener.assert_called_once()
    sent = opener.call_args[0][0]
    assert sent.endswith("PDF")
    assert "OD-0001" not in sent
    assert message == folder


if __name__ == "__main__":
    test_open_path_falls_back_to_explorer_on_startfile_error()
    test_open_path_reports_failure_when_explorer_also_fails()
    test_open_directory_uses_explorer_first()
    test_open_containing_folder_is_lexical()
    print("ok")
