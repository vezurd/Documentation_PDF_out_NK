"""Tests for UNC/local path open helpers."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.path_actions import open_containing_folder, open_directory, open_path

_POS2 = (
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\8950\22_POS2"
    r"\Для передачи\13_рев.03-AN04_от_2026.07.16"
)


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
    assert args[0] == "explorer.exe"
    assert args[1] == r"\\bcc\eng\PrDoc\folder"
    assert "/root," not in subprocess.list2cmdline(args)


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


def test_open_directory_uses_startfile() -> None:
    with (
        patch("rd_catalog.path_actions.os.startfile", return_value=None) as startfile,
        patch("rd_catalog.path_actions.subprocess.Popen") as popen,
    ):
        ok, message = open_directory(_POS2)
    assert ok is True
    assert message == _POS2
    startfile.assert_called_once_with(_POS2)
    popen.assert_not_called()


def test_open_directory_fallback_keeps_spaces_out_of_switches() -> None:
    with (
        patch("rd_catalog.path_actions.os.startfile", side_effect=OSError("nope")),
        patch("rd_catalog.path_actions.subprocess.Popen") as popen,
        patch("rd_catalog.path_actions.os.name", "nt"),
    ):
        ok, message = open_directory(_POS2)
    assert ok is True
    assert message == _POS2
    args = popen.call_args[0][0]
    assert args == ["explorer.exe", _POS2]
    cmdline = subprocess.list2cmdline(args)
    assert "/root," not in cmdline
    assert '"\\\\bcc\\eng\\' in cmdline or cmdline.startswith("explorer.exe ")


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
    test_open_directory_uses_startfile()
    test_open_directory_fallback_keeps_spaces_out_of_switches()
    test_open_containing_folder_is_lexical()
    print("ok")
