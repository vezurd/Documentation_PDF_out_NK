"""Workstation VPN tab: CursorBind + 3proxy + ProxiFyre status and launchers."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QShowEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ds_compare_center.vpn_status import (
    START_BAT,
    VPN_HOME,
    VpnSnapshot,
    collect_vpn_status,
    ips_look_correct,
    probe_public_ips,
    start_bat_elevated,
)
from utils.path import open_dir

_HEADERS = ("Проверка", "Статус", "Деталь")
_OK_FG = QColor("#1a7f37")
_OK_BG = QColor("#e6f4ea")
_BAD_FG = QColor("#b42318")
_BAD_BG = QColor("#fce8e6")
_BANNER_OK = "color: #287a3d; font-size: 13px; font-weight: 600;"
_BANNER_BAD = "color: #b42318; font-size: 13px; font-weight: 600;"
_ROW_HEIGHT = 24


class VpnPanel(QWidget):
    """Show whether the Cursor-only Amnezia split stack is up."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._seen = False

        intro = QLabel(
            "Стек CursorBind: AmneziaWG (Table=off) → 3proxy :1080 → ProxiFyre "
            "только для Cursor.exe. Браузер и UNC \\\\bcc\\eng остаются на Ethernet. "
            "H10 (полный туннель) должен быть выключен."
        )
        intro.setWordWrap(True)

        self._banner = QLabel("Нажмите «Проверить статус».")
        self._banner.setWordWrap(True)

        self._table = QTableWidget(0, 3)
        self._table.setHorizontalHeaderLabels(list(_HEADERS))
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self._table.verticalHeader().setVisible(False)
        self._table.setAlternatingRowColors(True)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)

        btn_status = QPushButton("Проверить статус")
        btn_status.clicked.connect(self.refresh_status)
        btn_ips = QPushButton("Проверить IP (curl)")
        btn_ips.clicked.connect(self._on_probe_ips)
        btn_admin = QPushButton("Запустить стек от имени администратора")
        btn_admin.clicked.connect(self._on_start_admin)
        btn_folder = QPushButton("Открыть папку Amnezia")
        btn_folder.clicked.connect(self._on_open_folder)

        buttons = QHBoxLayout()
        buttons.addWidget(btn_status)
        buttons.addWidget(btn_ips)
        buttons.addStretch(1)

        launch = QHBoxLayout()
        launch.addWidget(btn_admin)
        launch.addWidget(btn_folder)
        launch.addStretch(1)

        paths = QLabel(f"bat: {START_BAT}\nпапка: {VPN_HOME}")
        paths.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        paths.setWordWrap(True)

        box = QGroupBox("CursorBind / 3proxy / ProxiFyre")
        box_l = QVBoxLayout(box)
        box_l.addWidget(intro)
        box_l.addLayout(buttons)
        box_l.addLayout(launch)
        box_l.addWidget(self._banner)
        box_l.addWidget(self._table, 1)
        box_l.addWidget(paths)

        root = QVBoxLayout(self)
        root.addWidget(box)

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        if not self._seen:
            self._seen = True
            self.refresh_status()

    def refresh_status(self) -> None:
        """Re-read adapters/processes and fill the table."""
        try:
            snap = collect_vpn_status()
        except Exception as exc:
            self._banner.setText(f"Ошибка опроса: {exc}")
            self._banner.setStyleSheet(_BANNER_BAD)
            self._table.setRowCount(0)
            return
        self._apply_snapshot(snap)

    def _apply_snapshot(self, snap: VpnSnapshot) -> None:
        if snap.overall_ok:
            self._banner.setText("Стек CursorBind работает как задумано.")
            self._banner.setStyleSheet(_BANNER_OK)
        else:
            self._banner.setText("Есть замечания — см. таблицу. H10 должен быть выключен.")
            self._banner.setStyleSheet(_BANNER_BAD)
        self._table.setRowCount(len(snap.checks))
        for row, check in enumerate(snap.checks):
            status = "ОК" if check.ok else "нет"
            values = (check.name, status, check.detail)
            fg = _OK_FG if check.ok else _BAD_FG
            bg = _OK_BG if check.ok else _BAD_BG
            for col, text in enumerate(values):
                item = QTableWidgetItem(text)
                item.setForeground(QBrush(fg))
                item.setBackground(QBrush(bg))
                self._table.setItem(row, col, item)
            self._table.setRowHeight(row, _ROW_HEIGHT)

    def _on_start_admin(self) -> None:
        ok, detail = start_bat_elevated()
        if ok:
            QMessageBox.information(
                self,
                "VPN",
                "Запрос UAC отправлен. В окне bat: CursorBind уже включён, H10 выключен, "
                "затем Enter.\n\n" + detail,
            )
            return
        QMessageBox.warning(
            self,
            "VPN",
            f"{detail}\n\nОткройте папку и запустите start-cursor-vpn.bat вручную "
            "(ПКМ → от имени администратора).",
        )

    def _on_open_folder(self) -> None:
        if not VPN_HOME.is_dir():
            QMessageBox.warning(self, "VPN", f"Папка не найдена:\n{VPN_HOME}")
            return
        open_dir(str(VPN_HOME))

    def _on_probe_ips(self) -> None:
        self._banner.setText("curl api.ipify.org (direct и SOCKS)…")
        self._banner.setStyleSheet("")
        try:
            direct, socks = probe_public_ips()
        except Exception as exc:
            QMessageBox.warning(self, "VPN", str(exc))
            return
        ok, detail = ips_look_correct(direct, socks)
        QMessageBox.information(self, "IP", detail)
        if ok:
            self._banner.setText("IP split ок: " + detail)
            self._banner.setStyleSheet(_BANNER_OK)
        else:
            self._banner.setText("IP split не ок: " + detail)
            self._banner.setStyleSheet(_BANNER_BAD)
