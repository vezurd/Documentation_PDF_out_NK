"""Live split-tunnel schematic for the VPN tab (lamps + short comments)."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from ds_compare_center.vpn_status import FlowLamp, VpnSnapshot, exit_lamps, flow_lamps

_OK_BG = "#e6f4ea"
_BAD_BG = "#fce8e6"
_UNK_BG = "#f3f4f6"
_OK_FG = "#1a7f37"
_BAD_FG = "#b42318"
_UNK_FG = "#6b7280"
_LAMP = {
    True: "#1a7f37",
    False: "#b42318",
    None: "#9aa0a6",
}

# key, title, зачем этот шаг
_NODES: tuple[tuple[str, str, str], ...] = (
    ("cursor", "Cursor.exe", "чат и агент Cursor; в туннель только этот процесс"),
    ("proxifyre", "ProxiFyre", "перехват сокетов Cursor; LAN не трогает (bypassLan)"),
    ("socks", "3proxy :1080", "локальный SOCKS; исходящий bind на 10.95.175.58"),
    ("cursorbind", "CursorBind", "труба Amnezia в NL; Table=off + слабый default metric 9999"),
    ("socks_exit", "выход NL S4", "интернет Cursor; ожидаем ~31.58… (кнопка «Проверить IP»)"),
    ("other_apps", "браузер / python / git", "не Cursor.exe — ProxiFyre их не цепляет"),
    ("eth_default", "Ethernet default", "0.0.0.0/0 через 172.16.3.1 metric 5 — побеждает у всех, кроме bind"),
    ("office_exit", "офисный IP", "обычный интернет ПК; ~84.204…, не H10 195.218…"),
    ("unc", r"UNC \\bcc\eng", "SMB Windows, не сокет Cursor"),
    ("lan", "LAN 172.16.0.0/16", "узкий маршрут metric 1 — шары всегда по Ethernet"),
    ("lan_gw", "шлюз 172.16.3.1", "корпоративный корень, не VPN"),
)

_LANES: tuple[tuple[str, tuple[str, ...]], ...] = (
    (
        "Cursor → NL — обход закрытых портов офиса",
        ("cursor", "proxifyre", "socks", "cursorbind", "socks_exit"),
    ),
    (
        "Остальной интернет — весь ПК в VPN не тащим",
        ("other_apps", "eth_default", "office_exit"),
    ),
    (
        "Корп. LAN / UNC — шары не через туннель",
        ("unc", "lan", "lan_gw"),
    ),
)


class _Lamp(QLabel):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setFixedSize(12, 12)
        self.set_ok(None)

    def set_ok(self, ok: bool | None) -> None:
        color = _LAMP[ok]
        self.setStyleSheet(
            f"background: {color}; border-radius: 6px; border: 1px solid #555;"
        )


class VpnFlowNode(QFrame):
    """One box on the schematic: lamp, title, live status, why-comment."""

    def __init__(self, title: str, comment: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setMinimumWidth(118)

        self._lamp = _Lamp()
        head = QLabel(title)
        head.setStyleSheet("font-weight: 600; font-size: 11px;")
        head.setWordWrap(True)

        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.setSpacing(6)
        top.addWidget(self._lamp, 0, Qt.AlignmentFlag.AlignTop)
        top.addWidget(head, 1)

        self._status = QLabel("—")
        self._status.setWordWrap(True)
        self._status.setStyleSheet("font-size: 10px;")

        why = QLabel(comment)
        why.setWordWrap(True)
        why.setStyleSheet("font-size: 10px; color: #444;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(3)
        layout.addLayout(top)
        layout.addWidget(self._status)
        layout.addWidget(why)
        layout.addStretch(1)

    def apply_lamp(self, lamp: FlowLamp) -> None:
        """Paint lamp/background from a computed FlowLamp."""
        self._lamp.set_ok(lamp.ok)
        self._status.setText(lamp.status or "—")
        if lamp.ok is True:
            bg, fg = _OK_BG, _OK_FG
        elif lamp.ok is False:
            bg, fg = _BAD_BG, _BAD_FG
        else:
            bg, fg = _UNK_BG, _UNK_FG
        self._status.setStyleSheet(f"font-size: 10px; color: {fg};")
        self.setStyleSheet(
            f"background: {bg}; border: 1px solid #c5c5c5; border-radius: 6px;"
        )


class VpnFlowDiagram(QWidget):
    """Three-lane route schematic with status lamps."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._nodes: dict[str, VpnFlowNode] = {}
        for key, title, comment in _NODES:
            self._nodes[key] = VpnFlowNode(title, comment)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        legend = QLabel(
            "Схема маршрутов. Лампочка: зелёная — шаг жив, красная — нет, "
            "серая — не опрашивается или нужен curl."
        )
        legend.setWordWrap(True)
        legend.setStyleSheet("font-size: 11px; color: #333;")
        root.addWidget(legend)

        self._h10_lamp = _Lamp()
        self._h10_text = QLabel("H10 полный туннель — статус после проверки.")
        self._h10_text.setWordWrap(True)
        h10_row = QHBoxLayout()
        h10_row.setContentsMargins(8, 4, 8, 4)
        h10_row.setSpacing(8)
        h10_row.addWidget(self._h10_lamp, 0, Qt.AlignmentFlag.AlignTop)
        h10_row.addWidget(self._h10_text, 1)
        self._h10_bar = QFrame()
        self._h10_bar.setLayout(h10_row)
        self._h10_bar.setFrameShape(QFrame.Shape.StyledPanel)
        root.addWidget(self._h10_bar)

        for title, keys in _LANES:
            cap = QLabel(title)
            cap.setStyleSheet("font-weight: 600; font-size: 12px; color: #1f3a5f;")
            row = QHBoxLayout()
            row.setSpacing(2)
            for index, key in enumerate(keys):
                if index:
                    arrow = QLabel("→")
                    arrow.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    arrow.setStyleSheet("font-size: 18px; color: #555; padding: 0 2px;")
                    row.addWidget(arrow, 0)
                row.addWidget(self._nodes[key], 1)
            lane = QWidget()
            lane_l = QVBoxLayout(lane)
            lane_l.setContentsMargins(0, 0, 0, 0)
            lane_l.setSpacing(2)
            lane_l.addWidget(cap)
            lane_l.addLayout(row)
            root.addWidget(lane)

    def apply(
        self,
        snap: VpnSnapshot | None,
        *,
        direct: str = "",
        socks: str = "",
    ) -> None:
        """Repaint lamps from a snapshot and optional curl IPs."""
        if snap is None:
            empty = FlowLamp(None, "нет данных")
            for node in self._nodes.values():
                node.apply_lamp(empty)
            self._set_h10(FlowLamp(None, "нет данных"))
            office, nl = exit_lamps(direct, socks)
            self._nodes["office_exit"].apply_lamp(office)
            self._nodes["socks_exit"].apply_lamp(nl)
            return
        lamps = flow_lamps(snap, direct=direct, socks=socks)
        for key, node in self._nodes.items():
            node.apply_lamp(lamps.get(key, FlowLamp(None, "нет данных")))
        self._set_h10(lamps.get("h10_off", FlowLamp(None, "нет данных")))

    def _set_h10(self, lamp: FlowLamp) -> None:
        self._h10_lamp.set_ok(lamp.ok)
        why = (
            "H10 — полный туннель (весь интернет в NL, metric 0). "
            "Запасной режим; вместе с CursorBind не включать."
        )
        self._h10_text.setText(f"{lamp.status}\n{why}")
        if lamp.ok is True:
            bg, fg = _OK_BG, _OK_FG
        elif lamp.ok is False:
            bg, fg = _BAD_BG, _BAD_FG
        else:
            bg, fg = _UNK_BG, _UNK_FG
        self._h10_text.setStyleSheet(f"font-size: 11px; color: {fg};")
        self._h10_bar.setStyleSheet(
            f"QFrame {{ background: {bg}; border: 1px solid #c5c5c5; border-radius: 6px; }}"
        )
