"""Local CursorBind / 3proxy / ProxiFyre stack checks (workstation VPN)."""

from __future__ import annotations

import json
import os
import socket
import subprocess
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any

CREATE_NO_WINDOW = 0x08000000

VPN_HOME = Path(r"C:\YandexDisk\z_Docs\амнезия")
START_BAT = VPN_HOME / "start-cursor-vpn.bat"
PROXY_DIR = VPN_HOME / "3proxy" / "bin64"
PROXIFYRE_EXE = Path(r"C:\Program Files\ProxiFyre\ProxiFyre.exe")

CURSORBIND_ALIAS = "CursorBind"
CURSORBIND_IP = "10.95.175.58"
H10_ALIAS = "NetherlandsAmsterdamH10"
LAN_PREFIX = "172.16.0.0/16"
LAN_GATEWAY = "172.16.3.1"
SOCKS_HOST = "127.0.0.1"
SOCKS_PORT = 1080
UNC_SAMPLE = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД"
H10_EXIT_IP = "195.218.230.62"
SOCKS_EXIT_PREFIX = "31.58."


@dataclass(frozen=True)
class VpnCheck:
    """One row of the VPN status table."""

    name: str
    ok: bool
    detail: str
    required: bool = True
    key: str = ""


@dataclass(frozen=True)
class VpnSnapshot:
    """Evaluated CursorBind split-tunnel stack."""

    checks: tuple[VpnCheck, ...]
    overall_ok: bool
    raw: dict[str, Any]


@dataclass(frozen=True)
class FlowLamp:
    """One schematic node: lamp colour and a one-line live status."""

    ok: bool | None
    status: str


def _run_hidden(argv: list[str], *, timeout: int) -> subprocess.CompletedProcess[str]:
    flags = CREATE_NO_WINDOW if os.name == "nt" else 0
    return subprocess.run(
        argv,
        capture_output=True,
        text=True,
        timeout=timeout,
        creationflags=flags,
    )


def _ps_json(script: str) -> Any:
    proc = _run_hidden(
        ["powershell", "-NoProfile", "-Command", script],
        timeout=25,
    )
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(err or f"powershell exit {proc.returncode}")
    text = (proc.stdout or "").strip()
    if not text:
        return {}
    return json.loads(text)


def _task_running(image_name: str) -> bool:
    proc = _run_hidden(
        ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/NH"],
        timeout=10,
    )
    return image_name.lower() in (proc.stdout or "").lower()


def _socks_listening() -> bool:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(0.4)
    try:
        return sock.connect_ex((SOCKS_HOST, SOCKS_PORT)) == 0
    finally:
        sock.close()


def _unc_exists(path: str, timeout_s: float = 3.0) -> bool | None:
    box: list[bool | None] = [None]

    def _work() -> None:
        try:
            box[0] = Path(path).exists()
        except OSError:
            box[0] = False

    thread = threading.Thread(target=_work, daemon=True)
    thread.start()
    thread.join(timeout_s)
    if thread.is_alive():
        return None
    return box[0]


def _as_list(value: Any) -> list[dict[str, Any]]:
    if value is None:
        return []
    if isinstance(value, list):
        return [item if isinstance(item, dict) else {} for item in value]
    if isinstance(value, dict):
        return [value]
    return []


def evaluate_vpn_raw(raw: dict[str, Any]) -> VpnSnapshot:
    """Turn a collected adapter/process snapshot into pass/fail rows.

    Args:
        raw: Dict from ``collect_vpn_raw`` (or a test double).

    Returns:
        Evaluated snapshot. ``overall_ok`` is True when every required check passes.
    """
    addrs = _as_list(raw.get("addrs"))
    defaults = _as_list(raw.get("defaults"))
    lan = _as_list(raw.get("lan"))

    def _ips(alias: str) -> list[str]:
        out: list[str] = []
        for row in addrs:
            if str(row.get("InterfaceAlias") or "") == alias:
                ip = str(row.get("IPAddress") or "")
                if ip:
                    out.append(ip)
        return out

    cursor_ips = _ips(CURSORBIND_ALIAS)
    h10_ips = _ips(H10_ALIAS)
    socks_ok = bool(raw.get("socks"))
    p3 = bool(raw.get("proc3proxy"))
    pf = bool(raw.get("procProxiFyre"))
    unc = raw.get("unc")
    bat_ok = bool(raw.get("bat_exists"))

    eth_default = None
    h10_default = None
    for row in defaults:
        alias = str(row.get("InterfaceAlias") or "")
        hop = str(row.get("NextHop") or "")
        metric = row.get("RouteMetric")
        if alias == H10_ALIAS:
            h10_default = f"{hop} metric={metric}"
        if "Ethernet" in alias and hop == LAN_GATEWAY:
            eth_default = f"{alias} {hop} metric={metric}"

    lan_ok_row = None
    for row in lan:
        alias = str(row.get("InterfaceAlias") or "")
        hop = str(row.get("NextHop") or "")
        if hop == LAN_GATEWAY and "Ethernet" in alias:
            lan_ok_row = f"{alias} {hop} metric={row.get('RouteMetric')}"
            break

    cb_default_row = None
    cb_metric: int | None = None
    for row in defaults:
        if str(row.get("InterfaceAlias") or "") != CURSORBIND_ALIAS:
            continue
        hop = str(row.get("NextHop") or "")
        metric = row.get("RouteMetric")
        try:
            cb_metric = int(metric) if metric is not None else None
        except (TypeError, ValueError):
            cb_metric = None
        cb_default_row = f"{hop} metric={metric}"
        break
    cb_default_ok = cb_metric is not None and cb_metric >= 100

    unc_ok = unc is True
    if unc is None:
        unc_detail = f"таймаут {UNC_SAMPLE}"
    elif unc:
        unc_detail = UNC_SAMPLE
    else:
        unc_detail = f"нет доступа: {UNC_SAMPLE}"

    checks = (
        VpnCheck(
            "Файл start-cursor-vpn.bat",
            bat_ok,
            str(START_BAT) if bat_ok else f"нет файла: {START_BAT}",
            key="bat",
        ),
        VpnCheck(
            "Туннель CursorBind",
            CURSORBIND_IP in cursor_ips,
            ", ".join(cursor_ips) or "адаптер выключен",
            key="cursorbind",
        ),
        VpnCheck(
            "Слабый default CursorBind",
            cb_default_ok,
            cb_default_row or "нет 0.0.0.0/0 на CursorBind (metric 9999)",
            key="cb_default",
        ),
        VpnCheck(
            "H10 выключен",
            not h10_ips and h10_default is None,
            f"IP {', '.join(h10_ips) or '—'}; default {h10_default or 'нет'}",
            key="h10_off",
        ),
        VpnCheck(
            "Default через Ethernet",
            eth_default is not None and h10_default is None,
            eth_default or "нет маршрута 0.0.0.0/0 на Ethernet",
            key="eth_default",
        ),
        VpnCheck(
            f"LAN {LAN_PREFIX}",
            lan_ok_row is not None,
            lan_ok_row or "нет маршрута на Ethernet",
            key="lan",
        ),
        VpnCheck(
            f"SOCKS {SOCKS_HOST}:{SOCKS_PORT}",
            socks_ok,
            "слушает" if socks_ok else "не слушает (3proxy?)",
            key="socks",
        ),
        VpnCheck("Процесс 3proxy", p3, "запущен" if p3 else "не найден", key="proc3proxy"),
        VpnCheck("Процесс ProxiFyre", pf, "запущен" if pf else "не найден", key="proxifyre"),
        VpnCheck("UNC шара", unc_ok, unc_detail, required=True, key="unc"),
    )
    overall = all(item.ok for item in checks if item.required)
    return VpnSnapshot(checks=checks, overall_ok=overall, raw=raw)


def _check_by_key(snap: VpnSnapshot, key: str) -> VpnCheck | None:
    for item in snap.checks:
        if item.key == key:
            return item
    return None


def _lamp_from_check(snap: VpnSnapshot, key: str) -> FlowLamp:
    item = _check_by_key(snap, key)
    if item is None:
        return FlowLamp(None, "нет данных")
    return FlowLamp(item.ok, item.detail)


def exit_lamps(direct: str, socks: str) -> tuple[FlowLamp, FlowLamp]:
    """Office (direct) and NL (SOCKS) lamps from curl ipify strings.

    Args:
        direct: Public IP without proxy.
        socks: Public IP via SOCKS.

    Returns:
        ``(office_exit, socks_exit)``.
    """
    if direct:
        if direct == H10_EXIT_IP:
            office = FlowLamp(False, f"{direct} это выход H10")
        else:
            office = FlowLamp(True, f"{direct} (офис)")
    else:
        office = FlowLamp(None, "нажмите «Проверить IP»")
    if socks:
        if socks.startswith(SOCKS_EXIT_PREFIX):
            nl = FlowLamp(True, f"{socks} (S4)")
        else:
            nl = FlowLamp(False, f"{socks} ожидали префикс {SOCKS_EXIT_PREFIX}")
    else:
        nl = FlowLamp(None, "нажмите «Проверить IP»")
    return office, nl


def flow_lamps(
    snap: VpnSnapshot,
    *,
    direct: str = "",
    socks: str = "",
) -> dict[str, FlowLamp]:
    """Map a snapshot (+ optional curl IPs) onto schematic node lamps.

    Args:
        snap: Evaluated adapter/process snapshot.
        direct: Public IP without proxy (empty until curl).
        socks: Public IP via SOCKS (empty until curl).

    Returns:
        Node key → lamp. ``ok is None`` means not probed / not applicable.
    """
    socks_chk = _check_by_key(snap, "socks")
    p3_chk = _check_by_key(snap, "proc3proxy")
    if socks_chk is None and p3_chk is None:
        socks_lamp = FlowLamp(None, "нет данных")
    else:
        socks_ok = bool(socks_chk and socks_chk.ok and p3_chk and p3_chk.ok)
        if socks_ok:
            socks_lamp = FlowLamp(True, "слушает 127.0.0.1:1080")
        elif socks_chk and socks_chk.ok:
            socks_lamp = FlowLamp(False, "порт есть, процесс 3proxy не найден")
        else:
            socks_lamp = FlowLamp(False, "не слушает :1080")

    h10 = _lamp_from_check(snap, "h10_off")
    eth = _lamp_from_check(snap, "eth_default")
    if h10.ok is False:
        other = FlowLamp(False, "H10 забрал default — браузер тоже в NL")
    else:
        other = FlowLamp(eth.ok, "не Cursor.exe, идут на Ethernet")

    office, nl = exit_lamps(direct, socks)

    cb = _lamp_from_check(snap, "cursorbind")
    cbd = _lamp_from_check(snap, "cb_default")
    if cb.ok is False:
        tunnel = FlowLamp(False, cb.status)
    elif cbd.ok is False:
        tunnel = FlowLamp(False, "туннель есть, нет metric 9999 — SOCKS не выйдет")
    elif cb.ok and cbd.ok:
        tunnel = FlowLamp(True, f"{cb.status}; {cbd.status}")
    else:
        tunnel = cb

    lan = _lamp_from_check(snap, "lan")
    return {
        "cursor": FlowLamp(None, "не опрашивается"),
        "proxifyre": _lamp_from_check(snap, "proxifyre"),
        "socks": socks_lamp,
        "cursorbind": tunnel,
        "cb_default": cbd,
        "socks_exit": nl,
        "other_apps": other,
        "eth_default": eth,
        "office_exit": office,
        "unc": _lamp_from_check(snap, "unc"),
        "lan": lan,
        "lan_gw": FlowLamp(lan.ok, LAN_GATEWAY),
        "h10_off": (
            FlowLamp(True, "выключен — так и надо")
            if h10.ok
            else FlowLamp(False, "включён — выключить, иначе заберёт интернет")
            if h10.ok is False
            else h10
        ),
    }


def collect_vpn_raw() -> dict[str, Any]:
    """Read adapters, routes, processes, SOCKS port, UNC (Windows).

    Returns:
        JSON-able dict consumed by ``evaluate_vpn_raw``.

    Raises:
        RuntimeError: PowerShell query failed.
    """
    script = (
        "[pscustomobject]@{"
        " addrs = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue"
        " | Select-Object InterfaceAlias, IPAddress);"
        " defaults = @(Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue"
        " | Select-Object InterfaceAlias, NextHop, RouteMetric);"
        " lan = @(Get-NetRoute -DestinationPrefix '172.16.0.0/16' -ErrorAction SilentlyContinue"
        " | Select-Object InterfaceAlias, NextHop, RouteMetric)"
        "} | ConvertTo-Json -Compress -Depth 4"
    )
    data = _ps_json(script)
    if not isinstance(data, dict):
        data = {}
    data["socks"] = _socks_listening()
    data["proc3proxy"] = _task_running("3proxy.exe")
    data["procProxiFyre"] = _task_running("ProxiFyre.exe")
    data["unc"] = _unc_exists(UNC_SAMPLE)
    data["bat_exists"] = START_BAT.is_file()
    data["proxifyre_exe"] = PROXIFYRE_EXE.is_file()
    data["proxy_exe"] = (PROXY_DIR / "3proxy.exe").is_file()
    return data


def collect_vpn_status() -> VpnSnapshot:
    """Collect live status and evaluate the CursorBind stack.

    Returns:
        Evaluated snapshot for the GUI table.
    """
    return evaluate_vpn_raw(collect_vpn_raw())


def probe_public_ips(*, timeout: int = 15) -> tuple[str, str]:
    """Return (direct ipify, SOCKS ipify) strings (may be empty on timeout).

    Args:
        timeout: curl --max-time seconds.

    Returns:
        Pair of stdout strings from api.ipify.org.
    """

    def _curl(extra: list[str]) -> str:
        argv = ["curl.exe", "-s", "--max-time", str(timeout), *extra, "https://api.ipify.org"]
        proc = _run_hidden(argv, timeout=timeout + 5)
        return (proc.stdout or "").strip()

    direct = _curl([])
    socks = _curl(["--socks5-hostname", f"{SOCKS_HOST}:{SOCKS_PORT}"])
    return direct, socks


def ips_look_correct(direct: str, socks: str) -> tuple[bool, str]:
    """Check office vs NL split for the two ipify results.

    Args:
        direct: IP seen without proxy.
        socks: IP seen through 127.0.0.1:1080.

    Returns:
        (ok, detail).
    """
    if not direct or not socks:
        return False, f"direct={direct or '—'} socks={socks or 'пустой/таймаут'}"
    if direct == H10_EXIT_IP:
        return False, f"direct={direct} это выход H10, нужен офисный Ethernet"
    if not socks.startswith(SOCKS_EXIT_PREFIX):
        return False, f"socks={socks} ожидался префикс {SOCKS_EXIT_PREFIX}"
    if direct == socks:
        return False, f"оба IP одинаковые: {direct}"
    return True, f"direct={direct} socks={socks}"


def start_bat_elevated(path: Path | None = None) -> tuple[bool, str]:
    """Launch the start bat with a UAC prompt (Windows ShellExecute runas).

    Args:
        path: Bat file; default ``START_BAT``.

    Returns:
        (ok, detail). False if the file is missing or UAC was cancelled.
    """
    target = path or START_BAT
    if not target.is_file():
        return False, f"нет файла: {target}"
    if os.name != "nt":
        return False, "только Windows"
    import ctypes

    rc = int(
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            str(target),
            None,
            str(target.parent),
            1,
        )
    )
    if rc <= 32:
        return False, f"UAC/запуск не удался (код {rc})"
    return True, str(target)
