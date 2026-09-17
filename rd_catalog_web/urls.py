"""Listen URLs for the read-only catalog web monitor.

``0.0.0.0`` is a bind address, not a URL. This PC opens ``127.0.0.1``;
colleagues on the LAN open a non-VPN IPv4 (Ethernet/Wi-Fi), not a tunnel.
"""

from __future__ import annotations

import json
import socket
import subprocess
import sys
import time

DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 8765
_ALIAS_TTL_S = 45.0
_VPN_ALIAS_HINTS = (
    "vpn",
    "openvpn",
    "wireguard",
    "wintun",
    "nordlynx",
    "mullvad",
    "proton",
    "windscribe",
    "anyconnect",
    "cisco",
    "hamachi",
    "zerotier",
    "tailscale",
    "amsterdam",
    "netherlands",
    "outline",
    "tun",
    "tap-windows",
)
_alias_cache: tuple[float, dict[str, str]] | None = None


def _is_loopback(ip: str) -> bool:
    return ip.startswith("127.")


def _is_apipa(ip: str) -> bool:
    return ip.startswith("169.254.")


def _rfc1918_rank(ip: str) -> int:
    try:
        parts = [int(item) for item in ip.split(".")]
    except ValueError:
        return 9
    if len(parts) != 4:
        return 9
    first, second = parts[0], parts[1]
    if first == 192 and second == 168:
        return 0
    if first == 172 and 16 <= second <= 31:
        return 1
    if first == 10:
        return 2
    return 8


def _looks_like_vpn_alias(alias: str) -> bool:
    lowered = alias.casefold()
    return any(hint in lowered for hint in _VPN_ALIAS_HINTS)


def _windows_ip_aliases() -> dict[str, str]:
    """Return IPv4 → adapter alias on Windows. Empty when PowerShell fails."""

    global _alias_cache
    now = time.monotonic()
    cached = _alias_cache
    if cached is not None and now - cached[0] < _ALIAS_TTL_S:
        return cached[1]
    mapping: dict[str, str] = {}
    try:
        raw = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-NetIPAddress -AddressFamily IPv4 | "
                "Select-Object IPAddress,InterfaceAlias | ConvertTo-Json -Compress",
            ],
            timeout=3,
            stderr=subprocess.DEVNULL,
        )
        payload = json.loads(raw.decode("utf-8", errors="replace") or "[]")
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError, ValueError):
        _alias_cache = (now, mapping)
        return mapping
    rows = payload if isinstance(payload, list) else [payload]
    for row in rows:
        if not isinstance(row, dict):
            continue
        ip = str(row.get("IPAddress") or "").strip()
        alias = str(row.get("InterfaceAlias") or "").strip()
        if ip and alias:
            mapping[ip] = alias
    _alias_cache = (now, mapping)
    return mapping


def _share_sort_key(ip: str, aliases: dict[str, str]) -> tuple[int, int, str]:
    alias = aliases.get(ip, "")
    vpn = 1 if _looks_like_vpn_alias(alias) else 0
    return (vpn, _rfc1918_rank(ip), ip)


def lan_ipv4() -> list[str]:
    """Return non-loopback IPv4 addresses of this machine.

    APIPA ``169.254.*`` is skipped. VPN-named adapters are listed after
    Ethernet/Wi-Fi. ``192.168/16`` then ``172.16/12`` then ``10/8``.

    Returns:
        Deduplicated addresses, possibly empty when offline.
    """

    found: list[str] = []
    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, None, socket.AF_INET):
            ip = info[4][0]
            if _is_loopback(ip) or _is_apipa(ip) or ip in found:
                continue
            found.append(ip)
    except OSError:
        pass
    if not found:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
                sock.connect(("1.1.1.1", 80))
                ip = sock.getsockname()[0]
                if ip and not _is_loopback(ip) and not _is_apipa(ip):
                    found.append(ip)
        except OSError:
            pass
    aliases = _windows_ip_aliases() if sys.platform == "win32" else {}
    found.sort(key=lambda ip: _share_sort_key(ip, aliases))
    return found


def listen_urls(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> list[str]:
    """Return http URLs this process can be opened at.

    Args:
        host: Bind address passed to uvicorn (``0.0.0.0`` = all interfaces).
        port: TCP port.

    Returns:
        Localhost first, then LAN IPv4 URLs when bound on all interfaces.
        ``http://0.0.0.0:...`` is never returned.
    """

    if host in {"0.0.0.0", "::", ""}:
        urls = [f"http://127.0.0.1:{port}"]
        for ip in lan_ipv4():
            urls.append(f"http://{ip}:{port}")
        return urls
    display = "127.0.0.1" if host == "127.0.0.1" else host
    return [f"http://{display}:{port}"]


def local_open_url(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> str:
    """Return the URL to open in a browser on this PC.

    Args:
        host: Bind address.
        port: TCP port.

    Returns:
        Loopback URL (never ``0.0.0.0``).
    """

    return listen_urls(host, port)[0]


def lan_share_url(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> str:
    """Return the URL to give colleagues on the LAN.

    Args:
        host: Bind address.
        port: TCP port.

    Returns:
        First non-loopback listen URL after VPN deprioritize, else loopback.
    """

    urls = listen_urls(host, port)
    for url in urls:
        if "127.0.0.1" not in url:
            return url
    return urls[0]
