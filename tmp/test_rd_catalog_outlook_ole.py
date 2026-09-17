"""Local checks for Outlook OLE HGLOBAL / IStream readers."""

from __future__ import annotations

import ctypes
import sys
from ctypes import POINTER, byref, c_void_p, wintypes
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.outlook_ole import read_hglobal_handle, read_istream_ptr


def _global_from_bytes(payload: bytes) -> int:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    alloc = kernel32.GlobalAlloc
    alloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    alloc.restype = c_void_p
    lock = kernel32.GlobalLock
    lock.argtypes = [c_void_p]
    lock.restype = c_void_p
    unlock = kernel32.GlobalUnlock
    unlock.argtypes = [c_void_p]
    unlock.restype = wintypes.BOOL
    handle = alloc(0x0002, len(payload))  # GMEM_MOVEABLE
    assert handle
    pointer = lock(handle)
    ctypes.memmove(pointer, payload, len(payload))
    unlock(handle)
    return handle


def main() -> None:
    """Read a synthetic HGLOBAL and IStream."""

    payload = b"outlook-ole-msg"
    handle = _global_from_bytes(payload)
    assert read_hglobal_handle(handle) == payload

    ole32 = ctypes.WinDLL("ole32")
    create = ole32.CreateStreamOnHGlobal
    create.argtypes = [c_void_p, wintypes.BOOL, POINTER(c_void_p)]
    create.restype = ctypes.HRESULT
    stream = c_void_p()
    hr = create(handle, True, byref(stream))
    assert hr == 0 and stream.value
    assert read_istream_ptr(stream.value) == payload
    print("outlook_ole ok")


if __name__ == "__main__":
    main()
