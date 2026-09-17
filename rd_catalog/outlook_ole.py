"""Read Outlook virtual files from the live OLE ``IDataObject``.

Qt 6 ``QMimeData.data()`` only tries ``TYMED_HGLOBAL`` then ``TYMED_ISTREAM``
and treats an empty HGLOBAL as success. Outlook ``FileContents`` for a ``.msg``
is usually ``TYMED_ISTORAGE``. This module keeps the drag ``IDataObject`` by
wrapping the window ``IDropTarget`` (IAT hook of ``RegisterDragDrop`` on
``qwindows.dll``) and reads IStorage / IStream / HGLOBAL the same way as
EisenhowerMatrix ``OutlookDataObjectReader``.
"""

from __future__ import annotations

import ctypes
import sys
from ctypes import POINTER, Structure, byref, c_ushort, c_void_p, sizeof, wintypes

_TYMED_HGLOBAL = 1
_TYMED_ISTREAM = 4
_TYMED_ISTORAGE = 8
_DVASPECT_CONTENT = 1
_STGM_CREATE_RW_EXCL = 0x00001012
_S_OK = 0
_E_NOINTERFACE_SIGNED = ctypes.c_long(0x80004002).value
_IID_IUNKNOWN = "{00000000-0000-0000-C000-000000000046}"
_IID_IDROPTARGET = "{00000122-0000-0000-C000-000000000046}"
_MAX_MSG_BYTES = 80 * 1024 * 1024

_hook_installed = False
_hook_error = ""
_register_callback = None
_wrappers: list[object] = []
_last_data_object = 0


class FORMATETC(Structure):
    _fields_ = [
        ("cfFormat", c_ushort),
        ("ptd", c_void_p),
        ("dwAspect", wintypes.DWORD),
        ("lindex", wintypes.LONG),
        ("tymed", wintypes.DWORD),
    ]


class STGMEDIUM(Structure):
    _fields_ = [
        ("tymed", wintypes.DWORD),
        ("handle", c_void_p),
        ("pUnkForRelease", c_void_p),
    ]


class GUID(Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", ctypes.c_ubyte * 8),
    ]


class POINTL(Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


def ensure_drop_hook() -> str:
    """Install the ``RegisterDragDrop`` wrapper before Qt registers targets.

    Returns:
        Empty string on success, otherwise a short Russian reason.
    """

    global _hook_installed, _hook_error
    if _hook_installed and not _hook_error:
        return ""
    if _hook_installed and "ещё не загружен" not in _hook_error:
        return _hook_error
    if sys.platform != "win32":
        _hook_error = "OLE-хук только для Windows."
        _hook_installed = True
        return _hook_error
    try:
        _hook_error = _install_register_drag_drop_hook()
    except Exception as exc:
        _hook_error = f"Не удалось установить OLE-хук: {exc}"
    _hook_installed = True
    return _hook_error


def last_drop_data_object() -> int:
    """Return the captured ``IDataObject*`` from the current drop, or 0."""

    return _last_data_object


def hook_status() -> str:
    """Return the last hook install result (empty when the hook is active)."""

    if not _hook_installed:
        return "OLE-хук ещё не установлен."
    return _hook_error


def read_file_contents(data_object: int, index: int = 0) -> tuple[bytes, str]:
    """Read ``FileContents`` via IStorage, IStream, then HGLOBAL.

    Args:
        data_object: ``IDataObject*`` as an integer.
        index: Zero-based ``FORMATETC.lindex``.

    Returns:
        Payload and a short Russian note about the tymed that worked.
    """

    return read_ole_format(data_object, "FileContents", index=index)


def read_ole_format(
    data_object: int,
    format_name: str,
    *,
    index: int = -1,
) -> tuple[bytes, str]:
    """Get one clipboard format from ``IDataObject``.

    Args:
        data_object: ``IDataObject*`` as an integer.
        format_name: Registered clipboard format name.
        index: ``lindex``; ``-1`` means the default for non-file formats.

    Returns:
        Bytes and a note. Empty bytes when GetData fails.
    """

    if not data_object:
        return b"", "нет IDataObject"
    cf = _register_clipboard_format(format_name)
    if not cf:
        return b"", f"не зарегистрирован формат {format_name}"
    indexes = (index,) if index >= 0 else (0, -1)
    if format_name.casefold() != "filecontents":
        indexes = (index,)
    last_error = ""
    for lindex in indexes:
        for tymed, label in (
            (_TYMED_ISTORAGE, "ISTORAGE"),
            (_TYMED_ISTREAM, "ISTREAM"),
            (_TYMED_HGLOBAL, "HGLOBAL"),
        ):
            payload, error = _get_data(data_object, cf, lindex, tymed)
            if payload:
                return payload, f"{format_name} {label} index={lindex} {len(payload)} байт"
            if error:
                last_error = error
    return b"", last_error or f"{format_name}: пусто"


def read_hglobal_handle(handle: int) -> bytes:
    """Copy bytes from an HGLOBAL (tests and HGLOBAL tymed)."""

    if not handle:
        return b""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    lock = kernel32.GlobalLock
    lock.argtypes = [c_void_p]
    lock.restype = c_void_p
    size_fn = kernel32.GlobalSize
    size_fn.argtypes = [c_void_p]
    size_fn.restype = ctypes.c_size_t
    unlock = kernel32.GlobalUnlock
    unlock.argtypes = [c_void_p]
    unlock.restype = wintypes.BOOL
    pointer = lock(handle)
    if not pointer:
        return b""
    try:
        size = int(size_fn(handle))
        if size <= 0 or size > _MAX_MSG_BYTES:
            return b""
        return ctypes.string_at(pointer, size)
    finally:
        unlock(handle)


def read_istream_ptr(stream: int) -> bytes:
    """Read an ``IStream*`` to bytes."""

    if not stream:
        return b""
    _seek_stream(stream, 0)
    chunks: list[bytes] = []
    total = 0
    buffer = ctypes.create_string_buffer(81_920)
    read_fn = _vtable_func(
        stream,
        3,
        ctypes.HRESULT,
        c_void_p,
        c_void_p,
        wintypes.ULONG,
        POINTER(wintypes.ULONG),
    )
    while total < _MAX_MSG_BYTES:
        got = wintypes.ULONG(0)
        hr = read_fn(stream, buffer, len(buffer), byref(got))
        if hr < 0 or got.value <= 0:
            break
        chunks.append(buffer.raw[: got.value])
        total += got.value
        if got.value < len(buffer):
            break
    return b"".join(chunks)


def read_istorage_ptr(storage: int) -> bytes:
    """Serialize ``IStorage*`` to an OLE compound-file byte string (``.msg``)."""

    if not storage:
        return b""
    ole32 = ctypes.WinDLL("ole32")
    create_lock = ole32.CreateILockBytesOnHGlobal
    create_lock.argtypes = [c_void_p, wintypes.BOOL, POINTER(c_void_p)]
    create_lock.restype = ctypes.HRESULT
    create_doc = ole32.StgCreateDocfileOnILockBytes
    create_doc.argtypes = [c_void_p, wintypes.DWORD, wintypes.DWORD, POINTER(c_void_p)]
    create_doc.restype = ctypes.HRESULT
    lock_bytes = c_void_p()
    hr = create_lock(None, True, byref(lock_bytes))
    if hr < 0 or not lock_bytes.value:
        return b""
    dest = c_void_p()
    hr = create_doc(lock_bytes, _STGM_CREATE_RW_EXCL, 0, byref(dest))
    if hr < 0 or not dest.value:
        _release(lock_bytes.value)
        return b""
    try:
        copy_to = _vtable_func(
            storage,
            7,
            ctypes.HRESULT,
            c_void_p,
            wintypes.DWORD,
            c_void_p,
            c_void_p,
            c_void_p,
        )
        hr = copy_to(storage, 0, None, None, dest.value)
        if hr < 0:
            return b""
        commit = _vtable_func(dest.value, 9, ctypes.HRESULT, c_void_p, wintypes.DWORD)
        commit(dest.value, 0)
        flush = _vtable_func(lock_bytes.value, 5, ctypes.HRESULT, c_void_p)
        flush(lock_bytes.value)
        return _read_lock_bytes(lock_bytes.value)
    finally:
        _release(dest.value)
        _release(lock_bytes.value)


def _get_data(
    data_object: int,
    cf: int,
    lindex: int,
    tymed: int,
) -> tuple[bytes, str]:
    fmt = FORMATETC()
    fmt.cfFormat = cf
    fmt.dwAspect = _DVASPECT_CONTENT
    fmt.lindex = lindex
    fmt.tymed = tymed
    medium = STGMEDIUM()
    get_data = _vtable_func(
        data_object,
        3,
        ctypes.HRESULT,
        c_void_p,
        POINTER(FORMATETC),
        POINTER(STGMEDIUM),
    )
    hr = get_data(data_object, byref(fmt), byref(medium))
    if hr < 0:
        return b"", f"GetData hr=0x{hr & 0xFFFFFFFF:08X} tymed={tymed}"
    try:
        if medium.tymed == _TYMED_ISTORAGE:
            payload = read_istorage_ptr(medium.handle or 0)
            return payload, "" if payload else "ISTORAGE пуст"
        if medium.tymed == _TYMED_ISTREAM:
            payload = read_istream_ptr(medium.handle or 0)
            return payload, "" if payload else "ISTREAM пуст"
        if medium.tymed == _TYMED_HGLOBAL:
            payload = read_hglobal_handle(medium.handle or 0)
            return payload, "" if payload else "HGLOBAL пуст"
        return b"", f"неподдерживаемый tymed={medium.tymed}"
    finally:
        _release_stg_medium(medium)


def _read_lock_bytes(lock_bytes: int) -> bytes:
    class _ULARGE(Structure):
        _fields_ = [("quad", ctypes.c_ulonglong)]

    class _STATSTG(Structure):
        _fields_ = [
            ("pwcsName", c_void_p),
            ("type", wintypes.DWORD),
            ("cbSize", _ULARGE),
            ("_rest", ctypes.c_ubyte * 64),
        ]

    stat = _STATSTG()
    stat_fn = _vtable_func(
        lock_bytes,
        9,
        ctypes.HRESULT,
        c_void_p,
        POINTER(_STATSTG),
        wintypes.DWORD,
    )
    hr = stat_fn(lock_bytes, byref(stat), 1)
    if hr < 0:
        return b""
    size = int(stat.cbSize.quad)
    if size <= 0 or size > _MAX_MSG_BYTES:
        return b""
    buf = ctypes.create_string_buffer(size)
    read = wintypes.ULONG(0)
    read_at = _vtable_func(
        lock_bytes,
        3,
        ctypes.HRESULT,
        c_void_p,
        ctypes.c_ulonglong,
        c_void_p,
        wintypes.ULONG,
        POINTER(wintypes.ULONG),
    )
    hr = read_at(lock_bytes, 0, buf, size, byref(read))
    if hr < 0:
        return b""
    return buf.raw[: read.value]


def _seek_stream(stream: int, position: int) -> None:
    class _LARGE(Structure):
        _fields_ = [("quad", ctypes.c_longlong)]

    seek = _vtable_func(
        stream,
        5,
        ctypes.HRESULT,
        c_void_p,
        _LARGE,
        wintypes.DWORD,
        c_void_p,
    )
    seek(stream, _LARGE(position), 0, None)


def _register_clipboard_format(name: str) -> int:
    user32 = ctypes.WinDLL("user32")
    fn = user32.RegisterClipboardFormatW
    fn.argtypes = [wintypes.LPCWSTR]
    fn.restype = wintypes.UINT
    return int(fn(name))


def _release_stg_medium(medium: STGMEDIUM) -> None:
    ole32 = ctypes.WinDLL("ole32")
    fn = ole32.ReleaseStgMedium
    fn.argtypes = [POINTER(STGMEDIUM)]
    fn.restype = None
    fn(byref(medium))


def _vtable_func(punk: int, index: int, restype, *argtypes):
    vptr = ctypes.cast(punk, POINTER(c_void_p)).contents.value
    slot = ctypes.cast(vptr + index * sizeof(c_void_p), POINTER(c_void_p)).contents.value
    proto = ctypes.WINFUNCTYPE(restype, *argtypes)
    return proto(slot)


def _addref(punk: int) -> None:
    if punk:
        _vtable_func(punk, 1, wintypes.ULONG, c_void_p)(punk)


def _release(punk: int) -> None:
    if punk:
        _vtable_func(punk, 2, wintypes.ULONG, c_void_p)(punk)


def _capture_data_object(punk: int) -> None:
    global _last_data_object
    if _last_data_object == punk:
        return
    if _last_data_object:
        _release(_last_data_object)
        _last_data_object = 0
    if punk:
        _addref(punk)
        _last_data_object = punk


def _guid(text: str) -> GUID:
    value = text.strip("{}")
    data1, data2, data3, data4 = value.split("-", 3)
    raw = bytes.fromhex(data4)
    guid = GUID()
    guid.Data1 = int(data1, 16)
    guid.Data2 = int(data2, 16)
    guid.Data3 = int(data3, 16)
    for i, byte in enumerate(raw):
        guid.Data4[i] = byte
    return guid


def _guid_eq(left: GUID, text: str) -> bool:
    right = _guid(text)
    return (
        left.Data1 == right.Data1
        and left.Data2 == right.Data2
        and left.Data3 == right.Data3
        and bytes(left.Data4) == bytes(right.Data4)
    )


class _DropTargetObject(Structure):
    _fields_ = [
        ("lpVtbl", c_void_p),
        ("inner", c_void_p),
        ("refcount", wintypes.ULONG),
    ]


_QueryInterface = ctypes.WINFUNCTYPE(
    ctypes.HRESULT, c_void_p, POINTER(GUID), POINTER(c_void_p)
)
_AddRef = ctypes.WINFUNCTYPE(wintypes.ULONG, c_void_p)
_Release = ctypes.WINFUNCTYPE(wintypes.ULONG, c_void_p)
_DragEnter = ctypes.WINFUNCTYPE(
    ctypes.HRESULT, c_void_p, c_void_p, wintypes.DWORD, POINTL, POINTER(wintypes.DWORD)
)
_DragOver = ctypes.WINFUNCTYPE(
    ctypes.HRESULT, c_void_p, wintypes.DWORD, POINTL, POINTER(wintypes.DWORD)
)
_DragLeave = ctypes.WINFUNCTYPE(ctypes.HRESULT, c_void_p)
_Drop = ctypes.WINFUNCTYPE(
    ctypes.HRESULT, c_void_p, c_void_p, wintypes.DWORD, POINTL, POINTER(wintypes.DWORD)
)


class _DropTargetVtbl(Structure):
    _fields_ = [
        ("QueryInterface", _QueryInterface),
        ("AddRef", _AddRef),
        ("Release", _Release),
        ("DragEnter", _DragEnter),
        ("DragOver", _DragOver),
        ("DragLeave", _DragLeave),
        ("Drop", _Drop),
    ]


def _as_wrapper(this: int) -> _DropTargetObject:
    return ctypes.cast(this, POINTER(_DropTargetObject)).contents


def _query_interface(this, riid, ppv) -> int:
    iid = riid.contents
    if _guid_eq(iid, _IID_IUNKNOWN) or _guid_eq(iid, _IID_IDROPTARGET):
        ppv[0] = this
        _add_ref(this)
        return _S_OK
    ppv[0] = None
    return _E_NOINTERFACE_SIGNED


def _add_ref(this: int) -> int:
    obj = _as_wrapper(this)
    obj.refcount += 1
    return obj.refcount


def _release_wrapper(this: int) -> int:
    obj = _as_wrapper(this)
    obj.refcount -= 1
    return obj.refcount


def _drag_enter(this, p_data, key_state, point, effect) -> int:
    _capture_data_object(p_data or 0)
    inner = _as_wrapper(this).inner
    fn = _vtable_func(
        inner,
        3,
        ctypes.HRESULT,
        c_void_p,
        c_void_p,
        wintypes.DWORD,
        POINTL,
        POINTER(wintypes.DWORD),
    )
    return fn(inner, p_data, key_state, point, effect)


def _drag_over(this, key_state, point, effect) -> int:
    inner = _as_wrapper(this).inner
    fn = _vtable_func(
        inner,
        4,
        ctypes.HRESULT,
        c_void_p,
        wintypes.DWORD,
        POINTL,
        POINTER(wintypes.DWORD),
    )
    return fn(inner, key_state, point, effect)


def _drag_leave(this) -> int:
    inner = _as_wrapper(this).inner
    fn = _vtable_func(inner, 5, ctypes.HRESULT, c_void_p)
    return fn(inner)


def _drop(this, p_data, key_state, point, effect) -> int:
    _capture_data_object(p_data or 0)
    inner = _as_wrapper(this).inner
    fn = _vtable_func(
        inner,
        6,
        ctypes.HRESULT,
        c_void_p,
        c_void_p,
        wintypes.DWORD,
        POINTL,
        POINTER(wintypes.DWORD),
    )
    return fn(inner, p_data, key_state, point, effect)


_VTBL = _DropTargetVtbl(
    _QueryInterface(_query_interface),
    _AddRef(_add_ref),
    _Release(_release_wrapper),
    _DragEnter(_drag_enter),
    _DragOver(_drag_over),
    _DragLeave(_drag_leave),
    _Drop(_drop),
)


def _wrap_drop_target(inner: int) -> int:
    obj = _DropTargetObject()
    obj.lpVtbl = ctypes.addressof(_VTBL)
    obj.inner = inner
    obj.refcount = 1
    _addref(inner)
    _wrappers.append(obj)
    return ctypes.addressof(obj)


def _install_register_drag_drop_hook() -> str:
    global _register_callback
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    get_module = kernel32.GetModuleHandleW
    get_module.argtypes = [wintypes.LPCWSTR]
    get_module.restype = wintypes.HMODULE
    module = int(get_module("qwindows") or 0)
    if not module:
        return "qwindows.dll ещё не загружен."
    ole32 = ctypes.WinDLL("ole32")
    orig = ole32.RegisterDragDrop
    orig.argtypes = [wintypes.HWND, c_void_p]
    orig.restype = ctypes.HRESULT

    proto = ctypes.WINFUNCTYPE(ctypes.HRESULT, wintypes.HWND, c_void_p)

    def _on_register(hwnd, punk):
        wrapped = _wrap_drop_target(punk) if punk else punk
        return orig(hwnd, wrapped)

    _register_callback = proto(_on_register)
    if _patch_iat(module, "ole32.dll", "RegisterDragDrop", _register_callback):
        return ""
    return "В IAT qwindows.dll нет RegisterDragDrop."


class _IMAGE_DOS_HEADER(Structure):
    _fields_ = [("e_magic", wintypes.WORD), ("_skip", ctypes.c_ubyte * 58), ("e_lfanew", wintypes.LONG)]


class _IMAGE_DATA_DIRECTORY(Structure):
    _fields_ = [("VirtualAddress", wintypes.DWORD), ("Size", wintypes.DWORD)]


class _IMAGE_IMPORT_DESCRIPTOR(Structure):
    _fields_ = [
        ("OriginalFirstThunk", wintypes.DWORD),
        ("TimeDateStamp", wintypes.DWORD),
        ("ForwarderChain", wintypes.DWORD),
        ("Name", wintypes.DWORD),
        ("FirstThunk", wintypes.DWORD),
    ]


def _patch_iat(module: int, dll_name: str, func_name: str, callback) -> bool:
    dos = _IMAGE_DOS_HEADER.from_address(module)
    if dos.e_magic != 0x5A4D:
        return False
    nt = module + dos.e_lfanew
    magic = ctypes.c_ushort.from_address(nt + 4 + 16).value  # FileHeader + SizeOfOptionalHeader? 
    # PE signature 4, FileHeader 20, OptionalHeader Magic at start
    opt_magic = ctypes.c_ushort.from_address(nt + 24).value
    if opt_magic == 0x20B:
        import_rva = wintypes.DWORD.from_address(nt + 24 + 120).value
    elif opt_magic == 0x10B:
        import_rva = wintypes.DWORD.from_address(nt + 24 + 104).value
    else:
        return False
    if not import_rva:
        return False
    desc_size = sizeof(_IMAGE_IMPORT_DESCRIPTOR)
    index = 0
    wanted_dll = dll_name.casefold()
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    virtual_protect = kernel32.VirtualProtect
    virtual_protect.argtypes = [
        c_void_p,
        ctypes.c_size_t,
        wintypes.DWORD,
        POINTER(wintypes.DWORD),
    ]
    virtual_protect.restype = wintypes.BOOL
    while True:
        desc = _IMAGE_IMPORT_DESCRIPTOR.from_address(module + import_rva + index * desc_size)
        if desc.Name == 0 and desc.FirstThunk == 0:
            break
        name = ctypes.string_at(module + desc.Name).decode("ascii", "replace")
        if name.casefold() == wanted_dll:
            if _patch_thunks(module, desc, func_name, callback, virtual_protect):
                return True
        index += 1
    return False


def _patch_thunks(module: int, desc, func_name: str, callback, virtual_protect) -> bool:
    oft = desc.OriginalFirstThunk or desc.FirstThunk
    iat = desc.FirstThunk
    slot = 0
    while True:
        orig_thunk = ctypes.c_uint64.from_address(module + oft + slot * 8).value
        if orig_thunk == 0:
            return False
        if orig_thunk & (1 << 63):
            slot += 1
            continue
        import_name = ctypes.string_at(module + orig_thunk + 2).decode("ascii", "replace")
        if import_name == func_name:
            iat_addr = module + iat + slot * 8
            old = wintypes.DWORD()
            if not virtual_protect(iat_addr, 8, 0x40, byref(old)):
                return False
            ctypes.c_uint64.from_address(iat_addr).value = ctypes.cast(
                callback, c_void_p
            ).value
            virtual_protect(iat_addr, 8, old.value, byref(old))
            return True
        slot += 1
