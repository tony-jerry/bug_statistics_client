"""使用 Windows DPAPI 保存当前用户的客户端密码。"""

from __future__ import annotations

import ctypes
import json
import os
from ctypes import wintypes
from pathlib import Path


CRYPTPROTECT_UI_FORBIDDEN = 0x01


class _DataBlob(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


def _input_blob(data: bytes) -> tuple[_DataBlob, ctypes.Array[ctypes.c_char]]:
    buffer = ctypes.create_string_buffer(data)
    blob = _DataBlob(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)),
    )
    return blob, buffer


def _crypt32() -> ctypes.WinDLL:
    if os.name != "nt":
        raise OSError("密码安全保存仅支持 Windows")
    return ctypes.WinDLL("crypt32", use_last_error=True)


def _free_output(blob: _DataBlob) -> None:
    local_free = ctypes.WinDLL("kernel32", use_last_error=True).LocalFree
    local_free.argtypes = [ctypes.c_void_p]
    local_free.restype = ctypes.c_void_p
    local_free(ctypes.cast(blob.pbData, ctypes.c_void_p))


def _protect(data: bytes) -> bytes:
    crypt32 = _crypt32()
    protect = crypt32.CryptProtectData
    protect.argtypes = [
        ctypes.POINTER(_DataBlob),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    protect.restype = wintypes.BOOL

    input_blob, input_buffer = _input_blob(data)
    output_blob = _DataBlob()
    if not protect(
        ctypes.byref(input_blob),
        "BugStatisticsClient",
        None,
        None,
        None,
        CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(output_blob),
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        _free_output(output_blob)
        del input_buffer


def _unprotect(data: bytes) -> bytes:
    crypt32 = _crypt32()
    unprotect = crypt32.CryptUnprotectData
    unprotect.argtypes = [
        ctypes.POINTER(_DataBlob),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DataBlob),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DataBlob),
    ]
    unprotect.restype = wintypes.BOOL

    input_blob, input_buffer = _input_blob(data)
    output_blob = _DataBlob()
    if not unprotect(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        CRYPTPROTECT_UI_FORBIDDEN,
        ctypes.byref(output_blob),
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(output_blob.pbData, output_blob.cbData)
    finally:
        _free_output(output_blob)
        del input_buffer


def load_credentials(path: Path) -> dict[str, str]:
    """读取当前 Windows 用户加密保存的登录信息。"""
    try:
        payload = json.loads(_unprotect(path.read_bytes()).decode("utf-8"))
    except (OSError, UnicodeError, ValueError):
        return {}
    if not isinstance(payload, dict):
        return {}
    return {
        key: value
        for key in ("username", "password", "crm_password")
        if isinstance((value := payload.get(key)), str)
    }


def save_credentials(
    path: Path,
    username: str,
    password: str,
    crm_password: str,
) -> bool:
    """加密并保存登录信息；加密数据只能由当前 Windows 用户解密。"""
    payload = json.dumps(
        {
            "username": username,
            "password": password,
            "crm_password": crm_password,
        },
        ensure_ascii=False,
    ).encode("utf-8")
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path.write_bytes(_protect(payload))
        os.replace(temporary_path, path)
        return True
    except OSError:
        try:
            temporary_path.unlink(missing_ok=True)
        except OSError:
            pass
        return False
