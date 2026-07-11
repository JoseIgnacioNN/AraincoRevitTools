# -*- coding: utf-8 -*-
"""
Validación de acceso corporativo para herramientas BIMTools.

Comprueba un archivo JSON en el servidor de incidencias. Pensado para
reutilizarse desde script.py de cada pushbutton (una llamada al inicio).

Ruta de validación (producción):
  Y:\\00_SERVIDOR DE INCIDENCIAS\\RECURSOS COMPARTIDOS\\bimtools_access.json
"""

from __future__ import print_function

import json
import os
import time

try:
    import clr

    clr.AddReference("System")
    import System
    import System.IO as sio

    _USE_SYSTEM_IO = True
except Exception:
    _USE_SYSTEM_IO = False

ISSUES_DIR = u"Y:\\00_SERVIDOR DE INCIDENCIAS"
SHARED_RESOURCES_DIR = os.path.join(ISSUES_DIR, u"RECURSOS COMPARTIDOS")
ACCESS_FILE = os.path.join(SHARED_RESOURCES_DIR, u"bimtools_access.json")

DENIED_MESSAGE = u"Este equipo no tiene permiso para utilizar estas herramientas"
DEFAULT_DIALOG_TITLE = u"Arainco: BIMTools"

_APPDOMAIN_CACHE_KEY = u"BIMTools.CorporateAccess.Cache"
_DEFAULT_CACHE_TTL_SECONDS = 300


def _read_text_file(path):
    if _USE_SYSTEM_IO:
        try:
            if sio.File.Exists(path):
                return sio.File.ReadAllText(path, System.Text.Encoding.UTF8)
        except Exception:
            pass
    try:
        import codecs

        if os.path.isfile(path):
            with codecs.open(path, "r", encoding="utf-8") as handle:
                return handle.read()
    except Exception:
        pass
    return None


def _load_access_payload():
    raw = _read_text_file(ACCESS_FILE)
    if not raw:
        return None
    if raw.startswith(u"\ufeff"):
        raw = raw.lstrip(u"\ufeff")
    try:
        data = json.loads(raw)
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    return data


def _path_exists(path):
    if _USE_SYSTEM_IO:
        try:
            return bool(sio.Directory.Exists(path) or sio.File.Exists(path))
        except Exception:
            pass
    return os.path.exists(path)


def _get_access_file_signature():
    """
    Firma del archivo de acceso (mtime + tamaño).
    None si el archivo no existe o no es accesible.
    """
    if _USE_SYSTEM_IO:
        try:
            info = sio.FileInfo(ACCESS_FILE)
            if not info.Exists:
                return None
            return (int(info.LastWriteTimeUtc.Ticks), int(info.Length))
        except Exception:
            pass
    try:
        if not os.path.isfile(ACCESS_FILE):
            return None
        st = os.stat(ACCESS_FILE)
        return (int(st.st_mtime), int(st.st_size))
    except Exception:
        return None


def _parse_expires_at(value):
    if not value:
        return None
    text = unicode(value).strip()
    if not text:
        return None
    try:
        import datetime

        if text.endswith("Z"):
            text = text[:-1]
        if u"T" in text:
            return datetime.datetime.strptime(text[:19], u"%Y-%m-%dT%H:%M:%S")
        return datetime.datetime.strptime(text[:10], u"%Y-%m-%d")
    except Exception:
        return None


def _get_cached_entry(ttl_seconds):
    try:
        cached = System.AppDomain.CurrentDomain.GetData(_APPDOMAIN_CACHE_KEY)
    except Exception:
        cached = None
    if not isinstance(cached, dict):
        return None
    checked_at = cached.get("checked_at")
    if checked_at is None:
        return None
    try:
        age = time.time() - float(checked_at)
    except Exception:
        return None
    if age > float(ttl_seconds):
        return None
    return cached


def _set_cached_entry(ok, file_sig, reason):
    try:
        System.AppDomain.CurrentDomain.SetData(
            _APPDOMAIN_CACHE_KEY,
            {
                "ok": bool(ok),
                "file_sig": file_sig,
                "reason": reason,
                "checked_at": time.time(),
            },
        )
    except Exception:
        pass


def clear_access_cache():
    """Limpia la caché en memoria (útil para pruebas)."""
    try:
        System.AppDomain.CurrentDomain.SetData(_APPDOMAIN_CACHE_KEY, None)
    except Exception:
        pass


def check_corporate_access(use_cache=True, cache_ttl_seconds=_DEFAULT_CACHE_TTL_SECONDS):
    """
    Evalúa si el equipo puede usar herramientas BIMTools.

    Siempre comprueba rutas y existencia del archivo antes de usar caché,
    para que quitar el JSON del servidor bloquee de inmediato.

    Returns:
        tuple(bool ok, unicode reason_code)
    """
    if not _path_exists(ISSUES_DIR):
        _set_cached_entry(False, None, u"missing_server")
        return False, u"missing_server"

    if not _path_exists(SHARED_RESOURCES_DIR):
        _set_cached_entry(False, None, u"missing_folder")
        return False, u"missing_folder"

    file_sig = _get_access_file_signature()
    if file_sig is None:
        _set_cached_entry(False, None, u"missing_file")
        return False, u"missing_file"

    if use_cache:
        cached = _get_cached_entry(cache_ttl_seconds)
        if (
            cached is not None
            and cached.get("file_sig") == file_sig
            and cached.get("reason") not in (u"missing_server", u"missing_folder", u"missing_file")
        ):
            return bool(cached.get("ok")), cached.get("reason", u"cache")

    payload = _load_access_payload()
    if payload is None:
        _set_cached_entry(False, file_sig, u"invalid_file")
        return False, u"invalid_file"

    if not payload.get("enabled", False):
        _set_cached_entry(False, file_sig, u"disabled")
        return False, u"disabled"

    expires_at = _parse_expires_at(payload.get("expires_at"))
    if expires_at is not None:
        import datetime

        if datetime.datetime.now() > expires_at:
            _set_cached_entry(False, file_sig, u"expired")
            return False, u"expired"

    _set_cached_entry(True, file_sig, u"ok")
    return True, u"ok"


def ensure_corporate_access(dialog_title=None, uiapp=None, use_cache=True, cache_ttl_seconds=_DEFAULT_CACHE_TTL_SECONDS):
    """
    Valida acceso corporativo. Si falla, muestra diálogo WPF estilo BIMTools y devuelve False.
    """
    ok, _reason = check_corporate_access(
        use_cache=use_cache,
        cache_ttl_seconds=cache_ttl_seconds,
    )
    if ok:
        return True

    title = dialog_title or DEFAULT_DIALOG_TITLE
    _show_denied_dialog(title, uiapp)
    return False


def _show_denied_dialog(title, uiapp=None):
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass

    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            title,
            DENIED_MESSAGE,
            content=u"",
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass

    try:
        import clr

        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(title, DENIED_MESSAGE)
    except Exception:
        try:
            from pyrevit import forms

            forms.alert(DENIED_MESSAGE, title=title)
        except Exception:
            print(DENIED_MESSAGE)


def get_access_paths():
    """Rutas usadas por la validación (útil para diagnóstico)."""
    return {
        "issues_dir": ISSUES_DIR,
        "shared_resources_dir": SHARED_RESOURCES_DIR,
        "access_file": ACCESS_FILE,
    }
