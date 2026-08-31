# -*- coding: utf-8 -*-
"""
Instrumentación de canvas sketch — log de clics, snap, transform y rechazos.

Log de sesión en %%TEMP%%\\Arainco_FundacionCorridaFranja_canvas.log
Overlay en UI: Ctrl+Shift+D (AppDomain ``Arainco.FundacionCorridaFranja.CanvasDebug``).
"""

from __future__ import print_function

import os
import time

from System import AppDomain

_DEBUG_KEY = u"Arainco.FundacionCorridaFranja.CanvasDebug"
_DEFAULT_LOG = u"Arainco_FundacionCorridaFranja_canvas.log"
_MAX_UI_LINES = 10
_MAX_FILE_BYTES = 512 * 1024


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _now_stamp():
    t = time.time()
    ms = int((t - int(t)) * 1000.0)
    return time.strftime(u"%Y-%m-%d %H:%M:%S", time.localtime(t)) + u".{0:03d}".format(ms)


class CanvasSketchInstrument(object):
    """Log append-only + buffer corto para overlay de depuración."""

    def __init__(self, log_basename=_DEFAULT_LOG):
        temp = os.environ.get(u"TEMP") or os.environ.get(u"TMP") or u"."
        self._log_path = os.path.join(temp, log_basename)
        self._ui_lines = []
        self._session_id = 0

    @property
    def log_path(self):
        return self._log_path

    def is_ui_enabled(self):
        try:
            return bool(AppDomain.CurrentDomain.GetData(_DEBUG_KEY))
        except Exception:
            return False

    def set_ui_enabled(self, enabled):
        try:
            AppDomain.CurrentDomain.SetData(_DEBUG_KEY, bool(enabled))
        except Exception:
            pass

    def toggle_ui(self):
        self.set_ui_enabled(not self.is_ui_enabled())
        return self.is_ui_enabled()

    def begin_session(self, meta=None):
        self._session_id += 1
        self._write(u"=" * 72)
        self._write(
            u"SESSION {0} START {1}".format(self._session_id, _now_stamp())
        )
        if meta:
            for key in sorted(meta.keys()):
                self._write(u"  {0}={1}".format(key, _as_unicode(meta[key])))
        self._ui_lines = []

    def event(self, name, **fields):
        line = self._format_line(name, fields)
        self._write(line)
        self._ui_lines.append(line)
        if len(self._ui_lines) > _MAX_UI_LINES:
            self._ui_lines = self._ui_lines[-_MAX_UI_LINES:]
        return line

    def ui_text(self):
        if not self._ui_lines:
            return u"(sin eventos de canvas)"
        header = u"DEBUG · {0}".format(self._log_path)
        return header + u"\n" + u"\n".join(self._ui_lines[-_MAX_UI_LINES:])

    def _format_line(self, name, fields):
        parts = [u"[{0}] {1}".format(_now_stamp(), _as_unicode(name))]
        for key in sorted(fields.keys()):
            val = fields[key]
            if val is None:
                continue
            parts.append(u"{0}={1}".format(key, _as_unicode(val)))
        return u" | ".join(parts)

    def _write(self, line):
        try:
            self._rotate_if_needed()
            import codecs

            with codecs.open(self._log_path, u"a", u"utf-8") as handle:
                handle.write(line + u"\n")
        except Exception:
            pass

    def _rotate_if_needed(self):
        try:
            if not os.path.isfile(self._log_path):
                return
            if os.path.getsize(self._log_path) <= _MAX_FILE_BYTES:
                return
            backup = self._log_path + u".old"
            try:
                if os.path.isfile(backup):
                    os.remove(backup)
            except Exception:
                pass
            try:
                os.rename(self._log_path, backup)
            except Exception:
                pass
        except Exception:
            pass


def dist_mm(a, b):
    try:
        dx = float(a[0]) - float(b[0])
        dy = float(a[1]) - float(b[1])
        return (dx * dx + dy * dy) ** 0.5
    except Exception:
        return None


def dist_px(a, b):
    try:
        dx = float(a[0]) - float(b[0])
        dy = float(a[1]) - float(b[1])
        return (dx * dx + dy * dy) ** 0.5
    except Exception:
        return None
