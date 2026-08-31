# -*- coding: utf-8 -*-
"""Abre el Manual de modelado de hormigón (Arainco) en el navegador.

Revit 2024–2026 · IronPython (pyRevit)
"""

from __future__ import print_function

import os
import sys

_MANUAL_NAME = u"manual_modelado_hormigon.html"
_DIALOG_TITLE = u"Arainco: Manual modelado hormigón"


def _as_unicode(value):
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _show_message(uiapp, instruction, content=u""):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp else None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        from Autodesk.Revit.UI import TaskDialog

        msg = instruction
        if content:
            msg = instruction + u"\n\n" + content
        TaskDialog.Show(_DIALOG_TITLE, msg)
    except Exception:
        pass


def _resolve_manual_path():
    try:
        import bimtools_paths

        pushbutton_dir = bimtools_paths.get_pushbutton_dir()
    except Exception:
        pushbutton_dir = None
    if pushbutton_dir:
        candidate = os.path.join(pushbutton_dir, _MANUAL_NAME)
        if os.path.isfile(candidate):
            return candidate
    here = os.path.dirname(os.path.abspath(__file__))
    cursor = here
    for _ in range(16):
        for rel in (
            os.path.join(
                "BIMTools.tab",
                "Estandares.panel",
                "01_ManualModeladoHormigon.pushbutton",
                _MANUAL_NAME,
            ),
            os.path.join("Estandares.panel", "01_ManualModeladoHormigon.pushbutton", _MANUAL_NAME),
        ):
            candidate = os.path.join(cursor, rel)
            if os.path.isfile(candidate):
                return candidate
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def run(uiapp):
    manual_path = _resolve_manual_path()
    if not manual_path:
        _show_message(
            uiapp,
            u"No se encontró el manual web.",
            u"Archivo esperado: {}".format(_MANUAL_NAME),
        )
        return
    try:
        os.startfile(os.path.abspath(manual_path))
    except Exception as ex:
        _show_message(
            uiapp,
            u"No se pudo abrir el manual.",
            _as_unicode(ex),
        )
