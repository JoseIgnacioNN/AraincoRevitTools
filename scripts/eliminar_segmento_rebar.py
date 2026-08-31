# -*- coding: utf-8 -*-
"""
Eliminar segmento rebar — avisos y entrada.

Revit 2024+ | pyRevit | IronPython
"""

from __future__ import print_function

_TITULO = u"Arainco: Eliminar segmento rebar"
_ALREADY_RUNNING = u"La herramienta ya esta en ejecucion."


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def mostrar_aviso(uiapp, instruction, content=u""):
    """Aviso WPF BIMTools; respaldo a TaskDialog."""
    instruction = _as_unicode(instruction)
    content = _as_unicode(content)
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = None
        try:
            hwnd = revit_main_hwnd(uiapp)
        except Exception:
            hwnd = None
        show_message_dialog(
            _TITULO,
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
        import clr

        clr.AddReference("RevitAPIUI")
        from Autodesk.Revit.UI import TaskDialog

        msg = instruction
        if content:
            msg = u"{0}\n\n{1}".format(instruction, content)
        TaskDialog.Show(_TITULO, msg)
    except Exception:
        pass


def run(revit):
    from eliminar_segmento_rebar_ui import show_eliminar_segmento_window

    show_eliminar_segmento_window(revit)
