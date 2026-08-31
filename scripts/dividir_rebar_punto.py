# -*- coding: utf-8 -*-
"""
Dividir y Traslapar — lógica / avisos.

Revit 2024+ | pyRevit | IronPython
"""

from __future__ import print_function

_TITULO = u"Arainco: Dividir y Traslapar"


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


def grade_from_combo_text(text):
    """Mapea texto de combo a concrete_grade (G25 / G35 / G45; default G25)."""
    s = _as_unicode(text).strip().upper()
    if s in (u"G25", u"G35", u"G45"):
        return s
    return u"G25"


def run(revit):
    """Pick de Rebar → UI con la barra cargada (instancia única)."""
    from dividir_rebar_punto_ui import show_dividir_rebar_punto_window

    show_dividir_rebar_punto_window(revit)
