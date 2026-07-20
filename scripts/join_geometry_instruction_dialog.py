# -*- coding: utf-8 -*-
"""Diálogos WPF de la herramienta — shell estándar BIMTools."""

from bimtools_instruction_dialog import show_message_dialog, show_ok_cancel_dialog

_DIALOG = u"Arainco: Unir geometría (hormigón, vista)"


def _hwnd_revit(uiapp):
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            return revit_main_hwnd(uiapp)
    except Exception:
        pass
    return None


def show_selection_instructions(uiapp=None):
    """
    Pantalla de inicio al ejecutar el botón.
    Devuelve ``True`` si el usuario pulsa Aceptar; ``False`` si cancela.
    """
    hwnd = _hwnd_revit(uiapp)
    return show_ok_cancel_dialog(
        _DIALOG,
        u"Une Join Geometry entre elementos de hormigón visibles en la vista activa.",
        u"1. Abra una vista de modelo (planta, 3D, corte, etc.).\n"
        u"2. Se analizan muros, forjados, pilares y cimentación con material "
        u"estructural Concrete.\n"
        u"3. Los pares candidatos se detectan por solape de cajas; los forjados "
        u"quedan recortados cuando corresponde.\n\n"
        u"Pulse Aceptar para iniciar el proceso. Esc cancela.",
        ok_text=u"Aceptar",
        cancel_text=u"Cancelar",
        hwnd_revit=hwnd,
        uiapp=uiapp,
    )


def mostrar_aviso(uiapp, instruction, content=u""):
    """Aviso informativo (éxito / fracaso / validación). Respaldo: TaskDialog."""
    hwnd = _hwnd_revit(uiapp)
    try:
        show_message_dialog(
            _DIALOG,
            instruction,
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

        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_DIALOG, body)
    except Exception:
        pass
