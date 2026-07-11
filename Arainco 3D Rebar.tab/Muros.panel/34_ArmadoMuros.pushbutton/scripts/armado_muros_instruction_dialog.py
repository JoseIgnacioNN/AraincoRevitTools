# -*- coding: utf-8 -*-
"""Diálogo modal WPF — shell estándar BIMTools (cinta blanca + cuerpo oscuro)."""

from Autodesk.Revit.UI import TaskDialog

_DIALOG = u"Arainco: Armado Muros"

try:
    from bimtools_instruction_dialog import (
        show_message_dialog,
        show_ok_cancel_dialog,
    )
except Exception:
    show_message_dialog = None
    show_ok_cancel_dialog = None


def _hwnd_from_uiapp(uiapp):
    hwnd = None
    try:
        from revit_wpf_window_position import revit_main_hwnd

        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    return hwnd


def show_building_section_view_required(view, uiapp=None):
    """Aviso al inicio si la vista activa no es Building Section."""
    hwnd = _hwnd_from_uiapp(uiapp)
    try:
        from armado_muros_etiqueta_malla import texto_aviso_vista_building_section

        instruction, content = texto_aviso_vista_building_section(view)
    except Exception:
        instruction = (
            u"Esta herramienta solo puede ejecutarse en secciones "
            u"tipo Building Section."
        )
        content = u"Abra una sección de edificio antes de continuar."
    if show_message_dialog is not None:
        try:
            return show_message_dialog(
                _DIALOG,
                instruction,
                content,
                ok_text=u"Entendido",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )
        except Exception:
            pass
    TaskDialog.Show(
        _DIALOG,
        u"{0}\n\n{1}".format(instruction, content).strip(),
    )
    return True


def show_selection_instructions(uiapp=None):
    """
    Instrucciones previas a la selección en modelo (al ejecutar el botón).
    Devuelve ``True`` si el usuario pulsa Aceptar; ``False`` si cancela.
    """
    hwnd = _hwnd_from_uiapp(uiapp)
    instruction = u"Seleccione uno o más muros a armar."
    content = (
        u"Pulse Aceptar para iniciar la selección en el modelo. "
        u"Finalice con la cinta (Finalizar) o cancela con Esc."
    )
    if show_ok_cancel_dialog is not None:
        try:
            return show_ok_cancel_dialog(
                _DIALOG,
                instruction,
                content,
                ok_text=u"Aceptar",
                cancel_text=u"Cancelar",
                hwnd_revit=hwnd,
                uiapp=uiapp,
            )
        except Exception:
            pass
    from Autodesk.Revit.UI import TaskDialogCommonButtons, TaskDialogResult

    td = TaskDialog(_DIALOG)
    td.MainInstruction = instruction
    td.MainContent = content
    td.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
    return td.Show() == TaskDialogResult.Ok
