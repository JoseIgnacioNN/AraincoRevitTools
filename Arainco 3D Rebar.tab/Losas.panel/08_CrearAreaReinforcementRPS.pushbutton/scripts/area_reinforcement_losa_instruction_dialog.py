# -*- coding: utf-8 -*-
"""Instrucciones iniciales Malla en Losa — delega en bimtools_instruction_dialog."""

_DIALOG = u"Arainco: Malla en Losa"

_SELECTION_CONTENT = (
    u"Esta herramienta crea mallas de Area Reinforcement en losas de hormigón.\n\n"
    u"Flujo:\n"
    u"1. Tras Aceptar, elija losas (Floor) en el modelo. Finalice con Finish "
    u"en la barra de opciones (ESC cancela).\n"
    u"2. Configure malla superior e inferior: diámetro y espaciado.\n"
    u"3. Pulse «Colocar armaduras» para crear las mallas y etiquetarlas en planta.\n\n"
    u"Requisitos: losas con sketch cerrado, AreaReinforcementType, RebarBarType "
    u"y RebarHookType en el proyecto."
)


def show_ok_cancel_dialog(
    title,
    instruction,
    content=u"",
    ok_text=u"Aceptar",
    cancel_text=u"Cancelar",
    hwnd_revit=None,
    uiapp=None,
):
    from bimtools_instruction_dialog import show_ok_cancel_dialog as _show

    return _show(
        title,
        instruction,
        content=content,
        ok_text=ok_text,
        cancel_text=cancel_text,
        hwnd_revit=hwnd_revit,
        uiapp=uiapp,
    )


def show_message_dialog(title, instruction, ok_text=u"Entendido", uiapp=None):
    from bimtools_instruction_dialog import show_message_dialog as _show

    return _show(title, instruction, ok_text=ok_text, uiapp=uiapp)


def show_selection_instructions(uiapp=None, hwnd_revit=None):
    """
    Instrucciones previas a la selección de losas.
    Devuelve ``True`` si el usuario pulsa Aceptar; ``False`` si cancela.
    """
    if hwnd_revit is None and uiapp is not None:
        try:
            from revit_wpf_window_position import revit_main_hwnd

            hwnd_revit = revit_main_hwnd(uiapp)
        except Exception:
            hwnd_revit = None

    return show_ok_cancel_dialog(
        _DIALOG,
        u"Seleccione una o más losas a armar.",
        content=_SELECTION_CONTENT,
        ok_text=u"Aceptar",
        cancel_text=u"Cancelar",
        hwnd_revit=hwnd_revit,
        uiapp=uiapp,
    )
