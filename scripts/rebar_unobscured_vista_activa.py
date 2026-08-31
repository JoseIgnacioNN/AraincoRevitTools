# -*- coding: utf-8 -*-
"""
Ver / Ocultar Armadura en la vista activa — interruptor según icono.

Icono off («Ver Armadura») → aplica Unobscured.
Icono on («Ocultar Armadura») → quita Unobscured.
Sin diálogo de acción.

Revit 2024+ | pyRevit / IronPython.
Entrada: ``37_RebarUnobscuredVista.smartbutton/script.py`` (botón ligero).
"""

from __future__ import print_function

from Autodesk.Revit.DB import Transaction, View, ViewSchedule, ViewSheet
from Autodesk.Revit.UI import TaskDialog

from bimtools_rebar_3d_visibility import (
    apply_reinforcement_unobscured_in_view,
    collect_reinforcement_in_view,
)

__title__ = u"Arainco: Ver Armadura"
__title_apply__ = u"Arainco: Ver Armadura"
__title_remove__ = u"Arainco: Ocultar Armadura"

_TITLE_OFF = u"Ver\nArmadura"
_TITLE_ON = u"Ocultar\nArmadura"
_ENV_ON = u"BIMTOOLS_REBAR_UNOBSCURED_ON"
_AD_BUTTON = u"BIMTOOLS_REBAR_UNOBSCURED_UI_BUTTON"
_ICON_LARGE = 32


def _mostrar_aviso(uiapp, instruction, content=u"", title=None, ok_text=u"Entendido"):
    dlg_title = title or __title__
    hwnd = None
    try:
        if uiapp is not None:
            from revit_wpf_window_position import revit_main_hwnd

            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            dlg_title,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(dlg_title, body)
    except Exception:
        pass


def _resolve_active_model_view(uidoc):
    u"""
    Vista gráfica activa del documento, resuelta por ElementId.

    Prioriza ``ActiveGraphicalView`` (pestaña visual) frente a ``ActiveView``
    (puede desviarse con el navegador de proyecto o paneles UI). Solo se
    trabaja con esa instancia: nunca plantillas ni otras vistas abiertas.
    """
    if uidoc is None:
        return None
    doc = uidoc.Document
    if doc is None:
        return None

    view = None
    try:
        view = getattr(uidoc, "ActiveGraphicalView", None)
    except Exception:
        view = None
    if view is None:
        try:
            view = uidoc.ActiveView
        except Exception:
            view = None
    if view is None:
        return None

    try:
        resolved = doc.GetElement(view.Id)
    except Exception:
        resolved = None
    if isinstance(resolved, View):
        view = resolved

    if not isinstance(view, View):
        return None
    try:
        if view.IsTemplate:
            return None
    except Exception:
        pass
    return view


def _is_eye_open():
    try:
        from pyrevit import script

        return bool(script.get_envvar(_ENV_ON))
    except Exception:
        return False


def _ribbon_title(is_open):
    return _TITLE_ON if is_open else _TITLE_OFF


def _set_ribbon_titles(is_open):
    """Actualiza el texto del botón en la cinta (off=Ver, on=Ocultar)."""
    title = _ribbon_title(is_open)
    try:
        from System import AppDomain

        btn = AppDomain.CurrentDomain.GetData(_AD_BUTTON)
        if btn is not None and hasattr(btn, "set_title"):
            btn.set_title(title)
    except Exception:
        pass
    try:
        from pyrevit import script

        for btn in script.get_all_buttons() or []:
            if btn is not None and hasattr(btn, "set_title"):
                btn.set_title(title)
    except Exception:
        pass


def _set_eye_state(is_open):
    """Actualiza envvar + icono + título on/off. No lanza."""
    try:
        from pyrevit import script
    except Exception:
        script = None
    try:
        from pyrevit.coreutils.ribbon import ICON_LARGE

        icon_size = ICON_LARGE
    except Exception:
        icon_size = _ICON_LARGE

    is_open = bool(is_open)
    if script is not None:
        try:
            script.set_envvar(_ENV_ON, is_open)
            script.toggle_icon(is_open, icon_size=icon_size)
        except Exception:
            try:
                script.set_envvar(_ENV_ON, is_open)
            except Exception:
                pass
    _set_ribbon_titles(is_open)


def run(revit_app):
    uidoc = revit_app.ActiveUIDocument
    if uidoc is None:
        _mostrar_aviso(revit_app, u"No hay documento activo.")
        return
    doc = uidoc.Document
    view = _resolve_active_model_view(uidoc)

    if view is None:
        _mostrar_aviso(
            revit_app,
            u"No hay una vista de modelo activa (o la activa es una plantilla de vista).",
            content=u"Activa una planta, alzado, sección o 3D e inténtalo de nuevo.",
        )
        return

    if isinstance(view, (ViewSheet, ViewSchedule)):
        _mostrar_aviso(
            revit_app,
            u"Abre una vista de modelo (planta, alzado, sección, 3D…), no una lámina ni un cuadro.",
        )
        return

    target_view_id = view.Id
    refuerzos = collect_reinforcement_in_view(doc, view)
    if not refuerzos:
        _mostrar_aviso(
            revit_app,
            u"No hay barras ni armaduras en la vista activa «{0}».".format(
                getattr(view, "Name", None) or u"(vista)"
            ),
        )
        return

    # Off («Ver Armadura») → aplicar; on («Ocultar Armadura») → quitar.
    eye_open = _is_eye_open()
    unobscured = not eye_open
    tx_name = __title_apply__ if unobscured else __title_remove__

    try:
        view = doc.GetElement(target_view_id)
    except Exception:
        view = None
    if not isinstance(view, View) or isinstance(view, (ViewSheet, ViewSchedule)):
        _mostrar_aviso(
            revit_app,
            u"La vista activa ya no está disponible. Vuelve a ejecutar la herramienta.",
            title=tx_name,
        )
        return

    t = Transaction(doc, tx_name)
    t.Start()
    try:
        apply_reinforcement_unobscured_in_view(
            doc, refuerzos, view, unobscured=unobscured
        )
        t.Commit()
    except Exception as ex:
        t.RollBack()
        _mostrar_aviso(revit_app, u"Error: {0}".format(ex), title=tx_name)
        raise

    _set_eye_state(unobscured)
