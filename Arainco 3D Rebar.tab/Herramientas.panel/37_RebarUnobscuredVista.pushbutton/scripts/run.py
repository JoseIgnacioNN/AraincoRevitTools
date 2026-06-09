# -*- coding: utf-8 -*-
"""
View Unobscured en la vista activa: aplicar, quitar o consultar estado de barras y armaduras.

Revit 2024+ | pyRevit / IronPython.
"""

from __future__ import print_function

from Autodesk.Revit.DB import Transaction, ViewSchedule, ViewSheet

from lib.rebar_3d_visibility import (
    apply_reinforcement_unobscured_in_view,
    collect_reinforcement_in_view,
    summarize_reinforcement_unobscured_in_view,
)
from ui.action_dialog import (
    show_info_dialog,
    show_rebar_unobscured_action_dialog,
)

__title__ = u"Arainco: View Unobscured barras"
__title_apply__ = u"Arainco: Aplicar View Unobscured barras"
__title_remove__ = u"Arainco: Quitar View Unobscured barras"


def _count_by_type(refuerzos):
    n_rebar = 0
    n_in_system = 0
    n_area = 0
    try:
        from Autodesk.Revit.DB.Structure import AreaReinforcement, Rebar, RebarInSystem

        for el in refuerzos:
            if isinstance(el, Rebar):
                n_rebar += 1
            elif isinstance(el, RebarInSystem):
                n_in_system += 1
            elif isinstance(el, AreaReinforcement):
                n_area += 1
    except Exception:
        pass
    return n_rebar, n_in_system, n_area


def _estado_texto(summary):
    return (
        u"Estado actual en la vista:\n"
        u"  · Con View Unobscured: {0}\n"
        u"  · Sin View Unobscured: {1}\n"
        u"  · Sin dato / no aplica: {2}"
    ).format(
        summary.get("unobscured", 0),
        summary.get("obscured", 0),
        summary.get("unknown", 0),
    )


def _resumen_dialogo(view_name, refuerzos, summary):
    n_rebar, n_in_system, n_area = _count_by_type(refuerzos)
    return (
        u"Elementos en vista: {0}\n"
        u"  · Rebar: {1}\n"
        u"  · RebarInSystem: {2}\n"
        u"  · AreaReinforcement: {3}\n\n"
        u"{4}"
    ).format(
        len(refuerzos),
        n_rebar,
        n_in_system,
        n_area,
        _estado_texto(summary),
    )


def _preguntar_accion(view_name, refuerzos, summary, uiapp):
    return show_rebar_unobscured_action_dialog(
        title=__title__,
        subtitle=u"Barras en vista «{0}»".format(view_name),
        summary=_resumen_dialogo(view_name, refuerzos, summary),
        uiapp=uiapp,
    )


def run(revit_app):
    uidoc = revit_app.ActiveUIDocument
    if uidoc is None:
        show_info_dialog(__title__, u"No hay documento activo.", uiapp=revit_app)
        return
    doc = uidoc.Document
    view = uidoc.ActiveView

    if isinstance(view, (ViewSheet, ViewSchedule)):
        show_info_dialog(
            __title__,
            u"Abre una vista de modelo (planta, alzado, sección, 3D…), no una lámina ni un cuadro.",
            uiapp=revit_app,
        )
        return

    refuerzos = collect_reinforcement_in_view(doc, view)
    if not refuerzos:
        show_info_dialog(
            __title__,
            u"No hay barras ni armaduras en la vista activa «{0}».".format(
                getattr(view, "Name", None) or u"(vista)"
            ),
            uiapp=revit_app,
        )
        return

    view_name = getattr(view, "Name", None) or u"(vista)"
    summary = summarize_reinforcement_unobscured_in_view(doc, refuerzos, view)
    accion = _preguntar_accion(view_name, refuerzos, summary, revit_app)
    if accion is None:
        return

    if accion == "status":
        show_info_dialog(
            __title__,
            u"Vista: {0}\n\nElementos en vista: {1}\n\n{2}".format(
                view_name,
                len(refuerzos),
                _estado_texto(summary),
            ),
            uiapp=revit_app,
        )
        return

    unobscured = accion == "apply"
    tx_name = __title_apply__ if unobscured else __title_remove__

    t = Transaction(doc, tx_name)
    t.Start()
    try:
        n_ok = apply_reinforcement_unobscured_in_view(
            doc, refuerzos, view, unobscured=unobscured
        )
        t.Commit()
    except Exception as ex:
        t.RollBack()
        show_info_dialog(tx_name, u"Error: {0}".format(ex), uiapp=revit_app)
        raise

    summary_after = summarize_reinforcement_unobscured_in_view(doc, refuerzos, view)
    accion_txt = u"Aplicado" if unobscured else u"Quitado"
    msg = (
        u"Vista: {0}\n\n"
        u"{1} View Unobscured en {2} elemento(s).\n\n"
        u"{3}"
    ).format(
        view_name,
        accion_txt,
        n_ok,
        _estado_texto(summary_after),
    )
    show_info_dialog(tx_name, msg, uiapp=revit_app)
