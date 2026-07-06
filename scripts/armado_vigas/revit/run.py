# -*- coding: utf-8 -*-
"""Entrada pyRevit — bootstrap portable y apertura de UI."""

from __future__ import print_function

import os
import sys

from armado_vigas.bootstrap_paths import ensure_paths, set_pushbutton_dir


def launch(uiapp, pushbutton_dir=None):
    if pushbutton_dir:
        set_pushbutton_dir(pushbutton_dir)
    ensure_paths(pushbutton_dir)
    from armado_vigas.ui.window import show_armado_vigas_window
    return show_armado_vigas_window(uiapp, pushbutton_dir)


def _resolve_uiapp(revit_globals):
    """
    pyRevit inyecta ``__revit__`` como ``UIApplication`` (``ActiveUIDocument``, etc.).
    Algunos runners pasan un dict o un wrapper con ``.uiapp``.
    """
    if revit_globals is None:
        return None
    try:
        if hasattr(revit_globals, "ActiveUIDocument"):
            return revit_globals
    except Exception:
        pass
    try:
        uiapp = revit_globals.uiapp
        if uiapp is not None:
            return uiapp
    except Exception:
        pass
    if isinstance(revit_globals, dict):
        uiapp = revit_globals.get("uiapp")
        if uiapp is not None:
            return uiapp
    return None


def run_pyrevit(revit_globals):
    import clr

    clr.AddReference("RevitAPIUI")
    from Autodesk.Revit.UI import TaskDialog

    from armado_vigas.revit.selection import (
        pick_lote_inicial,
        show_selection_instructions,
        validate_initial_selection,
    )
    from armado_vigas.revit.session import SESSION

    uiapp = _resolve_uiapp(revit_globals)
    pb = os.environ.get("ARAINCO_ARMADO_VIGAS_PB_DIR")
    if not pb:
        try:
            pb = os.path.dirname(os.path.abspath(__file__))
            for _ in range(4):
                if os.path.basename(pb) == "scripts":
                    pb = os.path.dirname(pb)
                    break
                pb = os.path.dirname(pb)
        except NameError:
            pb = None
    if uiapp is None:
        raise Exception(u"No se pudo obtener UIApplication desde pyRevit.")

    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Arainco: Armado vigas", u"No hay documento activo con vista.")
        return

    from armado_vigas.ui.window import get_existing_armado_vigas_window, show_armado_vigas_window

    if get_existing_armado_vigas_window() is not None:
        return show_armado_vigas_window(uiapp, pb)

    SESSION.reset()
    if not show_selection_instructions(uiapp):
        return

    refs = pick_lote_inicial(uidoc)
    if not refs:
        return

    ok, msg = validate_initial_selection(uidoc.Document, refs, uidoc.ActiveView)
    if not ok:
        TaskDialog.Show(u"Arainco: Armado vigas", msg)
        return

    try:
        SESSION.set_selection(uidoc.Document, refs, uidoc.ActiveView)
    except Exception as ex:
        try:
            err = unicode(ex)
        except NameError:
            err = str(ex)
        TaskDialog.Show(u"Arainco: Armado vigas", u"Error al procesar la selección:\n\n{0}".format(err))
        return

    try:
        from armado_vigas.revit.direction_overlay import (
            _AUTO_SHOW_DIRECTION_OVERLAY_ON_LAUNCH,
            show_beam_direction_overlay,
        )

        if _AUTO_SHOW_DIRECTION_OVERLAY_ON_LAUNCH:
            n_markers = show_beam_direction_overlay(uidoc.Document, uidoc.ActiveView)
            if n_markers:
                SESSION.last_message += (
                    u" · Flechas modelo: eje 0→1 ({0} trazo(s))".format(n_markers)
                )
    except Exception:
        pass

    return launch(uiapp, pb)
