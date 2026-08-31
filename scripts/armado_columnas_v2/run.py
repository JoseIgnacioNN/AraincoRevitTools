# -*- coding: utf-8 -*-
"""Entrada de la herramienta Armado columnas V2."""

from __future__ import print_function


def _resolve_uiapp(revit_globals):
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


def run(revit_globals, pushbutton_dir=None):
    """
    Flujo:
      1. Pick columnas + fundaciones + vigas + losas en vista activa
      2. Proyectar a dominio de elevación
      3. Abrir UI (canvas a escala de vista)
    """
    import clr

    clr.AddReference("RevitAPIUI")
    from Autodesk.Revit.UI import TaskDialog

    from armado_columnas_v2.session import SESSION
    from armado_columnas_v2.ui.window import (
        get_existing_armado_columnas_v2_window,
        show_armado_columnas_v2_window,
    )

    _DIALOG = u"Arainco: Armado columnas V2"

    uiapp = _resolve_uiapp(revit_globals)
    if uiapp is None:
        TaskDialog.Show(_DIALOG, u"No se pudo obtener UIApplication desde pyRevit.")
        return None

    if get_existing_armado_columnas_v2_window() is not None:
        return show_armado_columnas_v2_window(uiapp, pushbutton_dir)

    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_DIALOG, u"No hay documento activo con vista.")
        return None

    view = uidoc.ActiveView
    if view is None:
        TaskDialog.Show(_DIALOG, u"No hay vista activa.")
        return None

    from armado_columnas_v2.revit.selection import (
        pick_lote_inicial,
        validate_initial_selection,
    )

    refs = pick_lote_inicial(uidoc)
    if not refs:
        return None

    ok, msg = validate_initial_selection(uidoc.Document, refs, view)
    if not ok:
        TaskDialog.Show(_DIALOG, msg)
        return None

    try:
        members = SESSION.set_selection(uidoc.Document, refs, view)
    except Exception as ex:
        try:
            err = unicode(ex)
        except NameError:
            err = str(ex)
        TaskDialog.Show(_DIALOG, u"Error al procesar la selección:\n\n{0}".format(err))
        return None

    if not members:
        TaskDialog.Show(
            _DIALOG,
            u"No se pudo proyectar ningún elemento en la vista activa.\n"
            u"Use un alzado o sección donde sean visibles.",
        )
        return None

    return show_armado_columnas_v2_window(uiapp, pushbutton_dir)
