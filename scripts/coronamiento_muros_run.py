# -*- coding: utf-8 -*-
"""
Entrada — Arainco: Coronamiento muros.

Multi-pick (máx. 2) → modo U libre / Empotrado → UI → Colocar.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Wall
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

_DIALOG_TITLE = u"Arainco: Coronamiento muros"
_MAX_WALLS = 2


class _WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Wall)

    def AllowReference(self, reference, point):
        return True


def _mostrar(uiapp, instruction, content=u""):
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        show_message_dialog(
            _DIALOG_TITLE,
            instruction=instruction,
            content=content or None,
            ok_text=u"Entendido",
            hwnd_revit=revit_main_hwnd(uiapp) if uiapp else None,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"{0}\n{1}".format(instruction, content or u"").strip(),
        )
    except Exception:
        pass


def pick_walls_max2(uidoc):
    """
    Multi-select de muros, máximo 2.

    Returns:
        list[Wall]: 0–2 muros. Cancelar / 0 → lista vacía.
        Si el usuario elige más de 2, se muestra aviso y se devuelve None
        (no abrir UI).
    """
    if uidoc is None:
        return []
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _WallSelectionFilter(),
            u"Seleccione 1 o 2 muros (máximo 2). Finish confirma; Esc cancela.",
        )
    except OperationCanceledException:
        return []
    except Exception:
        return []
    if not refs:
        return []
    doc = uidoc.Document
    walls = []
    seen = set()
    for ref in refs:
        if ref is None:
            continue
        try:
            el = doc.GetElement(ref)
        except Exception:
            el = None
        if not isinstance(el, Wall):
            continue
        try:
            wid = int(el.Id.IntegerValue)
        except Exception:
            continue
        if wid in seen:
            continue
        seen.add(wid)
        walls.append(el)
    if len(walls) > _MAX_WALLS:
        return None
    return walls


def resolve_pick_mode(walls):
    """
    Deriva modo geométrico del multi-pick.

    Returns:
        dict: ok, geom_mode, host, upper, walls_ord, voladizo_specs, message
    """
    from coronamiento_muros_place import resolve_coronamiento_pick

    return resolve_coronamiento_pick(walls)


def run(uiapp):
    """Entrada pyRevit / RPS: ``run(__revit__)``."""
    if uiapp is None:
        return
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        _mostrar(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _mostrar(uiapp, u"No hay documento activo.")
        return

    try:
        walls = pick_walls_max2(uidoc)
    except OperationCanceledException:
        return
    except Exception as ex:
        try:
            msg = unicode(ex)
        except Exception:
            msg = str(ex)
        _mostrar(uiapp, u"No se pudo seleccionar el muro.", msg)
        return

    if walls is None:
        _mostrar(
            uiapp,
            u"Máximo 2 muros.",
            u"Seleccione 1 muro (U libre) o 2 muros apilados (Empotrado).",
        )
        return
    if not walls:
        # 0 / cancel → salir sin UI
        return

    ctx = resolve_pick_mode(walls)
    if not ctx.get(u"ok"):
        _mostrar(
            uiapp,
            ctx.get(u"message")
            or u"Los muros seleccionados no están apilados (sin contacto en Z).",
            u"Seleccione 1 muro (U libre) o 2 muros apilados (Empotrado).",
        )
        return

    from coronamiento_muros_ui import show_coronamiento_window

    show_coronamiento_window(
        uiapp,
        uidoc,
        doc,
        ctx.get(u"host"),
        geom_mode=ctx.get(u"geom_mode") or u"u_libre",
        upper_wall=ctx.get(u"upper"),
        walls_ord=ctx.get(u"walls_ord"),
        voladizo_specs=ctx.get(u"voladizo_specs"),
        overhang_mm=ctx.get(u"overhang_mm"),
        embed_side=ctx.get(u"embed_side"),
    )


def run_pyrevit(revit_app):
    run(revit_app)
