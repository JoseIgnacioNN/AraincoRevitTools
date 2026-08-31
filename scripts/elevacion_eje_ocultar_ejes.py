# -*- coding: utf-8 -*-
"""
Oculta ejes innecesarios en una Elevación Eje.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

En la vista, recolecta muros y Structural Framing cuya LocationCurve no es
paralela al plano (producto punto con ViewDirection no cercano a 0). Extrae
sus sólidos y oculta los Grid cuya curva ViewSpecific no intersecta ninguno.

El eje que originó la elevación no se oculta.
No abre transacción: el caller ya tiene una TX abierta.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    DatumExtentType,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    Line,
    LocationCurve,
    Options,
    Solid,
    SolidCurveIntersectionOptions,
)

try:
    from System.Collections.Generic import List as DotNetList
except Exception:
    DotNetList = None

_TOL_PARALELO = 0.01


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _eid_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _solidos_elemento(element, opt):
    """Sólidos con volumen > 0 (geometría directa e instancias)."""
    solids = []
    if element is None or opt is None:
        return solids
    try:
        geom_elem = element.get_Geometry(opt)
    except Exception:
        geom_elem = None
    if not geom_elem:
        return solids
    try:
        for g_obj in geom_elem:
            if isinstance(g_obj, Solid):
                try:
                    if g_obj.Volume > 0:
                        solids.append(g_obj)
                except Exception:
                    pass
            elif isinstance(g_obj, GeometryInstance):
                try:
                    inst_geom = g_obj.GetInstanceGeometry()
                except Exception:
                    inst_geom = None
                if not inst_geom:
                    continue
                for i_obj in inst_geom:
                    if isinstance(i_obj, Solid):
                        try:
                            if i_obj.Volume > 0:
                                solids.append(i_obj)
                        except Exception:
                            pass
    except Exception:
        return solids
    return solids


def _paralelo_al_plano_vista(element, view_dir):
    """
    True si la LocationCurve (Line) es paralela al plano de la vista
    (producto punto con la normal cercano a 0).
    """
    if element is None or view_dir is None:
        return False
    try:
        loc = element.Location
    except Exception:
        return False
    if not isinstance(loc, LocationCurve):
        return False
    try:
        curve = loc.Curve
    except Exception:
        return False
    if not isinstance(curve, Line):
        return False
    try:
        direction = curve.Direction
        dot_product = abs(float(direction.DotProduct(view_dir)))
        return dot_product < _TOL_PARALELO
    except Exception:
        return False


def _collector_categoria_vista(document, view, bic):
    try:
        return (
            FilteredElementCollector(document, view.Id)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return None


def _curvas_eje_en_vista(grid, view):
    if grid is None or view is None:
        return []
    try:
        curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
    except Exception:
        curves = None
    if not curves:
        return []
    out = []
    try:
        for curve in curves:
            if curve is not None:
                out.append(curve)
    except Exception:
        return []
    return out


def _eje_intersecta_solidos(grid, view, solids, inter_options):
    curves = _curvas_eje_en_vista(grid, view)
    if not curves or not solids:
        return False
    for curve in curves:
        for solid in solids:
            try:
                inter_result = solid.IntersectWithCurve(curve, inter_options)
                if inter_result is not None and int(inter_result.SegmentCount) > 0:
                    return True
            except Exception:
                continue
    return False


def _puede_ocultar(view, elem):
    if view is None or elem is None:
        return False
    try:
        if elem.IsHidden(view):
            return False
    except Exception:
        pass
    try:
        if hasattr(elem, "CanBeHidden") and not elem.CanBeHidden(view):
            return False
    except Exception:
        return False
    return True


def ocultar_ejes_sin_interseccion(document, view, exclude_grid=None):
    """
    Oculta grids de ``view`` que no intersectan sólidos de muros/vigas
    no paralelos al plano.

    Args:
        document: Document
        view: vista de elevación recién creada
        exclude_grid: Grid origen (no se oculta)

    Returns:
        dict: ok, n_hidden, n_grids, n_solids, skipped, error
    """
    info = {
        u"ok": False,
        u"n_hidden": 0,
        u"n_grids": 0,
        u"n_solids": 0,
        u"skipped": False,
        u"error": u"",
    }
    if document is None or view is None:
        info[u"error"] = u"vista o documento nulo"
        return info

    col_grids = _collector_categoria_vista(
        document, view, BuiltInCategory.OST_Grids
    )
    if col_grids is None:
        info[u"error"] = u"no se pudieron recolectar ejes"
        return info
    grids = [g for g in col_grids if g is not None]
    info[u"n_grids"] = len(grids)
    if not grids:
        info[u"ok"] = True
        info[u"skipped"] = True
        return info

    walls_col = _collector_categoria_vista(
        document, view, BuiltInCategory.OST_Walls
    )
    framing_col = _collector_categoria_vista(
        document, view, BuiltInCategory.OST_StructuralFraming
    )
    elements_to_check = []
    if walls_col is not None:
        elements_to_check.extend([e for e in walls_col if e is not None])
    if framing_col is not None:
        elements_to_check.extend([e for e in framing_col if e is not None])

    try:
        view_dir = view.ViewDirection
    except Exception:
        view_dir = None
    if view_dir is None:
        info[u"error"] = u"sin ViewDirection"
        return info

    geom_options = Options()
    try:
        geom_options.View = view
        geom_options.ComputeReferences = False
    except Exception:
        pass

    target_solids = []
    for elem in elements_to_check:
        if _paralelo_al_plano_vista(elem, view_dir):
            continue
        target_solids.extend(_solidos_elemento(elem, geom_options))

    info[u"n_solids"] = len(target_solids)
    if not target_solids:
        info[u"ok"] = True
        info[u"skipped"] = True
        return info

    exclude_id = _eid_int(getattr(exclude_grid, u"Id", None))
    inter_options = SolidCurveIntersectionOptions()
    ids_hide = []
    for grid in grids:
        gid = _eid_int(getattr(grid, u"Id", None))
        if exclude_id is not None and gid == exclude_id:
            continue
        if _eje_intersecta_solidos(grid, view, target_solids, inter_options):
            continue
        if not _puede_ocultar(view, grid):
            continue
        try:
            ids_hide.append(grid.Id)
        except Exception:
            continue

    if not ids_hide:
        info[u"ok"] = True
        return info

    if DotNetList is None:
        info[u"error"] = u"List[ElementId] no disponible"
        return info

    try:
        net_ids = DotNetList[ElementId]()
        for eid in ids_hide:
            net_ids.Add(eid)
        view.HideElements(net_ids)
        info[u"n_hidden"] = int(net_ids.Count)
        info[u"ok"] = True
        return info
    except Exception as ex:
        info[u"error"] = _as_unicode(ex)
        return info
