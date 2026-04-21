# -*- coding: utf-8 -*-
"""
Leyenda de Rebar Shapes (solo formas).

Genera/actualiza una `Drafting View` con nombre fijo `VIEW_NAME` y dibuja:
1) Rectángulos de grilla (celdas).
2) Geometría 2D de cada `RebarShape` realmente usada en el proyecto.

Requiere Revit 2024+ (pyRevit / IronPython).
"""

import clr
import math

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Arc,
    CurveElement,
    DetailCurve,
    ElementId,
    FilteredElementCollector,
    Line,
    Transaction,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    View,
    XYZ,
)
from Autodesk.Revit.DB.Structure import RebarShape
from Autodesk.Revit.UI import TaskDialog


VIEW_NAME = u"TIPO DE BARRAS"

# Grilla (ajustable)
GRID_COLS = 4
GRID_ROWS_MIN = 3
CELL_W_MM = 60.0
CELL_H_MM = 40.0
MARGIN_MM = 2.0


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _get_drafting_view_family_type(doc):
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft and vft.ViewFamily == ViewFamily.Drafting:
            return vft
    return None


def _find_drafting_view_by_name(doc, view_name):
    # Vista se identifica por nombre exacto.
    for v in FilteredElementCollector(doc).OfClass(ViewDrafting):
        try:
            if v and v.Name == view_name:
                return v
        except Exception:
            pass
    return None


def _create_or_get_view(doc, view_name):
    view = _find_drafting_view_by_name(doc, view_name)
    if view:
        return view

    # Si existe una vista con el mismo nombre pero no es Drafting,
    # evitamos crear una nueva sin poder garantizar el nombre fijo.
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if v and v.Name == view_name and not isinstance(v, ViewDrafting):
                raise Exception(u'Ya existe una vista con el nombre "{}" que no es Drafting.'.format(view_name))
        except Exception:
            # Si no se puede inspeccionar, se ignora.
            pass

    vft = _get_drafting_view_family_type(doc)
    if not vft:
        raise Exception(u"No se encontró ViewFamilyType para Drafting.")

    view = ViewDrafting.Create(doc, vft.Id)
    view.Name = view_name
    return view


def _clear_detail_curves_in_view(doc, drafting_view):
    # Borrar solo el contenido de esa vista. OfClass(DetailCurve) falla en algunas APIs;
    # se usa CurveElement y se filtra por DetailCurve.
    view_id = drafting_view.Id
    to_delete = []
    for ce in FilteredElementCollector(doc).OfClass(CurveElement):
        try:
            if ce is None or ce.OwnerViewId != view_id:
                continue
            if not isinstance(ce, DetailCurve):
                continue
            to_delete.append(ce.Id)
        except Exception:
            pass
    for did in to_delete:
        try:
            doc.Delete(did)
        except Exception:
            pass


def _collect_project_rebar_shapes(doc):
    """
    Devuelve un dict:
      {ElementId(shape_id): RebarShape}
    con TODOS los RebarShape existentes en el proyecto.
    """
    shapes = {}
    try:
        for rs in FilteredElementCollector(doc).OfClass(RebarShape):
            if not rs:
                continue
            try:
                sid = rs.Id
                if sid and sid != ElementId.InvalidElementId and hasattr(rs, "GetCurvesForBrowser"):
                    shapes[sid] = rs
            except Exception:
                pass
    except Exception:
        pass
    return shapes


def _bbox_from_curves_local(curves):
    """
    Calcula bbox (minx,maxx,miny,maxy) tomando puntos evaluados en:
    inicio, 25%, 50%, 75% y fin del dominio de cada curva.

    Asume que las curvas de `GetCurvesForBrowser()` están en un plano
    cuyo eje local relevante es X/Y (Z cercano a 0).
    """
    min_x = 1e99
    min_y = 1e99
    max_x = -1e99
    max_y = -1e99
    any_pts = False

    for c in curves:
        if c is None:
            continue
        try:
            p0 = c.GetEndParameter(0)
            p1 = c.GetEndParameter(1)
            dom = p1 - p0
            if abs(dom) < 1e-12:
                dom = 0.0
            sample_params = [p0, p0 + 0.25 * dom, p0 + 0.5 * dom, p0 + 0.75 * dom, p1]
            for sp in sample_params:
                pt = c.Evaluate(sp, False)  # False = parámetros no normalizados
                if pt is None:
                    continue
                any_pts = True
                x = float(pt.X)
                y = float(pt.Y)
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)
        except Exception:
            # Fallback: endpoints (si hay tipo de curva no evaluable por Evaluate)
            try:
                pt_a = c.GetEndPoint(0)
                pt_b = c.GetEndPoint(1)
                any_pts = True
                min_x = min(min_x, float(pt_a.X), float(pt_b.X))
                min_y = min(min_y, float(pt_a.Y), float(pt_b.Y))
                max_x = max(max_x, float(pt_a.X), float(pt_b.X))
                max_y = max(max_y, float(pt_a.Y), float(pt_b.Y))
            except Exception:
                pass

    if not any_pts:
        return None
    return (min_x, max_x, min_y, max_y)


def _map_local_to_view(view, x_local, y_local):
    """Mapea coordenadas 2D locales al plano de la vista Drafting."""
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin
    return XYZ(
        o.X + right.X * x_local + up.X * y_local,
        o.Y + right.Y * x_local + up.Y * y_local,
        o.Z + right.Z * x_local + up.Z * y_local,
    )


def _draw_curve_safely(doc, view, curve, cell_center_x, cell_center_y, bb_cx, bb_cy, scale):
    """
    Dibuja curvas de forma segura evitando transformaciones no conformes que
    pueden inestabilizar Revit en algunos tipos de curva.
    """
    if curve is None:
        return

    def map_pt(pt):
        x = cell_center_x + (float(pt.X) - bb_cx) * scale
        y = cell_center_y + (float(pt.Y) - bb_cy) * scale
        return _map_local_to_view(view, x, y)

    # Línea
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        p0 = None
        p1 = None

    # Arco
    if isinstance(curve, Arc):
        try:
            pm = curve.Evaluate(0.5, True)  # parámetro normalizado
            a = map_pt(p0)
            b = map_pt(pm)
            c = map_pt(p1)
            arc = Arc.Create(a, c, b)
            doc.Create.NewDetailCurve(view, arc)
            return
        except Exception:
            # fallback a tessellation
            pass


def _draw_safe_shape_glyph(doc, view, shape, x_min, x_max, y_min, y_max):
    """
    Dibujo ultra-estable para evitar crashes:
    - SimpleLine: una línea horizontal.
    - SimpleArc: un arco semicircular.
    - Otros: zig-zag simple.
    """
    w = x_max - x_min
    h = y_max - y_min
    if w <= 1e-9 or h <= 1e-9:
        return

    pad = min(w, h) * 0.15
    lx0 = x_min + pad
    lx1 = x_max - pad
    ly0 = y_min + pad
    ly1 = y_max - pad
    cy = 0.5 * (ly0 + ly1)
    cx = 0.5 * (lx0 + lx1)

    is_simple_line = False
    is_simple_arc = False
    try:
        is_simple_line = bool(getattr(shape, "SimpleLine", False))
    except Exception:
        pass
    try:
        is_simple_arc = bool(getattr(shape, "SimpleArc", False))
    except Exception:
        pass

    # 1) Línea simple
    if is_simple_line:
        p0 = _pt_on_view_plane(view, lx0, cy)
        p1 = _pt_on_view_plane(view, lx1, cy)
        _draw_detail_line(doc, view, p0, p1)
        return

    # 2) Arco simple
    if is_simple_arc:
        try:
            left = _pt_on_view_plane(view, lx0, cy)
            right = _pt_on_view_plane(view, lx1, cy)
            top = _pt_on_view_plane(view, cx, ly1)
            arc = Arc.Create(left, right, top)
            doc.Create.NewDetailCurve(view, arc)
            return
        except Exception:
            pass

    # 3) Shape complejo: zig-zag estable
    p1 = _pt_on_view_plane(view, lx0, ly0)
    p2 = _pt_on_view_plane(view, x_min + w * 0.35, ly1)
    p3 = _pt_on_view_plane(view, x_min + w * 0.65, ly0)
    p4 = _pt_on_view_plane(view, lx1, ly1)
    _draw_detail_line(doc, view, p1, p2)
    _draw_detail_line(doc, view, p2, p3)
    _draw_detail_line(doc, view, p3, p4)

    # Línea simple
    if isinstance(curve, Line):
        try:
            doc.Create.NewDetailCurve(view, Line.CreateBound(map_pt(p0), map_pt(p1)))
            return
        except Exception:
            pass

    # Fallback universal: tessellar y dibujar segmentos lineales
    try:
        pts = list(curve.Tessellate() or [])
    except Exception:
        pts = []
    if len(pts) < 2 and p0 is not None and p1 is not None:
        pts = [p0, p1]

    for i in range(len(pts) - 1):
        try:
            a = map_pt(pts[i])
            b = map_pt(pts[i + 1])
            doc.Create.NewDetailCurve(view, Line.CreateBound(a, b))
        except Exception:
            pass


def _pt_on_view_plane(view, x_local, y_local):
    """
    Convierte (x_local,y_local) en coordenadas globales sobre el plano de la vista.
    """
    right = view.RightDirection
    up = view.UpDirection
    o = view.Origin
    return XYZ(
        o.X + right.X * x_local + up.X * y_local,
        o.Y + right.Y * x_local + up.Y * y_local,
        o.Z + right.Z * x_local + up.Z * y_local,
    )


def _draw_detail_line(doc, view, p1, p2):
    line = Line.CreateBound(p1, p2)
    if line is None:
        return None
    return doc.Create.NewDetailCurve(view, line)


def _run_internal(revit):
    doc = revit.ActiveUIDocument.Document
    uidoc = revit.ActiveUIDocument

    used_shapes = _collect_project_rebar_shapes(doc)
    if not used_shapes:
        TaskDialog.Show(VIEW_NAME, u"No se encontraron RebarShape en el proyecto.")
        return

    # Orden estable: por nombre y luego por Id.
    shapes = list(used_shapes.values())
    try:
        shapes.sort(key=lambda s: (s.Name or u"", int(s.Id.IntegerValue)))
    except Exception:
        pass

    cols = int(GRID_COLS)
    if cols < 1:
        cols = 1
    n = len(shapes)
    rows_calc = int(math.ceil(float(n) / float(cols))) if n else 1
    rows = max(GRID_ROWS_MIN, rows_calc)

    cell_w_ft = _mm_to_ft(CELL_W_MM)
    cell_h_ft = _mm_to_ft(CELL_H_MM)
    margin_ft = _mm_to_ft(MARGIN_MM)

    if cell_w_ft <= 0 or cell_h_ft <= 0:
        raise Exception(u"Config de celda inválida.")

    view = _create_or_get_view(doc, VIEW_NAME)

    # Dibujo en una sola transacción.
    t = Transaction(doc, u"Actualizar Leyenda de Rebar Shapes")
    t.Start()
    try:
        _clear_detail_curves_in_view(doc, view)

        # Grilla centrada en el origen local de la vista.
        total_w = cols * cell_w_ft
        total_h = rows * cell_h_ft
        grid_min_x = -0.5 * total_w
        grid_min_y = -0.5 * total_h

        # Rectángulos de celdas (siempre para consistencia)
        for idx in range(n):
            r = idx // cols
            c = idx % cols
            if r >= rows:
                break

            x_min = grid_min_x + c * cell_w_ft
            x_max = x_min + cell_w_ft
            y_min = grid_min_y + r * cell_h_ft
            y_max = y_min + cell_h_ft

            # 4 lados
            p1 = _pt_on_view_plane(view, x_min, y_min)
            p2 = _pt_on_view_plane(view, x_max, y_min)
            p3 = _pt_on_view_plane(view, x_max, y_max)
            p4 = _pt_on_view_plane(view, x_min, y_max)

            # Usar un mismo estilo gráfico: el tipo de línea por defecto de DetailCurve.
            _draw_detail_line(doc, view, p1, p2)
            _draw_detail_line(doc, view, p2, p3)
            _draw_detail_line(doc, view, p3, p4)
            _draw_detail_line(doc, view, p4, p1)

        # Dibujar cada forma dentro de su celda (modo ultra-estable).
        for idx, shape in enumerate(shapes):
            if not shape:
                continue
            r = idx // cols
            c = idx % cols
            if r >= rows:
                break

            x_min = grid_min_x + c * cell_w_ft
            x_max = x_min + cell_w_ft
            y_min = grid_min_y + r * cell_h_ft
            y_max = y_min + cell_h_ft

            cell_center_x = 0.5 * (x_min + x_max)
            cell_center_y = 0.5 * (y_min + y_max)

            # Mantener margen interno para no pegar al borde.
            gx_min = x_min + margin_ft
            gx_max = x_max - margin_ft
            gy_min = y_min + margin_ft
            gy_max = y_max - margin_ft
            _draw_safe_shape_glyph(doc, view, shape, gx_min, gx_max, gy_min, gy_max)

        t.Commit()
        return
    except Exception as ex:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show(VIEW_NAME, u"Error al crear leyenda:\n\n{}".format(str(ex)))


def run(revit, close_on_finish=False):
    # close_on_finish se ignora (no hay ventana WPF)
    _run_internal(revit)


