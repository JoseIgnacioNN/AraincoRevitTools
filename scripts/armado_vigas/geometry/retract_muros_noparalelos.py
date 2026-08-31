# -*- coding: utf-8 -*-
"""
Ajustes de extremos de línea fusionada (pre-troceo / post-fusión colineal).

Reglas de negocio:
  · Tras unificar la fibra de la corrida, se evalúan **GetEndPoint(0/1)**
    contra elementos de la **selección inicial** que **no** son vigas paralelas
    al plano de la vista activa.
  · Si un extremo colisiona con un **muro no paralelo** a la vista, se estira
    ese extremo **+(ancho/2 − 25 mm)** y se marca para **pata L**.
  · Si un extremo colisiona con una **viga** no // (selección o unida a la
    cadena), se estira **+(ancho/2 − 25 mm)** y se marca para **pata L**.
  · Si un extremo colisiona con una **columna** de hormigón, se toma el
    **ancho o alto** de sección según el eje de la fibra y se estira
    **+(dim/2 − 25 mm)** + **pata L**.
  · Si un extremo colisiona con un **muro paralelo** a la vista, se marca
    para **empotramiento** según Ø de barra.

No abre transacciones. Pata L / empotramiento se cierran en extremos / colocación.
"""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    Line,
    LocationCurve,
    Options,
    ViewDetailLevel,
    XYZ,
)

from geometria_colision_vigas import obtener_solidos_elemento, solidos_intersectan_por_booleana

try:
    from evaluacion_curva_puntos_obstaculos import _punto_en_volumen_solido
except Exception:
    _punto_en_volumen_solido = None

try:
    from geometria_empotramiento_extremos import _solid_esfera_en_centro
except Exception:
    _solid_esfera_en_centro = None

try:
    from armado_vigas.revit.view_order import (
        beam_axis_parallel_to_view_plane,
        beam_axis_tangent,
        view_normal_unit,
    )
except Exception:
    beam_axis_parallel_to_view_plane = None
    beam_axis_tangent = None
    view_normal_unit = None

_FRAMING_CAT = int(BuiltInCategory.OST_StructuralFraming)
_WALL_CAT = int(BuiltInCategory.OST_Walls)
_COL_CAT = int(BuiltInCategory.OST_StructuralColumns)

_PROBE_RADIUS_MM = 4.0
_TOL_VOLUMEN_FT3 = 1e-12
_MIN_LINE_LEN_FT = 5.0 / 304.8
_TOL_DOT_EJE_PARALELO = 0.0349  # sin(2°)
# Estiramiento en extremo vs muro/viga/columna: +(b/2 − clearance).
_END_CLEARANCE_MM = 25.0
_BEAM_END_CLEARANCE_MM = _END_CLEARANCE_MM  # alias compat
_WALL_END_CLEARANCE_MM = _END_CLEARANCE_MM
_COL_END_CLEARANCE_MM = _END_CLEARANCE_MM


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _geometry_options():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    return opts


def _eid_int(el_or_id):
    try:
        if hasattr(el_or_id, "IntegerValue"):
            return int(el_or_id.IntegerValue)
        if hasattr(el_or_id, "Id"):
            return int(el_or_id.Id.IntegerValue)
        return int(el_or_id)
    except Exception:
        return None


def _id_en_set(eid, id_set):
    if eid is None or not id_set:
        return False
    try:
        ei = int(eid)
    except Exception:
        return False
    for x in id_set:
        try:
            if int(x) == ei:
                return True
        except Exception:
            continue
    return False


def _resolve_view(document, view):
    if view is not None:
        return view
    if document is None:
        return None
    try:
        return document.ActiveView
    except Exception:
        return None


def _element_category_int(el):
    try:
        return int(el.Category.Id.IntegerValue)
    except Exception:
        return None


def _is_structural_framing(el):
    return _element_category_int(el) == _FRAMING_CAT


def _is_wall(el):
    return _element_category_int(el) == _WALL_CAT


def _is_column(el):
    return _element_category_int(el) == _COL_CAT


def _wall_axis_tangent(wall):
    """Tangente del muro (eje en planta)."""
    if wall is None:
        return None
    try:
        loc = wall.Location
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            crv = loc.Curve
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            d = p1 - p0
            if float(d.GetLength()) > 1e-9:
                return d.Normalize()
    except Exception:
        pass
    try:
        ori = wall.Orientation
        if ori is None or float(ori.GetLength()) < 1e-12:
            return None
        n = ori.Normalize()
        # Eje ≈ horizontal ⊥ normal del muro.
        axis = XYZ(-n.Y, n.X, 0.0)
        if float(axis.GetLength()) < 1e-9:
            return None
        return axis.Normalize()
    except Exception:
        return None


def _axis_parallel_to_view_plane(tang, view):
    if tang is None:
        return False
    n_view = view_normal_unit(view) if view_normal_unit is not None else None
    if n_view is None:
        return True
    try:
        return abs(float(tang.DotProduct(n_view))) <= _TOL_DOT_EJE_PARALELO
    except Exception:
        return False


def element_is_parallel_to_view_plane(el, view):
    """Viga/muro cuyo eje es // al plano de vista."""
    if el is None or view is None:
        return False
    if _is_structural_framing(el):
        if beam_axis_parallel_to_view_plane is not None:
            try:
                return bool(beam_axis_parallel_to_view_plane(el, view))
            except Exception:
                pass
        tang = beam_axis_tangent(el) if beam_axis_tangent is not None else None
        return _axis_parallel_to_view_plane(tang, view)
    if _is_wall(el):
        return _axis_parallel_to_view_plane(_wall_axis_tangent(el), view)
    return False


def wall_thickness_mm(wall):
    """Espesor del muro (mm) — ``Wall.Width`` o parámetro de tipo."""
    if wall is None:
        return None
    try:
        w_ft = float(wall.Width)
        if w_ft > 1e-9:
            return w_ft * 304.8
    except Exception:
        pass
    for bip in (BuiltInParameter.WALL_ATTR_WIDTH_PARAM,):
        try:
            p = wall.get_Parameter(bip)
            if p is not None and p.HasValue:
                return float(p.AsDouble()) * 304.8
        except Exception:
            continue
    try:
        t = wall.LookupParameter(u"Width") or wall.LookupParameter(u"Espesor")
        if t is not None and t.HasValue:
            return float(t.AsDouble()) * 304.8
    except Exception:
        pass
    return None


def beam_width_mm(document, beam):
    """Ancho de sección de viga (mm) — parámetros de tipo / bbox."""
    if beam is None:
        return None
    curve = None
    try:
        loc = beam.Location
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            curve = loc.Curve
    except Exception:
        curve = None
    if document is not None:
        try:
            from armado_vigas.revit.adapters import _read_width_depth_ft

            w_ft, _d_ft = _read_width_depth_ft(document, beam, curve)
            w_ft = float(w_ft or 0.0)
            if w_ft > 1e-9:
                return w_ft * 304.8
        except Exception:
            pass
    try:
        et = document.GetElement(beam.GetTypeId()) if document is not None else None
    except Exception:
        et = None
    if et is not None:
        for n in (u"Width", u"Ancho", u"Ancho nominal", u"b", u"B"):
            try:
                p = et.LookupParameter(n)
                if p is not None and p.HasValue:
                    w = float(p.AsDouble()) * 304.8
                    if w > 1.0:
                        return w
            except Exception:
                continue
    return None


def _column_type_width_depth_ft(document, column):
    """(Width, Depth) de sección de columna en pies, o ``(None, None)``."""
    if column is None:
        return None, None
    w_ft, d_ft = None, None
    if document is not None:
        try:
            from armado_vigas.revit.adapters import _read_width_depth_ft

            w_ft, d_ft = _read_width_depth_ft(document, column, None)
            w_ft = float(w_ft or 0.0) or None
            d_ft = float(d_ft or 0.0) or None
        except Exception:
            w_ft, d_ft = None, None
    if w_ft and d_ft:
        return w_ft, d_ft
    try:
        et = document.GetElement(column.GetTypeId()) if document is not None else None
    except Exception:
        et = None
    if et is not None:
        if not w_ft:
            for n in (u"Width", u"Ancho", u"Ancho nominal", u"b", u"B"):
                try:
                    p = et.LookupParameter(n)
                    if p is not None and p.HasValue:
                        w_ft = float(p.AsDouble())
                        break
                except Exception:
                    continue
        if not d_ft:
            for n in (u"Depth", u"Height", u"Profundidad", u"h", u"H", u"d"):
                try:
                    p = et.LookupParameter(n)
                    if p is not None and p.HasValue:
                        d_ft = float(p.AsDouble())
                        break
                except Exception:
                    continue
    return w_ft, d_ft


def _horiz_unit(v):
    """Proyección horizontal unitaria de un vector, o None."""
    if v is None:
        return None
    try:
        h = XYZ(float(v.X), float(v.Y), 0.0)
        if float(h.GetLength()) < 1e-12:
            return None
        return h.Normalize()
    except Exception:
        return None


def _column_section_axes(column):
    """Ejes locales de sección (BasisX / BasisY o Hand/Facing), horizontales."""
    if column is None:
        return None, None
    bx = by = None
    if isinstance(column, FamilyInstance):
        try:
            tr = column.GetTransform()
            if tr is not None:
                bx = _horiz_unit(tr.BasisX)
                by = _horiz_unit(tr.BasisY)
        except Exception:
            bx = by = None
        if bx is None:
            try:
                bx = _horiz_unit(column.HandOrientation)
            except Exception:
                bx = None
        if by is None:
            try:
                by = _horiz_unit(column.FacingOrientation)
            except Exception:
                by = None
    return bx, by


def _bbox_extent_along_axis_mm(el, axis):
    """Extensión del bbox del elemento proyectada sobre ``axis`` (mm)."""
    if el is None or axis is None:
        return None
    try:
        u = axis.Normalize()
    except Exception:
        return None
    try:
        bb = el.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return None
    try:
        corners = (
            XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
            XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
            XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
            XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
            XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
            XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
            XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
            XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
        )
        vals = [float(c.DotProduct(u)) for c in corners]
        span_ft = max(vals) - min(vals)
        if span_ft > 1e-9:
            return span_ft * 304.8
    except Exception:
        pass
    return None


def column_section_dim_along_axis_mm(document, column, axis):
    """
    Ancho o alto de sección de columna (mm) **según el eje de la fibra**.

    Elige Width vs Depth según qué eje local de sección alinea mejor con
    ``axis`` (tangente de la línea fusionada). Fallback: proyección del bbox.
    """
    if column is None:
        return None
    ax = _horiz_unit(axis)
    if ax is None:
        try:
            ax = axis.Normalize() if axis is not None else None
        except Exception:
            ax = None
    w_ft, d_ft = _column_type_width_depth_ft(document, column)
    bx, by = _column_section_axes(column)
    if ax is not None and w_ft and d_ft and bx is not None and by is not None:
        try:
            if abs(float(ax.DotProduct(bx))) >= abs(float(ax.DotProduct(by))):
                dim = float(w_ft) * 304.8
            else:
                dim = float(d_ft) * 304.8
            if dim >= 1.0:
                return dim
        except Exception:
            pass
    # Una sola dimensión de tipo conocida.
    if w_ft and (not d_ft or abs(float(w_ft) - float(d_ft or 0.0)) < 1e-9):
        dim = float(w_ft) * 304.8
        if dim >= 1.0:
            return dim
    if ax is not None:
        span = _bbox_extent_along_axis_mm(column, ax)
        if span is not None and span >= 1.0:
            return span
    if w_ft:
        return float(w_ft) * 304.8
    if d_ft:
        return float(d_ft) * 304.8
    return _bbox_extent_along_axis_mm(column, ax or XYZ.BasisX)


def _point_collides_element(pt, element, opts=None):
    """True si el punto (esfera de sonda) intersecta un sólido del elemento."""
    if pt is None or element is None:
        return False
    opts = opts or _geometry_options()
    r_ft = _mm_to_ft(_PROBE_RADIUS_MM)
    sphere = None
    if _solid_esfera_en_centro is not None:
        try:
            sphere = _solid_esfera_en_centro(pt, r_ft)
        except Exception:
            sphere = None
    try:
        solids = list(obtener_solidos_elemento(element, options=opts) or [])
    except Exception:
        solids = []
    for s in solids:
        if s is None:
            continue
        if sphere is not None:
            try:
                if solidos_intersectan_por_booleana(
                    sphere, s, tol_volumen=_TOL_VOLUMEN_FT3
                ):
                    return True
            except Exception:
                pass
        if _punto_en_volumen_solido is not None:
            try:
                if _punto_en_volumen_solido(s, pt):
                    return True
            except Exception:
                pass
    return False


def _selection_collision_candidates(document, ids_seleccion, excluir_ids, view):
    """
    Elementos de la selección a evaluar en extremos:

    · Excluye host de la cadena.
    · Excluye **vigas paralelas** al plano de la vista activa.
    · Incluye muros, columnas y framing no // a la vista.
    """
    out = []
    if document is None:
        return out
    for raw in ids_seleccion or []:
        eid = _eid_int(raw)
        if eid is None or _id_en_set(eid, excluir_ids):
            continue
        try:
            el = document.GetElement(ElementId(eid))
        except Exception:
            el = None
        if el is None:
            continue
        try:
            if not el.IsValidObject:
                continue
        except Exception:
            pass
        cat = _element_category_int(el)
        if cat == _FRAMING_CAT:
            if view is not None and element_is_parallel_to_view_plane(el, view):
                continue
            out.append(el)
            continue
        if cat in (_WALL_CAT, _COL_CAT):
            out.append(el)
            continue
        # Otros: solo si no son framing // (ya filtrado).
        out.append(el)
    return out


def _joined_npar_framing_candidates(document, host_chain_elements, view, excluir_ids):
    """
    Vigas hormigón unidas a la cadena, no // a la vista (tip. transversales).

    El filtro de selección solo admite framing //; estas unidas no suelen estar
    en ``ids_seleccion`` y son las que disparan el estirón de extremo.
    """
    out = []
    if document is None:
        return out
    hosts = []
    for h in host_chain_elements or []:
        if h is None:
            continue
        # Puede venir Element o id.
        if _is_structural_framing(h):
            hosts.append(h)
            continue
        eid = _eid_int(h)
        if eid is None:
            continue
        try:
            el = document.GetElement(ElementId(eid))
        except Exception:
            el = None
        if el is not None and _is_structural_framing(el):
            hosts.append(el)
    if not hosts:
        return out
    try:
        from armado_vigas.revit.joined_framing import detect_joined_concrete_framing

        result = detect_joined_concrete_framing(document, hosts, view)
    except Exception:
        return out
    seen = set()
    for rec in (result or {}).get("not_parallel") or []:
        if bool(rec.get("parallelToView")):
            continue
        el = rec.get("element")
        eid = _eid_int(el) if el is not None else _eid_int(rec.get("elementIdInt"))
        if eid is None or _id_en_set(eid, excluir_ids) or eid in seen:
            continue
        if el is None:
            try:
                el = document.GetElement(ElementId(eid))
            except Exception:
                el = None
        if el is None or not _is_structural_framing(el):
            continue
        seen.add(eid)
        out.append(el)
    return out


def _framing_stretch_candidates(
    document, ids_seleccion, host_chain_elements, view, excluir_ids
):
    """Candidatos de viga para estirón: selección no// + unidas no// a la cadena."""
    out = []
    seen = set()
    for el in _selection_collision_candidates(
        document, ids_seleccion, excluir_ids, view
    ):
        if not _is_structural_framing(el):
            continue
        eid = _eid_int(el)
        if eid is None or eid in seen:
            continue
        seen.add(eid)
        out.append(el)
    for el in _joined_npar_framing_candidates(
        document, host_chain_elements, view, excluir_ids
    ):
        eid = _eid_int(el)
        if eid is None or eid in seen:
            continue
        seen.add(eid)
        out.append(el)
    return out


def _first_colliding_wall_noparallel(pt, candidates, view):
    """Primer muro no // a la vista que colisiona con el punto."""
    for el in candidates or []:
        if not _is_wall(el):
            continue
        if view is not None and element_is_parallel_to_view_plane(el, view):
            continue
        if _point_collides_element(pt, el):
            return el
    return None


def _first_colliding_wall_parallel(pt, candidates, view):
    """Primer muro // al plano de la vista que colisiona con el punto."""
    for el in candidates or []:
        if not _is_wall(el):
            continue
        if view is not None and not element_is_parallel_to_view_plane(el, view):
            continue
        if _point_collides_element(pt, el):
            return el
    return None


def _element_ref_point(el):
    """Punto de referencia (centro) del elemento para proximidad a extremos."""
    if el is None:
        return None
    try:
        loc = el.Location
        if loc is not None and hasattr(loc, "Point") and loc.Point is not None:
            return loc.Point
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            return loc.Curve.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        bb = el.get_BoundingBox(None)
        if bb is not None:
            return XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                (bb.Min.Z + bb.Max.Z) * 0.5,
            )
    except Exception:
        pass
    return None


def _wall_belongs_to_endpoint(wall, p0, p1, end_index):
    """
    True si el muro pertenece a este extremo (más cerca de él que del opuesto).

    Evita retractar ambos extremos por el mismo muro cuando solo uno colisiona
    de forma significativa (o ambos disparan sonda por geometría compartida).
    """
    if wall is None or p0 is None or p1 is None:
        return False
    ref = _element_ref_point(wall)
    if ref is None:
        # Sin ref: confiar solo en colisión del extremo (ya filtrada).
        return True
    try:
        d0 = float(ref.DistanceTo(p0))
        d1 = float(ref.DistanceTo(p1))
    except Exception:
        return True
    if end_index == 0:
        return d0 <= d1 + 1e-9
    return d1 <= d0 + 1e-9


def _first_colliding_framing_noparallel(pt, candidates, view):
    """Primera viga (framing) no // a la vista que colisiona con el punto."""
    for el in candidates or []:
        if not _is_structural_framing(el):
            continue
        if view is not None and element_is_parallel_to_view_plane(el, view):
            continue
        if _point_collides_element(pt, el):
            return el
    return None


def _first_colliding_column(pt, candidates):
    """Primera columna estructural que colisiona con el punto."""
    for el in candidates or []:
        if not _is_column(el):
            continue
        if _point_collides_element(pt, el):
            return el
    return None


def _retract_endpoint(p0, p1, end_index, retract_ft):
    """Mueve el extremo ``end_index`` (0|1) hacia el interior de la línea."""
    try:
        v = p1 - p0
        L = float(v.GetLength())
    except Exception:
        return p0, p1
    if L < _MIN_LINE_LEN_FT or retract_ft <= 1e-12:
        return p0, p1
    max_retract = max(0.0, L * 0.45)
    d = min(float(retract_ft), max_retract)
    if d <= 1e-12:
        return p0, p1
    try:
        u = v.Normalize()
    except Exception:
        return p0, p1
    if end_index == 0:
        p0 = p0 + u.Multiply(d)
    else:
        p1 = p1 - u.Multiply(d)
    return p0, p1


def _extend_endpoint(p0, p1, end_index, extend_ft):
    """Mueve el extremo ``end_index`` (0|1) hacia el exterior de la línea."""
    try:
        v = p1 - p0
        L = float(v.GetLength())
    except Exception:
        return p0, p1
    if L < _MIN_LINE_LEN_FT or extend_ft <= 1e-12:
        return p0, p1
    try:
        u = v.Normalize()
    except Exception:
        return p0, p1
    d = float(extend_ft)
    if end_index == 0:
        p0 = p0 - u.Multiply(d)
    else:
        p1 = p1 + u.Multiply(d)
    return p0, p1


def _rebuild_line(p0, p1, fallback_line):
    try:
        L = float((p1 - p0).GetLength())
        if L < _MIN_LINE_LEN_FT:
            return fallback_line
        new_line = Line.CreateBound(p0, p1)
    except Exception:
        return fallback_line
    return new_line if new_line is not None else fallback_line


def aplicar_retracto_extremos_muros_noparalelos(
    document,
    line,
    ids_seleccion,
    host_chain_elements=None,
    view=None,
):
    """
    Pre-troceo, post-fusión: estira **solo** el extremo que colisiona con muro no // vista.

    Estiramiento = ``ancho/2 − 25 mm`` (solo si es > 0). Un mismo muro no estira
    ambos extremos (solo el más cercano). La pata L se marca en longitudinales.

    Nota: el nombre histórico habla de «retracto»; el comportamiento actual es
    estirón positivo (misma regla que vigas transversales).

    Returns:
        ``(line_out, meta)`` — meta por extremo si se estira
        (``stretch_mm``, ``width_mm`` / ``thickness_mm``).
    """
    meta = {
        u"start": None,
        u"end": None,
        u"applied": False,
    }
    if document is None or line is None:
        return line, meta
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
    except Exception:
        return line, meta

    view = _resolve_view(document, view)
    excluir = set()
    for h in host_chain_elements or []:
        eid = _eid_int(h)
        if eid is not None:
            excluir.add(eid)

    candidates = _selection_collision_candidates(
        document, ids_seleccion, excluir, view
    )
    if not candidates:
        return line, meta

    # Evaluar colisiones sobre extremos originales (antes de mover).
    p0_orig, p1_orig = p0, p1
    changed = False
    for end_idx, key in ((0, u"start"), (1, u"end")):
        pt = p0_orig if end_idx == 0 else p1_orig
        wall = _first_colliding_wall_noparallel(pt, candidates, view)
        if wall is None:
            continue
        if not _wall_belongs_to_endpoint(wall, p0_orig, p1_orig, end_idx):
            continue
        th_mm = wall_thickness_mm(wall)
        if th_mm is None or th_mm < 1.0:
            continue
        stretch_mm = 0.5 * float(th_mm) - float(_WALL_END_CLEARANCE_MM)
        if stretch_mm <= 1e-6:
            continue
        p0, p1 = _extend_endpoint(p0, p1, end_idx, _mm_to_ft(stretch_mm))
        try:
            eid = int(wall.Id.IntegerValue)
        except Exception:
            eid = None
        meta[key] = {
            u"kind": u"muro_noparalelo",
            u"wall_id": eid,
            u"width_mm": float(th_mm),
            u"thickness_mm": float(th_mm),
            u"stretch_mm": float(stretch_mm),
            u"clearance_mm": float(_WALL_END_CLEARANCE_MM),
        }
        meta[u"applied"] = True
        changed = True

    if not changed:
        return line, meta

    return _rebuild_line(p0, p1, line), meta


def detectar_extremos_muros_paralelos(
    document,
    line,
    ids_seleccion,
    host_chain_elements=None,
    view=None,
    skip_ends=None,
):
    """
    Detecta extremos que colisionan con muro **paralelo** al plano de la vista.

    No modifica la línea. El estirón de empotramiento (según Ø) se aplica después
    al conocer el diámetro de capa.

    ``skip_ends``: iterable ``('start',)`` / ``('end',)`` ya reservados (p. ej.
    estirón muro/viga no //).

    Returns:
        meta ``{start, end, applied}`` con ``kind: muro_paralelo``.
    """
    meta = {
        u"start": None,
        u"end": None,
        u"applied": False,
    }
    if document is None or line is None:
        return meta
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
    except Exception:
        return meta

    skip = set()
    for s in skip_ends or []:
        try:
            skip.add(unicode(s))
        except Exception:
            skip.add(str(s))

    view = _resolve_view(document, view)
    excluir = set()
    for h in host_chain_elements or []:
        eid = _eid_int(h)
        if eid is not None:
            excluir.add(eid)

    candidates = _selection_collision_candidates(
        document, ids_seleccion, excluir, view
    )
    if not candidates:
        return meta

    for end_idx, key in ((0, u"start"), (1, u"end")):
        if key in skip:
            continue
        pt = p0 if end_idx == 0 else p1
        wall = _first_colliding_wall_parallel(pt, candidates, view)
        if wall is None:
            continue
        if not _wall_belongs_to_endpoint(wall, p0, p1, end_idx):
            continue
        try:
            eid = int(wall.Id.IntegerValue)
        except Exception:
            eid = None
        meta[key] = {
            u"kind": u"muro_paralelo",
            u"wall_id": eid,
        }
        meta[u"applied"] = True

    return meta


def aplicar_estiramiento_extremos_columnas(
    document,
    line,
    ids_seleccion,
    host_chain_elements=None,
    view=None,
):
    """
    Pre-troceo, post-fusión: estira start/end si colisionan con **columna**.

    Dimensión = ancho o alto de sección según el eje de la fibra.
    Estiramiento = ``dim/2 − 25 mm`` (solo si es > 0). Una misma columna no
    estira ambos extremos (solo el más cercano). Pata L en longitudinales.

    Returns:
        ``(line_out, meta)`` — meta por extremo si se estira
        (``stretch_mm``, ``width_mm`` / ``section_mm``).
    """
    meta = {
        u"start": None,
        u"end": None,
        u"applied": False,
    }
    if document is None or line is None:
        return line, meta
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
    except Exception:
        return line, meta

    try:
        tang = (p1 - p0).Normalize()
    except Exception:
        tang = None

    view = _resolve_view(document, view)
    excluir = set()
    for h in host_chain_elements or []:
        eid = _eid_int(h)
        if eid is not None:
            excluir.add(eid)

    candidates = _selection_collision_candidates(
        document, ids_seleccion, excluir, view
    )
    if not candidates:
        return line, meta

    p0_orig, p1_orig = p0, p1
    changed = False
    for end_idx, key in ((0, u"start"), (1, u"end")):
        pt = p0_orig if end_idx == 0 else p1_orig
        col = _first_colliding_column(pt, candidates)
        if col is None:
            continue
        if not _wall_belongs_to_endpoint(col, p0_orig, p1_orig, end_idx):
            continue
        dim_mm = column_section_dim_along_axis_mm(document, col, tang)
        if dim_mm is None or dim_mm < 1.0:
            continue
        stretch_mm = 0.5 * float(dim_mm) - float(_COL_END_CLEARANCE_MM)
        if stretch_mm <= 1e-6:
            continue
        p0, p1 = _extend_endpoint(p0, p1, end_idx, _mm_to_ft(stretch_mm))
        try:
            eid = int(col.Id.IntegerValue)
        except Exception:
            eid = None
        meta[key] = {
            u"kind": u"columna",
            u"column_id": eid,
            u"width_mm": float(dim_mm),
            u"section_mm": float(dim_mm),
            u"stretch_mm": float(stretch_mm),
            u"clearance_mm": float(_COL_END_CLEARANCE_MM),
        }
        meta[u"applied"] = True
        changed = True

    if not changed:
        return line, meta

    return _rebuild_line(p0, p1, line), meta


# Alias semántico (estirón, no retracto).
aplicar_estiramiento_extremos_muros_noparalelos = (
    aplicar_retracto_extremos_muros_noparalelos
)


def aplicar_estiramiento_extremos_vigas_noparalelas(
    document,
    line,
    ids_seleccion,
    host_chain_elements=None,
    view=None,
):
    """
    Pre-troceo, post-fusión: estira start/end si colisionan con viga no // vista.

    Estiramiento = ``ancho/2 − 25 mm`` (solo si es > 0). Considera framing no
    paralelo de la selección y vigas unidas no // a la cadena (el filtro de
    selección suele excluir transversales). No considera las de
    ``host_chain_elements``.

    Returns:
        ``(line_out, meta)`` — meta por extremo si se estira.
    """
    meta = {
        u"start": None,
        u"end": None,
        u"applied": False,
    }
    if document is None or line is None:
        return line, meta
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
    except Exception:
        return line, meta

    view = _resolve_view(document, view)
    excluir = set()
    for h in host_chain_elements or []:
        eid = _eid_int(h)
        if eid is not None:
            excluir.add(eid)

    candidates = _framing_stretch_candidates(
        document, ids_seleccion, host_chain_elements, view, excluir
    )
    if not candidates:
        return line, meta

    changed = False
    for end_idx, key in ((0, u"start"), (1, u"end")):
        pt = p0 if end_idx == 0 else p1
        beam = _first_colliding_framing_noparallel(pt, candidates, view)
        if beam is None:
            continue
        w_mm = beam_width_mm(document, beam)
        if w_mm is None or w_mm < 1.0:
            continue
        stretch_mm = 0.5 * float(w_mm) - float(_BEAM_END_CLEARANCE_MM)
        if stretch_mm <= 1e-6:
            continue
        p0, p1 = _extend_endpoint(p0, p1, end_idx, _mm_to_ft(stretch_mm))
        try:
            eid = int(beam.Id.IntegerValue)
        except Exception:
            eid = None
        meta[key] = {
            u"kind": u"viga_noparalela",
            u"beam_id": eid,
            u"width_mm": float(w_mm),
            u"stretch_mm": float(stretch_mm),
            u"clearance_mm": float(_BEAM_END_CLEARANCE_MM),
        }
        meta[u"applied"] = True
        changed = True

    if not changed:
        return line, meta

    return _rebuild_line(p0, p1, line), meta