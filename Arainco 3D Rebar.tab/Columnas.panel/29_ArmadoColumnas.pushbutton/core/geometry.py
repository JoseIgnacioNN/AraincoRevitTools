# -*- coding: utf-8 -*-
"""
Extracción de dimensiones y ejes locales de columnas estructurales.

REGLA DE CAPA:
- No abre transacciones ni escribe en el modelo.
- No importa módulos de ui/ ni creators/.
- Devuelve tipos Python puros y tipos de Revit DB (XYZ, etc.); nunca WPF.

Toda la lógica está portada literalmente desde column_reinforcement_layout_rps.py
para que el .pushbutton sea 100% self-contained.
"""
from __future__ import print_function

import math

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    GeometryInstance,
    JoinGeometryUtils,
    Options,
    Solid,
    StorageType,
    UnitTypeId,
    UnitUtils,
    ViewDetailLevel,
    XYZ,
)


# ---------------------------------------------------------------------------
# Constantes de configuración (idénticas al monolito)
# ---------------------------------------------------------------------------

LAYOUT_BAR_NOMINAL_DIAM_MM = 12.0
LAYOUT_EMBED_CONCRETE_GRADE = None
COLUMN_REBAR_L_SHAPE_DISPLAY_NAME = u"02"
COLUMN_ARMA_UBICACION_PARAM = u"Armadura_Ubicacion"

_FOUNDATION_STRETCH_DEDUCTION_MM = 50.0
_FOUNDATION_JOIN_FACE_Z_TOLERANCE_MM = 35.0
_FOUNDATION_JOIN_OVERLAP_XY_MM = 75.0
_CAT_STRUCT_FOUNDATION_IV = int(BuiltInCategory.OST_StructuralFoundation)

_EMBED_PROBE_XY_MARGIN_MM = 1.0
_EMBED_PROBE_MIN_HALF_SIDE_MM = 2.0
_TOL_VOL_INTERSECCION_EMBED_FT3 = 5e-8
_REVOKE_EMBED_EXTRA_SHRINK_MM = 25.0

XY_KEY_DECIMALS_DEFAULT = 9

_NUMERACION_COLUMNA_PARAM_CANDIDATES = (
    u"Numeracion Columna",
    u"Numeración Columna",
)

# ---------------------------------------------------------------------------
# Helpers de ElementId
# ---------------------------------------------------------------------------

def _element_id_iv(elem):
    u"""
    Id numérico estable para ElementId.
    Usa Id.Value si != 0 (Revit 2024+); fallback a IntegerValue.
    """
    if elem is None:
        return -1
    try:
        rid = elem.Id
    except Exception:
        return -1
    try:
        v = getattr(rid, "Value", None)
        if v is not None:
            iv = int(v)
            if iv != 0:
                return iv
    except Exception:
        pass
    try:
        return int(rid.IntegerValue)
    except Exception:
        return -1


def _element_id_from_int(iv):
    """ElementId desde entero de API (Value / IntegerValue)."""
    from Autodesk.Revit.DB import ElementId
    import System
    if iv is None:
        return None
    try:
        return ElementId(int(iv))
    except Exception:
        try:
            return ElementId(System.Convert.ToInt64(int(iv)))
        except Exception:
            return None


# ---------------------------------------------------------------------------
# Clave de sección canónica
# ---------------------------------------------------------------------------

def _canonical_section_mm_key(width_ft, depth_ft):
    """
    Tupla estable ``(lado_corto_mm, lado_largo_mm)`` independiente del eje proyecto.
    """
    try:
        wx = UnitUtils.ConvertFromInternalUnits(abs(float(width_ft)), UnitTypeId.Millimeters)
        dx = UnitUtils.ConvertFromInternalUnits(abs(float(depth_ft)), UnitTypeId.Millimeters)
    except Exception:
        wx = abs(float(width_ft)) * 304.8
        dx = abs(float(depth_ft)) * 304.8
    s = float(min(wx, dx))
    L = float(max(wx, dx))
    return int(round(s)), int(round(L))


# ---------------------------------------------------------------------------
# Opciones de geometría
# ---------------------------------------------------------------------------

def _geometry_options_structure_solids():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    return opts


# ---------------------------------------------------------------------------
# Iteración de sólidos y vértices
# ---------------------------------------------------------------------------

def _iter_solids_revit_element(elem, opts):
    if elem is None:
        return
    try:
        ge = elem.get_Geometry(opts)
    except Exception:
        return
    if ge is None:
        return
    for obj in ge:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if float(obj.Volume) < 1e-11:
                    continue
            except Exception:
                continue
            yield obj
        elif isinstance(obj, GeometryInstance):
            try:
                sub = obj.GetInstanceGeometry()
            except Exception:
                continue
            if sub is None:
                continue
            for g2 in sub:
                if isinstance(g2, Solid):
                    try:
                        if float(g2.Volume) < 1e-11:
                            continue
                    except Exception:
                        continue
                    yield g2


def _iter_vertices_from_solid(solid):
    if solid is None:
        return
    try:
        if float(solid.Volume) < 1e-11:
            return
    except Exception:
        return
    try:
        edges = solid.Edges
    except Exception:
        return
    seen = set()
    for edge in edges:
        if edge is None:
            continue
        for i in (0, 1):
            try:
                pt = edge.AsCurve().GetEndPoint(i)
                key = (round(pt.X, 6), round(pt.Y, 6), round(pt.Z, 6))
                if key not in seen:
                    seen.add(key)
                    yield pt
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Rangos de vértices
# ---------------------------------------------------------------------------

def _solid_aggregate_vertex_ranges_ft(elem, opts=None):
    if elem is None:
        return None
    if opts is None:
        opts = _geometry_options_structure_solids()
    min_x = min_y = min_z = float("inf")
    max_x = max_y = max_z = float("-inf")
    found = False
    for solid in _iter_solids_revit_element(elem, opts):
        for pt in _iter_vertices_from_solid(solid):
            found = True
            x, y, z = float(pt.X), float(pt.Y), float(pt.Z)
            if x < min_x: min_x = x
            if x > max_x: max_x = x
            if y < min_y: min_y = y
            if y > max_y: max_y = y
            if z < min_z: min_z = z
            if z > max_z: max_z = z
    if not found:
        return None
    return (min_x, max_x, min_y, max_y, min_z, max_z)


# ---------------------------------------------------------------------------
# Parámetros de sección (b / h)
# ---------------------------------------------------------------------------

_PARAM_WIDTH_NAMES  = (u"b", u"Width",  u"Ancho", u"Base")
_PARAM_DEPTH_NAMES  = (u"h", u"Height", u"Profundidad", u"Depth", u"Altura")


def _lookup_column_section_width_depth_ft(column):
    """Devuelve (width_ft, depth_ft) desde parámetros de tipo/instancia; None si no existe."""
    def _try_param(elem, names):
        for pname in names:
            try:
                p = elem.LookupParameter(pname)
                if p and p.HasValue:
                    return p.AsDouble()
            except Exception:
                pass
        return None

    w = _try_param(column, _PARAM_WIDTH_NAMES)
    d = _try_param(column, _PARAM_DEPTH_NAMES)
    if w is None or d is None:
        try:
            sym = column.Symbol
            if sym is not None:
                if w is None:
                    w = _try_param(sym, _PARAM_WIDTH_NAMES)
                if d is None:
                    d = _try_param(sym, _PARAM_DEPTH_NAMES)
        except Exception:
            pass
    return w, d


# ---------------------------------------------------------------------------
# Curva de ubicación y eje
# ---------------------------------------------------------------------------

def _column_curve_endpoints_axis_ft(column, tol_curve):
    try:
        loc = column.Location
        cr = getattr(loc, "Curve", None)
        if cr is not None:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            ln = float(v.GetLength())
            if ln >= tol_curve:
                return p0, p1, v.Normalize(), ln
    except Exception:
        pass
    return None, None, None, 0.0


# ---------------------------------------------------------------------------
# Transform local del símbolo
# ---------------------------------------------------------------------------

def _plan_axes_from_family_transform_ft(column, width_ft, depth_ft):
    """
    Ejes en planta del símbolo (v_short, v_long). Devuelve (None, None) si no aplica.
    """
    try:
        trans = column.GetTransform()
        bx = trans.BasisX.Normalize()
        by = trans.BasisY.Normalize()
        try:
            bz = trans.BasisZ.Normalize()
        except Exception:
            bz = XYZ.BasisZ
        bx_is_up = abs(float(bx.Z)) > 0.9
        by_is_up = abs(float(by.Z)) > 0.9
        if bx_is_up or by_is_up:
            return None, None
        short_on_x = width_ft <= depth_ft
        if short_on_x:
            return bx, by
        else:
            return by, bx
    except Exception:
        return None, None


# ---------------------------------------------------------------------------
# Dimensión orientada desde vértices
# ---------------------------------------------------------------------------

def _oriented_section_width_depth_ft(origin, axis_u, verts):
    """
    Proyecta los vértices del sólido en el plano perp. al eje de la columna
    y mide extensión en los dos ejes locales.
    """
    try:
        ax = axis_u.Normalize()
    except Exception:
        return 0.0, 0.0
    try:
        perp = XYZ.BasisX
        if abs(float(ax.DotProduct(XYZ.BasisX))) > 0.9:
            perp = XYZ.BasisY
        perp_u = ax.CrossProduct(perp).Normalize()
        perp_v = ax.CrossProduct(perp_u).Normalize()
    except Exception:
        return 0.0, 0.0
    us_list, vs_list = [], []
    for pt in verts:
        try:
            dv = pt - origin
            us_list.append(float(dv.DotProduct(perp_u)))
            vs_list.append(float(dv.DotProduct(perp_v)))
        except Exception:
            pass
    if not us_list:
        return 0.0, 0.0
    w = max(us_list) - min(us_list)
    d = max(vs_list) - min(vs_list)
    return float(w), float(d)


# ---------------------------------------------------------------------------
# Mallas trianguladas (preferencia máxima)
# ---------------------------------------------------------------------------

def _try_column_dimensions_transform_mesh_ft(column, tol_curve):
    """
    Intenta medir la sección usando GetTransform() + mallas trianguladas
    de caras del sólido. Devuelve 6-tuple o None si no aplica.
    """
    try:
        trans = column.GetTotalTransform()
    except Exception:
        try:
            trans = column.GetTransform()
        except Exception:
            return None
    try:
        bx = trans.BasisX.Normalize()
        by = trans.BasisY.Normalize()
        bz = trans.BasisZ.Normalize()
        if abs(float(bz.Z)) < 0.5:
            return None
    except Exception:
        return None
    opts = _geometry_options_structure_solids()
    us_list, vs_list, ws_list = [], [], []
    origin = trans.Origin
    for solid in _iter_solids_revit_element(column, opts):
        for pt in _iter_vertices_from_solid(solid):
            dv = pt - origin
            us_list.append(float(dv.DotProduct(bx)))
            vs_list.append(float(dv.DotProduct(by)))
            ws_list.append(float(dv.DotProduct(bz)))
    if not us_list:
        return None
    u_ext = max(us_list) - min(us_list)
    v_ext = max(vs_list) - min(vs_list)
    w_ext = max(ws_list) - min(ws_list)
    if u_ext < tol_curve or v_ext < tol_curve or w_ext < tol_curve:
        return None
    width  = float(u_ext)
    depth  = float(v_ext)
    height = float(w_ext)
    min_z = float(origin.Z) + min(ws_list)
    cx = float(origin.X) + (max(us_list) + min(us_list)) / 2.0
    cy = float(origin.Y) + (max(vs_list) + min(vs_list)) / 2.0
    center = XYZ(cx, cy, min_z)
    short_on_x = width <= depth
    vs = bx if short_on_x else by
    vl = by if short_on_x else bx
    return width, depth, height, center, vs, vl


# ---------------------------------------------------------------------------
# API pública principal
# ---------------------------------------------------------------------------

def get_column_dimensions(column):
    """
    Dimensiones y centro base sin ``Element.GetBoundingBox``.

    Devuelve ``(width_ft, depth_ft, height_ft, center_xyz, v_short, v_long)``.

    Preferencia:
    1. Transform local + mallas trianguladas (``_try_column_dimensions_transform_mesh_ft``).
    2. Parámetros b/h + sólidos + curva de ubicación como respaldo.
    """
    doc = getattr(column, "Document", None)
    tol_curve = 1e-9
    try:
        if doc is not None:
            tol_curve = float(doc.Application.ShortCurveTolerance)
    except Exception:
        pass
    tol_curve = max(tol_curve, 1e-12)

    mesh_dims = _try_column_dimensions_transform_mesh_ft(column, tol_curve)
    if mesh_dims is not None:
        return mesh_dims

    p0, p1, axis_u, curve_len = _column_curve_endpoints_axis_ft(column, tol_curve)
    opts = _geometry_options_structure_solids()
    rng = _solid_aggregate_vertex_ranges_ft(column, opts)
    if rng is None:
        raise Exception(
            u"No se pudo obtener geometría sólida de la columna "
            u"(¿sin sólidos o elemento mal configurado?)."
        )
    min_x, max_x, min_y, max_y, min_z, max_z = rng
    w_par, d_par = _lookup_column_section_width_depth_ft(column)

    verts = []
    for solid in _iter_solids_revit_element(column, opts):
        for pt in _iter_vertices_from_solid(solid):
            verts.append(pt)

    origin_for_section = p0
    if origin_for_section is None:
        origin_for_section = XYZ(
            0.5 * (float(min_x) + float(max_x)),
            0.5 * (float(min_y) + float(max_y)),
            0.5 * (float(min_z) + float(max_z)),
        )

    if axis_u is None:
        axis_u = XYZ.BasisZ
    try:
        axis_u = axis_u.Normalize()
    except Exception:
        axis_u = XYZ.BasisZ

    if (
        w_par is not None and d_par is not None
        and float(w_par) > 1e-12 and float(d_par) > 1e-12
    ):
        width, depth = float(w_par), float(d_par)
    else:
        if not verts:
            raise Exception(
                u"No se pudieron medir lados de sección: sin vértices en sólidos "
                u"y sin parámetros de ancho/profundidad."
            )
        width, depth = _oriented_section_width_depth_ft(origin_for_section, axis_u, verts)
        if width <= 1e-12 or depth <= 1e-12:
            raise Exception(u"No se pudo determinar la sección en planta desde la geometría sólida.")

    if curve_len > tol_curve:
        height = float(curve_len)
    else:
        height = None
        if verts:
            dots = []
            for p in verts:
                try:
                    dots.append(float((p - origin_for_section).DotProduct(axis_u)))
                except Exception:
                    continue
            if dots:
                cand = float(max(dots) - min(dots))
                if cand > tol_curve:
                    height = cand
        if height is None or height <= tol_curve:
            height = abs(float(max_z) - float(min_z))
        if height <= tol_curve:
            raise Exception(u"No se pudo determinar la altura del pilar.")

    if p0 is not None and p1 is not None:
        mid = p0 + 0.5 * (p1 - p0)
        cx, cy = float(mid.X), float(mid.Y)
    else:
        cx = 0.5 * (float(min_x) + float(max_x))
        cy = 0.5 * (float(min_y) + float(max_y))

    center = XYZ(cx, cy, float(min_z))
    vs, vl = _plan_axes_from_family_transform_ft(column, width, depth)
    return width, depth, height, center, vs, vl


# ---------------------------------------------------------------------------
# Nivel y etiqueta de pilar
# ---------------------------------------------------------------------------

def _column_reference_level_name(elem):
    if elem is None:
        return None
    try:
        lid = elem.LevelId
        if lid is None:
            return None
        try:
            if int(lid.IntegerValue) < 0:
                return None
        except Exception:
            pass
        lv = elem.Document.GetElement(lid)
        return lv.Name.ToString() if lv is not None and lv.Name is not None else None
    except Exception:
        return None


def _raw_numeracion_columna_value(elem):
    if elem is None:
        return None
    for pname in _NUMERACION_COLUMNA_PARAM_CANDIDATES:
        try:
            p = elem.LookupParameter(pname)
        except Exception:
            continue
        if p is None:
            continue
        try:
            if not p.HasValue:
                continue
        except Exception:
            pass
        raw = None
        try:
            st = p.StorageType
            if st == StorageType.String:
                raw = p.AsString() or p.AsValueString()
            elif st == StorageType.Integer:
                raw = u"{}".format(int(p.AsInteger()))
            elif st == StorageType.Double:
                raw = u"{:.0f}".format(float(p.AsDouble()))
            else:
                raw = p.AsValueString()
        except Exception:
            try:
                raw = p.AsValueString()
            except Exception:
                raw = None
        if raw is None:
            continue
        t = u"{}".format(raw).strip()
        if t:
            return t
    return None


def _column_pilar_conjunto_label(elem):
    tok = _raw_numeracion_columna_value(elem)
    if not tok:
        return None
    s = u"{}".format(tok).strip()
    if not s:
        return None
    su = s.upper()
    if not su.startswith(u"P"):
        s = u"P{}".format(s)
    return u"Pilar {}".format(s)


# ---------------------------------------------------------------------------
# Filas del esquema de troceo
# ---------------------------------------------------------------------------

def build_troceo_scheme_rows(columns_ordered):
    """
    Tuplas ``(elemento, z_mm, id, height_mm, level_name, pilar_label)``
    ordenadas por z_mm ascendente para el TroceoSchemeController.
    """
    rows = []
    for col in columns_ordered or []:
        z_ft = 0.0
        h_ft = None
        try:
            dims = get_column_dimensions(col)
            z_ft = float(dims[3].Z)
            h_ft = float(dims[2])
        except Exception:
            pass
        try:
            z_mm = UnitUtils.ConvertFromInternalUnits(z_ft, UnitTypeId.Millimeters)
        except Exception:
            z_mm = float(z_ft) * 304.8
        h_mm = None
        if h_ft is not None:
            try:
                h_mm = UnitUtils.ConvertFromInternalUnits(h_ft, UnitTypeId.Millimeters)
            except Exception:
                try:
                    h_mm = float(h_ft) * 304.8
                except Exception:
                    h_mm = None
        eid = _element_id_iv(col)
        if eid < 0:
            continue
        rows.append((
            col,
            float(z_mm),
            eid,
            h_mm,
            _column_reference_level_name(col),
            _column_pilar_conjunto_label(col),
        ))
    rows.sort(key=lambda r: r[1])
    if not rows and columns_ordered:
        for col in columns_ordered:
            eid = _element_id_iv(col)
            if eid < 0:
                continue
            rows.append((col, 0.0, eid, None,
                         _column_reference_level_name(col),
                         _column_pilar_conjunto_label(col)))
        rows.sort(key=lambda r: r[1])
    return rows


# ---------------------------------------------------------------------------
# Overrides de estribos por columna
# ---------------------------------------------------------------------------

def stirrup_spacing_override(spacing_by_col_id, col):
    """Espaciamiento mm desde mapa {element_id_iv: float}. None si no hay override."""
    d = spacing_by_col_id or {}
    if not d or col is None:
        return None
    iv = _element_id_iv(col)
    try:
        if iv >= 0 and iv in d:
            return float(d[iv])
    except Exception:
        pass
    try:
        iv_i = int(col.Id.IntegerValue)
        if iv_i in d:
            return float(d[iv_i])
    except Exception:
        pass
    return None


def stirrup_bar_type_override(bar_type_by_col_id, col):
    """RebarBarType desde mapa {element_id_iv: RebarBarType}. None si no hay override."""
    d = bar_type_by_col_id or {}
    if not d or col is None:
        return None
    iv = _element_id_iv(col)
    try:
        if iv >= 0 and iv in d:
            return d[iv]
    except Exception:
        pass
    try:
        iv_i = int(col.Id.IntegerValue)
        if iv_i in d:
            return d[iv_i]
    except Exception:
        pass
    return None
