# -*- coding: utf-8 -*-
u"""
Lógica pura de creación de estribos y trabas para columnas.

Extraído de ``30_EstribosColumnaCanvas.pushbutton/script.py``.
Sin UI ni transacción; requiere una transacción activa al llamar a
``create_stirrups_for_column``.
"""
import clr

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Curve,
    ElementId,
    FilteredElementCollector,
    GeometryInstance,
    JoinGeometryUtils,
    Line,
    Options,
    Solid,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarStyle,
)


# ---------------------------------------------------------------------------
# Geometría de sólidos
# ---------------------------------------------------------------------------

def _geometry_options_column_solids():
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
        pass
    try:
        edges = solid.Edges
        ne = int(edges.Size)
    except Exception:
        return
    for i in range(ne):
        try:
            edge = edges.get_Item(i)
            crv = edge.AsCurve()
            if crv is None:
                continue
            for k in (0, 1):
                try:
                    yield crv.GetEndPoint(k)
                except Exception:
                    pass
        except Exception:
            continue


def _join_geometry_joined_element_ids(document, element):
    if document is None or element is None:
        return []
    raw = None
    for getter in (
        lambda: JoinGeometryUtils.GetJoinedElements(document, element),
        lambda: JoinGeometryUtils.GetJoinedElements(document, element.Id),
    ):
        try:
            raw = getter()
        except Exception:
            raw = None
        if raw is not None:
            break
    out = []
    if raw is None:
        return out
    try:
        for jid in raw:
            if jid is not None and jid != ElementId.InvalidElementId:
                out.append(jid)
    except (TypeError, AttributeError):
        try:
            n = int(raw.Count)
        except Exception:
            n = 0
        for i in range(n):
            jid = None
            try:
                jid = raw[i]
            except Exception:
                try:
                    jid = raw.get_Item(i)
                except Exception:
                    jid = None
            if jid is not None and jid != ElementId.InvalidElementId:
                out.append(jid)
    return out


def _column_has_joined_structural_foundation(document, column):
    if document is None or column is None:
        return False
    try:
        bic_f = int(BuiltInCategory.OST_StructuralFoundation)
    except Exception:
        return False
    for jid in _join_geometry_joined_element_ids(document, column):
        el = document.GetElement(jid)
        if el is None:
            continue
        try:
            cat = el.Category
            if cat is None or int(cat.Id.IntegerValue) != bic_f:
                continue
        except Exception:
            continue
        return True
    return False


# ---------------------------------------------------------------------------
# Ejes locales
# ---------------------------------------------------------------------------

def _column_axis_unit(column, lz_fallback):
    if column is None or lz_fallback is None:
        return XYZ.BasisZ if lz_fallback is None else lz_fallback
    lc = getattr(column.Location, "Curve", None)
    if lc is None:
        try:
            return lz_fallback.Normalize()
        except Exception:
            return lz_fallback
    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        vdir = p1.Subtract(p0)
        denom = float(vdir.GetLength())
        if denom < 1e-11:
            return lz_fallback.Normalize()
        return vdir.Multiply(1.0 / denom)
    except Exception:
        try:
            return lz_fallback.Normalize()
        except Exception:
            return lz_fallback


def _section_lx_ly_lz_orthogonal(column, trans):
    u"""Ejes locales ⟂ entre sí: ``lz`` según línea del pilar, ``lx`` proyectado desde ``BasisX``."""
    try:
        lz = _column_axis_unit(column, trans.BasisZ)
    except Exception:
        lz = XYZ.BasisZ
    try:
        lx_raw = trans.BasisX.Normalize()
    except Exception:
        lx_raw = XYZ.BasisX
    try:
        h = float(lx_raw.DotProduct(lz))
        lx_proj = lx_raw.Subtract(lz.Multiply(h))
        mx = float(lx_proj.GetLength())
        if mx < 1e-10:
            try:
                ly0 = trans.BasisY.Normalize()
            except Exception:
                ly0 = XYZ.BasisY
            lx_proj = ly0.CrossProduct(lz)
            mx = float(lx_proj.GetLength())
            if mx < 1e-10:
                lx = XYZ.BasisX
                ly = lz.CrossProduct(lx).Normalize()
                return lx, ly, lz
        lx = lx_proj.Multiply(1.0 / mx)
        ly = lz.CrossProduct(lx).Normalize()
    except Exception:
        lx, ly, lz = XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ
    return lx, ly, lz


def _column_axis_up_aligned_global_z(lz):
    u"""Eje del pilar normalizado con componente Z ≥ 0."""
    try:
        if lz is None or float(lz.GetLength()) < 1e-12:
            return XYZ.BasisZ
        u = lz.Normalize()
        if float(u.DotProduct(XYZ.BasisZ)) < 0.0:
            return u.Negate()
        return u
    except Exception:
        return XYZ.BasisZ


# ---------------------------------------------------------------------------
# Punto de anclaje en planta
# ---------------------------------------------------------------------------

def _nearest_on_bounded_location_curve_ft(column, q_xyz):
    if column is None or q_xyz is None:
        return None
    try:
        lc = getattr(column.Location, "Curve", None)
        if lc is None:
            return None
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        vdir = p1.Subtract(p0)
        denom = float(vdir.GetLength())
        if denom < 1e-11:
            return p0
        vu = vdir.Multiply(1.0 / denom)
        tv = float(q_xyz.Subtract(p0).DotProduct(vu))
        tv = max(0.0, min(tv, denom))
        return p0.Add(vu.Multiply(tv))
    except Exception:
        return None


def _anchor_plan_on_column_z_axis(trans, p_ref):
    if trans is None or p_ref is None:
        return p_ref
    try:
        o = trans.Origin
        vz = trans.BasisZ.Normalize()
    except Exception:
        return p_ref
    try:
        if vz.GetLength() < 1e-12:
            return p_ref
        t_ax = float(p_ref.Subtract(o).DotProduct(vz))
        return o.Add(vz.Multiply(t_ax))
    except Exception:
        return p_ref


def _plan_anchor_for_column_ft(column, trans, p_ref):
    q = _nearest_on_bounded_location_curve_ft(column, p_ref)
    if q is not None:
        return q
    return _anchor_plan_on_column_z_axis(trans, p_ref)


# ---------------------------------------------------------------------------
# Geometría de sección por sólidos
# ---------------------------------------------------------------------------

def _column_geom_anchor_and_sizes_ft(column, lx, ly, lz, ref_on_plane, band_half_candidates_ft=None):
    u"""Centro y dimensiones ``sa``, ``sb`` en pies proyectando vértices de sólidos en ``lx/ly``."""
    if column is None or ref_on_plane is None:
        return None
    if band_half_candidates_ft is None:
        band_half_candidates_ft = (
            350.0 / 304.8,
            600.0 / 304.8,
            1000.0 / 304.8,
            None,
        )
    try:
        lz_u = lz.Normalize()
    except Exception:
        lz_u = lz
    opts = _geometry_options_column_solids()

    def _accumulate_filtered(use_slice, band_ft_inf):
        min_u = max_u = min_w = max_w = None
        n = 0
        for solid in _iter_solids_revit_element(column, opts):
            for pt in _iter_vertices_from_solid(solid):
                try:
                    if use_slice:
                        h = abs(float(pt.Subtract(ref_on_plane).DotProduct(lz_u)))
                        if h > band_ft_inf:
                            continue
                    r = pt.Subtract(ref_on_plane)
                    u = float(r.DotProduct(lx))
                    w = float(r.DotProduct(ly))
                except Exception:
                    continue
                n += 1
                min_u = u if min_u is None else min(min_u, u)
                max_u = u if max_u is None else max(max_u, u)
                min_w = w if min_w is None else min(min_w, w)
                max_w = w if max_w is None else max(max_w, w)
        return n, min_u, max_u, min_w, max_w

    for cand in band_half_candidates_ft:
        use_slice = cand is not None
        band_lim = float(cand) if use_slice else 0.0
        n, min_u, max_u, min_w, max_w = _accumulate_filtered(use_slice, band_lim)
        min_pts = 16 if use_slice else 6
        if n < min_pts or min_u is None:
            continue
        span_u = float(max_u - min_u)
        span_w = float(max_w - min_w)
        if span_u < 1e-4 or span_w < 1e-4:
            continue
        uc = 0.5 * float(min_u + max_u)
        wc = 0.5 * float(min_w + max_w)
        try:
            anchor = ref_on_plane.Add(lx.Multiply(uc)).Add(ly.Multiply(wc))
        except Exception:
            continue
        return anchor, span_u, span_w
    return None


# ---------------------------------------------------------------------------
# Orientaciones de gancho
# ---------------------------------------------------------------------------

def _column_tie_hook_orientations(pt1, pt2, plane_normal, axis_point_xyz):
    u"""Par de orientaciones de gancho para una traba recta, orientando hacia el interior."""
    dvec = pt2.Subtract(pt1)
    ln = dvec.GetLength()
    if ln < 1e-12:
        return RebarHookOrientation.Left, RebarHookOrientation.Left
    t = dvec.Multiply(1.0 / ln)
    pn_len = plane_normal.GetLength()
    if pn_len < 1e-12:
        return RebarHookOrientation.Left, RebarHookOrientation.Left
    n = plane_normal.Multiply(1.0 / pn_len)
    mid = pt1.Add(pt2.Subtract(pt1).Multiply(0.5))
    to_axis = axis_point_xyz.Subtract(mid)
    h = float(to_axis.DotProduct(n))
    to_plane = to_axis.Subtract(n.Multiply(h))
    tpl = float(to_plane.GetLength())
    if tpl < 1e-12:
        return RebarHookOrientation.Left, RebarHookOrientation.Left
    d_in = to_plane.Multiply(1.0 / tpl)

    def _side(tang):
        lat = n.CrossProduct(tang)
        l = float(lat.GetLength())
        if l < 1e-12:
            return RebarHookOrientation.Left
        lat_u = lat.Multiply(1.0 / l)
        return (
            RebarHookOrientation.Right
            if float(lat_u.DotProduct(d_in)) < 0.0
            else RebarHookOrientation.Left
        )

    return _side(t), _side(t.Negate())


# ---------------------------------------------------------------------------
# Creación de Rebar
# ---------------------------------------------------------------------------

def _rebar_stirrup_create_try_normals(
    curves_list, host, bar_type, hook_t0, hook_t1, normals, o_start, o_end, doc
):
    u"""``Rebar.CreateFromCurves`` estilo estribo; prueba cada normal candidata."""
    last_ex = None
    for lz_try in normals:
        if lz_try is None:
            continue
        try:
            if float(lz_try.GetLength()) < 1e-12:
                continue
        except Exception:
            continue
        try:
            r = Rebar.CreateFromCurves(
                doc,
                RebarStyle.StirrupTie,
                bar_type,
                hook_t0,
                hook_t1,
                host,
                lz_try,
                curves_list,
                o_start,
                o_end,
                True,
                True,
            )
            if r is not None:
                return r, None
        except Exception as ex:
            try:
                last_ex = str(ex)
            except Exception:
                last_ex = u"(sin mensaje)"
    return None, last_ex


# ---------------------------------------------------------------------------
# Patrón de confinamiento
# ---------------------------------------------------------------------------

def stirrup_pattern_options(val):
    u"""Opciones de patrón para un lado con *val* barras (idéntico a BIMColumnApp.update_combo_logic)."""
    if val == 2:
        return [u"Perimetral Únicamente"]
    elif val == 3:
        return [u"Perimetral + Traba (Rojo)", u"Perimetral Únicamente"]
    elif val == 4:
        return [
            u"Perimetral + Estribo Int. (Verde)",
            u"Perimetral + Traba (Rojo)",
            u"Perimetral Únicamente",
        ]
    elif val == 5:
        return [
            u"Perimetral + Estribo Int. (Verde)",
            u"Perimetral + Estribo + Traba",
            u"Perimetral + Traba Central (Rojo)",
        ]
    elif val == 6:
        return [u"Perimetral + 2 Estribos Int.", u"Perimetral + 1 Estribo Central"]
    elif val == 7:
        return [
            u"Perimetral + 2 Estribos Int.",
            u"Perimetral + 1 Estribo Central",
            u"Perimetral + 2 Estribos + Traba",
        ]
    elif val == 8:
        return [
            u"Perimetral + 2 Estribos (3 barras/est)",
            u"Perimetral + 2 Estribos + 2 Trabas",
            u"Perimetral + Estribo (2-5) + Traba (4)",
        ]
    elif val == 9:
        return [
            u"Perimetral + 2 Estribos Int. (1-3, 5-7)",
            u"Perimetral + 2 Estribos + 3 Trabas (2, 4, 6)",
        ]
    elif val == 10:
        return [
            u"Perimetral + 3 Estribos (1-2, 4-5, 7-8)",
            u"Perimetral + 2 Estribos (2-4, 5-7)",
        ]
    else:
        return [u"Perimetral Únicamente"]


def tie_axis_shift_toward_section_center(
    center_coord,
    long_bar_radius,
    tie_index=None,
    bar_count=None,
):
    u"""
    Desplaza la coordenada del eje de la traba desde el **centro** de la barra longitudinal
    hacia la **tangente interior** respecto al eje de la sección (origen local 0).

    Si la barra cae en el eje de simetría (``center_coord`` ≈ 0), hace falta ``tie_index`` y
    ``bar_count`` para elegir un lado determinista (antes no se desplazaba y la traba quedaba
    sobre el centro del círculo).

    ``center_coord`` y ``long_bar_radius`` deben usar las mismas unidades
    (pies en modelo o píxeles en preview lineal respecto al centro).
    """
    r = float(long_bar_radius)
    if r <= 0.0:
        return float(center_coord)
    c = float(center_coord)
    tol = max(1e-12, r * 1e-6)
    on_axis = abs(c) <= tol
    if on_axis:
        if tie_index is None or bar_count is None:
            return c
        bc = int(bar_count)
        ti = int(tie_index)
        if bc <= 1 or not (0 <= ti < bc):
            return c
        hi = float(bc - 1)
        tif = float(ti)
        if tif * 2.0 < hi:
            return -r
        if tif * 2.0 > hi:
            return r
        return r
    sign = 1.0 if c > 0.0 else -1.0
    shifted = c - sign * r
    if sign > 0.0 and shifted < 0.0:
        return 0.0
    if sign < 0.0 and shifted > 0.0:
        return 0.0
    return shifted


def build_stirrup_rect_and_tie_defs(val_a, val_b, sel_a_text, sel_b_text):
    u"""
    Devuelve ``(rect_defs, tie_defs)`` según el texto de selección de patrón.

    ``rect_defs`` — lista de ``(idx_a, idx_b, sp_a, sp_b)``; perimetral siempre en posición 0.
    ``tie_defs``  — lista de ``(idx, is_a)`` para cada traba.
    """
    rect_defs = [(0, 0, val_a - 1, val_b - 1)]
    tie_defs = []

    def _parse(val, text, is_a):
        if not text or u"\u00da" + u"nicamente" in text or u"nicamente" in text:
            return
        spans = []
        if val == 10:
            if u"3 Estribos" in text:
                spans = [(1, 1), (4, 1), (7, 1)]
            elif u"2 Estribos" in text:
                spans = [(2, 2), (5, 2)]
        elif val == 9:
            if u"1-3, 5-7" in text:
                spans = [(1, 2), (5, 2)]
        elif val == 8:
            if u"3 barras/est" in text:
                spans = [(1, 2), (4, 2)]
            elif u"2-5" in text:
                spans = [(2, 3)]
        elif val == 7:
            if u"2 Estribos" in text:
                spans = [(1, 1), (4, 1)]
            elif u"1 Estribo" in text:
                spans = [(2, 2)]
        elif val == 6:
            if u"2 Estribos" in text:
                spans = [(1, 1), (3, 1)]
            elif u"1 Estribo" in text:
                spans = [(2, 1)]
        elif val <= 5:
            if u"Estribo" in text:
                spans = [(1, val - 3 if val > 3 else 1)]

        for idx, sp in spans:
            if is_a:
                rect_defs.append((idx, 0, sp, val_b - 1))
            else:
                rect_defs.append((0, idx, val_a - 1, sp))

        if u"Traba" in text:
            t_idx = []
            if val == 9 and u"3 Trabas" in text:
                t_idx = [2, 4, 6]
            elif val == 7 and u"Traba" in text:
                t_idx = [3]
            elif val == 5 and u"Traba" in text:
                t_idx = [2]
            elif val == 4 and u"Traba" in text:
                t_idx = [1]
            elif val == 3:
                t_idx = [1]
            for idx in t_idx:
                tie_defs.append((idx, is_a))

    _parse(val_a, sel_a_text, True)
    _parse(val_b, sel_b_text, False)
    return rect_defs, tie_defs


# ---------------------------------------------------------------------------
# Punto de entrada principal
# ---------------------------------------------------------------------------

def column_bar_geometry(col, stirrup_bar_type=None, cover_mm=25.0, long_bar_diam_mm=16.0):
    u"""
    Geometría de sección alineada al pipeline de estribos.

    Devuelve ``(plan_anchor, lx, ly, sa, sb, offset_long_ft)`` donde:

    * ``plan_anchor`` — punto en la cara inferior del pilar usado como origen de la rejilla.
    * ``lx``, ``ly`` — ejes locales ortogonales en planta (mismos que usa ``create_stirrups_for_column``).
    * ``sa``, ``sb`` — dimensiones de la sección en pies proyectadas sobre ``lx`` y ``ly``
      respectivamente (en ese orden, sin reordenar por longitud).
    * ``offset_long_ft`` — distancia cara → eje de barra longitudinal
      ``= cover_ft + Ø_estribo + radio_barra_long``.

    ``long_bar_diam_mm``: diámetro nominal de la barra longitudinal en mm (para el radio).
    Devuelve ``None`` si no se puede obtener la geometría mínima del pilar.
    """
    try:
        trans = col.GetTotalTransform()
    except Exception:
        try:
            trans = col.GetTransform()
        except Exception:
            return None

    lx, ly, lz = _section_lx_ly_lz_orthogonal(col, trans)

    loc = col.Location
    try:
        if hasattr(loc, "Curve"):
            p1 = loc.Curve.Evaluate(0.0, True)
            p2 = loc.Curve.Evaluate(1.0, True)
            p_z_base = p1 if p1.Z < p2.Z else p2
        else:
            bb = col.get_BoundingBox(None)
            z_base = bb.Min.Z if bb else trans.Origin.Z
            origin = loc.Point if hasattr(loc, "Point") else trans.Origin
            p_z_base = XYZ(origin.X, origin.Y, z_base)
    except Exception:
        p_z_base = trans.Origin

    plan_axis_pt = _plan_anchor_for_column_ft(col, trans, p_z_base)
    geom_sec = _column_geom_anchor_and_sizes_ft(col, lx, ly, lz, plan_axis_pt)

    if geom_sec is not None:
        plan_anchor, sa, sb = geom_sec
    else:
        sa_ft = None
        sb_ft = None
        for pname in (u"b", u"Width", u"Base"):
            try:
                p = col.Symbol.LookupParameter(pname)
                if p and p.HasValue:
                    sa_ft = p.AsDouble()
                    break
            except Exception:
                pass
        for pname in (u"h", u"Height", u"Altura"):
            try:
                p = col.Symbol.LookupParameter(pname)
                if p and p.HasValue:
                    sb_ft = p.AsDouble()
                    break
            except Exception:
                pass
        if sa_ft is None and sb_ft is None:
            return None
        sa = sa_ft if sa_ft is not None else (300.0 / 304.8)
        sb = sb_ft if sb_ft is not None else (400.0 / 304.8)
        plan_anchor = plan_axis_pt

    cover_ft = float(cover_mm) / 304.8
    if stirrup_bar_type is not None:
        try:
            stirrup_diam_ft = stirrup_bar_type.BarModelDiameter
        except AttributeError:
            stirrup_diam_ft = stirrup_bar_type.BarDiameter
    else:
        stirrup_diam_ft = 8.0 / 304.8
    long_bar_r_ft = float(long_bar_diam_mm) / 2.0 / 304.8
    offset_long_ft = cover_ft + stirrup_diam_ft + long_bar_r_ft

    return plan_anchor, lx, ly, sa, sb, offset_long_ft


def create_stirrups_for_column(
    doc,
    col,
    val_a,
    val_b,
    sel_a_text,
    sel_b_text,
    stirrup_bar_type,
    hook_135_type,
    spacing_mm,
    cover_mm=25.0,
    long_bar_diam_mm=16.0,
    collect_rebars=None,
):
    u"""
    Crea estribos y trabas para *col* dentro de una transacción activa.

    ``long_bar_diam_mm``: diámetro **modelo** (Revit) del longitudinal en mm,
    coherente con ``RebarBarType.BarModelDiameter`` de las barras creadas.

    ``collect_rebars``: si es una lista, se añade cada ``Rebar`` creado (rectos y trabas).

    Devuelve el número de ``Rebar`` creados.
    Propaga excepciones para que el llamante pueda registrarlas y hacer rollback.
    """
    n_created = 0
    # Let exceptions propagate so the caller can report them.
    try:
        trans = col.GetTotalTransform()
    except Exception:
        trans = col.GetTransform()
    lx, ly, lz = _section_lx_ly_lz_orthogonal(col, trans)
    try:

        lz_norm_candidates = []
        try:
            if lz is not None and float(lz.GetLength()) > 1e-12:
                try:
                    bz_u = trans.BasisZ.Normalize()
                    if float(lz.DotProduct(bz_u)) >= 0.0:
                        lz_norm_candidates.extend([lz, lz.Negate()])
                    else:
                        lz_norm_candidates.extend([lz.Negate(), lz])
                except Exception:
                    lz_norm_candidates.extend([lz, lz.Negate()])
        except Exception:
            pass
        if not lz_norm_candidates:
            lz_norm_candidates = [lz]

        loc = col.Location
        if hasattr(loc, "Curve"):
            p1 = loc.Curve.Evaluate(0.0, True)
            p2 = loc.Curve.Evaluate(1.0, True)
            p_z_base = p1 if p1.Z < p2.Z else p2
            h_col = p1.DistanceTo(p2)
        else:
            bb = col.get_BoundingBox(None)
            z_base = bb.Min.Z if bb else trans.Origin.Z
            origin = loc.Point if hasattr(loc, "Point") else trans.Origin
            p_z_base = XYZ(origin.X, origin.Y, z_base)
            h_col = (bb.Max.Z - bb.Min.Z) if bb else 10.0

        plan_axis_pt = _plan_anchor_for_column_ft(col, trans, p_z_base)
        geom_sec = _column_geom_anchor_and_sizes_ft(col, lx, ly, lz, plan_axis_pt)

        try:
            stirrup_diam_ft = stirrup_bar_type.BarModelDiameter
        except AttributeError:
            stirrup_diam_ft = stirrup_bar_type.BarDiameter

        cover_ft = cover_mm / 304.8
        long_bar_r_ft = float(long_bar_diam_mm) / 2.0 / 304.8
        stirrup_r_ft = stirrup_diam_ft * 0.5
        offset_axis = cover_ft
        offset_long = cover_ft + stirrup_diam_ft + long_bar_r_ft
        # Desde centro de barra hasta línea de centro del estribo en patas tangentes.
        tangent_inset_ft = long_bar_r_ft + stirrup_r_ft

        if geom_sec is not None:
            plan_anchor, sa, sb = geom_sec
        else:
            sa_ft = None
            sb_ft = None
            for pname in (u"b", u"Width", u"Base"):
                try:
                    p = col.Symbol.LookupParameter(pname)
                    if p and p.HasValue:
                        sa_ft = p.AsDouble()
                        break
                except Exception:
                    pass
            for pname in (u"h", u"Height", u"Altura"):
                try:
                    p = col.Symbol.LookupParameter(pname)
                    if p and p.HasValue:
                        sb_ft = p.AsDouble()
                        break
                except Exception:
                    pass
            sa = sa_ft if sa_ft is not None else (300.0 / 304.8)
            sb = sb_ft if sb_ft is not None else (400.0 / 304.8)
            plan_anchor = plan_axis_pt

        foundation_drop_ft = 0.0
        if _column_has_joined_structural_foundation(doc, col):
            foundation_drop_ft = 300.0 / 304.8
        lz_up = _column_axis_up_aligned_global_z(lz)

        plan_anchor_rebar = plan_anchor
        if foundation_drop_ft > 1e-12:
            try:
                plan_anchor_rebar = plan_anchor.Add(lz_up.Multiply(-foundation_drop_ft))
            except Exception:
                pass

        array_length = h_col + foundation_drop_ft
        if array_length <= 0.0:
            array_length = 1.0
        spacing_ft = spacing_mm / 304.8

        val_a = int(val_a)
        val_b = int(val_b)
        gap_a = (sa - 2.0 * offset_long) / (val_a - 1) if val_a > 1 else 0.0
        gap_b = (sb - 2.0 * offset_long) / (val_b - 1) if val_b > 1 else 0.0

        rect_defs, tie_defs = build_stirrup_rect_and_tie_defs(val_a, val_b, sel_a_text, sel_b_text)

        def _center_along_a(i):
            return -sa / 2.0 + offset_long + float(i) * gap_a

        def _center_along_b(i):
            return -sb / 2.0 + offset_long + float(i) * gap_b

        for idx_a, idx_b, sp_a, sp_b in rect_defs:
            # Eje A (lx): parcial → patas tangentes a centros de barras; completo → perimetral.
            if sp_a < val_a - 1:
                x_left = _center_along_a(idx_a) - tangent_inset_ft
                x_right = _center_along_a(idx_a + sp_a) + tangent_inset_ft
                w_span = x_right - x_left
                x_s = x_left
            else:
                w_span = sa - 2.0 * offset_axis
                x_s = -sa / 2.0 + offset_axis + float(idx_a) * gap_a

            # Eje B (ly): ídem.
            if sp_b < val_b - 1:
                y_low = _center_along_b(idx_b) - tangent_inset_ft
                y_high = _center_along_b(idx_b + sp_b) + tangent_inset_ft
                h_span = y_high - y_low
                y_s = y_low
            else:
                h_span = sb - 2.0 * offset_axis
                y_s = -sb / 2.0 + offset_axis + float(idx_b) * gap_b

            pt1 = plan_anchor_rebar + lx.Multiply(x_s) + ly.Multiply(y_s)
            pt2 = pt1 + lx.Multiply(w_span)
            pt3 = pt2 + ly.Multiply(h_span)
            pt4 = pt1 + ly.Multiply(h_span)
            # Path CW desde pt2 (BR): BR→BL→TL→TR→BR.
            # El gancho cae en la esquina inferior (pt2), dejando ambas
            # esquinas superiores con doblez recto 90° simétricas entre sí.
            pts4 = [pt2, pt1, pt4, pt3, pt2]
            curves = List[Curve]()
            for k in range(4):
                pa_k, pb_k = pts4[k], pts4[k + 1]
                if pa_k.DistanceTo(pb_k) > 0.01:
                    curves.Add(Line.CreateBound(pa_k, pb_k))
            if curves.Count > 0:
                rebar, _ = _rebar_stirrup_create_try_normals(
                    curves, col, stirrup_bar_type, hook_135_type, hook_135_type,
                    lz_norm_candidates,
                    RebarHookOrientation.Right,
                    RebarHookOrientation.Right,
                    doc,
                )
                if rebar is not None:
                    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(
                        spacing_ft, array_length, True, True, False
                    )
                    n_created += 1
                    if collect_rebars is not None:
                        try:
                            collect_rebars.append(rebar)
                        except Exception:
                            pass

        for idx, is_a in tie_defs:
            if is_a:
                x_tie = tie_axis_shift_toward_section_center(
                    _center_along_a(idx),
                    long_bar_r_ft,
                    tie_index=idx,
                    bar_count=val_a,
                )
                pt1 = (
                    plan_anchor_rebar
                    + lx.Multiply(x_tie)
                    + ly.Multiply(-sb / 2.0 + offset_axis)
                )
                pt2 = pt1 + ly.Multiply(sb - 2.0 * offset_axis)
            else:
                y_tie = tie_axis_shift_toward_section_center(
                    _center_along_b(idx),
                    long_bar_r_ft,
                    tie_index=idx,
                    bar_count=val_b,
                )
                pt1 = (
                    plan_anchor_rebar
                    + lx.Multiply(-sa / 2.0 + offset_axis)
                    + ly.Multiply(y_tie)
                )
                pt2 = pt1 + lx.Multiply(sa - 2.0 * offset_axis)
            if pt1.DistanceTo(pt2) > 0.01:
                curves = List[Curve]()
                curves.Add(Line.CreateBound(pt1, pt2))
                mid_tie = pt1.Add(pt2.Subtract(pt1).Multiply(0.5))
                axis_ref = _plan_anchor_for_column_ft(col, trans, mid_tie)
                o_start, o_end = _column_tie_hook_orientations(pt1, pt2, lz, axis_ref)
                rebar, _ = _rebar_stirrup_create_try_normals(
                    curves, col, stirrup_bar_type, hook_135_type, hook_135_type,
                    lz_norm_candidates, o_start, o_end, doc,
                )
                if rebar is not None:
                    rebar.GetShapeDrivenAccessor().SetLayoutAsMaximumSpacing(
                        spacing_ft, array_length, True, True, False
                    )
                    n_created += 1
                    if collect_rebars is not None:
                        try:
                            collect_rebars.append(rebar)
                        except Exception:
                            pass

    except Exception:
        raise
    return n_created
