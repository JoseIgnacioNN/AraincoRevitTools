# -*- coding: utf-8 -*-
u"""
Creación de armadura longitudinal para columnas.

REGLA DE CAPA:
- ``run_longitudinal_layout`` abre su propio TransactionGroup + Transaction.
- ``create_vertical_rebar`` / ``create_longitudinal_rebar_with_optional_patas``
  requieren transacción activa.
- No importa módulos de ui/.

Portado desde column_reinforcement_layout_rps.py.
"""
from __future__ import print_function

import math

import clr
clr.AddReference("RevitAPI")

import System
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    ElementId,
    FilteredElementCollector,
    Line,
    Plane,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarShape,
    RebarStyle,
)

from core.geometry import (
    _element_id_iv,
    _canonical_section_mm_key,
    _solid_aggregate_vertex_ranges_ft,
    COLUMN_REBAR_L_SHAPE_DISPLAY_NAME,
    COLUMN_ARMA_UBICACION_PARAM,
    LAYOUT_EMBED_CONCRETE_GRADE,
)
from core.tables import hook_length_mm_from_nominal_diameter_mm
from core.troceo_engine import (
    column_bottom_joined_foundation_stretch_down_mm,
    embed_stretch_collides_any_column_solids,
    _iter_solids_revit_element,
    _geometry_options_structure_solids,
)
from core.jobs import (
    generate_bar_points,
    fuse_vertical_world_intervals_from_jobs,
    _linea_fierro_nombre_alfabetico,
    _arma_len_mm_round_from_internal_ft,
)


# ---------------------------------------------------------------------------
# Resolución de RebarBarType
# ---------------------------------------------------------------------------

def _rebar_nominal_diameter_mm(bt):
    if bt is None:
        return None
    for attr in ("BarNominalDiameter", "BarDiameter"):
        try:
            raw = getattr(bt, attr)
            if raw is not None:
                return float(
                    UnitUtils.ConvertFromInternalUnits(float(raw), UnitTypeId.Millimeters)
                )
        except Exception:
            pass
    return None


def _resolve_rebar_bar_type_by_diameter_mm(doc, target_mm):
    """Busca el RebarBarType más cercano al diámetro nominal solicitado."""
    if doc is None:
        return None, False, None
    best = None
    best_delta = None
    target = float(target_mm)
    try:
        coll = FilteredElementCollector(doc).OfClass(RebarBarType)
    except Exception:
        coll = []
    for bt in coll:
        dmm = _rebar_nominal_diameter_mm(bt)
        if dmm is None:
            continue
        delta = abs(float(dmm) - target)
        if best is None or delta < best_delta:
            best = bt
            best_delta = delta
    if best is None:
        return None, False, None
    return best, (float(best_delta) <= 0.25), float(best_delta)


# ---------------------------------------------------------------------------
# Creación de curvas
# ---------------------------------------------------------------------------

def _curve_ilist_for_rebar(curves):
    lst = List[Curve]()
    for crv in curves:
        lst.Add(crv)
    return lst


def _curve_clr_array_curve_host(curves):
    n = len(curves)
    clr_arr = System.Array.CreateInstance(Curve, n)
    for i, crv in enumerate(curves):
        clr_arr[i] = crv
    return clr_arr


# ---------------------------------------------------------------------------
# Formas de Rebar
# ---------------------------------------------------------------------------

def _shape_display_name_normalized(value):
    if value is None:
        return u""
    try:
        t = u"{}".format(value)
    except Exception:
        return u""
    try:
        return t.replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def _rebar_shape_visible_label(shape):
    if shape is None:
        return u""
    for bip in (BuiltInParameter.SYMBOL_NAME_PARAM, BuiltInParameter.ALL_MODEL_TYPE_NAME):
        try:
            p = shape.get_Parameter(bip)
            if p is not None and p.HasValue:
                s = _shape_display_name_normalized(p.AsString())
                if s:
                    return s
        except Exception:
            continue
    try:
        return _shape_display_name_normalized(getattr(shape, "Name", None))
    except Exception:
        return u""


def _document_rebar_shape_by_visible_name(doc, nombre_visible):
    if doc is None or not nombre_visible:
        return None
    key = _shape_display_name_normalized(nombre_visible)
    if not key:
        return None
    try:
        key_lower = key.lower()
    except Exception:
        key_lower = key
    key_digits = u"".join(ch for ch in key if ch in u"0123456789")
    candidates = []
    try:
        for sh in FilteredElementCollector(doc).OfClass(RebarShape):
            try:
                sn = _rebar_shape_visible_label(sh)
                if not sn:
                    continue
                try:
                    sn_low = sn.lower()
                except Exception:
                    sn_low = sn
                dig = u"".join(ch for ch in sn if ch in u"0123456789")
                candidates.append((sh, sn, sn_low, dig))
            except Exception:
                continue
    except Exception:
        return None
    for sh, sn, _, _dig in candidates:
        if sn == key:
            return sh
    for sh, _, sn_low, _dig in candidates:
        if sn_low == key_lower:
            return sh
    for sh, sn, _, dig in candidates:
        if dig and dig == key:
            return sh
    for sh, _sn, _sn_low, dig in candidates:
        if key_digits and dig == key_digits:
            return sh
    return None


# ---------------------------------------------------------------------------
# CreateFromCurves helpers
# ---------------------------------------------------------------------------

def _try_create_from_curves_and_shape(doc, host, bar_type, curves, normal_vec, shape_name):
    if doc is None or host is None or bar_type is None or not curves:
        return None
    shape = _document_rebar_shape_by_visible_name(doc, shape_name)
    if shape is None:
        return None
    try:
        cl = _curve_ilist_for_rebar(curves)
    except Exception:
        return None
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    norms = [normal_vec]
    try:
        norms.append(normal_vec.Negate())
    except Exception:
        pass
    invalid = ElementId.InvalidElementId
    for nvec in norms:
        if nvec is None:
            continue
        for so, eo in orient_pairs:
            rb = None
            try:
                rb = Rebar.CreateFromCurvesAndShape(
                    doc, shape, bar_type, None, None,
                    host, nvec, cl, so, eo, 0.0, 0.0, invalid, invalid,
                )
            except Exception:
                try:
                    rb = Rebar.CreateFromCurvesAndShape(
                        doc, shape, bar_type, None, None,
                        host, nvec, cl, so, eo,
                    )
                except Exception:
                    rb = None
            if rb is not None:
                try:
                    rb.Style = RebarStyle.Standard
                except Exception:
                    pass
                return rb
    return None


def _create_rebar_from_curves_no_hooks(doc, host, bar_type, curves, normal_vec):
    if doc is None or host is None or bar_type is None or not curves:
        return None
    try:
        arr_list = _curve_ilist_for_rebar(curves)
    except Exception:
        return None
    n_cv = len(curves)
    curve_inputs = [(arr_list, "List")]
    if n_cv > 1:
        try:
            curve_inputs.append((_curve_clr_array_curve_host(curves), "Arr"))
        except Exception:
            pass
    if n_cv > 1:
        use_pairs = ((False, False), (False, True), (True, False), (True, True))
    else:
        use_pairs = ((True, True), (True, False), (False, True), (False, False))
    norms = [normal_vec]
    try:
        norms.append(normal_vec.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    for curve_arg, _ in curve_inputs:
        for use_existing, create_new in use_pairs:
            for nvec in norms:
                if nvec is None:
                    continue
                for so, eo in orient_pairs:
                    try:
                        rb = Rebar.CreateFromCurves(
                            doc, RebarStyle.Standard, bar_type, None, None,
                            host, nvec, curve_arg, so, eo, use_existing, create_new,
                        )
                        if rb is not None:
                            return rb
                    except Exception:
                        continue
    return None


# ---------------------------------------------------------------------------
# Comentarios / Armadura_Ubicacion
# ---------------------------------------------------------------------------

def _set_rebar_comment_text(doc, rebar_elem, txt):
    if rebar_elem is None or not txt:
        return
    try:
        p = rebar_elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p is not None and not p.IsReadOnly:
            p.Set(u"{}".format(txt))
    except Exception:
        pass


def _aplicar_armadura_ubicacion(rebar_element, valor_texto):
    if rebar_element is None or valor_texto is None:
        return
    try:
        txt = u"{}".format(valor_texto)
        p = rebar_element.LookupParameter(COLUMN_ARMA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return
        p.Set(txt)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Centro en planta de columna contribuyente
# ---------------------------------------------------------------------------

def _nearest_contrib_column_plan_center_xy(doc, bar_x, bar_y, contrib_elem_ids):
    best_xy = None
    best_d2 = None
    bx = float(bar_x)
    by = float(bar_y)
    for eid in contrib_elem_ids or []:
        try:
            el = doc.GetElement(eid)
        except Exception:
            continue
        if el is None:
            continue
        rng = _solid_aggregate_vertex_ranges_ft(el)
        if rng is None:
            continue
        cx = 0.5 * (float(rng[0]) + float(rng[1]))
        cy = 0.5 * (float(rng[2]) + float(rng[3]))
        d2 = (bx - cx) ** 2 + (by - cy) ** 2
        if best_d2 is None or d2 < best_d2:
            best_d2 = d2
            best_xy = (cx, cy)
    return best_xy


# ---------------------------------------------------------------------------
# Pata horizontal
# ---------------------------------------------------------------------------

def _pata_horizontal_leg_at_z(doc, bx, by, z_level, contrib_elem_ids, pata_len_ft):
    bx = float(bx)
    by = float(by)
    zr = float(z_level)
    lf = float(pata_len_ft)
    if lf <= 1e-12:
        return None
    nxy = _nearest_contrib_column_plan_center_xy(doc, bx, by, contrib_elem_ids)
    if nxy is None:
        return None
    cx, cy = nxy
    dx = float(cx) - bx
    dy = float(cy) - by
    hpl = math.hypot(dx, dy)
    if hpl <= 1e-9:
        return None
    ux = dx / hpl
    uy = dy / hpl
    pa = XYZ(bx, by, zr)
    pb = XYZ(bx + ux * lf, by + uy * lf, zr)
    return (pa, pb, ux, uy)


def _rebar_plane_normal_vertical_with_xy_leg(ux, uy):
    try:
        n = XYZ(float(uy), -float(ux), 0.0)
        lm = float(n.GetLength())
        if lm < 1e-12:
            return None
        return n.Normalize()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Creación de una barra longitudinal
# ---------------------------------------------------------------------------

def create_vertical_rebar(doc, pt, span_along_z, host, bar_type, comment_text=None):
    dz = float(span_along_z)
    tol = doc.Application.ShortCurveTolerance
    if abs(dz) < tol:
        return None
    start_pt = XYZ(pt.X, pt.Y, pt.Z)
    end_pt   = XYZ(pt.X, pt.Y, pt.Z + dz)
    try:
        curve = Line.CreateBound(start_pt, end_pt)
    except Exception:
        return None
    rb = _create_rebar_from_curves_no_hooks(doc, host, bar_type, [curve], XYZ.BasisX)
    if rb is not None and comment_text:
        _set_rebar_comment_text(doc, rb, comment_text)
    return rb


def create_longitudinal_rebar_with_optional_patas(
    doc,
    bx,
    by,
    zs,
    span_seg_z,
    host,
    bar_type,
    comment_text,
    want_bottom_pata,
    want_top_pata,
    pata_len_ft,
    contrib_ids,
):
    """
    Un solo Rebar: tramo vertical con pata inferior y/o superior en la misma cadena de curvas.
    Devuelve ``(rebar, did_bottom_pata, did_top_pata)``.
    """
    bx = float(bx)
    by = float(by)
    zs = float(zs)
    span = float(span_seg_z)
    tol = abs(float(doc.Application.ShortCurveTolerance))
    if span <= tol:
        return None, False, False

    z_top = zs + span
    foot  = XYZ(bx, by, zs)
    head  = XYZ(bx, by, z_top)

    bot_leg = None
    top_leg = None
    if want_bottom_pata:
        bot_leg = _pata_horizontal_leg_at_z(doc, bx, by, zs, contrib_ids, pata_len_ft)
    if want_top_pata:
        top_leg = _pata_horizontal_leg_at_z(doc, bx, by, z_top, contrib_ids, pata_len_ft)

    curves = []
    ux_plane = None
    uy_plane = None

    if bot_leg and top_leg:
        pa_b, pb_b, ux_b, uy_b = bot_leg
        pa_t, pb_t, ux_t, uy_t = top_leg
        ux_plane = ux_b
        uy_plane = uy_b
        try:
            curves.append(Line.CreateBound(pb_b, pa_b))
            curves.append(Line.CreateBound(pa_b, head))
            curves.append(Line.CreateBound(pa_t, pb_t))
        except Exception:
            return None, False, False
    elif bot_leg:
        pa_b, pb_b, ux_b, uy_b = bot_leg
        ux_plane = ux_b
        uy_plane = uy_b
        try:
            curves.append(Line.CreateBound(head, pa_b))
            curves.append(Line.CreateBound(pa_b, pb_b))
        except Exception:
            bot_leg = None
            curves = []
    elif top_leg:
        pa_t, pb_t, ux_t, uy_t = top_leg
        ux_plane = ux_t
        uy_plane = uy_t
        try:
            curves.append(Line.CreateBound(foot, head))
            curves.append(Line.CreateBound(pa_t, pb_t))
        except Exception:
            top_leg = None
            curves = []

    if not curves:
        try:
            curves.append(Line.CreateBound(foot, head))
        except Exception:
            return None, False, False

    nvec = XYZ.BasisX
    if ux_plane is not None and uy_plane is not None:
        n_pl = _rebar_plane_normal_vertical_with_xy_leg(ux_plane, uy_plane)
        if n_pl is not None:
            nvec = n_pl

    rb = None
    if len(curves) == 2 and ((bot_leg is not None) ^ (top_leg is not None)):
        rb = _try_create_from_curves_and_shape(
            doc, host, bar_type, curves, nvec, COLUMN_REBAR_L_SHAPE_DISPLAY_NAME
        )
    if rb is None:
        rb = _create_rebar_from_curves_no_hooks(doc, host, bar_type, curves, nvec)
    if rb is not None and comment_text:
        _set_rebar_comment_text(doc, rb, comment_text)

    did_bot = bool(bot_leg is not None and want_bottom_pata)
    did_top = bool(top_leg is not None and want_top_pata)
    return rb, did_bot, did_top


# ---------------------------------------------------------------------------
# Planos de troceo desde Location.Curve de columna
# ---------------------------------------------------------------------------

def _column_solid_vertex_z_min_max(col):
    rng = _solid_aggregate_vertex_ranges_ft(col)
    if rng is None:
        return None, None
    try:
        return float(rng[4]), float(rng[5])
    except Exception:
        return None, None


def _horizontal_plane_z_snap_to_bbox(pl, col, use_top):
    if pl is None or col is None:
        return pl
    try:
        n = pl.Normal
        if abs(float(n.Z)) < 0.99:
            return pl
        zm, zM = _column_solid_vertex_z_min_max(col)
        if zm is None or zM is None:
            return pl
        zt = float(zM if use_top else zm)
        o = pl.Origin
        return Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(float(o.X), float(o.Y), zt))
    except Exception:
        return pl


def _plane_from_column_location_curve(col, short_curve_tolerance_ft):
    tol = max(abs(float(short_curve_tolerance_ft)), 1e-12)
    if col is None:
        return None
    try:
        loc = col.Location
        cr = getattr(loc, "Curve", None)
        if cr is not None:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            if v.GetLength() >= tol:
                pl = Plane.CreateByNormalAndOrigin(v.Normalize(), p0)
            else:
                pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p0)
        else:
            pt_loc = getattr(loc, "Point", None)
            if pt_loc is not None:
                pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, pt_loc)
            else:
                return None
    except Exception:
        return None
    return _horizontal_plane_z_snap_to_bbox(pl, col, use_top=False)


def _plane_from_column_location_curve_at_end(col, short_curve_tolerance_ft):
    tol = max(abs(float(short_curve_tolerance_ft)), 1e-12)
    if col is None:
        return None
    try:
        loc = col.Location
        cr = getattr(loc, "Curve", None)
        if cr is not None:
            p0 = cr.GetEndPoint(0)
            p1 = cr.GetEndPoint(1)
            v = p1 - p0
            if v.GetLength() >= tol:
                pl = Plane.CreateByNormalAndOrigin(v.Normalize(), p1)
            else:
                pl = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, p1)
        else:
            pt_loc = getattr(loc, "Point", None)
            if pt_loc is not None:
                _, zM = _column_solid_vertex_z_min_max(col)
                if zM is None:
                    return None
                pl = Plane.CreateByNormalAndOrigin(
                    XYZ.BasisZ, XYZ(float(pt_loc.X), float(pt_loc.Y), zM)
                )
            else:
                return None
    except Exception:
        return None
    return _horizontal_plane_z_snap_to_bbox(pl, col, use_top=True)


# ---------------------------------------------------------------------------
# Visibilidad 3D
# ---------------------------------------------------------------------------

def _apply_rebar_3d_visibility(doc, rebar_elements):
    """Aplica visibilidad 3D a los Rebar creados (sin transacción)."""
    for rb in rebar_elements or []:
        if rb is None:
            continue
        try:
            rb.SetUnobscuredInView(
                doc.ActiveView,
                True,
            )
        except Exception:
            pass
        try:
            acc = rb.GetShapeDrivenAccessor()
            if acc is not None:
                acc.SetLayoutAsSingle()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Punto de entrada de capa
# ---------------------------------------------------------------------------

def run_longitudinal_layout(doc, fused_world, wiz, dims_cache, columns_ordered):
    """
    Crea toda la armadura longitudinal en un TransactionGroup.

    ``fused_world``: lista de ``(base_xyz, span_z_ft, contrib_elem_ids, bar_enum)``
    de ``fuse_vertical_world_intervals_from_jobs``.

    ``wiz`` debe exponer:
    - ``section_grid_config``: dict {(s,L): {bars_a, bars_b, ...}}
    - ``global_long_bar_diam_mm``: float
    - ``troceo_cut_planes_a``: list de Plane (planos de troceo tipo A)
    - ``troceo_diams_by_label``: dict {label: float_mm}
    - ``concrete_grade``: str|None
    """
    from Autodesk.Revit.DB import TaskDialog

    nominal_diam_mm = float(getattr(wiz, "global_long_bar_diam_mm", 12.0))
    concrete_grade  = getattr(wiz, "concrete_grade", LAYOUT_EMBED_CONCRETE_GRADE)
    tol_ft = abs(float(doc.Application.ShortCurveTolerance))

    bar_type, _, _ = _resolve_rebar_bar_type_by_diameter_mm(doc, nominal_diam_mm)
    if bar_type is None:
        TaskDialog.Show(
            u"Armado Columnas",
            u"No se encontró RebarBarType para Ø{:.0f} mm. No se creó armadura longitudinal.".format(
                nominal_diam_mm
            ),
        )
        return

    try:
        hook_mm = hook_length_mm_from_nominal_diameter_mm(nominal_diam_mm, concrete_grade)
        hook_ft = UnitUtils.ConvertToInternalUnits(float(hook_mm), UnitTypeId.Millimeters)
    except Exception:
        hook_ft = float(nominal_diam_mm) / 304.8 * 10.0

    try:
        pata_mm = hook_mm
        pata_ft = hook_ft
    except Exception:
        pata_ft = hook_ft

    # Recopilar sólidos de todas las columnas para prueba de colisión de empotramiento
    opts = _geometry_options_structure_solids()
    all_column_solids = []
    for col in columns_ordered or []:
        for solid in _iter_solids_revit_element(col, opts):
            all_column_solids.append(solid)

    # Construir mapa de columnas contribuyentes por element_id
    col_by_iv = {}
    for col in columns_ordered or []:
        iv = _element_id_iv(col)
        if iv >= 0:
            col_by_iv[iv] = col

    created_rebars = []
    errors = []

    tg = TransactionGroup(doc, u"Arainco: Armado Columnas — Longitudinal")
    tg.Start()
    try:
        t = Transaction(doc, u"Arainco: Armado Columnas — Barras longitudinales")
        t.Start()
        try:
            for entry in fused_world or []:
                try:
                    base_xyz, span_ft, contrib_ids, bar_enum = entry
                except (ValueError, TypeError):
                    continue
                bx = float(base_xyz.X)
                by = float(base_xyz.Y)
                zs = float(base_xyz.Z)

                # Host: primera columna contribuyente disponible
                host = None
                for civ in contrib_ids or []:
                    host = col_by_iv.get(civ)
                    if host is not None:
                        break
                if host is None:
                    continue

                # Selección de bar_type por diámetro de troceo (si está definido)
                seg_diam_mm = nominal_diam_mm
                try:
                    troceo_diams = getattr(wiz, "troceo_diams_by_label", {})
                    if troceo_diams and bar_enum in troceo_diams:
                        seg_diam_mm = float(troceo_diams[bar_enum])
                except Exception:
                    pass
                seg_bar_type = bar_type
                if abs(seg_diam_mm - nominal_diam_mm) > 0.25:
                    seg_bar_type, ok, _ = _resolve_rebar_bar_type_by_diameter_mm(doc, seg_diam_mm)
                    if seg_bar_type is None:
                        seg_bar_type = bar_type

                try:
                    seg_hook_mm = hook_length_mm_from_nominal_diameter_mm(
                        seg_diam_mm, concrete_grade
                    )
                    seg_hook_ft = UnitUtils.ConvertToInternalUnits(seg_hook_mm, UnitTypeId.Millimeters)
                    seg_pata_ft = seg_hook_ft
                except Exception:
                    seg_hook_ft = hook_ft
                    seg_pata_ft = pata_ft

                # Estiramiento +Z (empotramiento superior)
                pt_top = XYZ(bx, by, zs + float(span_ft))
                try:
                    embed_top_collides = embed_stretch_collides_any_column_solids(
                        pt_top, seg_hook_ft, seg_diam_mm, all_column_solids
                    )
                except Exception:
                    embed_top_collides = False

                want_top_pata   = not embed_top_collides
                top_stretch_ft  = seg_hook_ft if embed_top_collides else 0.0
                total_span      = float(span_ft)
                if not embed_top_collides:
                    total_span += seg_hook_ft

                # Estiramiento −Z (fundación o empotramiento)
                found_mm = 0.0
                for civ in contrib_ids or []:
                    col_cand = col_by_iv.get(civ)
                    if col_cand is None:
                        continue
                    try:
                        fm = column_bottom_joined_foundation_stretch_down_mm(doc, col_cand)
                    except Exception:
                        fm = 0.0
                    if fm > found_mm:
                        found_mm = fm
                        break

                want_bot_pata  = False
                bot_stretch_ft = 0.0
                if found_mm > 0.0:
                    try:
                        bot_stretch_ft = UnitUtils.ConvertToInternalUnits(
                            found_mm, UnitTypeId.Millimeters
                        )
                    except Exception:
                        bot_stretch_ft = found_mm / 304.8
                    zs -= bot_stretch_ft
                    total_span += bot_stretch_ft
                else:
                    pt_bot = XYZ(bx, by, zs)
                    try:
                        embed_bot_collides = embed_stretch_collides_any_column_solids(
                            pt_bot, -seg_hook_ft, seg_diam_mm, all_column_solids
                        )
                    except Exception:
                        embed_bot_collides = False
                    want_bot_pata = embed_bot_collides

                comment = u"{}".format(bar_enum)
                rb, did_bot, did_top = create_longitudinal_rebar_with_optional_patas(
                    doc, bx, by, zs, total_span, host, seg_bar_type,
                    comment, want_bot_pata, want_top_pata, seg_pata_ft, contrib_ids,
                )
                if rb is not None:
                    _aplicar_armadura_ubicacion(rb, bar_enum)
                    created_rebars.append(rb)

            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            try:
                errors.append(u"Error en creación longitudinal: {}".format(ex))
            except Exception:
                errors.append(u"Error en creación longitudinal")

        tg.Assimilate()
    except Exception:
        try:
            tg.RollBack()
        except Exception:
            pass
        raise

    if errors:
        from Autodesk.Revit.DB import TaskDialog
        TaskDialog.Show(
            u"Armado Columnas — Longitudinal",
            u"\n".join(errors),
        )
