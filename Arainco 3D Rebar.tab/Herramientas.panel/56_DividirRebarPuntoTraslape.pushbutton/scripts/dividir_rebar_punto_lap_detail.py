# -*- coding: utf-8 -*-
"""
Detail Item de traslape / empalme para barras divididas.

Familia canónica BIMTools: ``EST_D_DEATIL ITEM_EMPALME`` / tipo ``Empalme``
(line-based Detail Component en la vista activa).

Tras colocar cada detail, se intenta acotar la longitud de empalme con el mismo
criterio que Armado Vigas / cabezal muros: referencias nombradas Left/Right del
FamilyInstance + ``NewDimension`` desplazada en el plano de la vista.
"""

from __future__ import print_function

import math
import os
import sys

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    DimensionType,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    Line,
    View,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar

_LAP_DETAIL_DEFAULT_FAMILY_NAME = u"EST_D_DEATIL ITEM_EMPALME"
_LAP_DETAIL_DEFAULT_TYPE_NAME = u"Empalme"
_LAP_DETAIL_ALT_FAMILY_NAMES = (
    u"EST_D_DEATIL ITEM_EMPALME",
    u"EST_D_DETAIL ITEM_EMPALME",
)

# Misma calibración que Armado Vigas (cota de traslape @ 1/50).
_LAP_DIM_SCALE_REFERENCE = 50
_LAP_DIM_OFFSET_MM_AT_REF_SCALE = 450.0
_FIXED_DIMSTYLE_NAME = u"Linear - 2.5mm Arial"
_MAX_DIM_WARNINGS = 8
_MIN_SEG_FT = 1.0 / 304.8

_shared_create_overlap_dim = None
_shared_get_left_right_refs = None
_shared_view_accepts_dim = None
_shared_helpers_tried = False


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _norm_name(s):
    try:
        t = _as_unicode(s or u"")
    except Exception:
        t = u""
    try:
        t = t.replace(u"\u00A0", u" ")
    except Exception:
        pass
    return u" ".join([p for p in t.strip().lower().split() if p])


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _eid_int(element_id):
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except AttributeError:
        try:
            return int(element_id.IntegerValue)
        except Exception:
            return None


def _find_extension_scripts_dir():
    """Localiza ``BIMTools.extension/scripts`` (tiene ``enfierrado_shaft_hashtag.py``)."""
    cursor = os.path.dirname(os.path.abspath(__file__))
    for _ in range(24):
        marker = os.path.join(cursor, u"enfierrado_shaft_hashtag.py")
        if os.path.isfile(marker):
            return cursor
        nested = os.path.join(cursor, u"scripts", u"enfierrado_shaft_hashtag.py")
        if os.path.isfile(nested):
            return os.path.join(cursor, u"scripts")
        parent = os.path.dirname(cursor)
        if parent == cursor:
            break
        cursor = parent
    return None


def _try_load_shared_dim_helpers():
    """
    Preferir helpers canónicos de ``enfierrado_shaft_hashtag`` si el path de
    extensión está disponible; si no, queda el fallback local.
    """
    global _shared_create_overlap_dim
    global _shared_get_left_right_refs
    global _shared_view_accepts_dim
    global _shared_helpers_tried
    if _shared_helpers_tried:
        return
    _shared_helpers_tried = True
    ext_scripts = _find_extension_scripts_dir()
    if ext_scripts and ext_scripts not in sys.path:
        try:
            sys.path.append(ext_scripts)
        except Exception:
            pass
    try:
        from enfierrado_shaft_hashtag import (
            _create_overlap_dimension_from_detail_refs,
            _get_named_left_right_refs_from_detail_instance,
            _view_accepts_overlap_dimension,
        )

        _shared_create_overlap_dim = _create_overlap_dimension_from_detail_refs
        _shared_get_left_right_refs = _get_named_left_right_refs_from_detail_instance
        _shared_view_accepts_dim = _view_accepts_overlap_dimension
    except Exception:
        _shared_create_overlap_dim = None
        _shared_get_left_right_refs = None
        _shared_view_accepts_dim = None


def find_lap_detail_symbol(doc):
    """
    Busca el FamilySymbol del detail de empalme.

    Returns:
        (FamilySymbol|None, mensaje_aviso_o_None)
    """
    if doc is None:
        return None, u"No hay documento activo."
    fam_alt = set()
    for nm in _LAP_DETAIL_ALT_FAMILY_NAMES:
        t = _norm_name(nm)
        if t:
            fam_alt.add(t)
    fam_target = _norm_name(_LAP_DETAIL_DEFAULT_FAMILY_NAME)
    typ_target = _norm_name(_LAP_DETAIL_DEFAULT_TYPE_NAME)
    if fam_target:
        fam_alt.add(fam_target)
    try:
        syms = list(
            FilteredElementCollector(doc)
            .OfClass(FamilySymbol)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
        )
    except Exception:
        syms = []
    if not syms:
        return None, u"No hay Detail Components en el proyecto."
    exact = None
    fam_any = None
    for sym in syms:
        if sym is None:
            continue
        fam = u""
        typ = u""
        try:
            fam = _norm_name(getattr(sym, u"FamilyName", None))
        except Exception:
            fam = u""
        if not fam:
            try:
                if sym.Family is not None:
                    fam = _norm_name(sym.Family.Name)
            except Exception:
                fam = u""
        try:
            typ = _norm_name(getattr(sym, u"Name", None))
        except Exception:
            typ = u""
        if fam not in fam_alt:
            continue
        if typ == typ_target:
            exact = sym
            break
        if fam_any is None:
            fam_any = sym
    if exact is not None:
        return exact, None
    if fam_any is not None:
        return fam_any, (
            u"Tipo exacto «{0}» no encontrado; se usó otro tipo de «{1}»."
            .format(_LAP_DETAIL_DEFAULT_TYPE_NAME, _LAP_DETAIL_DEFAULT_FAMILY_NAME)
        )
    return None, (
        u"No se encontró Detail Component «{0} : {1}»."
        .format(_LAP_DETAIL_DEFAULT_FAMILY_NAME, _LAP_DETAIL_DEFAULT_TYPE_NAME)
    )


def view_accepts_detail_components(view):
    if view is None or not isinstance(view, View):
        return False
    try:
        vt = view.ViewType
        if vt is None:
            return True
        blocked = (
            ViewType.ThreeD,
            ViewType.DrawingSheet,
            ViewType.Schedule,
            ViewType.ProjectBrowser,
            ViewType.SystemBrowser,
        )
        if vt in blocked:
            return False
        name = vt.ToString() or u""
        for b in (u"ThreeD", u"DrawingSheet", u"Schedule", u"Browser"):
            if b in name:
                return False
    except Exception:
        pass
    return True


def _cantidad_posiciones_rebar(rebar):
    best = 1
    for getter in (
        lambda: int(rebar.NumberOfBarPositions),
        lambda: int(rebar.GetNumberOfBarPositions()),
        lambda: int(rebar.Quantity),
    ):
        try:
            n = int(getter())
            if n > best:
                best = n
        except Exception:
            pass
    return best


def _bar_exists_at(rebar, idx):
    i = int(idx)
    try:
        if hasattr(rebar, u"DoesBarExistAtPosition"):
            return bool(rebar.DoesBarExistAtPosition(i))
    except Exception:
        pass
    try:
        if hasattr(rebar, u"IsBarIncluded"):
            return bool(rebar.IsBarIncluded(i))
    except Exception:
        pass
    return True


def _bar_position_transform(rebar, bar_index):
    bi = int(bar_index)
    try:
        return rebar.GetBarPositionTransform(bi)
    except Exception:
        pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, u"GetBarPositionTransform"):
            return acc.GetBarPositionTransform(bi)
    except Exception:
        pass
    return None


def _map_point_to_bar_index(rebar, point, from_idx, to_idx):
    """Mapea un XYZ de la posición ``from_idx`` a la posición ``to_idx`` del set."""
    if point is None:
        return None
    fi, ti = int(from_idx), int(to_idx)
    if fi == ti:
        return point
    t0 = _bar_position_transform(rebar, fi)
    t1 = _bar_position_transform(rebar, ti)
    if t0 is None or t1 is None:
        return point
    try:
        return t1.OfPoint(t0.Inverse.OfPoint(point))
    except Exception:
        try:
            return point + (t1.Origin - t0.Origin)
        except Exception:
            return point


def _presentation_mode_name(rebar, view):
    if rebar is None or view is None:
        return u""
    try:
        mode = rebar.GetPresentationMode(view)
    except Exception:
        return u""
    try:
        return _as_unicode(mode.ToString() or u"")
    except Exception:
        return u""


def middle_bar_index(rebar):
    """Índice visible con Show Middle (NumberOfBarPositions // 2, existente)."""
    npos = _cantidad_posiciones_rebar(rebar)
    if npos <= 1:
        return 0
    mid = int(npos / 2)
    if _bar_exists_at(rebar, mid):
        return mid
    for delta in range(1, npos):
        for idx in (mid - delta, mid + delta):
            if 0 <= idx < npos and _bar_exists_at(rebar, idx):
                return int(idx)
    return mid


def represented_bar_indices(rebar, view):
    """
    Índices de barra representados en ``view`` según PresentationMode.

    - Middle → barra central
    - FirstLast → primera y última existentes
    - Select → no ocultas
    - All / sin vista → solo índice 0 (un detail; evita N×cortes en sets grandes)
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return [0]
    n = _cantidad_posiciones_rebar(rebar)
    if n <= 1:
        return [0]

    mode_name = _presentation_mode_name(rebar, view)
    idxs = []

    if mode_name == u"Middle" or u"Middle" in mode_name:
        idxs = [middle_bar_index(rebar)]
    elif mode_name == u"FirstLast" or u"FirstLast" in mode_name:
        for i in (0, n - 1):
            if _bar_exists_at(rebar, i) and i not in idxs:
                idxs.append(int(i))
    elif mode_name == u"Select" or u"Select" in mode_name:
        for i in range(n):
            hidden = False
            try:
                hidden = bool(rebar.IsBarHidden(view, int(i)))
            except Exception:
                hidden = False
            if not hidden and _bar_exists_at(rebar, i):
                idxs.append(int(i))
    else:
        # All u otros: un solo detail en la primera barra existente
        for i in range(n):
            if _bar_exists_at(rebar, i):
                idxs = [int(i)]
                break
        if not idxs:
            idxs = [0]

    return idxs or [0]


def expand_lap_segments_for_presentation(
    rebar, view, segments, source_bar_index=0
):
    """
    Duplica/desplaza segmentos de empalme desde ``source_bar_index`` (centerline
    usada al construir) hacia las barras representadas en la vista.
    """
    base = list(segments or [])
    if not base or rebar is None:
        return base
    indices = represented_bar_indices(rebar, view)
    src = int(source_bar_index)
    if len(indices) == 1 and int(indices[0]) == src:
        return base

    out = []
    for bi in indices:
        bi = int(bi)
        for seg in base:
            if not seg or len(seg) < 2:
                continue
            p0 = _map_point_to_bar_index(rebar, seg[0], src, bi)
            p1 = _map_point_to_bar_index(rebar, seg[1], src, bi)
            if p0 is not None and p1 is not None:
                out.append((p0, p1))
    return out if out else base


def _view_accepts_overlap_dimension_local(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        vt = view.ViewType
        if vt == ViewType.ThreeD:
            return False
        name = (vt.ToString() if vt is not None else u"") or u""
        if u"ThreeD" in name:
            return False
    except Exception:
        pass
    return True


def _project_point_to_view_plane(view, p):
    if view is None or p is None:
        return p
    try:
        origin = view.Origin
        normal = view.ViewDirection
        if origin is None or normal is None:
            return p
        n = normal.Normalize()
        v = p.Subtract(origin)
        dist = v.DotProduct(n)
        return p.Subtract(n.Multiply(dist))
    except Exception:
        return p


def _view_scale_denominator(view):
    if view is None:
        return _LAP_DIM_SCALE_REFERENCE
    try:
        s = int(view.Scale)
        if s > 0:
            return s
    except Exception:
        pass
    return _LAP_DIM_SCALE_REFERENCE


def _lap_dim_offset_mm(view):
    ratio = float(_view_scale_denominator(view)) / float(_LAP_DIM_SCALE_REFERENCE)
    return float(_LAP_DIM_OFFSET_MM_AT_REF_SCALE) * ratio


def _unit_3d(v):
    if v is None:
        return None
    try:
        ln = float(v.GetLength())
    except Exception:
        try:
            ln = math.sqrt(
                float(v.X) * float(v.X)
                + float(v.Y) * float(v.Y)
                + float(v.Z) * float(v.Z)
            )
        except Exception:
            return None
    if ln < 1e-12:
        return None
    try:
        return v.Normalize()
    except Exception:
        return XYZ(float(v.X) / ln, float(v.Y) / ln, float(v.Z) / ln)


def _get_named_left_right_refs_local(detail_inst):
    if detail_inst is None:
        return None, None, u"No existe instancia de detail para extraer referencias."
    left_candidates = (u"Left", u"LEFT", u"Izquierda", u"IZQUIERDA")
    right_candidates = (u"Right", u"RIGHT", u"Derecha", u"DERECHA")
    ref_left = None
    ref_right = None
    for nm in left_candidates:
        try:
            r = detail_inst.GetReferenceByName(nm)
        except Exception:
            r = None
        if r is not None:
            ref_left = r
            break
    for nm in right_candidates:
        try:
            r = detail_inst.GetReferenceByName(nm)
        except Exception:
            r = None
        if r is not None:
            ref_right = r
            break
    if ref_left is None or ref_right is None:
        return None, None, u"La familia de empalme no expone referencias nombradas Left/Right."
    return ref_left, ref_right, None


def _try_apply_fixed_dimension_type(doc, dim):
    if doc is None or dim is None:
        return
    target = _norm_name(_FIXED_DIMSTYLE_NAME)
    found_id = None
    try:
        for dt in FilteredElementCollector(doc).OfClass(DimensionType):
            try:
                nm = _norm_name(getattr(dt, u"Name", None))
            except Exception:
                nm = u""
            if nm == target:
                found_id = dt.Id
                break
    except Exception:
        found_id = None
    if found_id is None:
        return
    try:
        dim.ChangeTypeId(found_id)
    except Exception:
        pass


def _dim_line_endpoints_view_plane(
    view, lap_start, lap_end, line_offset_mm, prefer_dir=None
):
    """Extremos de cota paralelos al empalme, desplazados en el plano de la vista."""
    if lap_start is None or lap_end is None or view is None:
        return None, None
    try:
        n = view.ViewDirection
        if n is None or n.GetLength() < 1e-12:
            return None, None
        n = n.Normalize()
    except Exception:
        return None, None
    u_raw = _unit_3d(lap_end.Subtract(lap_start))
    if u_raw is None:
        return None, None
    try:
        du = float(u_raw.DotProduct(n))
        u_in = u_raw.Subtract(n.Multiply(du))
        if u_in.GetLength() < 1e-12:
            return None, None
        u_in = u_in.Normalize()
    except Exception:
        return None, None
    try:
        tdir = n.CrossProduct(u_in)
        if tdir.GetLength() < 1e-12:
            tdir = u_in.CrossProduct(n)
        if tdir.GetLength() < 1e-12:
            return None, None
        tdir = tdir.Normalize()
    except Exception:
        return None, None
    # Preferir un lado (p. ej. Up de la vista = sobre barras horizontales).
    if prefer_dir is not None:
        try:
            pref = _unit_3d(prefer_dir)
            if pref is not None:
                # Quitar componente paralela al eje del empalme
                pu = float(pref.DotProduct(u_in))
                pref_perp = pref.Subtract(u_in.Multiply(pu))
                # Quitar componente hacia la cámara
                pn = float(pref_perp.DotProduct(n))
                pref_perp = pref_perp.Subtract(n.Multiply(pn))
                if pref_perp.GetLength() > 1e-12:
                    if float(tdir.DotProduct(pref_perp)) < 0.0:
                        tdir = tdir.Negate()
        except Exception:
            pass
    try:
        off_ft = _mm_to_ft(float(line_offset_mm))
    except Exception:
        off_ft = _mm_to_ft(450.0)
    try:
        out = tdir.Multiply(off_ft)
        a = lap_start.Add(out)
        b = lap_end.Add(out)
        return a, b
    except Exception:
        return None, None


def _create_overlap_dimension_local(
    doc,
    view,
    ref_left,
    ref_right,
    lap_start,
    lap_end,
    line_offset_mm,
    prefer_dir=None,
):
    if (
        doc is None
        or view is None
        or ref_left is None
        or ref_right is None
        or lap_start is None
        or lap_end is None
    ):
        return False, u"Parámetros incompletos para cota desde referencias del detail.", None
    a, b = _dim_line_endpoints_view_plane(
        view, lap_start, lap_end, line_offset_mm, prefer_dir=prefer_dir
    )
    if a is None or b is None:
        return False, u"No se pudo construir la línea de cota en el plano de la vista.", None
    try:
        if a.DistanceTo(b) <= _MIN_SEG_FT:
            return False, u"Segmento de cota demasiado corto en la vista.", None
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ReferenceArray

        dim_line = Line.CreateBound(a, b)
        ra = ReferenceArray()
        ra.Append(ref_left)
        ra.Append(ref_right)
        dim = doc.Create.NewDimension(view, dim_line, ra)
    except Exception:
        dim = None
    if dim is None:
        return False, u"Revit no permitió crear cota de traslapo desde referencias del detail.", None
    try:
        nrefs = int(dim.References.Size)
    except Exception:
        nrefs = None
    if nrefs is not None and nrefs != 2:
        try:
            doc.Delete(dim.Id)
        except Exception:
            pass
        return False, u"Cota de traslapo inválida (referencias extra).", None
    _try_apply_fixed_dimension_type(doc, dim)
    return True, None, {u"dim_id": _eid_int(dim.Id)}


def create_lap_dimension_for_detail(
    doc, view, lap_inst, p0, p1, prefer_above=False
):
    """
    Cota lineal entre refs Left/Right del detail de empalme.

    ``prefer_above=True``: desplaza la cota hacia ``view.UpDirection``
    (sobre barras horizontales en alzado).

    Returns:
        (ElementId|None, error_unicode|None)
    """
    _try_load_shared_dim_helpers()
    if lap_inst is None or view is None or doc is None:
        return None, None

    accepts = _shared_view_accepts_dim or _view_accepts_overlap_dimension_local
    if not accepts(view):
        return None, u"La vista no admite cotas de traslapo."

    get_refs = _shared_get_left_right_refs or _get_named_left_right_refs_local
    ref_l, ref_r, ref_err = get_refs(lap_inst)
    if ref_l is None or ref_r is None:
        return None, ref_err

    pa = _project_point_to_view_plane(view, p0)
    pb = _project_point_to_view_plane(view, p1)
    if pa is None or pb is None:
        return None, u"Extremos de empalme inválidos para acotar."
    try:
        if pa.DistanceTo(pb) <= _MIN_SEG_FT:
            return None, u"Segmento de traslape demasiado corto para acotar."
    except Exception:
        return None, u"Segmento de traslape inválido para acotar."

    axis_u = _unit_3d(pb.Subtract(pa))
    offset_mm = _lap_dim_offset_mm(view)

    prefer_dir = None
    inward_3d = None
    if prefer_above:
        try:
            prefer_dir = view.UpDirection
            if prefer_dir is not None and prefer_dir.GetLength() > 1e-12:
                # Interior del muro bajo el coronamiento → cota hacia afuera = Up
                inward_3d = prefer_dir.Negate()
        except Exception:
            prefer_dir = None
            inward_3d = None

    if _shared_create_overlap_dim is not None:
        ok_dim, msg_dim, dim_data = _shared_create_overlap_dim(
            doc,
            view,
            ref_l,
            ref_r,
            pa,
            pb,
            axis_u,
            lateral_hint=None,
            line_offset_mm=offset_mm,
            inward_dir_xy=None,
            inward_dir_3d=inward_3d,
            use_view_plane_dim_line=True,
            flip_dimension_side=False,
        )
    else:
        ok_dim, msg_dim, dim_data = _create_overlap_dimension_local(
            doc,
            view,
            ref_l,
            ref_r,
            pa,
            pb,
            offset_mm,
            prefer_dir=prefer_dir,
        )

    if not ok_dim:
        return None, msg_dim
    try:
        if dim_data and dim_data.get(u"dim_id") is not None:
            return ElementId(int(dim_data[u"dim_id"])), None
    except Exception:
        pass
    return None, msg_dim


def place_line_based_lap_detail(doc, view, family_symbol, p0, p1):
    """
    Coloca Detail Component line-based entre ``p0`` y ``p1`` (XYZ modelo).

    Returns:
        (ok, err, instance|None)
    """
    if doc is None or view is None or family_symbol is None:
        return False, u"Parámetros incompletos para detail de traslape.", None
    if p0 is None or p1 is None:
        return False, u"Segmento de traslape inválido.", None
    try:
        p0p = _project_point_to_view_plane(view, p0)
        p1p = _project_point_to_view_plane(view, p1)
        if p0p is not None and p1p is not None:
            if p0p.DistanceTo(p1p) > _MIN_SEG_FT:
                p0, p1 = p0p, p1p
    except Exception:
        pass
    try:
        if p0.DistanceTo(p1) <= _MIN_SEG_FT:
            return False, u"Segmento de traslape demasiado corto en la vista.", None
        ln = Line.CreateBound(p0, p1)
    except Exception as ex:
        return False, _as_unicode(ex), None
    try:
        if not bool(getattr(family_symbol, u"IsActive", True)):
            family_symbol.Activate()
            try:
                doc.Regenerate()
            except Exception:
                pass
    except Exception:
        pass
    try:
        inst = doc.Create.NewFamilyInstance(ln, family_symbol, view)
        return True, None, inst
    except Exception as ex:
        # Respaldo: XY + Z del plano de vista (API / vistas antiguas).
        try:
            origin = view.Origin
            z = float(origin.Z) if origin is not None else float(p0.Z)
            p0f = XYZ(float(p0.X), float(p0.Y), z)
            p1f = XYZ(float(p1.X), float(p1.Y), z)
            if p0f.DistanceTo(p1f) <= _MIN_SEG_FT:
                return False, _as_unicode(ex), None
            ln_p = Line.CreateBound(p0f, p1f)
            inst = doc.Create.NewFamilyInstance(ln_p, family_symbol, view)
            return True, None, inst
        except Exception:
            return False, _as_unicode(ex), None


def place_lap_details_for_segments(
    doc, view, segments, place_dims=True, prefer_dims_above=False
):
    """
    Coloca un detail por cada segmento ``(p0, p1)`` XYZ y acota cada uno.

    ``place_dims=False``: solo Detail Item (sin cotas).
    ``prefer_dims_above=True``: cotas hacia Up de la vista (sobre barras).

    Returns:
        dict: n_ok, n_fail, n_dims_ok, n_dims_fail, ids, dim_ids, errors, warning
    """
    result = {
        u"n_ok": 0,
        u"n_fail": 0,
        u"n_dims_ok": 0,
        u"n_dims_fail": 0,
        u"ids": [],
        u"dim_ids": [],
        u"errors": [],
        u"warning": None,
    }
    segments = list(segments or [])
    if doc is None or not segments:
        return result
    if not view_accepts_detail_components(view):
        result[u"n_fail"] = len(segments)
        result[u"errors"].append(
            u"La vista activa no admite Detail Components "
            u"(use planta, alzado o sección; no 3D ni lámina)."
        )
        return result
    sym, warn = find_lap_detail_symbol(doc)
    result[u"warning"] = warn
    if sym is None:
        result[u"n_fail"] = len(segments)
        if warn:
            result[u"errors"].append(warn)
        return result

    aviso_refs = None
    do_dims = bool(place_dims)
    for seg in segments:
        if not seg or len(seg) < 2:
            result[u"n_fail"] += 1
            continue
        p0, p1 = seg[0], seg[1]
        ok, err, inst = place_line_based_lap_detail(doc, view, sym, p0, p1)
        if not (ok and inst is not None):
            result[u"n_fail"] += 1
            if err:
                result[u"errors"].append(err)
            continue

        result[u"n_ok"] += 1
        try:
            result[u"ids"].append(inst.Id)
        except Exception:
            pass

        if not do_dims:
            continue

        # Acotar sin abortar la división si la cota falla.
        try:
            dim_eid, dim_err = create_lap_dimension_for_detail(
                doc,
                view,
                inst,
                p0,
                p1,
                prefer_above=bool(prefer_dims_above),
            )
        except Exception as ex_dim:
            dim_eid, dim_err = None, _as_unicode(ex_dim)
        if dim_eid is not None:
            result[u"n_dims_ok"] += 1
            try:
                result[u"dim_ids"].append(dim_eid)
            except Exception:
                pass
        elif dim_err:
            result[u"n_dims_fail"] += 1
            if u"Left/Right" in (dim_err or u"") and aviso_refs is None:
                aviso_refs = dim_err
            elif len(result[u"errors"]) < _MAX_DIM_WARNINGS:
                result[u"errors"].append(u"Cota empalme: {0}".format(dim_err))

    if aviso_refs and len(result[u"errors"]) < _MAX_DIM_WARNINGS:
        result[u"errors"].append(aviso_refs)
    return result


def build_lap_segments_from_cuts(
    curves, cuts_ft, lap_ft, point_at_distance_fn, lap_mode=None
):
    """
    Segmentos de empalme en XYZ según ``lap_mode`` (default ±L/2 centrado).
    """
    out = []
    if not curves or not cuts_ft or lap_ft is None or float(lap_ft) <= 0:
        return out
    if point_at_distance_fn is None:
        return out
    try:
        from dividir_rebar_punto_geom import lap_zone_around_cut
    except Exception:
        lap_zone_around_cut = None
    for cut in cuts_ft:
        try:
            c = float(cut)
        except Exception:
            continue
        if lap_zone_around_cut is not None:
            a, b = lap_zone_around_cut(c, lap_ft, lap_mode)
        else:
            half = 0.5 * float(lap_ft)
            a, b = c - half, c + half
        p0 = point_at_distance_fn(curves, a)
        p1 = point_at_distance_fn(curves, b)
        if p0 is not None and p1 is not None:
            out.append((p0, p1))
    return out
