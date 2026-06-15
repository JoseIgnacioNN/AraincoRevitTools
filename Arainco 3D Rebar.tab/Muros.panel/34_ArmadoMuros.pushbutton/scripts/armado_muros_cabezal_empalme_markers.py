# -*- coding: utf-8 -*-
"""
Detail Items de empalme en cabezal muros (y cotas de traslape en capas 0 y 1).

No se registra vínculo en Extensible Storage ni en el DMU global de empalmes
(``lap_detail_updater_dmu``): los símbolos quedan fijos tras la transacción de armado.

Las cotas de traslape se crean solo en capas de índice 0 y 1; en capas superiores
solo se coloca el Detail Item (sin cota). En vista 1/50: offset 300 mm (capa 0) y
450 mm (capa 1) respecto al empalme, hacia fuera del muro; otras escalas, proporcional.
"""

from __future__ import print_function

import os
import sys

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    ElementId,
    FamilySymbol,
    Transaction,
    UnitTypeId,
    UnitUtils,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarBarType

_MAX_WARNINGS = 8


def _ensure_import_paths():
    here = os.path.dirname(os.path.abspath(__file__))
    try:
        import bootstrap_paths
        return bootstrap_paths.pin_local_scripts_first()
    except Exception:
        if here and here not in sys.path:
            sys.path.insert(0, here)
        return here


_ensure_import_paths()

try:
    from armado_muros_lap_detail_shared import _find_fixed_lap_detail_symbol_id
except Exception:
    _find_fixed_lap_detail_symbol_id = None

try:
    from armado_muros_rebar_params import stamp_armadura_conjunto_guid
except Exception:
    stamp_armadura_conjunto_guid = None

try:
    from enfierrado_shaft_hashtag import (
        _create_overlap_dimension_from_detail_refs,
        _get_named_left_right_refs_from_detail_instance,
        _place_line_based_detail_component,
        _view_accepts_overlap_dimension,
    )
except Exception:
    _create_overlap_dimension_from_detail_refs = None
    _get_named_left_right_refs_from_detail_instance = None
    _place_line_based_detail_component = None
    _view_accepts_overlap_dimension = None

# Distancia modelo cota ↔ empalme (calibrado en vista 1/50; capa 1 +150 mm alinea con paso ~150 mm).
_LAP_DIM_SCALE_REFERENCE = 50
_LAP_DIM_OFFSET_MM_AT_REF_SCALE = {0: 300.0, 1: 450.0}
# Cotas de traslape solo en las dos primeras capas (índices 0 y 1); el resto solo detail.
_LAP_DIM_LAYER_INDICES = (0, 1)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _bar_nominal_diameter_mm(bar_type):
    """Ø nominal (mm) del ``RebarBarType`` — mismo criterio que troceo cabezal / shaft."""
    if bar_type is None or not isinstance(bar_type, RebarBarType):
        return None
    try:
        d = bar_type.BarNominalDiameter
        return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
    except Exception:
        pass
    try:
        d = bar_type.BarModelDiameter
        return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
    except Exception:
        return None


def _lap_length_mm_from_bar_type(bar_type, concrete_grade=None):
    d_mm = _bar_nominal_diameter_mm(bar_type)
    if d_mm is None or float(d_mm) <= 0.0:
        return None
    try:
        from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm

        lap_mm = traslape_mm_from_nominal_diameter_mm(d_mm, concrete_grade)
        if lap_mm is not None and float(lap_mm) > 50.0:
            return float(lap_mm)
    except Exception:
        pass
    return None


def _lap_length_mm_from_rebar_element(rebar, doc, concrete_grade=None):
    if rebar is None or doc is None or not isinstance(rebar, Rebar):
        return None
    try:
        bt = doc.GetElement(rebar.GetTypeId())
    except Exception:
        bt = None
    return _lap_length_mm_from_bar_type(bt, concrete_grade)


def _lap_length_ft_from_rebar(doc, rebar_id, bar_type_hint=None, concrete_grade=None):
    """
    Longitud del detail line-based de empalme (pies).

    Prioridad: ``bar_type_hint`` del ``seg_job`` (ø usado al trocear). Respaldo: tipo
    de la ``Rebar`` ya creada. Último recurso: 860 mm (ø12 tabla base).
    """
    lap_mm = _lap_length_mm_from_bar_type(bar_type_hint, concrete_grade)
    if lap_mm is None and doc is not None and rebar_id is not None:
        try:
            rb = doc.GetElement(rebar_id)
        except Exception:
            rb = None
        lap_mm = _lap_length_mm_from_rebar_element(rb, doc, concrete_grade)
    if lap_mm is not None:
        return _mm_to_internal(lap_mm)
    return _mm_to_internal(860.0)


def _lap_length_ft_for_empalme_pair(
    document,
    rebar_id_a,
    rebar_id_b,
    bar_type_a=None,
    bar_type_b=None,
    concrete_grade=None,
):
    """Longitud del detail de traslape: mayor L entre las dos barras (ø mayor)."""
    la = _lap_length_ft_from_rebar(
        document, rebar_id_a, bar_type_hint=bar_type_a, concrete_grade=concrete_grade,
    )
    lb = _lap_length_ft_from_rebar(
        document, rebar_id_b, bar_type_hint=bar_type_b, concrete_grade=concrete_grade,
    )
    return la if la >= lb else lb


def _build_empalme_pairs_from_seg_jobs(seg_jobs):
    """Agrupa segmentos troceados y empareja rebars consecutivos (ra, rb)."""
    groups = {}
    for sj in seg_jobs or []:
        if sj is None:
            continue
        try:
            n_seg = int(sj.get(u"n_segments", 1) or 1)
        except Exception:
            n_seg = 1
        if n_seg < 2:
            continue
        fk = sj.get(u"fusion_key")
        if fk is None:
            continue
        rid = sj.get(u"rebar_id")
        if rid is None:
            continue
        groups.setdefault(fk, []).append(sj)

    pairs = []
    for jobs in groups.values():
        jobs.sort(key=lambda j: int(j.get(u"seg_index", 0) or 0))
        for i in range(len(jobs) - 1):
            j_lo = jobs[i]
            j_hi = jobs[i + 1]
            ra = j_lo.get(u"rebar_id")
            rb = j_hi.get(u"rebar_id")
            if ra is None or rb is None:
                continue
            try:
                z_joint = float(j_hi.get(u"zs", 0.0))
            except Exception:
                continue
            pairs.append({
                u"ra": ra,
                u"rb": rb,
                u"bar_type_a": j_lo.get(u"bar_type"),
                u"bar_type_b": j_hi.get(u"bar_type"),
                u"bx": float(j_lo.get(u"bx", 0.0)),
                u"by": float(j_lo.get(u"by", 0.0)),
                u"z_joint": z_joint,
                u"wall": j_lo.get(u"wall"),
                u"extremo": j_lo.get(u"extremo"),
                u"layer_index": int(j_lo.get(u"layer_index", 0) or 0),
                u"vec_long": j_lo.get(u"vec_long"),
                u"normal_muro": j_lo.get(u"normal_muro"),
            })
    return pairs


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


def _lap_dim_offset_mm(layer_index, view):
    """Offset modelo (mm) cota ↔ empalme; proporcional a 1/50 en otras escalas de vista."""
    try:
        li = int(layer_index)
    except Exception:
        return None
    base = _LAP_DIM_OFFSET_MM_AT_REF_SCALE.get(li)
    if base is None:
        return None
    scale = _view_scale_denominator(view)
    ratio = float(scale) / float(_LAP_DIM_SCALE_REFERENCE)
    return float(base) * ratio


def _xy_toward_wall_interior(wall, pt):
    """Unitario en planta desde el empalme hacia el interior del host (centroide XY)."""
    if wall is None or pt is None:
        return None
    try:
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return None
        cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
        cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
        v = XYZ(cx - float(pt.X), cy - float(pt.Y), 0.0)
        if float(v.GetLength()) < 1e-9:
            return None
        return v.Normalize()
    except Exception:
        return None


def _inward_dirs_for_cabezal_lap(wall, extremo, lap_pt, normal_muro=None, vec_long=None):
    """
    Dirección hacia el interior del muro; la cota se desplaza al lado opuesto (fuera del host).

    Combina centroide del muro, normal de cara y eje longitudinal del cabezal.
    """
    inward_xy = _xy_toward_wall_interior(wall, lap_pt)
    inward_3d = None

    if normal_muro is not None:
        try:
            nm = normal_muro.Normalize()
            # Orientation Revit ≈ exterior; interior del espesor ≈ −normal.
            cand = nm.Negate()
            if inward_xy is not None:
                nxy = XYZ(float(nm.X), float(nm.Y), 0.0)
                if float(nxy.GetLength()) > 1e-12:
                    nxy = nxy.Normalize()
                    if float(inward_xy.DotProduct(nxy)) > 0.0:
                        cand = nm
            inward_3d = cand
        except Exception:
            inward_3d = None

    if inward_3d is None and vec_long is not None:
        try:
            vl = vec_long.Normalize()
            if (extremo or u"") == u"fin":
                vl = vl.Negate()
            if float(vl.GetLength()) > 1e-12:
                inward_3d = vl
                if inward_xy is None:
                    inward_xy = XYZ(float(vl.X), float(vl.Y), 0.0)
                    if float(inward_xy.GetLength()) > 1e-12:
                        inward_xy = inward_xy.Normalize()
                    else:
                        inward_xy = None
        except Exception:
            pass

    if inward_3d is None and inward_xy is not None:
        inward_3d = XYZ(float(inward_xy.X), float(inward_xy.Y), 0.0)

    if inward_3d is None and wall is not None:
        try:
            orient = wall.Orientation
            if orient is not None and float(orient.GetLength()) > 1e-9:
                inward_3d = orient.Normalize().Negate()
                inward_xy = XYZ(float(inward_3d.X), float(inward_3d.Y), 0.0)
                if float(inward_xy.GetLength()) > 1e-12:
                    inward_xy = inward_xy.Normalize()
                else:
                    inward_xy = None
        except Exception:
            pass

    return inward_xy, inward_3d


def _axis_u_vertical(pa, pb):
    try:
        dv = pb.Subtract(pa)
        if dv.GetLength() > 1e-9:
            return dv.Normalize()
    except Exception:
        pass
    return XYZ.BasisZ


def _create_lap_dimension_for_detail(document, view, lap_inst, pa, pb, spec):
    if (
        lap_inst is None
        or _get_named_left_right_refs_from_detail_instance is None
        or _create_overlap_dimension_from_detail_refs is None
        or _view_accepts_overlap_dimension is None
        or not _view_accepts_overlap_dimension(view)
    ):
        return None, None

    ref_l, ref_r, ref_err = _get_named_left_right_refs_from_detail_instance(lap_inst)
    if ref_l is None or ref_r is None:
        return None, ref_err

    layer_index = int(spec.get(u"layer_index", 0) or 0)
    line_offset_mm = _lap_dim_offset_mm(layer_index, view)
    if line_offset_mm is None:
        return None, u"Capa {0} sin offset de cota definido.".format(layer_index)

    axis_u = _axis_u_vertical(pa, pb)
    lap_pt = XYZ(
        0.5 * (float(pa.X) + float(pb.X)),
        0.5 * (float(pa.Y) + float(pb.Y)),
        0.5 * (float(pa.Z) + float(pb.Z)),
    )
    inward_xy, inward_3d = _inward_dirs_for_cabezal_lap(
        spec.get(u"wall"),
        spec.get(u"extremo"),
        lap_pt,
        normal_muro=spec.get(u"normal_muro"),
        vec_long=spec.get(u"vec_long"),
    )
    lateral_hint = spec.get(u"vec_long")
    ok_dm, msg_dm, dim_data = _create_overlap_dimension_from_detail_refs(
        document,
        view,
        ref_l,
        ref_r,
        pa,
        pb,
        axis_u,
        lateral_hint=lateral_hint,
        line_offset_mm=line_offset_mm,
        inward_dir_xy=inward_xy,
        inward_dir_3d=inward_3d,
        use_view_plane_dim_line=True,
        flip_dimension_side=False,
    )
    if not ok_dm:
        return None, msg_dm
    try:
        if dim_data and dim_data.get(u"dim_id") is not None:
            return ElementId(int(dim_data[u"dim_id"])), None
    except Exception:
        pass
    return None, msg_dm


def _view_accepts_detail(view):
    if view is None:
        return False
    try:
        from enfierrado_shaft_hashtag import _vista_admite_detail_components

        return bool(_vista_admite_detail_components(view))
    except Exception:
        pass
    try:
        vt = view.ViewType
        if vt is None:
            return True
        name = vt.ToString() or u""
        blocked = (u"ThreeD", u"DrawingSheet", u"ProjectBrowser", u"SystemBrowser")
        for b in blocked:
            if b in name:
                return False
        return True
    except Exception:
        return True


def colocar_marcadores_empalme_cabezal(document, view, seg_jobs):
    """
    Coloca Detail Items line-based en juntas de troceo y los vincula a las dos barras.

    ``seg_jobs``: lista de jobs longitudinales ya creados (con ``rebar_id``).
    """
    result = {
        u"n_ok": 0,
        u"n_fail": 0,
        u"n_dims_ok": 0,
        u"n_dims_fail": 0,
        u"messages": [],
    }
    if (
        document is None
        or view is None
        or _place_line_based_detail_component is None
        or _find_fixed_lap_detail_symbol_id is None
    ):
        result[u"messages"].append(u"Empalme: módulo de marcadores no disponible.")
        return result

    pairs = _build_empalme_pairs_from_seg_jobs(seg_jobs)
    if not pairs:
        return result

    if not _view_accepts_detail(view):
        result[u"messages"].append(
            u"Empalme: la vista activa no admite detail components "
            u"(use planta, alzado o sección; no plantilla ni 3D).",
        )
        result[u"n_fail"] = len(pairs)
        return result

    sid, sym_err = _find_fixed_lap_detail_symbol_id(document)
    if sid is None:
        if sym_err:
            result[u"messages"].append(sym_err)
        result[u"n_fail"] = len(pairs)
        return result

    lap_sym = document.GetElement(sid)
    if lap_sym is None or not isinstance(lap_sym, FamilySymbol):
        result[u"messages"].append(u"Empalme: símbolo Detail Item no válido.")
        result[u"n_fail"] = len(pairs)
        return result

    t = Transaction(document, u"Arainco: Cabezal muros — marcadores empalme")
    t.Start()
    try:
        for spec in pairs:
            ra = spec[u"ra"]
            rb = spec[u"rb"]
            z_joint = float(spec[u"z_joint"])
            bx = float(spec[u"bx"])
            by = float(spec[u"by"])
            dz = _lap_length_ft_for_empalme_pair(
                document,
                ra,
                rb,
                bar_type_a=spec.get(u"bar_type_a"),
                bar_type_b=spec.get(u"bar_type_b"),
            )
            if dz < _mm_to_internal(50.0):
                dz = _mm_to_internal(50.0)
            pa = XYZ(bx, by, z_joint)
            pb = XYZ(bx, by, z_joint + dz)

            ok_d, err_d, lap_inst = _place_line_based_detail_component(
                document, view, lap_sym, pa, pb,
            )
            if not ok_d or lap_inst is None:
                result[u"n_fail"] += 1
                if len(result[u"messages"]) < _MAX_WARNINGS:
                    result[u"messages"].append(
                        err_d or u"Empalme: no se pudo colocar el detail.",
                    )
                continue

            if stamp_armadura_conjunto_guid is not None:
                try:
                    stamp_armadura_conjunto_guid(lap_inst)
                except Exception:
                    pass

            dim_eid = None
            layer_index = int(spec.get(u"layer_index", 0) or 0)
            if layer_index in _LAP_DIM_LAYER_INDICES:
                dim_eid, dim_err = _create_lap_dimension_for_detail(
                    document, view, lap_inst, pa, pb, spec,
                )
                if dim_eid is not None:
                    result[u"n_dims_ok"] += 1
                elif dim_err:
                    result[u"n_dims_fail"] += 1
                    if len(result[u"messages"]) < _MAX_WARNINGS:
                        result[u"messages"].append(
                            u"Cota empalme (capa {0}): {1}".format(
                                layer_index, dim_err,
                            ),
                        )

            result[u"n_ok"] += 1
        t.Commit()
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        result[u"messages"].append(u"Empalme: {0}".format(ex))
        result[u"n_fail"] = max(int(result.get(u"n_fail", 0)), len(pairs))
    return result
