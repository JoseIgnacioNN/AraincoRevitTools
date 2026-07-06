# -*- coding: utf-8 -*-
"""
Etiquetado de barras longitudinales (Armado vigas).

Familia ``EST_A_STRUCTURAL REBAR TAG_HORIZONTAL``: el tipo de etiqueta se resuelve
por nombre del ``RebarShape`` de cada barra (p. ej. «03», «02»), igual que en
``enfierrado_shaft_hashtag`` / Armado muros cabezal.

Las cabeceras superiores se alinean hacia arriba en la vista; las inferiores hacia
abajo (dirección opuesta), con margen exterior al hormigón y sin solapar otras
etiquetas (p. ej. confinamiento).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, IndependentTag, XYZ

LONGITUDINAL_REBAR_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_HORIZONTAL"

_TAG_ALIGN_EXTRA_FT = 0.04
_TAG_OUTSIDE_FT = 0.35
_TAG_INF_CLEARANCE_MM = 80.0
_TAG_NUDGE_STEP_MM = 28.0
_TAG_NUDGE_MAX_STEPS = 48

try:
    from enfierrado_shaft_hashtag import etiquetar_rebars_creados_en_vista
except Exception:
    etiquetar_rebars_creados_en_vista = None

try:
    from geometria_viga_cara_superior_detalle import (
        _alinear_etiquetas_rebar_mismo_lote,
        _collect_independent_tags_for_rebar_lote,
        _framing_host_desde_lote_rebars,
        _proyectar_vector_en_plano_perp_normal,
        _separar_etiquetas_rebar_solapadas_lote,
        _tags_overlap_with_clearance,
        _vec_dot,
        _vec_normalize_xyz,
    )
except Exception:
    _alinear_etiquetas_rebar_mismo_lote = None
    _collect_independent_tags_for_rebar_lote = None
    _framing_host_desde_lote_rebars = None
    _proyectar_vector_en_plano_perp_normal = None
    _separar_etiquetas_rebar_solapadas_lote = None
    _tags_overlap_with_clearance = None
    _vec_dot = None
    _vec_normalize_xyz = None


def _rebar_element_ids(rebars):
    ids = []
    seen = set()
    for rb in rebars or []:
        if rb is None:
            continue
        try:
            rid = rb.Id
        except Exception:
            rid = rb
        if rid is None:
            continue
        try:
            key = int(rid.IntegerValue)
        except Exception:
            key = rid
        if key in seen:
            continue
        seen.add(key)
        ids.append(rid)
    return ids


def _mm_to_ft(mm):
    try:
        return float(mm) / 304.8
    except Exception:
        return 0.0


def _view_vdir(view):
    if view is None:
        return None
    try:
        return _vec_normalize_xyz(view.ViewDirection)
    except Exception:
        return None


def _perp_exterior_viga_en_vista(view, es_cara_inferior):
    """
    Eje «exterior» en el plano de la vista: sup = ``UpDirection``; inf = opuesto.
    """
    vdir = _view_vdir(view)
    if vdir is None:
        return None
    try:
        up = view.UpDirection
    except Exception:
        up = None
    if up is None:
        return None
    perp = None
    if _proyectar_vector_en_plano_perp_normal is not None:
        perp = _vec_normalize_xyz(
            _proyectar_vector_en_plano_perp_normal(up, vdir)
        )
    if perp is None:
        return None
    if es_cara_inferior:
        try:
            perp = XYZ(-float(perp.X), -float(perp.Y), -float(perp.Z))
        except Exception:
            pass
    return perp


def _beam_half_depth_along_perp_ft(document, rebar_ids, perp):
    """Semicanto de la viga medido a lo largo de ``perp`` (margen extra inf)."""
    if document is None or perp is None:
        return _mm_to_ft(120.0)
    host = None
    if _framing_host_desde_lote_rebars is not None:
        try:
            host = _framing_host_desde_lote_rebars(document, rebar_ids)
        except Exception:
            host = None
    if host is None:
        return _mm_to_ft(120.0)
    try:
        bb = host.get_BoundingBox(None)
        if bb is None:
            return _mm_to_ft(120.0)
        c = (bb.Min + bb.Max) * 0.5
        mn = bb.Min
        mx = bb.Max
        corners = (
            XYZ(mn.X, mn.Y, mn.Z),
            XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z),
            XYZ(mx.X, mx.Y, mn.Z),
            XYZ(mn.X, mn.Y, mx.Z),
            XYZ(mx.X, mn.Y, mx.Z),
            XYZ(mn.X, mx.Y, mx.Z),
            XYZ(mx.X, mx.Y, mx.Z),
        )
        scalars = [float(_vec_dot(p.Subtract(c), perp)) for p in corners]
        span = max(scalars) - min(scalars)
        if span < 1e-9:
            return _mm_to_ft(120.0)
        return 0.5 * float(span) + _mm_to_ft(_TAG_INF_CLEARANCE_MM)
    except Exception:
        return _mm_to_ft(120.0)


def _other_independent_tags_in_view(document, view, exclude_tags):
    if document is None or view is None:
        return []
    try:
        vid = view.Id
    except Exception:
        return []
    skip = set()
    for tg in exclude_tags or []:
        if tg is None:
            continue
        try:
            skip.add(int(tg.Id.IntegerValue))
        except Exception:
            pass
    out = []
    try:
        coll = (
            FilteredElementCollector(document)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []
    for el in coll or []:
        if el is None or not isinstance(el, IndependentTag):
            continue
        try:
            if int(el.Id.IntegerValue) in skip:
                continue
            if el.OwnerViewId != vid:
                continue
        except Exception:
            continue
        out.append(el)
    return out


def _tag_overlaps_any(tag, others, view, clearance_mm):
    if tag is None or not others:
        return False
    if _tags_overlap_with_clearance is None:
        return False
    for ob in others:
        try:
            if _tags_overlap_with_clearance(tag, ob, view, clearance_mm):
                return True
        except Exception:
            continue
    return False


def _alinear_cabeceras_exterior_viga(
    document, view, rebar_ids, es_cara_inferior,
):
    """Alinea cabeceras en una fila común hacia el exterior del hormigón en la vista."""
    if (
        document is None
        or view is None
        or not rebar_ids
        or _collect_independent_tags_for_rebar_lote is None
    ):
        return
    perp = _perp_exterior_viga_en_vista(view, es_cara_inferior)
    if perp is None:
        return
    tags = _collect_independent_tags_for_rebar_lote(document, view, rebar_ids)
    if not tags:
        return
    extra_ft = float(_TAG_ALIGN_EXTRA_FT) + float(_TAG_OUTSIDE_FT)
    if es_cara_inferior:
        extra_ft += _beam_half_depth_along_perp_ft(document, rebar_ids, perp)
    projs = []
    for tag in tags:
        try:
            head = tag.TagHeadPosition
            projs.append((tag, head, float(_vec_dot(head, perp))))
        except Exception:
            continue
    if not projs:
        return
    try:
        ref_s = max(s for _, _, s in projs) + extra_ft
    except Exception:
        return
    for tag, head, s0 in projs:
        try:
            shift = ref_s - s0
            tag.TagHeadPosition = XYZ(
                head.X + shift * perp.X,
                head.Y + shift * perp.Y,
                head.Z + shift * perp.Z,
            )
        except Exception:
            continue
    try:
        document.Regenerate()
    except Exception:
        pass


def _empujar_etiquetas_fuera_de_otras(document, view, tags, perp):
    """Desplaza cabeceras aún más en ``perp`` mientras solapen otras etiquetas."""
    if not tags or perp is None:
        return
    others = _other_independent_tags_in_view(document, view, tags)
    if not others:
        return
    step_ft = _mm_to_ft(_TAG_NUDGE_STEP_MM)
    clr_mm = float(_TAG_INF_CLEARANCE_MM)
    for tg in tags:
        if tg is None:
            continue
        for _ in range(int(_TAG_NUDGE_MAX_STEPS)):
            if not _tag_overlaps_any(tg, others, view, clr_mm):
                break
            try:
                h = tg.TagHeadPosition
                tg.TagHeadPosition = XYZ(
                    h.X + step_ft * perp.X,
                    h.Y + step_ft * perp.Y,
                    h.Z + step_ft * perp.Z,
                )
            except Exception:
                break
            try:
                document.Regenerate()
            except Exception:
                pass


def _corregir_etiquetas_inferiores_viga(document, view, rebar_ids):
    """
    Refuerzo para capa inferior: dirección explícita hacia abajo en la vista,
    margen al canto y separación de otras etiquetas (confinamiento, etc.).
    """
    if not rebar_ids:
        return
    _alinear_cabeceras_exterior_viga(
        document, view, rebar_ids, es_cara_inferior=True,
    )
    if _separar_etiquetas_rebar_solapadas_lote is not None:
        try:
            _separar_etiquetas_rebar_solapadas_lote(
                document, view, rebar_ids, es_cara_inferior=True,
            )
        except Exception:
            pass
    perp = _perp_exterior_viga_en_vista(view, es_cara_inferior=True)
    if perp is None or _collect_independent_tags_for_rebar_lote is None:
        return
    tags = _collect_independent_tags_for_rebar_lote(document, view, rebar_ids)
    _empujar_etiquetas_fuera_de_otras(document, view, tags, perp)


def _align_longitudinal_tags_by_side(document, view, rebar_ids, es_cara_inferior):
    if not rebar_ids:
        return
    if es_cara_inferior:
        if _alinear_etiquetas_rebar_mismo_lote is not None:
            try:
                _alinear_etiquetas_rebar_mismo_lote(
                    document, view, rebar_ids, es_cara_inferior=True,
                )
            except Exception:
                pass
        _corregir_etiquetas_inferiores_viga(document, view, rebar_ids)
        return
    if _alinear_etiquetas_rebar_mismo_lote is not None:
        try:
            _alinear_etiquetas_rebar_mismo_lote(
                document, view, rebar_ids, es_cara_inferior=False,
            )
        except Exception:
            pass
    if _separar_etiquetas_rebar_solapadas_lote is not None:
        try:
            _separar_etiquetas_rebar_solapadas_lote(
                document, view, rebar_ids, es_cara_inferior=False,
            )
        except Exception:
            pass


def _align_longitudinal_tags_sup_inf(document, view, rebars_by_side):
    if not rebars_by_side:
        return
    sup_ids = _rebar_element_ids(rebars_by_side.get(u"sup"))
    inf_ids = _rebar_element_ids(rebars_by_side.get(u"inf"))
    _align_longitudinal_tags_by_side(document, view, sup_ids, es_cara_inferior=False)
    _align_longitudinal_tags_by_side(document, view, inf_ids, es_cara_inferior=True)


def realinear_longitudinales_inf_tras_confinamiento(document, view, rebars_by_side):
    """Tras etiquetar confinamiento, vuelve a empujar las inferiores hacia abajo."""
    if document is None or view is None or not rebars_by_side:
        return
    inf_ids = _rebar_element_ids(rebars_by_side.get(u"inf"))
    _corregir_etiquetas_inferiores_viga(document, view, inf_ids)


def etiquetar_longitudinales_en_vista(
    document,
    view,
    rebars,
    use_transaction=False,
    rebars_by_side=None,
):
    """
    Crea ``IndependentTag`` por barra longitudinal en ``view``.

    ``rebars_by_side``: opcional, ``{"sup": [...], "inf": [...]}``; si se indica,
    tras crear las etiquetas alinea cabeceras sup/inf hacia lados opuestos del hormigón.

    Returns:
        ``(n_etiquetas, avisos, err)`` — ``err`` no nulo solo si falla el bloque global.
    """
    ids = _rebar_element_ids(rebars)
    if not ids:
        return 0, [], None
    if document is None or view is None:
        return 0, [], u"Sin documento o vista activa para etiquetar longitudinales."
    if etiquetar_rebars_creados_en_vista is None:
        return (
            0,
            [],
            u"No se cargó enfierrado_shaft_hashtag (etiquetar_rebars_creados_en_vista).",
        )

    n_tags, avisos, err = etiquetar_rebars_creados_en_vista(
        document,
        view,
        ids,
        family_name=LONGITUDINAL_REBAR_TAG_FAMILY,
        fixed_type_name=None,
        use_transaction=use_transaction,
    )
    if n_tags > 0 and rebars_by_side:
        _align_longitudinal_tags_sup_inf(document, view, rebars_by_side)
    return n_tags, avisos, err
