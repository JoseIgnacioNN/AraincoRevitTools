# -*- coding: utf-8 -*-
"""
Etiquetado de barras longitudinales (Armado vigas).

Familia ``EST_A_STRUCTURAL REBAR TAG_HORIZONTAL``: el tipo de etiqueta se resuelve
por nombre del ``RebarShape`` de cada barra (p. ej. «03», «02»), igual que en
``enfierrado_shaft_hashtag`` / Armado muros cabezal.

Posición de cabeceras respecto a las barras en la vista de alzado/sección:
  - **Superiores**: arriba de las barras (``+View.UpDirection``)
  - **Inferiores**: abajo de las barras (``−View.UpDirection``)
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    ElementId,
    ElementOwnerViewFilter,
    ExclusionFilter,
    FilteredElementCollector,
    IndependentTag,
    XYZ,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar

LONGITUDINAL_REBAR_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_HORIZONTAL"

# Separación mínima cabecera ↔ barra a lo largo de View.Up (mm / ft).
_TAG_BAR_CLEARANCE_MM = 95.0
_TAG_ALIGN_EXTRA_FT = 0.05
_TAG_NUDGE_STEP_MM = 28.0
_TAG_NUDGE_MAX_STEPS = 48
_TAG_OVERLAP_CLEARANCE_MM = 50.0

try:
    from enfierrado_shaft_hashtag import etiquetar_rebars_creados_en_vista
except Exception:
    etiquetar_rebars_creados_en_vista = None

try:
    from geometria_viga_cara_superior_detalle import (
        _alinear_etiquetas_rebar_mismo_lote,
        _collect_independent_tags_for_rebar_lote,
        _proyectar_vector_en_plano_perp_normal,
        _separar_etiquetas_rebar_solapadas_lote,
        _tags_overlap_with_clearance,
        _vec_dot,
        _vec_normalize_xyz,
    )
except Exception:
    _alinear_etiquetas_rebar_mismo_lote = None
    _collect_independent_tags_for_rebar_lote = None
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


def _eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid)
        except Exception:
            return None


def _view_up_in_plane(view):
    """``View.UpDirection`` proyectado en el plano de la vista (arriba en pantalla)."""
    if view is None:
        return None
    try:
        up = view.UpDirection
        vdir = view.ViewDirection
    except Exception:
        return None
    if up is None:
        return None
    if _proyectar_vector_en_plano_perp_normal is not None and vdir is not None:
        try:
            p = _proyectar_vector_en_plano_perp_normal(up, vdir)
            if p is not None and _vec_normalize_xyz is not None:
                p = _vec_normalize_xyz(p)
            if p is not None:
                return p
        except Exception:
            pass
    if _vec_normalize_xyz is not None:
        try:
            return _vec_normalize_xyz(up)
        except Exception:
            pass
    return up


def _side_axis_for_face(view, es_cara_inferior):
    """
    Eje unitario hacia el lado donde debe ir la etiqueta:

    - SUP → arriba en la vista
    - INF → abajo en la vista
    """
    up = _view_up_in_plane(view)
    if up is None:
        return None
    if not es_cara_inferior:
        return up
    try:
        return XYZ(-float(up.X), -float(up.Y), -float(up.Z))
    except Exception:
        return None


def _rebar_mid_point(document, rebar_or_id):
    """Punto representativo de la barra (eje / bbox) para medir «arriba/abajo»."""
    if document is None or rebar_or_id is None:
        return None
    rb = rebar_or_id
    if not isinstance(rb, Rebar):
        try:
            rb = document.GetElement(rebar_or_id)
        except Exception:
            rb = None
    if rb is None or not isinstance(rb, Rebar):
        return None
    try:
        curves = rb.GetCenterlineCurves(
            False, False, False, MultiplanarOption.IncludeOnlyPlanarCurves, 0
        )
        if curves and curves.Count > 0:
            c0 = curves[0]
            return c0.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        curves = rb.GetCenterlineCurves(False, False, False)
        if curves and len(list(curves)) > 0:
            c0 = list(curves)[0]
            return c0.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        bb = rb.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _tag_tagged_element_id(tag):
    if tag is None:
        return None
    try:
        ids = tag.GetTaggedLocalElementIds()
        if ids is not None:
            for eid in ids:
                if eid is not None:
                    return eid
    except Exception:
        pass
    try:
        refs = tag.GetTaggedReferences()
        if refs is not None:
            for ref in refs:
                try:
                    return ref.ElementId
                except Exception:
                    continue
    except Exception:
        pass
    return None


def _dot(a, b):
    if a is None or b is None:
        return 0.0
    if _vec_dot is not None:
        try:
            return float(_vec_dot(a, b))
        except Exception:
            pass
    try:
        return (
            float(a.X) * float(b.X)
            + float(a.Y) * float(b.Y)
            + float(a.Z) * float(b.Z)
        )
    except Exception:
        return 0.0


def _other_independent_tags_in_view(document, view, exclude_tags):
    """
    IndependentTag **owned** by the view (``ElementOwnerViewFilter``, C# API).

    Exclusión de tags recién creados vía ``ExclusionFilter`` cuando hay ids.
    """
    if document is None or view is None:
        return []
    try:
        vid = view.Id
    except Exception:
        return []
    skip_ids = []
    for tg in exclude_tags or []:
        if tg is None:
            continue
        try:
            skip_ids.append(tg.Id)
        except Exception:
            pass
    try:
        coll = (
            FilteredElementCollector(document)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .WherePasses(ElementOwnerViewFilter(vid))
        )
        if skip_ids:
            try:
                from System.Collections.Generic import List

                id_list = List[ElementId]()
                for eid in skip_ids:
                    if eid is not None and eid != ElementId.InvalidElementId:
                        id_list.Add(eid)
                if id_list.Count > 0:
                    coll = coll.WherePasses(ExclusionFilter(id_list))
            except Exception:
                # Fallback: excluir en Python (misma semántica).
                skip_iv = set()
                for eid in skip_ids:
                    try:
                        v = getattr(eid, u"Value", None)
                        skip_iv.add(int(v if v is not None else eid.IntegerValue))
                    except Exception:
                        pass
                out = []
                for el in coll:
                    if el is None:
                        continue
                    try:
                        v = getattr(el.Id, u"Value", None)
                        iv = int(v if v is not None else el.Id.IntegerValue)
                    except Exception:
                        continue
                    if iv in skip_iv:
                        continue
                    out.append(el)
                return out
        return list(coll)
    except Exception:
        return []


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


def _empujar_etiquetas_fuera_de_otras(document, view, tags, side_axis):
    """Desplaza cabeceras en ``side_axis`` mientras solapen otras etiquetas."""
    if not tags or side_axis is None:
        return
    others = _other_independent_tags_in_view(document, view, tags)
    if not others:
        return
    step_ft = _mm_to_ft(_TAG_NUDGE_STEP_MM)
    clr_mm = float(_TAG_OVERLAP_CLEARANCE_MM)
    for tg in tags:
        if tg is None:
            continue
        for _ in range(int(_TAG_NUDGE_MAX_STEPS)):
            if not _tag_overlaps_any(tg, others, view, clr_mm):
                break
            try:
                h = tg.TagHeadPosition
                tg.TagHeadPosition = XYZ(
                    h.X + step_ft * side_axis.X,
                    h.Y + step_ft * side_axis.Y,
                    h.Z + step_ft * side_axis.Z,
                )
            except Exception:
                break
            try:
                document.Regenerate()
            except Exception:
                pass


def _posicionar_etiquetas_lado_barras(
    document, view, rebar_ids, es_cara_inferior,
):
    """
    Coloca cada cabecera **sobre** (SUP) o **bajo** (INF) la barra etiquetada.

    1. Eje = ± ``View.UpDirection`` en plano de vista.
    2. Por etiqueta: ``dot(head, axis) >= dot(barra, axis) + holgura``.
    3. Alinea todas las cabeceras del lado en una fila exterior común.
    """
    if document is None or view is None or not rebar_ids:
        return
    if _collect_independent_tags_for_rebar_lote is None:
        return
    side = _side_axis_for_face(view, es_cara_inferior)
    if side is None:
        return
    tags = _collect_independent_tags_for_rebar_lote(document, view, rebar_ids)
    if not tags:
        return

    clear_ft = _mm_to_ft(_TAG_BAR_CLEARANCE_MM)
    allowed = set()
    for rid in rebar_ids:
        k = _eid_int(rid)
        if k is not None:
            allowed.add(k)

    projs = []
    for tag in tags:
        if tag is None:
            continue
        try:
            head = tag.TagHeadPosition
        except Exception:
            continue
        if head is None:
            continue

        bar_pt = None
        tid = _tag_tagged_element_id(tag)
        if tid is not None:
            tk = _eid_int(tid)
            if allowed and tk is not None and tk not in allowed:
                continue
            bar_pt = _rebar_mid_point(document, tid)
        if bar_pt is None and rebar_ids:
            # Respaldo: punto de primera barra del lote.
            bar_pt = _rebar_mid_point(document, rebar_ids[0])
        if bar_pt is None:
            continue

        bar_s = _dot(bar_pt, side)
        head_s = _dot(head, side)
        min_s = bar_s + clear_ft
        if head_s < min_s - 1e-9:
            shift = min_s - head_s
            try:
                head = XYZ(
                    head.X + shift * side.X,
                    head.Y + shift * side.Y,
                    head.Z + shift * side.Z,
                )
                tag.TagHeadPosition = head
                head_s = min_s
            except Exception:
                pass
        projs.append((tag, head, head_s))

    if not projs:
        return

    # Fila común: la más exterior del lote + margen.
    try:
        ref_s = max(s for _, _, s in projs) + float(_TAG_ALIGN_EXTRA_FT)
    except Exception:
        return
    for tag, head, s0 in projs:
        try:
            shift = ref_s - s0
            if abs(shift) < 1e-12:
                continue
            tag.TagHeadPosition = XYZ(
                head.X + shift * side.X,
                head.Y + shift * side.Y,
                head.Z + shift * side.Z,
            )
        except Exception:
            continue
    try:
        document.Regenerate()
    except Exception:
        pass


def _corregir_etiquetas_cara(
    document, view, rebar_ids, es_cara_inferior,
):
    """Alinea lote + fuerza lado de las barras + empuja solapes con otras etiq."""
    if not rebar_ids:
        return
    # Alineado por normal de cara (si disponible).
    if _alinear_etiquetas_rebar_mismo_lote is not None:
        try:
            _alinear_etiquetas_rebar_mismo_lote(
                document, view, rebar_ids, es_cara_inferior=es_cara_inferior,
            )
        except Exception:
            pass
    # Regla de negocio Armado vigas: SUP arriba / INF abajo de las barras.
    _posicionar_etiquetas_lado_barras(
        document, view, rebar_ids, es_cara_inferior,
    )
    if _separar_etiquetas_rebar_solapadas_lote is not None:
        try:
            _separar_etiquetas_rebar_solapadas_lote(
                document, view, rebar_ids, es_cara_inferior=es_cara_inferior,
            )
        except Exception:
            pass
    # Tras separar (a veces mueve a lo largo del eje), reponer lado de barras.
    _posicionar_etiquetas_lado_barras(
        document, view, rebar_ids, es_cara_inferior,
    )
    side = _side_axis_for_face(view, es_cara_inferior)
    if side is not None and _collect_independent_tags_for_rebar_lote is not None:
        tags = _collect_independent_tags_for_rebar_lote(document, view, rebar_ids)
        _empujar_etiquetas_fuera_de_otras(document, view, tags, side)
        # Último pase: no dejar etiquetas "del lado equivocado" tras el nudge.
        _posicionar_etiquetas_lado_barras(
            document, view, rebar_ids, es_cara_inferior,
        )


def _align_longitudinal_tags_by_side(document, view, rebar_ids, es_cara_inferior):
    if not rebar_ids:
        return
    _corregir_etiquetas_cara(
        document, view, rebar_ids, es_cara_inferior=es_cara_inferior,
    )


def _align_longitudinal_tags_sup_inf(document, view, rebars_by_side):
    if not rebars_by_side:
        return
    sup_ids = _rebar_element_ids(rebars_by_side.get(u"sup"))
    inf_ids = _rebar_element_ids(rebars_by_side.get(u"inf"))
    # SUP primero (arriba); INF después (abajo).
    _align_longitudinal_tags_by_side(document, view, sup_ids, es_cara_inferior=False)
    _align_longitudinal_tags_by_side(document, view, inf_ids, es_cara_inferior=True)


def realinear_longitudinales_inf_tras_confinamiento(document, view, rebars_by_side):
    """Tras confinamiento: reponer lados SUP/INF de las longitudinales."""
    if document is None or view is None or not rebars_by_side:
        return
    _align_longitudinal_tags_sup_inf(document, view, rebars_by_side)


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
    tras crear las etiquetas:

    - cabeceras SUP → **arriba** de las barras
    - cabeceras INF → **abajo** de las barras

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
    elif n_tags > 0 and rebars:
        # Sin side: no se puede distinguir SUP/INF; no forzar side.
        pass
    return n_tags, avisos, err
