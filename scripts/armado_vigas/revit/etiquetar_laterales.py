# -*- coding: utf-8 -*-
"""
Etiquetado de barras laterales (Armado vigas).

Familias de etiqueta:

- ``EST_A_STRUCTURAL REBAR TAG_LATERAL`` — una barra lateral en la viga.
- ``EST_A_STRUCTURAL REBAR TAG_VIGA_LATERAL_MULTI HOST`` — caras ±ancho (2+ Rebar)
  en la misma viga.

Tipo fijo «Laterales» en ambas familias. Sin leader; **una etiqueta por viga**
del lote inicial, cabecera en el **centroide de esa viga**.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Reference, Transaction, XYZ
from System.Collections.Generic import List

LATERAL_REBAR_TAG_FAMILY = u"EST_A_STRUCTURAL REBAR TAG_LATERAL"
LATERAL_MULTIHOST_TAG_FAMILY = (
    u"EST_A_STRUCTURAL REBAR TAG_VIGA_LATERAL_MULTI HOST"
)
LATERAL_REBAR_TAG_TYPE = u"Laterales"

try:
    from enfierrado_shaft_hashtag import (
        _collect_rebar_tag_symbol_map,
        _rebar_centerline_midpoint_xyz,
        _rebar_reference_candidates_for_tag,
        _resolve_fixed_tag_type_id,
        _view_ok_for_rebar_tags,
    )
except Exception:
    _collect_rebar_tag_symbol_map = None
    _rebar_centerline_midpoint_xyz = None
    _rebar_reference_candidates_for_tag = None
    _resolve_fixed_tag_type_id = None
    _view_ok_for_rebar_tags = None

try:
    from geometria_colision_vigas import obtener_solidos_elemento
except Exception:
    obtener_solidos_elemento = None

try:
    from armado_vigas.revit.colocar_rebar import _elemento_contiene_punto
except Exception:
    _elemento_contiene_punto = None

try:
    from armado_vigas.revit.etiquetar_confinamiento import (
        _aplicar_estilo_tag_confinamiento_viga,
        _crear_tag_confinamiento_sin_leader,
        _load_cabezal_tags_module,
        _referencias_tag_rebar_viga,
    )
except Exception:
    _aplicar_estilo_tag_confinamiento_viga = None
    _crear_tag_confinamiento_sin_leader = None
    _load_cabezal_tags_module = None
    _referencias_tag_rebar_viga = None


def _centroide_puntos_xyz(pts):
    if not pts:
        return None
    try:
        sx = sy = sz = 0.0
        n = 0
        for p in pts:
            if p is None:
                continue
            sx += float(p.X)
            sy += float(p.Y)
            sz += float(p.Z)
            n += 1
        if n == 0:
            return None
        inv = 1.0 / float(n)
        return XYZ(sx * inv, sy * inv, sz * inv)
    except Exception:
        return None


def _centroide_framing_beam(beam):
    """Centroide del sólido de la viga; fallback al centro del bounding box."""
    if beam is None:
        return None
    pts = []
    if obtener_solidos_elemento is not None:
        try:
            for solid in obtener_solidos_elemento(beam) or []:
                if solid is None:
                    continue
                try:
                    faces = solid.Faces
                except Exception:
                    continue
                if faces is None:
                    continue
                for face in faces:
                    try:
                        mesh = face.Triangulate()
                    except Exception:
                        continue
                    if mesh is None:
                        continue
                    try:
                        n_tri = int(mesh.NumTriangles)
                    except Exception:
                        n_tri = 0
                    for i in range(n_tri):
                        try:
                            tri = mesh.get_Triangle(i)
                            for j in range(3):
                                pts.append(tri.get_Vertex(j))
                        except Exception:
                            continue
        except Exception:
            pts = []
    c = _centroide_puntos_xyz(pts)
    if c is not None:
        return c
    try:
        bb = beam.get_BoundingBox(None)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    return None


def _beam_element_id_int(beam):
    if beam is None:
        return None
    try:
        return int(beam.Id.IntegerValue)
    except Exception:
        return None


def _rebar_element_id_int(rebar):
    if rebar is None:
        return None
    try:
        return int(rebar.Id.IntegerValue)
    except Exception:
        return None


def _resolve_lateral_rebars_for_beam(document, beam, rebars_laterales):
    """
    Todas las Rebar laterales de ``beam``: host directo, fibra en sólido o más cercana.
    """
    if beam is None or not rebars_laterales:
        return []
    bid = _beam_element_id_int(beam)
    hosted = []
    otros = []
    for rb in rebars_laterales or []:
        if rb is None:
            continue
        try:
            hid = int(rb.GetHostId().IntegerValue)
            if bid is not None and hid == bid:
                hosted.append(rb)
                continue
        except Exception:
            pass
        otros.append(rb)

    out = []
    seen = set()
    for rb in hosted:
        rid = _rebar_element_id_int(rb)
        if rid is None or rid in seen:
            continue
        seen.add(rid)
        out.append(rb)

    if _elemento_contiene_punto is not None and _rebar_centerline_midpoint_xyz is not None:
        for rb in otros:
            rid = _rebar_element_id_int(rb)
            if rid is None or rid in seen:
                continue
            try:
                pt = _rebar_centerline_midpoint_xyz(rb)
            except Exception:
                pt = None
            if pt is not None and _elemento_contiene_punto(pt, beam):
                seen.add(rid)
                out.append(rb)

    if out:
        return out

    centro = _centroide_framing_beam(beam)
    if centro is None or _rebar_centerline_midpoint_xyz is None:
        return []
    ranked = []
    for rb in otros:
        try:
            pt = _rebar_centerline_midpoint_xyz(rb)
        except Exception:
            pt = None
        if pt is None:
            continue
        try:
            d = float(pt.DistanceTo(centro))
        except Exception:
            continue
        ranked.append((d, rb))
    if not ranked:
        return []
    ranked.sort(key=lambda x: x[0])
    best_d = ranked[0][0]
    tol = 25.0 / 304.8
    for d, rb in ranked:
        if float(d) > float(best_d) + tol:
            break
        rid = _rebar_element_id_int(rb)
        if rid is None or rid in seen:
            continue
        seen.add(rid)
        out.append(rb)
    return out


def _dedupe_framing_beams(framing_elements):
    out = []
    seen = set()
    for el in framing_elements or []:
        if el is None:
            continue
        eid = _beam_element_id_int(el)
        if eid is not None:
            if eid in seen:
                continue
            seen.add(eid)
        out.append(el)
    return out


def _lateral_tag_groups_per_beam(document, beams, rebars_laterales):
    """
    Una entrada por viga seleccionada con laterales asociados.

    Returns:
        ``[(beam, [rebar, ...]), ...]`` — orden del lote inicial.
    """
    groups = []
    for beam in _dedupe_framing_beams(beams):
        beam_rebars = _resolve_lateral_rebars_for_beam(
            document, beam, rebars_laterales,
        )
        if beam_rebars:
            groups.append((beam, beam_rebars))
    return groups


def _ordered_rebar_ids(rebars):
    items = []
    for rb in rebars or []:
        rid = _rebar_element_id_int(rb)
        if rid is None:
            continue
        items.append((rid, rb))
    items.sort(key=lambda x: x[0])
    out = []
    seen = set()
    for rid, rb in items:
        if rid in seen:
            continue
        seen.add(rid)
        try:
            out.append(rb.Id)
        except Exception:
            pass
    return out


def _etiquetar_lateral_en_centroide_viga(
    document, view, beam, rebar, tag_type_id, tags_mod=None,
):
    if rebar is None or tag_type_id is None:
        return False, u"Rebar o tipo de etiqueta inválido"
    head = _centroide_framing_beam(beam)
    if head is None:
        return False, u"sin centroide de viga"
    if _crear_tag_confinamiento_sin_leader is None:
        return False, u"módulo de etiquetado confinamiento no disponible"
    tag, err = _crear_tag_confinamiento_sin_leader(
        document, view, rebar, tag_type_id, head, tags_mod,
    )
    if tag is not None:
        return True, None
    return False, err or u"etiqueta no creada"


def _etiquetar_multihost_lateral_viga(
    document, view, rebars, tag_type_id, head, tags_mod=None,
):
    ids = _ordered_rebar_ids(rebars)
    if not ids:
        return False, u"grupo multihost vacío"
    if tag_type_id is None:
        return False, u"tipo multihost inválido"
    if head is None:
        return False, u"sin centroide para etiqueta"
    if _crear_tag_confinamiento_sin_leader is None:
        return False, u"módulo de etiquetado confinamiento no disponible"

    primary = document.GetElement(ids[-1])
    if primary is None:
        return False, u"barra principal no encontrada"

    tag, err = _crear_tag_confinamiento_sin_leader(
        document, view, primary, tag_type_id, head, tags_mod,
    )
    if tag is None:
        return False, err or u"multihost no creado"
    if len(ids) < 2:
        return True, None

    add_fn = getattr(tag, u"AddReferences", None)
    if add_fn is None:
        return False, u"Revit sin AddReferences para multihost"

    extra_refs = []
    if tags_mod is not None and hasattr(tags_mod, u"_refs_multihost_para_rebars"):
        try:
            extra_refs = tags_mod._refs_multihost_para_rebars(document, view, ids[:-1])
        except Exception:
            extra_refs = []

    if not extra_refs and _rebar_reference_candidates_for_tag is not None:
        seen = set()
        for rid in ids[:-1]:
            rb = document.GetElement(rid)
            for ref in _rebar_reference_candidates_for_tag(document, view, rb) or []:
                try:
                    key = ref.ConvertToStableRepresentation(document)
                except Exception:
                    key = id(ref)
                if key in seen:
                    continue
                seen.add(key)
                extra_refs.append(ref)

    if not extra_refs and _referencias_tag_rebar_viga is not None:
        seen = set()
        for rid in ids[:-1]:
            rb = document.GetElement(rid)
            for ref in _referencias_tag_rebar_viga(document, rb, view, tags_mod) or []:
                try:
                    key = ref.ConvertToStableRepresentation(document)
                except Exception:
                    key = id(ref)
                if key in seen:
                    continue
                seen.add(key)
                extra_refs.append(ref)

    if not extra_refs:
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        return False, u"sin referencias adicionales multihost"

    refs_add = List[Reference]()
    for ref in extra_refs:
        refs_add.Add(ref)
    try:
        add_fn(refs_add)
    except Exception as ex:
        try:
            document.Delete(tag.Id)
        except Exception:
            pass
        try:
            return False, unicode(ex)
        except Exception:
            return False, str(ex)

    if _aplicar_estilo_tag_confinamiento_viga is not None:
        _aplicar_estilo_tag_confinamiento_viga(tag, head)
    return True, None


def _beam_label(beam):
    if beam is None:
        return u"?"
    try:
        return unicode(beam.Id.IntegerValue)
    except Exception:
        return u"?"


def etiquetar_laterales_en_vista(
    document,
    view,
    rebars,
    framing_elements=None,
    use_transaction=False,
):
    """
    Una etiqueta por viga del lote, sin leader, en el centroide de cada viga.

    - 2+ Rebar en la misma viga (caras ±ancho) → multihost en centroide de esa viga.
    - 1 Rebar → familia single-host en el mismo punto.

    Returns:
        ``(n_etiquetas, avisos, err)``
    """
    beams = _dedupe_framing_beams(framing_elements)
    if not beams:
        return 0, [], u"Sin vigas en el lote para etiquetar laterales."
    if not rebars:
        return 0, [], None
    if document is None or view is None:
        return 0, [], u"Sin documento o vista activa para etiquetar laterales."
    if _view_ok_for_rebar_tags is not None:
        ok_view, msg_view = _view_ok_for_rebar_tags(view)
        if not ok_view:
            return 0, [msg_view], None
    if _collect_rebar_tag_symbol_map is None or _resolve_fixed_tag_type_id is None:
        return (
            0,
            [],
            u"No se cargó enfierrado_shaft_hashtag (mapa de etiquetas).",
        )

    tag_map = _collect_rebar_tag_symbol_map(document, LATERAL_REBAR_TAG_FAMILY)
    if not tag_map:
        return 0, [], u"No se encontraron tipos en familia '{0}'.".format(
            LATERAL_REBAR_TAG_FAMILY
        )
    single_tag_type_id = _resolve_fixed_tag_type_id(tag_map, LATERAL_REBAR_TAG_TYPE)
    if single_tag_type_id is None:
        return 0, [], u"No se encontró tipo '{0}' en familia '{1}'.".format(
            LATERAL_REBAR_TAG_TYPE, LATERAL_REBAR_TAG_FAMILY
        )

    mh_tag_map = _collect_rebar_tag_symbol_map(
        document, LATERAL_MULTIHOST_TAG_FAMILY,
    )
    if not mh_tag_map:
        return 0, [], u"No se encontraron tipos en familia '{0}'.".format(
            LATERAL_MULTIHOST_TAG_FAMILY
        )
    multihost_tag_type_id = _resolve_fixed_tag_type_id(
        mh_tag_map, LATERAL_REBAR_TAG_TYPE,
    )
    if multihost_tag_type_id is None:
        return 0, [], u"No se encontró tipo '{0}' en familia '{1}'.".format(
            LATERAL_REBAR_TAG_TYPE, LATERAL_MULTIHOST_TAG_FAMILY
        )

    tags_mod = _load_cabezal_tags_module() if _load_cabezal_tags_module else None
    groups = _lateral_tag_groups_per_beam(document, beams, rebars)
    beams_with_tag = set(_beam_element_id_int(b) for b, _ in groups)

    def _run():
        n_ok = 0
        avisos_loc = []
        for beam in beams:
            bid = _beam_element_id_int(beam)
            if bid is not None and bid not in beams_with_tag:
                if len(avisos_loc) < 12:
                    avisos_loc.append(
                        u"Viga {0}: sin Rebar lateral asociado.".format(
                            _beam_label(beam)
                        )
                    )

        for beam, grp_rebars in groups:
            head = _centroide_framing_beam(beam)
            label = _beam_label(beam)
            if len(grp_rebars) >= 2:
                ok, err = _etiquetar_multihost_lateral_viga(
                    document,
                    view,
                    grp_rebars,
                    multihost_tag_type_id,
                    head,
                    tags_mod=tags_mod,
                )
            else:
                ok, err = _etiquetar_lateral_en_centroide_viga(
                    document,
                    view,
                    beam,
                    grp_rebars[0],
                    single_tag_type_id,
                    tags_mod=tags_mod,
                )
            if ok:
                n_ok += 1
            elif err and len(avisos_loc) < 12:
                avisos_loc.append(u"Viga {0} etiqueta lateral: {1}".format(label, err))
        return n_ok, avisos_loc

    if use_transaction:
        t = Transaction(document, u"Arainco: Etiquetar laterales vigas")
        t.Start()
        try:
            n_ok, avisos_loc = _run()
            t.Commit()
            return n_ok, avisos_loc, None
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            try:
                msg = unicode(ex)
            except Exception:
                msg = str(ex)
            return 0, [], msg
    n_ok, avisos_loc = _run()
    return n_ok, avisos_loc, None
