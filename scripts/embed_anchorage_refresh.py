# -*- coding: utf-8 -*-
"""
DMU cotas de empotramiento: ante cambios en el armado se **borran** el marcador y la cota
y se **crean de nuevo**. El marcador se coloca con la **distancia real** extremo–cara en el
modelo actual (no el mm tabulado guardado al crear), para alinear la cota con la geometría
actualizada de la barra.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import DetailCurve, ElementId, Line, PlanarFace, Reference, XYZ
from Autodesk.Revit.DB.Structure import Rebar

from embed_anchorage_link_schema import set_embed_anchorage_link


def refresh_embed_anchorage_marker(doc, marker, link):
    """
    Elimina el marcador y la dimensión existentes y crea un marcador nuevo + cota nueva
    según la geometría actual del Rebar y el vínculo guardado (cara, extremo, mm).
    """
    if doc is None or marker is None or link is None:
        return False
    if not isinstance(marker, DetailCurve):
        return False

    rebar = doc.GetElement(link["rb"])
    view = doc.GetElement(link["vw"])
    dim_id = link.get("dim")
    fst = link.get("fst") or u""
    amm = float(link.get("amm") or 0)
    endi = int(link.get("endi") or 0)

    if rebar is None or not isinstance(rebar, Rebar) or view is None:
        _delete_dim_marker(doc, marker, dim_id)
        return False

    if not fst:
        _delete_dim_marker(doc, marker, dim_id)
        return False

    from enfierrado_shaft_hashtag import (
        _collect_solids_from_host_element,
        _create_dimension_face_to_marker,
        _create_marker_detailcurve,
        _dimension_reference_count,
        _face_info_entry_from_face,
        _face_normal_xy,
        _face_origin,
        _ft_to_mm,
        _marker_point_for_table_anchorage,
        _mm_to_ft,
        _pick_end_for_anchorage,
        _point_plane_distance_ft_xy,
        _rebar_plan_endpoints_for_embed_anchorage,
        _rotate90_xy,
        _unit_xy,
        _xyz_add,
        _xyz_scale,
        _xyz_sub,
    )

    try:
        doc.Regenerate()
    except Exception:
        pass
    rebar = doc.GetElement(link["rb"])
    view = doc.GetElement(link["vw"])
    if rebar is None or not isinstance(rebar, Rebar) or view is None:
        _delete_dim_marker(doc, marker, dim_id)
        return False
    try:
        face_ref = Reference.ParseFromStableRepresentation(doc, fst)
        host = doc.GetElement(face_ref.ElementId)
        face = host.GetGeometryObjectFromReference(face_ref)
    except Exception:
        _delete_dim_marker(doc, marker, dim_id)
        return False
    if host is None or not isinstance(face, PlanarFace):
        _delete_dim_marker(doc, marker, dim_id)
        return False

    solids = _collect_solids_from_host_element(host)
    try:
        nxy = _face_normal_xy(face)
    except Exception:
        nxy = None
    if nxy is None:
        _delete_dim_marker(doc, marker, dim_id)
        return False

    face_infos = []
    try:
        fi_one = _face_info_entry_from_face(face, solids)
        if fi_one:
            face_infos.append(fi_one)
    except Exception:
        face_infos = []
    if not face_infos:
        _delete_dim_marker(doc, marker, dim_id)
        return False

    # Quitar anotaciones viejas primero; luego Regenerate y volver a resolver cara + eje del Rebar
    # (objetos de geometría previos pueden invalidarse tras Regenerate).
    _delete_dim_marker(doc, marker, dim_id)
    try:
        doc.Regenerate()
    except Exception:
        pass

    try:
        face_ref = Reference.ParseFromStableRepresentation(doc, fst)
        host = doc.GetElement(face_ref.ElementId)
        face = host.GetGeometryObjectFromReference(face_ref)
    except Exception:
        return False
    if host is None or not isinstance(face, PlanarFace):
        return False
    solids = _collect_solids_from_host_element(host)
    try:
        nxy = _face_normal_xy(face)
    except Exception:
        nxy = None
    if nxy is None:
        return False
    face_infos = []
    try:
        fi_one = _face_info_entry_from_face(face, solids)
        if fi_one:
            face_infos.append(fi_one)
    except Exception:
        face_infos = []
    if not face_infos:
        return False

    rebar = doc.GetElement(link["rb"])
    if rebar is None or not isinstance(rebar, Rebar):
        return False

    q0e, q1e = _rebar_plan_endpoints_for_embed_anchorage(rebar, view, face_nxy=nxy)
    if q0e is None or q1e is None:
        return False

    try:
        dxy = _unit_xy(_xyz_sub(q1e, q0e))
    except Exception:
        dxy = None
    if dxy is None:
        return False

    # Misma regla que al crear el enfierrado: extremo según distancia al plano y mm tabulados.
    end_idx = int(endi)
    face_pick = None
    try:
        end_idx, face_pick = _pick_end_for_anchorage(
            q0e, q1e, dxy, face_infos, expected_mm=amm, tol_mm=5.0
        )
    except Exception:
        end_idx, face_pick = int(endi), None
    if face_pick is None:
        end_idx = int(endi)
    end_idx = int(end_idx)
    if end_idx not in (0, 1):
        end_idx = 0

    end_pt = q0e if end_idx == 0 else q1e
    face_for_dim = face_pick.get("face") if isinstance(face_pick, dict) else face
    face_nxy = None
    try:
        if isinstance(face_pick, dict) and face_pick.get("nxy") is not None:
            face_nxy = face_pick.get("nxy")
    except Exception:
        face_nxy = None
    if face_nxy is None:
        face_nxy = nxy
    try:
        face_ref = face_pick.get("ref") if isinstance(face_pick, dict) else None
    except Exception:
        face_ref = None
    if face_ref is None:
        try:
            fr = getattr(face_for_dim, "Reference", None)
            if fr is not None:
                face_ref = fr
        except Exception:
            pass
    if face_ref is None:
        try:
            face_ref = Reference.ParseFromStableRepresentation(doc, fst)
        except Exception:
            return False

    face_o = None
    try:
        face_o = _face_origin(face_for_dim)
    except Exception:
        face_o = None

    # Distancia real en planta del extremo elegido al plano de la cara (geometría actual).
    # El mm guardado en el esquema (tabla al crear) no refleja alargos/recortes posteriores.
    anchorage_mm_geom = None
    try:
        d_ft = _point_plane_distance_ft_xy(end_pt, face_for_dim)
        if d_ft is not None and float(d_ft) > 1e-12:
            anchorage_mm_geom = _ft_to_mm(float(d_ft))
    except Exception:
        anchorage_mm_geom = None
    if anchorage_mm_geom is None:
        anchorage_mm_geom = float(amm)
    else:
        anchorage_mm_geom = max(0.0, float(anchorage_mm_geom))

    base_on_plane, marker_pt = _marker_point_for_table_anchorage(
        end_pt,
        dxy,
        end_idx,
        face_o,
        face_nxy,
        anchorage_mm=anchorage_mm_geom,
    )

    tdir = _rotate90_xy(face_nxy)
    if tdir is None:
        tdir = XYZ(1.0, 0.0, 0.0)

    try:
        ln_template = Line.CreateBound(q0e, q1e)
    except Exception:
        ln_template = None

    for shift_mm in (2.0, 20.0, 50.0):
        try:
            marker_pt_shift = _xyz_add(
                marker_pt, _xyz_scale(tdir, _mm_to_ft(float(shift_mm)))
            )
        except Exception:
            marker_pt_shift = marker_pt
        dc, marker_ref = _create_marker_detailcurve(
            doc, view, marker_pt_shift, face_nxy, length_mm=5.0
        )
        if marker_ref is None or face_ref is None:
            try:
                if dc is not None:
                    doc.Delete(dc.Id)
            except Exception:
                pass
            continue
        dim = None
        try:
            dim = _create_dimension_face_to_marker(
                doc,
                view,
                face_ref,
                marker_ref,
                base_on_plane,
                face_nxy,
                face_obj=face_for_dim,
                outside_offset_mm=450.0,
                line_len_mm=450.0,
                solids=solids,
                dim_line_template=ln_template,
            )
        except Exception:
            dim = None
        nrefs = _dimension_reference_count(dim)
        if (nrefs is None) or (nrefs != 2) or dim is None:
            try:
                if dim is not None:
                    doc.Delete(dim.Id)
            except Exception:
                pass
            try:
                if dc is not None:
                    doc.Delete(dc.Id)
            except Exception:
                pass
            continue
        try:
            ok = set_embed_anchorage_link(
                dc,
                link["rb"],
                dim.Id,
                link["vw"],
                fst,
                int(round(anchorage_mm_geom)),
                end_idx,
            )
        except Exception:
            ok = False
        if not ok:
            try:
                doc.Delete(dim.Id)
            except Exception:
                pass
            try:
                doc.Delete(dc.Id)
            except Exception:
                pass
            continue
        return True

    return False


def _delete_dim_marker(doc, marker, dim_id):
    if dim_id is not None and dim_id != ElementId.InvalidElementId:
        try:
            de = doc.GetElement(dim_id)
            if de is not None:
                doc.Delete(dim_id)
        except Exception:
            pass
    try:
        if marker is not None:
            doc.Delete(marker.Id)
    except Exception:
        pass
