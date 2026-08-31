# -*- coding: utf-8 -*-
"""Vigas de hormigón unidas a la selección principal.

Detecta Structural Framing de hormigón conectados a las vigas del lote
(principalmente vía ``JoinGeometryUtils``; respaldo por extremos cercanos)
y las clasifica según su eje respecto al plano de la vista activa:

- **paralelas**: ``|T · N_view| ≈ 0`` (mismo criterio de armado en alzado).
- **no paralelas**: transversales u oblicuas al alzado (típicas “en punta”
  o perpendiculares que empotran en la viga principal).

Todo el trabajo usa la API de Revit en el hilo de la UI (sin threads).
"""

from __future__ import division

import math

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BoundingBoxIntersectsFilter,
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    JoinGeometryUtils,
    LocationCurve,
    Outline,
    XYZ,
)

from armado_vigas.revit.view_order import (
    beam_axis_parallel_to_view_plane,
    beam_axis_tangent,
    _beam_endpoints,
)

# ~ 50 mm entre extremos (uniones por cara sin Join Geometry explícito).
_END_JOIN_TOL_FT = 50.0 / 304.8
# Coseno para colineales: |T_a · T_b| >= cos(8°) → se trata como misma cadena, no “transversal”.
_COLIN_DOT_MIN = math.cos(math.radians(8.0))

_FRAMING_CAT = int(BuiltInCategory.OST_StructuralFraming)


def _element_id_int(el):
    try:
        return int(el.Id.IntegerValue)
    except Exception:
        return None


def _is_structural_framing(el):
    try:
        if el is None or not isinstance(el, FamilyInstance):
            return False
        return int(el.Category.Id.IntegerValue) == _FRAMING_CAT
    except Exception:
        return False


def _is_concrete(el):
    try:
        from geometria_colision_vigas import material_estructural_es_concrete

        return material_estructural_es_concrete(el)
    except Exception:
        return False


def _label_beam(el):
    try:
        p = el.LookupParameter(u"Mark")
        if p and p.HasValue and p.AsString():
            s = (p.AsString() or u"").strip()
            if s:
                return s
    except Exception:
        pass
    eid = _element_id_int(el)
    return u"V-{0}".format(eid if eid is not None else u"?")


def _joined_element_ids(document, element):
    """IDs unidos por Join Geometry (API Revit 2024–2026)."""
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


def _xyz_dist(a, b):
    try:
        return float(a.DistanceTo(b))
    except Exception:
        try:
            return float((a - b).GetLength())
        except Exception:
            return 1e9


def _endpoint_join_score(host_el, cand_el, tol_ft=_END_JOIN_TOL_FT):
    """Distancia mínima extremo–extremo; True si ≤ tol (unión por encuentro)."""
    h0, h1 = _beam_endpoints(host_el)
    c0, c1 = _beam_endpoints(cand_el)
    if h0 is None or c0 is None:
        return False, None
    d_min = 1e9
    for hp in (h0, h1):
        for cp in (c0, c1):
            if hp is None or cp is None:
                continue
            d_min = min(d_min, _xyz_dist(hp, cp))
            # También: extremo del candidato sobre curva del host (empotra a mitad).
    # Distancia extremo candidato a curva host.
    try:
        loc = host_el.Location
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            crv = loc.Curve
            for cp in (c0, c1):
                if cp is None:
                    continue
                proj = crv.Project(cp)
                if proj is not None and proj.XYZPoint is not None:
                    d_min = min(d_min, _xyz_dist(cp, proj.XYZPoint))
    except Exception:
        pass
    # Simétrico: extremo host a curva candidato.
    try:
        loc = cand_el.Location
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            crv = loc.Curve
            for hp in (h0, h1):
                if hp is None:
                    continue
                proj = crv.Project(hp)
                if proj is not None and proj.XYZPoint is not None:
                    d_min = min(d_min, _xyz_dist(hp, proj.XYZPoint))
    except Exception:
        pass
    ok = d_min <= float(tol_ft)
    return ok, d_min if ok else None


def _axes_colinear(host_el, cand_el):
    """True si las tangentes son casi paralelas (misma cadena)."""
    th = beam_axis_tangent(host_el)
    tc = beam_axis_tangent(cand_el)
    if th is None or tc is None:
        return False
    try:
        return abs(float(th.DotProduct(tc))) >= _COLIN_DOT_MIN
    except Exception:
        return False


def _prefer_non_collinear_endpoint(host_el, cand_el, method_was_geometry):
    """
    Uniones por extremo colineales suelen ser continuación del tramo
    (ya capturables en la selección). Priorizamos transversales:
    si es solo endpoint y colineal, sigue válida pero se marca ``collinearJoin``.
    """
    if method_was_geometry:
        return False
    return _axes_colinear(host_el, cand_el)

def _framing_near_hosts(document, host_elements, pad_ft=None):
    """Recolector de Structural Framing en el entorno del lote (bbox unido)."""
    if document is None:
        return []
    pad = pad_ft if pad_ft is not None else (max(2.0, _END_JOIN_TOL_FT * 8.0))
    xs, ys, zs = [], [], []
    for el in host_elements or []:
        try:
            bb = el.get_BoundingBox(None)
            if bb is None:
                continue
            xs.extend([float(bb.Min.X), float(bb.Max.X)])
            ys.extend([float(bb.Min.Y), float(bb.Max.Y)])
            zs.extend([float(bb.Min.Z), float(bb.Max.Z)])
        except Exception:
            continue
    if not xs:
        # Fallback: todas las framing del documento es demasiado caro; vacío.
        return []
    try:
        outline = Outline(
            XYZ(min(xs) - pad, min(ys) - pad, min(zs) - pad),
            XYZ(max(xs) + pad, max(ys) + pad, max(zs) + pad),
        )
        filt = BoundingBoxIntersectsFilter(outline)
        return list(
            FilteredElementCollector(document)
            .OfCategory(BuiltInCategory.OST_StructuralFraming)
            .WhereElementIsNotElementType()
            .WherePasses(filt)
        )
    except Exception:
        try:
            return list(
                FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return []


def _empty_result():
    return {
        "all": [],
        "parallel": [],
        "not_parallel": [],
        "by_element_id": {},
        "counts": {
            "all": 0,
            "parallel": 0,
            "not_parallel": 0,
        },
    }


def _make_record(el, parallel, host_id, host_label, method, dist_ft=None, collinear=False):
    return {
        "elementIdInt": _element_id_int(el),
        "id": _label_beam(el),
        "element": el,
        "parallelToView": bool(parallel),
        "joinMethod": method,
        "collinearJoin": bool(collinear),
        "sourceBeamIdInts": [host_id] if host_id is not None else [],
        "sourceBeamLabels": [host_label] if host_label else [],
        "joinDistanceFt": dist_ft,
    }


def detect_joined_concrete_framing(document, host_framing, view=None):
    """
    Detecta vigas de hormigón unidas a ``host_framing`` (selección principal).

    :param document: Documento Revit
    :param host_framing: lista de ``FamilyInstance`` estructurales del lote
    :param view: vista activa (criterio paralelo al plano)
    :returns: dict con ``all``, ``parallel``, ``not_parallel``, ``counts``
    """
    result = _empty_result()
    hosts = [el for el in (host_framing or []) if _is_structural_framing(el)]
    if document is None or not hosts:
        return result

    host_ids = set()
    for el in hosts:
        eid = _element_id_int(el)
        if eid is not None:
            host_ids.add(eid)

    # elementIdInt → record (fusiona fuentes host).
    bucket = {}

    def _merge(record):
        eid = record.get("elementIdInt")
        if eid is None or eid in host_ids:
            return
        cur = bucket.get(eid)
        if cur is None:
            bucket[eid] = record
            return
        for hid in record.get("sourceBeamIdInts") or []:
            if hid not in cur["sourceBeamIdInts"]:
                cur["sourceBeamIdInts"].append(hid)
        for lab in record.get("sourceBeamLabels") or []:
            if lab not in cur["sourceBeamLabels"]:
                cur["sourceBeamLabels"].append(lab)
        # Preferir join geometry; distancia más corta si ambas son endpoint.
        if record.get("joinMethod") == u"geometry":
            cur["joinMethod"] = u"geometry"
        if record.get("joinDistanceFt") is not None:
            prev = cur.get("joinDistanceFt")
            if prev is None or float(record["joinDistanceFt"]) < float(prev):
                cur["joinDistanceFt"] = record["joinDistanceFt"]

    # 1) Join Geometry
    for host in hosts:
        host_id = _element_id_int(host)
        host_label = _label_beam(host)
        for jid in _joined_element_ids(document, host):
            try:
                el = document.GetElement(jid)
            except Exception:
                el = None
            if not _is_structural_framing(el):
                continue
            if not _is_concrete(el):
                continue
            eid = _element_id_int(el)
            if eid is None or eid in host_ids:
                continue
            parallel = beam_axis_parallel_to_view_plane(el, view)
            _merge(
                _make_record(
                    el, parallel, host_id, host_label, u"geometry", None,
                )
            )

    # 2) Extremos cercanos (uniones sin Join Geometry)
    candidates = _framing_near_hosts(document, hosts)
    for cand in candidates:
        if not _is_structural_framing(cand):
            continue
        if not _is_concrete(cand):
            continue
        ceid = _element_id_int(cand)
        if ceid is None or ceid in host_ids:
            continue
        for host in hosts:
            ok, dist = _endpoint_join_score(host, cand)
            if not ok:
                continue
            parallel = beam_axis_parallel_to_view_plane(cand, view)
            colin = _prefer_non_collinear_endpoint(host, cand, False)
            _merge(
                _make_record(
                    cand,
                    parallel,
                    _element_id_int(host),
                    _label_beam(host),
                    u"endpoint",
                    dist,
                    collinear=colin,
                )
            )

    all_recs = list(bucket.values())
    parallel = [r for r in all_recs if r.get("parallelToView")]
    not_par = [r for r in all_recs if not r.get("parallelToView")]

    result["all"] = all_recs
    result["parallel"] = parallel
    result["not_parallel"] = not_par
    result["by_element_id"] = {r["elementIdInt"]: r for r in all_recs if r.get("elementIdInt") is not None}
    result["counts"] = {
        "all": len(all_recs),
        "parallel": len(parallel),
        "not_parallel": len(not_par),
    }
    return result


def format_joined_summary(joined_result):
    """Texto corto para status / UI."""
    if not joined_result:
        return u""
    c = joined_result.get("counts") or {}
    n_all = int(c.get("all") or 0)
    if n_all <= 0:
        return u""
    n_p = int(c.get("parallel") or 0)
    n_np = int(c.get("not_parallel") or 0)
    return (
        u" · unidas: {0} ({1} paralela(s) vista · {2} no paralela(s))"
        .format(n_all, n_p, n_np)
    )


def not_parallel_joined_labels(joined_result, limit=8):
    """Lista de marcas de unidas no paralelas (para resumen UI)."""
    recs = (joined_result or {}).get("not_parallel") or []
    labels = []
    for r in recs[: max(0, int(limit))]:
        labels.append(r.get("id") or u"?")
    more = len(recs) - len(labels)
    if more > 0:
        labels.append(u"+{0}".format(more))
    return labels
