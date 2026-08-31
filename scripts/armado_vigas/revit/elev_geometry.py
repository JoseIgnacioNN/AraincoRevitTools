# -*- coding: utf-8 -*-
"""Geometría de alzado proyectada sobre la vista activa (Right × Up).

Cada elemento se representa por su AABB proyectado en el plano de la vista:
- U: escalar sobre ``view.RightDirection`` (eje horizontal del canvas)
- V: escalar sobre ``view.UpDirection`` (eje vertical del canvas; V mayor = más alto)
"""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import LocationCurve, XYZ

from armado_vigas.revit.view_order import (
    _beam_endpoints,
    _scalar_on_axis,
    _unit_vector,
    view_right_unit,
)

_FT_TO_MM = 304.8


def view_up_unit(view):
    if view is None:
        return None
    try:
        return _unit_vector(view.UpDirection)
    except Exception:
        return None


def _bbox_for_view(el, view):
    if el is None:
        return None
    bb = None
    if view is not None:
        try:
            bb = el.get_BoundingBox(view)
        except Exception:
            bb = None
    if bb is None:
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
    return bb


def _bbox_corners(bb):
    if bb is None:
        return []
    try:
        mn, mx = bb.Min, bb.Max
        xs = (float(mn.X), float(mx.X))
        ys = (float(mn.Y), float(mx.Y))
        zs = (float(mn.Z), float(mx.Z))
    except Exception:
        return []
    pts = []
    for x in xs:
        for y in ys:
            for z in zs:
                try:
                    pts.append(XYZ(x, y, z))
                except Exception:
                    pass
    return pts


def element_view_extents(el, view):
    """
    Extensión proyectada del sólido en la vista activa.

    Returns dict o ``None``::
        uMin, uMax, vMin, vMax  (pies, modelo)
        uMid, vMid, uSpanFt, vSpanFt
        widthMm, heightMm
    """
    if el is None or view is None:
        return None
    right = view_right_unit(view)
    up = view_up_unit(view)
    if right is None or up is None:
        return None

    bb = _bbox_for_view(el, view)
    corners = _bbox_corners(bb)
    if not corners:
        return None

    us = []
    vs = []
    for pt in corners:
        u = _scalar_on_axis(pt, right)
        v = _scalar_on_axis(pt, up)
        if u is not None:
            us.append(u)
        if v is not None:
            vs.append(v)
    if not us or not vs:
        return None

    u_min, u_max = min(us), max(us)
    v_min, v_max = min(vs), max(vs)
    u_span = max(0.0, u_max - u_min)
    v_span = max(0.0, v_max - v_min)
    return {
        "uMin": u_min,
        "uMax": u_max,
        "vMin": v_min,
        "vMax": v_max,
        "uMid": (u_min + u_max) * 0.5,
        "vMid": (v_min + v_max) * 0.5,
        "uSpanFt": u_span,
        "vSpanFt": v_span,
        "widthMm": int(round(u_span * _FT_TO_MM)),
        "heightMm": int(round(v_span * _FT_TO_MM)),
    }


def enrich_apoyo_view_geometry(apoyo, view):
    """Rellena métricas proyectadas en un dict de apoyo.

    En muros: ``thicknessMm`` = espesor real (``Wall.Width``); proyectado en
    ``widthMm``. ``parallelToView`` = eje // al plano de la vista activa.

    En losas: ``thicknessMm`` ≈ espesor (altura V o compound structure);
    se proyecta como banda horizontal en el canvas de elevación.
    """
    if not apoyo:
        return apoyo
    el = apoyo.get("element")
    ext = element_view_extents(el, view)
    if not ext:
        return apoyo
    apoyo["uMin"] = ext["uMin"]
    apoyo["uMax"] = ext["uMax"]
    apoyo["vMin"] = ext["vMin"]
    apoyo["vMax"] = ext["vMax"]
    apoyo["uView"] = ext["uMid"]
    apoyo["widthMm"] = max(1, int(ext["widthMm"] or 1))
    apoyo["heightMm"] = max(1, int(ext["heightMm"] or 1))
    kind = unicode(apoyo.get("kind") or u"").lower()
    if kind == "wall":
        th_mm = None
        try:
            if el is not None:
                th_mm = float(el.Width) * _FT_TO_MM
        except Exception:
            th_mm = None
        if th_mm is None or th_mm < 1.0:
            th_mm = float(apoyo["widthMm"])
        apoyo["thicknessMm"] = max(1, int(round(th_mm)))
        # Eje del muro // plano de vista (misma convención que framing).
        try:
            from armado_vigas.geometry.retract_muros_noparalelos import (
                element_is_parallel_to_view_plane,
            )

            apoyo["parallelToView"] = bool(
                element_is_parallel_to_view_plane(el, view)
            )
        except Exception:
            try:
                from armado_vigas.revit.view_order import beam_axis_parallel_to_view_plane

                # Fallback: si hay LocationCurve del muro.
                apoyo["parallelToView"] = bool(
                    beam_axis_parallel_to_view_plane(el, view)
                )
            except Exception:
                apoyo["parallelToView"] = False
    elif kind in (u"floor", u"losa", u"slab"):
        th_mm = _floor_thickness_mm(el)
        if th_mm is None or th_mm < 1.0:
            th_mm = float(apoyo.get("heightMm") or 0) or float(apoyo.get("widthMm") or 0)
        apoyo["thicknessMm"] = max(1, int(round(th_mm)))
        apoyo.setdefault("parallelToView", False)
    else:
        apoyo.setdefault("parallelToView", False)
    return apoyo


def _floor_thickness_mm(el):
    """Espesor de losa (mm) desde compound structure o type parameter."""
    if el is None:
        return None
    try:
        from Autodesk.Revit.DB import Floor

        if not isinstance(el, Floor):
            # Categoría OST_Floors sin tipo Floor manejado: intentar GetTypeId.
            pass
    except Exception:
        pass
    try:
        ft = el.FloorType if hasattr(el, u"FloorType") else None
        if ft is None:
            try:
                tid = el.GetTypeId()
                if tid is not None:
                    ft = el.Document.GetElement(tid)
            except Exception:
                ft = None
        if ft is not None:
            try:
                cs = ft.GetCompoundStructure()
                if cs is not None:
                    return float(cs.GetWidth()) * _FT_TO_MM
            except Exception:
                pass
            try:
                from Autodesk.Revit.DB import BuiltInParameter

                p = ft.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM)
                if p is not None and p.HasValue:
                    return float(p.AsDouble()) * _FT_TO_MM
            except Exception:
                pass
    except Exception:
        pass
    return None


def enrich_beam_view_geometry(beam, view, document=None):
    """Rellena u/v de solid + sección; conserva uStart/uEnd de LocationCurve si ya existen."""
    if not beam:
        return beam
    el = beam.get("element")
    ext = element_view_extents(el, view)
    if ext:
        beam["solidUMin"] = ext["uMin"]
        beam["solidUMax"] = ext["uMax"]
        beam["vMin"] = ext["vMin"]
        beam["vMax"] = ext["vMax"]
        beam["widthMm"] = max(1, int(ext["widthMm"] or 1))
        beam["heightMm"] = max(1, int(ext["heightMm"] or 1))
        # Si aún no hay rangos de eje, usar proyección del sólido.
        if beam.get("uStart") is None and beam.get("uEnd") is None:
            beam["uStart"] = ext["uMin"]
            beam["uEnd"] = ext["uMax"]

    # Ajuste longitudinal: extremos LocationCurve en Right (más fiel al eje estructural).
    if el is not None and view is not None:
        right = view_right_unit(view)
        p0, p1 = _beam_endpoints(el)
        if right is not None and p0 is not None and p1 is not None:
            u0 = _scalar_on_axis(p0, right)
            u1 = _scalar_on_axis(p1, right)
            if u0 is not None and u1 is not None:
                beam["uStart"] = min(u0, u1)
                beam["uEnd"] = max(u0, u1)
                beam["axisReversed"] = bool(u0 > u1)

    # Sección transversal nominal (cm) para etiqueta / fallback vertical.
    try:
        from armado_vigas.revit.adapters import _beam_type_label, _read_width_depth_ft

        curve = None
        try:
            loc = el.Location if el is not None else None
            if isinstance(loc, LocationCurve):
                curve = loc.Curve
        except Exception:
            curve = None
        if document is not None and el is not None:
            w_ft, d_ft = _read_width_depth_ft(document, el, curve)
            beam["sectionWidthMm"] = int(round(float(w_ft) * _FT_TO_MM))
            beam["sectionDepthMm"] = int(round(float(d_ft) * _FT_TO_MM))
            if not beam.get("type"):
                beam["type"] = _beam_type_label(document, el, curve)
    except Exception:
        pass

    return beam


def enrich_joined_framing_geometry(joined_result, view, document=None):
    """Proyecta cada viga unida al plano de vista (U/V) para dibujar en alzado."""
    if not joined_result:
        return joined_result
    recs = joined_result.get("all") if isinstance(joined_result, dict) else joined_result
    for rec in recs or []:
        el = rec.get("element")
        ext = element_view_extents(el, view)
        if ext:
            rec["solidUMin"] = ext["uMin"]
            rec["solidUMax"] = ext["uMax"]
            rec["uMin"] = ext["uMin"]
            rec["uMax"] = ext["uMax"]
            rec["vMin"] = ext["vMin"]
            rec["vMax"] = ext["vMax"]
            rec["widthMm"] = max(1, int(ext["widthMm"] or 1))
            rec["heightMm"] = max(1, int(ext["heightMm"] or 1))
        # Eje LocationCurve en Right (útil si es // a la vista).
        if el is not None and view is not None:
            right = view_right_unit(view)
            p0, p1 = _beam_endpoints(el)
            if right is not None and p0 is not None and p1 is not None:
                u0 = _scalar_on_axis(p0, right)
                u1 = _scalar_on_axis(p1, right)
                if u0 is not None and u1 is not None:
                    rec["uStart"] = min(u0, u1)
                    rec["uEnd"] = max(u0, u1)
                    rec["axisReversed"] = bool(u0 > u1)
        try:
            from armado_vigas.revit.adapters import _beam_type_label, _read_width_depth_ft

            curve = None
            try:
                loc = el.Location if el is not None else None
                if isinstance(loc, LocationCurve):
                    curve = loc.Curve
            except Exception:
                curve = None
            if document is not None and el is not None:
                w_ft, d_ft = _read_width_depth_ft(document, el, curve)
                rec["sectionWidthMm"] = int(round(float(w_ft) * _FT_TO_MM))
                rec["sectionDepthMm"] = int(round(float(d_ft) * _FT_TO_MM))
                rec["type"] = _beam_type_label(document, el, curve)
        except Exception:
            pass
    return joined_result


def enrich_session_elev_geometry(beams, apoyos, view, document=None):
    """Enriquece listados de dominio con proyección en vista."""
    for ap in apoyos or []:
        enrich_apoyo_view_geometry(ap, view)
    for beam in beams or []:
        enrich_beam_view_geometry(beam, view, document=document)
    return beams, apoyos


def _is_vertical_support_apoyo(ap):
    """Columnas/muros definen extremos; las losas solo se dibujan en alzado."""
    if not ap:
        return False
    try:
        kind = unicode(ap.get("kind") or u"").lower()
    except Exception:
        kind = u""
    if kind in (u"floor", u"losa", u"slab"):
        return False
    return True


def assign_beam_supports_by_proximity(beams, apoyos, view=None):
    """
    Asigna colStart/colEnd al apoyo más cercano a cada extremo de viga
    (proyección U), en lugar de encadenar por rango de selección.

    Las losas (kind floor) se ignoran: solo afectan el dibujo del canvas.
    """
    beams = list(beams or [])
    apoyos = [a for a in (apoyos or []) if _is_vertical_support_apoyo(a)]
    if not beams:
        return beams
    if not apoyos:
        for beam in beams:
            beam["colStart"] = u""
            beam["colEnd"] = u""
        return beams

    # Asegurar uView / uMid en apoyos
    ap_pts = []
    for ap in apoyos:
        u = ap.get("uView")
        if u is None:
            umin, umax = ap.get("uMin"), ap.get("uMax")
            if umin is not None and umax is not None:
                u = (float(umin) + float(umax)) * 0.5
        if u is None:
            continue
        ap_pts.append((ap.get("id"), float(u)))
    if not ap_pts:
        for beam in beams:
            beam["colStart"] = u""
            beam["colEnd"] = u""
        return beams

    def nearest_id(u_target):
        best_id = ap_pts[0][0]
        best_d = abs(ap_pts[0][1] - u_target)
        for aid, uu in ap_pts:
            d = abs(uu - u_target)
            if d < best_d:
                best_d = d
                best_id = aid
        return best_id

    for beam in beams:
        try:
            u0 = float(beam.get("uStart"))
            u1 = float(beam.get("uEnd"))
        except (TypeError, ValueError):
            beam["colStart"] = ap_pts[0][0]
            beam["colEnd"] = ap_pts[-1][0]
            continue
        # Extremos en pantalla: u menor = izquierda (ini), u mayor = fin
        beam["colStart"] = nearest_id(min(u0, u1))
        beam["colEnd"] = nearest_id(max(u0, u1))
    return beams


def model_uv_span(beams, apoyos):
    """Rango global [u_min,u_max] y [v_min,v_max] del lote en la vista."""
    us = []
    vs = []
    for beam in beams or []:
        for key in ("uStart", "uEnd", "solidUMin", "solidUMax"):
            v = beam.get(key)
            if v is None:
                continue
            try:
                us.append(float(v))
            except (TypeError, ValueError):
                pass
        for key in ("vMin", "vMax"):
            v = beam.get(key)
            if v is None:
                continue
            try:
                vs.append(float(v))
            except (TypeError, ValueError):
                pass
    for ap in apoyos or []:
        for key in ("uMin", "uMax", "uView"):
            v = ap.get(key)
            if v is None:
                continue
            try:
                us.append(float(v))
            except (TypeError, ValueError):
                pass
        for key in ("vMin", "vMax"):
            v = ap.get(key)
            if v is None:
                continue
            try:
                vs.append(float(v))
            except (TypeError, ValueError):
                pass
    if not us:
        return None
    u_min, u_max = min(us), max(us)
    if u_max - u_min < 1e-9:
        u_max = u_min + 1e-9
    if vs:
        v_min, v_max = min(vs), max(vs)
        if v_max - v_min < 1e-9:
            # Fallback: usar profundidad de sección típica
            v_max = v_min + 2.0  # ~60 cm
    else:
        v_min, v_max = 0.0, 2.0
    return {
        "uMin": u_min,
        "uMax": u_max,
        "vMin": v_min,
        "vMax": v_max,
    }
