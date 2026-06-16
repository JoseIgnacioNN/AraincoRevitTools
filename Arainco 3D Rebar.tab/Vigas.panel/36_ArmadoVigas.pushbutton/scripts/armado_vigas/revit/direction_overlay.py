# -*- coding: utf-8 -*-
"""Marcadores temporales de dirección de eje de viga (LocationCurve 0 → 1)."""

from __future__ import division

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Line, LocationCurve, Plane, SketchPlane, Transaction, XYZ
from Autodesk.Revit.UI import IExternalEventHandler

from armado_vigas.domain.tramos import sort_beams
from armado_vigas.revit.session import SESSION

_MM_TO_FT = 1.0 / 304.8
_ARROW_LEG_MM = 140.0
_START_TICK_MM = 90.0
_TOP_CLEAR_MM = 60.0
_SHAFT_LEN_MM = 360.0
_SHAFT_MIN_MM = 80.0
# Solo dibujar ``ModelCurve`` en el modelo si se invoca explícitamente
# ``show_beam_direction_overlay`` (p. ej. acción futura en UI). No al abrir la herramienta.
_AUTO_SHOW_DIRECTION_OVERLAY_ON_LAUNCH = False


def _mm_to_ft(mm):
    return float(mm) * _MM_TO_FT


def _sketch_plane_for_line(line, n_hint=None):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        t = (p1 - p0).Normalize()
    except Exception:
        return None
    bn = None
    if n_hint is not None:
        try:
            bn = t.CrossProduct(n_hint)
            if bn.GetLength() < 1e-12:
                bn = None
        except Exception:
            bn = None
    if bn is None or bn.GetLength() < 1e-12:
        try:
            bn = t.CrossProduct(XYZ.BasisZ)
            if bn.GetLength() < 1e-12:
                bn = t.CrossProduct(XYZ.BasisX)
        except Exception:
            return None
    try:
        bn = bn.Normalize()
        return Plane.CreateByNormalAndOrigin(bn, p0)
    except Exception:
        return None


def _create_model_curve(document, line, n_hint=None):
    pl = _sketch_plane_for_line(line, n_hint)
    if pl is None:
        return None
    try:
        sp = SketchPlane.Create(document, pl)
        mc = document.Create.NewModelCurve(line, sp)
        return mc.Id if mc is not None else None
    except Exception:
        return None


def _perp_in_plane(axis, n_hint):
    try:
        perp = axis.CrossProduct(n_hint)
        if perp.GetLength() < 1e-9:
            perp = axis.CrossProduct(XYZ.BasisX)
        return perp.Normalize()
    except Exception:
        return None


def _beam_top_axis_line(document, elem):
    try:
        loc = elem.Location
        if not isinstance(loc, LocationCurve):
            return None
        curve = loc.Curve
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        try:
            from armadura_vigas_capas import _beam_frame, _read_width_depth_ft

            frame = _beam_frame(curve)
            if frame is not None:
                axis, width_dir, depth_dir, p0c, p1c, beam_len = frame
                if beam_len > 1e-9:
                    _w_ft, d_ft = _read_width_depth_ft(document, elem, curve)
                    lift = depth_dir * (float(d_ft) * 0.5 + _mm_to_ft(_TOP_CLEAR_MM))
                    return Line.CreateBound(p0c + lift, p1c + lift)
        except Exception:
            pass
        lift = XYZ(0.0, 0.0, _mm_to_ft(_TOP_CLEAR_MM))
        return Line.CreateBound(p0 + lift, p1 + lift)
    except Exception:
        return None


def _arrow_wing_lines(p_tip, axis, n_hint, leg_ft):
    perp = _perp_in_plane(axis, n_hint)
    if perp is None:
        return []
    back = axis.Negate()
    wing1 = (back + perp * 0.42).Normalize()
    wing2 = (back - perp * 0.42).Normalize()
    return [
        Line.CreateBound(p_tip, p_tip + wing1 * leg_ft),
        Line.CreateBound(p_tip, p_tip + wing2 * leg_ft),
    ]


def _start_tick_line(p0, axis, n_hint, half_len_ft):
    perp = _perp_in_plane(axis, n_hint)
    if perp is None:
        return None
    return Line.CreateBound(p0 - perp * half_len_ft, p0 + perp * half_len_ft)


def _markers_for_beam(document, elem):
    axis_line = _beam_top_axis_line(document, elem)
    if axis_line is None:
        return []
    p0 = axis_line.GetEndPoint(0)
    p1 = axis_line.GetEndPoint(1)
    raw = p1 - p0
    if raw.GetLength() < 1e-9:
        return []
    axis = raw.Normalize()
    n_hint = XYZ.BasisZ
    leg_ft = _mm_to_ft(_ARROW_LEG_MM)
    beam_len = float(raw.GetLength())
    shaft_len_ft = min(_mm_to_ft(_SHAFT_LEN_MM), beam_len * 0.45)
    if shaft_len_ft < _mm_to_ft(_SHAFT_MIN_MM):
        shaft_len_ft = min(beam_len * 0.5, _mm_to_ft(_SHAFT_MIN_MM))
    half_shaft_ft = shaft_len_ft * 0.5
    p_mid = p0 + axis * (beam_len * 0.5)
    p_start = p_mid - axis * half_shaft_ft
    p_tip = p_mid + axis * half_shaft_ft
    arrow_back_ft = min(leg_ft * 0.55, half_shaft_ft * 0.35)
    p_shaft_end = p_tip - axis * arrow_back_ft

    ids = []
    shaft_id = _create_model_curve(document, Line.CreateBound(p_start, p_shaft_end), n_hint)
    if shaft_id is not None:
        ids.append(shaft_id)
    for wing in _arrow_wing_lines(p_tip, axis, n_hint, leg_ft):
        wid = _create_model_curve(document, wing, n_hint)
        if wid is not None:
            ids.append(wid)
    tick = _start_tick_line(p_start, axis, n_hint, _mm_to_ft(_START_TICK_MM) * 0.5)
    if tick is not None:
        tid = _create_model_curve(document, tick, n_hint)
        if tid is not None:
            ids.append(tid)
    return ids


def clear_beam_direction_overlay(document):
    ids = list(getattr(SESSION, "direction_overlay_ids", None) or [])
    SESSION.direction_overlay_ids = []
    SESSION.direction_overlay_view_id = None
    if not ids or document is None:
        return 0
    deleted = 0
    t = Transaction(document, u"Arainco: Limpiar dirección vigas")
    t.Start()
    try:
        for eid in ids:
            try:
                el = document.GetElement(eid)
                if el is not None and el.IsValidObject:
                    document.Delete(eid)
                    deleted += 1
            except Exception:
                pass
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
    return deleted


def show_beam_direction_overlay(document, view=None):
    """
    Dibuja ``ModelCurve`` sobre la cara superior de cada viga del lote.

    Convención: punto 0 de ``LocationCurve`` (marca ⊥ en inicio del indicador) → punto 1
    (punta de flecha). El trazo es corto y centrado sobre el eje de la viga.
    """
    if document is None:
        return 0
    clear_beam_direction_overlay(document)
    beams = sort_beams(list(SESSION.domain_beams or []))
    if not beams:
        return 0

    ids = []
    t = Transaction(document, u"Arainco: Dirección vigas (overlay)")
    t.Start()
    try:
        for beam in beams:
            el = beam.get("element")
            if el is None:
                continue
            ids.extend(_markers_for_beam(document, el))
        SESSION.direction_overlay_ids = ids
        try:
            SESSION.direction_overlay_view_id = view.Id if view is not None else None
        except Exception:
            SESSION.direction_overlay_view_id = None
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        SESSION.direction_overlay_ids = []
        raise
    return len(ids)


class ClearDirectionOverlayHandler(IExternalEventHandler):
    def Execute(self, uiapp):
        uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
        if uidoc is None:
            return
        clear_beam_direction_overlay(uidoc.Document)

    def GetName(self):
        return u"ArmadoVigasClearDirectionOverlay"
