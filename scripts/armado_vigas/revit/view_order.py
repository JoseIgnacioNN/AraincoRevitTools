# -*- coding: utf-8 -*-
"""Orden de vigas en canvas según la vista activa (patrón Armado Muros).

Opciones evaluadas:
  A) Proyectar el punto medio (o u_start) de cada LocationCurve sobre
     ``view.RightDirection`` y ordenar por escalar — elegida: replica
     ``compute_stacked_wall_layout(view_right_xy)`` de Armado Muros y alinea
     izquierda/derecha del canvas con la vista de Revit.
  B) Parámetro sobre eje unificado de cadena colineal — más preciso para
     tramos largos pero requiere detectar cadenas; innecesario para slots del canvas.
  C) Extraer helper de Muros a módulo compartido — acoplamiento y Vigas necesita
     proyección 3D (no solo XY).
  D) Comparar u0 vs u1 para coherencia de flecha 0→1 — se usa como
     ``axisReversed`` por viga, no como criterio global de orden.
"""

from __future__ import division

import math

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import LocationCurve

# Eje ⟂ normal de vista ⟺ eje paralelo al plano de la vista (≈ 2°).
_TOL_DOT_EJE_PARALELO_PLANO_VISTA = math.sin(math.radians(2.0))


def _unit_vector(vec):
    try:
        ln = float(vec.GetLength())
        if ln < 1e-12:
            return None
        return vec.Divide(ln)
    except Exception:
        return None


def view_right_unit(view):
    """``RightDirection`` normalizado de la vista, o ``None`` si no aplica."""
    if view is None:
        return None
    try:
        return _unit_vector(view.RightDirection)
    except Exception:
        return None


def view_normal_unit(view):
    """Normal unitaria del plano de la vista activa (``ViewDirection``)."""
    if view is None:
        return None
    try:
        vd = view.ViewDirection
        if vd is not None and float(vd.GetLength()) > 1e-12:
            return vd.Normalize()
    except Exception:
        pass
    return None


def beam_axis_tangent(elem):
    """Tangente unitaria del eje de la viga (``LocationCurve``)."""
    p0, p1 = _beam_endpoints(elem)
    if p0 is None or p1 is None:
        return None
    try:
        return _unit_vector(p1 - p0)
    except Exception:
        return None


def beam_axis_parallel_to_view_plane(elem, view):
    """
    True si el eje de la viga es paralelo al plano de la vista activa.

    Criterio: ``|T · N_view| <= sin(2°)`` con ``T`` tangente unitaria y ``N_view`` normal de vista.
    """
    n_view = view_normal_unit(view)
    if n_view is None:
        return True
    tang = beam_axis_tangent(elem)
    if tang is None:
        return False
    try:
        return abs(float(tang.DotProduct(n_view))) <= _TOL_DOT_EJE_PARALELO_PLANO_VISTA
    except Exception:
        return False


def _beam_endpoints(elem):
    try:
        loc = elem.Location
        if not isinstance(loc, LocationCurve):
            return None, None
        crv = loc.Curve
        return crv.GetEndPoint(0), crv.GetEndPoint(1)
    except Exception:
        return None, None


def _scalar_on_axis(point, axis):
    try:
        return (
            float(point.X) * float(axis.X)
            + float(point.Y) * float(axis.Y)
            + float(point.Z) * float(axis.Z)
        )
    except Exception:
        return None


def beam_layout_on_view(elem, view):
    """
    Posición de la viga sobre el eje horizontal de la vista.

    Devuelve ``u_start`` (extremo izquierdo en vista), ``u_mid`` y si el
    punto 0 de LocationCurve queda a la derecha del 1 (``axis_reversed``).
    """
    axis = view_right_unit(view)
    if axis is None:
        return None
    p0, p1 = _beam_endpoints(elem)
    if p0 is None or p1 is None:
        return None
    u0 = _scalar_on_axis(p0, axis)
    u1 = _scalar_on_axis(p1, axis)
    if u0 is None or u1 is None:
        return None
    u_start = min(u0, u1)
    u_end = max(u0, u1)
    return {
        "u_start": u_start,
        "u_end": u_end,
        "u_mid": (u0 + u1) * 0.5,
        "u0": u0,
        "u1": u1,
        "axis_reversed": u0 > u1,
    }


def assign_beam_view_order(beams, view=None):
    """
    Asigna ``beam['u']`` para :func:`armado_vigas.domain.tramos.sort_beams`.

    Con vista: orden físico izquierda→derecha según ``RightDirection``.
    Sin vista: conserva el orden de entrada (índice de selección).
    """
    beams = list(beams or [])
    if not beams:
        return beams

    axis = view_right_unit(view)
    if axis is None:
        for i, beam in enumerate(beams):
            beam["u"] = i
            beam["axisReversed"] = False
            beam["uStart"] = None
            beam["uEnd"] = None
        return beams

    layouts = []
    for i, beam in enumerate(beams):
        el = beam.get("element")
        layout = beam_layout_on_view(el, view) if el is not None else None
        if layout is None:
            layout = {
                "u_start": 1e9 + float(i),
                "u_mid": 1e9 + float(i),
                "u_end": 1e9 + float(i),
                "axis_reversed": False,
            }
        layouts.append(layout)
        beam["axisReversed"] = bool(layout.get("axis_reversed"))
        beam["uStart"] = layout.get("u_start")
        beam["uEnd"] = layout.get("u_end")

    order = sorted(
        range(len(beams)),
        key=lambda idx: (layouts[idx]["u_start"], layouts[idx]["u_mid"], idx),
    )
    for rank, idx in enumerate(order):
        beams[idx]["u"] = rank
    return beams


def _apoyo_reference_point(el):
    """Punto de referencia del apoyo para proyección en vista."""
    if el is None:
        return None
    try:
        loc = el.Location
        if loc is not None and hasattr(loc, "Point"):
            return loc.Point
        if isinstance(loc, LocationCurve):
            crv = loc.Curve
            return crv.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        bb = el.get_BoundingBox(None)
        if bb is not None:
            from Autodesk.Revit.DB import XYZ

            return XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                (bb.Min.Z + bb.Max.Z) * 0.5,
            )
    except Exception:
        pass
    return None


def assign_apoyo_view_order(apoyos, view=None):
    """Asigna ``apoyo['u']`` según posición en ``view.RightDirection``."""
    apoyos = list(apoyos or [])
    if not apoyos:
        return apoyos

    axis = view_right_unit(view)
    if axis is None:
        for i, apoyo in enumerate(apoyos):
            apoyo["u"] = i
            apoyo["uView"] = None
        return apoyos

    scalars = []
    for i, apoyo in enumerate(apoyos):
        pt = _apoyo_reference_point(apoyo.get("element"))
        u = _scalar_on_axis(pt, axis) if pt is not None else None
        scalars.append(u if u is not None else 1e9 + float(i))
        apoyo["uView"] = u

    order = sorted(range(len(apoyos)), key=lambda idx: (scalars[idx], idx))
    for rank, idx in enumerate(order):
        apoyos[idx]["u"] = rank
    return apoyos


def assign_beam_col_endpoints(beams, apoyos, view=None):
    """
    Asigna ``colStart`` / ``colEnd`` por orden espacial (rank ``u``), no por
    índice de selección. Debe llamarse después de :func:`assign_beam_view_order`.
    """
    beams = list(beams or [])
    apoyos = assign_apoyo_view_order(list(apoyos or []), view)
    sorted_apoyos = sorted(apoyos, key=lambda a: a.get("u", 0))
    sorted_beams = sorted(beams, key=lambda b: b.get("u", 0))
    n = len(sorted_apoyos)
    if not n:
        for beam in sorted_beams:
            beam["colStart"] = u""
            beam["colEnd"] = u""
        return beams

    for rank, beam in enumerate(sorted_beams):
        beam["colStart"] = sorted_apoyos[rank % n]["id"]
        if n > 1:
            beam["colEnd"] = sorted_apoyos[(rank + 1) % n]["id"]
        else:
            beam["colEnd"] = sorted_apoyos[0]["id"]
    return beams
