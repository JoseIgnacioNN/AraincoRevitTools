# -*- coding: utf-8 -*-
"""Proyección de elementos sobre el plano de la vista activa (elevación fiel)."""

from __future__ import division

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import XYZ

_FT_TO_M = 0.3048
_FT_TO_CM = 30.48
_FT_TO_MM = 304.8


def ft_to_m(ft):
    return float(ft) * _FT_TO_M


def ft_to_cm(ft):
    return float(ft) * _FT_TO_CM


def ft_to_mm(ft):
    return float(ft) * _FT_TO_MM


def _unit(vec):
    if vec is None:
        return None
    try:
        ln = float(vec.GetLength())
        if ln < 1e-12:
            return None
        return vec.Divide(ln)
    except Exception:
        return None


def view_basis(view):
    """
    Ejes de la vista activa: origen + Right (horizontal canvas) + Up (vertical canvas).

    Devuelve ``None`` si la vista no aporta ejes (p. ej. sin documento).
    """
    if view is None:
        return None
    try:
        right = _unit(view.RightDirection)
        up = _unit(view.UpDirection)
        origin = view.Origin
    except Exception:
        return None
    if right is None or up is None or origin is None:
        return None
    return {
        "origin": origin,
        "right": right,
        "up": up,
    }


def project_point_uv(point, basis):
    """Proyecta un punto 3D a ``(u, v)`` en pies sobre Right/Up de la vista."""
    if point is None or not basis:
        return None
    try:
        o = basis["origin"]
        r = basis["right"]
        up = basis["up"]
        dx = float(point.X) - float(o.X)
        dy = float(point.Y) - float(o.Y)
        dz = float(point.Z) - float(o.Z)
        u = dx * float(r.X) + dy * float(r.Y) + dz * float(r.Z)
        v = dx * float(up.X) + dy * float(up.Y) + dz * float(up.Z)
        return u, v
    except Exception:
        return None


def _bbox_corners(bb):
    if bb is None:
        return []
    try:
        x0, y0, z0 = float(bb.Min.X), float(bb.Min.Y), float(bb.Min.Z)
        x1, y1, z1 = float(bb.Max.X), float(bb.Max.Y), float(bb.Max.Z)
    except Exception:
        return []
    pts = []
    for x in (x0, x1):
        for y in (y0, y1):
            for z in (z0, z1):
                pts.append(XYZ(x, y, z))
    return pts


def element_bbox(elem, view=None):
    """BoundingBox en vista (si hay) o del modelo."""
    if elem is None:
        return None
    bb = None
    if view is not None:
        try:
            bb = elem.get_BoundingBox(view)
        except Exception:
            bb = None
    if bb is None:
        try:
            bb = elem.get_BoundingBox(None)
        except Exception:
            bb = None
    return bb


def elevation_rect_from_element(elem, view):
    """
    Rectángulo de elevación del elemento en coordenadas de vista (pies).

    ``u`` = RightDirection, ``v`` = UpDirection.
    Devuelve dict con uMin/uMax/vMin/vMax o ``None``.
    """
    basis = view_basis(view)
    if basis is None:
        return None
    bb = element_bbox(elem, view)
    corners = _bbox_corners(bb)
    if not corners:
        return None
    us = []
    vs = []
    for pt in corners:
        uv = project_point_uv(pt, basis)
        if uv is None:
            continue
        us.append(uv[0])
        vs.append(uv[1])
    if not us or not vs:
        return None
    u_min, u_max = min(us), max(us)
    v_min, v_max = min(vs), max(vs)
    # Evitar degenerados
    if abs(u_max - u_min) < 1e-6:
        u_max = u_min + 0.1
    if abs(v_max - v_min) < 1e-6:
        v_max = v_min + 0.1
    return {
        "uMin": u_min,
        "uMax": u_max,
        "vMin": v_min,
        "vMax": v_max,
        "uMid": 0.5 * (u_min + u_max),
        "vMid": 0.5 * (v_min + v_max),
        "spanU_ft": u_max - u_min,
        "spanV_ft": v_max - v_min,
    }


def union_uv_span(rects):
    """``(u_min, u_max, v_min, v_max)`` del conjunto, o ``None``."""
    us0, us1, vs0, vs1 = [], [], [], []
    for r in rects or []:
        if not r:
            continue
        try:
            us0.append(float(r["uMin"]))
            us1.append(float(r["uMax"]))
            vs0.append(float(r["vMin"]))
            vs1.append(float(r["vMax"]))
        except Exception:
            continue
    if not us0:
        return None
    return min(us0), max(us1), min(vs0), max(vs1)
