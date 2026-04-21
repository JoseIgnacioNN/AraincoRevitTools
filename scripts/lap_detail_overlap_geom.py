# -*- coding: utf-8 -*-
"""
Geometría del tramo de solape entre dos Rebar colineales (troceo con traslapo).
Usado para reposicionar Detail Components line-based al cambiar las barras base.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import XYZ
from Autodesk.Revit.DB.Structure import Rebar

# Longitud mínima del solape (~1 mm en pies) para considerar el tramo válido.
_MIN_OVERLAP_FT = 1.0 / 304.8


def _unit_3d(v):
    if v is None:
        return None
    try:
        ln = float(v.GetLength())
    except Exception:
        return None
    if ln < 1e-12:
        return None
    try:
        return v.Normalize()
    except Exception:
        return None


def _dominant_centerline_curve(rebar):
    """Curva de centro dominante (misma idea que enfierrado_shaft_hashtag._rebar_centerline_dominant_curve)."""
    if rebar is None or not isinstance(rebar, Rebar):
        return None
    center_curves = []
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        mpo = getattr(MultiplanarOption, "IncludeOnlyPlanarCurves", None)
        if mpo is None:
            mpo = getattr(MultiplanarOption, "IncludeAllMultiplanarCurves", None)
        if mpo is not None:
            center_curves = list(rebar.GetCenterlineCurves(False, False, False, mpo, 0))
    except Exception:
        center_curves = []

    if not center_curves:
        try:
            center_curves = list(rebar.GetCenterlineCurves(False, False, False))
        except Exception:
            center_curves = []

    if center_curves:
        best = None
        best_len = -1.0
        for c in center_curves:
            if c is None:
                continue
            try:
                ln = float(c.Length)
            except Exception:
                ln = 0.0
            if ln > best_len:
                best = c
                best_len = ln
        if best is not None:
            return best

    try:
        loc = getattr(rebar, "Location", None)
        lc = getattr(loc, "Curve", None) if loc is not None else None
        if lc is not None:
            return lc
    except Exception:
        pass

    return None


def _z_view(view):
    if view is None:
        return 0.0
    try:
        o = view.Origin
        if o is not None:
            return float(o.Z)
    except Exception:
        pass
    return 0.0


def _proyectar_punto_al_plano_vista(view, pt):
    """
    Proyección ortogonal al plano de la **vista dada** (origen ``view.Origin``,
    normal ``view.ViewDirection``). Misma lógica que al colocar el detail por primera vez;
    el DMU debe actualizar en el plano de ``inst.OwnerViewId``, no forzar Z de planta fija.
    """
    if view is None or pt is None:
        return None
    try:
        o = view.Origin
        vd = getattr(view, "ViewDirection", None)
        if o is not None and vd is not None and vd.GetLength() > 1e-12:
            vd = vd.Normalize()
            w = pt - o
            d = float(w.DotProduct(vd))
            return pt - vd.Multiply(d)
    except Exception:
        pass
    try:
        z = _z_view(view)
        return XYZ(float(pt.X), float(pt.Y), float(z))
    except Exception:
        return None


def compute_lap_segment_endpoints(rebar_a, rebar_b, view, min_len_ft=None):
    """
    Calcula el solape entre dos barras en **3D** (ejes de centro) y proyecta los extremos al
    **plano de la vista** pasada (la del detail: ``OwnerViewId`` al actualizar vía DMU).

    Returns:
        (p0, p1) en el plano de trabajo de esa vista, o (None, None) si no aplica.
    """
    if min_len_ft is None:
        min_len_ft = _MIN_OVERLAP_FT
    if rebar_a is None or rebar_b is None or view is None:
        return None, None
    ca = _dominant_centerline_curve(rebar_a)
    cb = _dominant_centerline_curve(rebar_b)
    if ca is None or cb is None:
        return None, None
    try:
        a0, a1 = ca.GetEndPoint(0), ca.GetEndPoint(1)
        b0, b1 = cb.GetEndPoint(0), cb.GetEndPoint(1)
    except Exception:
        return None, None

    u = _unit_3d(a1 - a0)
    if u is None:
        u = _unit_3d(b1 - b0)
    if u is None:
        return None, None

    try:
        vB = _unit_3d(b1 - b0)
        if vB is not None and abs(float(u.DotProduct(vB))) < 0.85:
            return None, None
    except Exception:
        pass

    def t_on_u(p):
        try:
            return float((p - a0).DotProduct(u))
        except Exception:
            return 0.0

    tA0 = t_on_u(a0)
    tA1 = t_on_u(a1)
    tB0 = t_on_u(b0)
    tB1 = t_on_u(b1)

    lo_a, hi_a = (min(tA0, tA1), max(tA0, tA1))
    lo_b, hi_b = (min(tB0, tB1), max(tB0, tB1))
    lo = max(lo_a, lo_b)
    hi = min(hi_a, hi_b)
    if hi - lo < float(min_len_ft):
        return None, None

    try:
        p0_3d = a0 + u.Multiply(float(lo))
        p1_3d = a0 + u.Multiply(float(hi))
    except Exception:
        try:
            p0_3d = a0 + u * float(lo)
            p1_3d = a0 + u * float(hi)
        except Exception:
            return None, None

    try:
        if float((p1_3d - p0_3d).GetLength()) < float(min_len_ft):
            return None, None
    except Exception:
        return None, None

    p0 = _proyectar_punto_al_plano_vista(view, p0_3d)
    p1 = _proyectar_punto_al_plano_vista(view, p1_3d)
    if p0 is None or p1 is None:
        return None, None

    try:
        if float((p1 - p0).GetLength()) < float(min_len_ft):
            return None, None
    except Exception:
        return None, None

    return p0, p1


def compute_lap_segment_endpoints_vigas(rebar_a, rebar_b, view, min_len_ft=None):
    """
    Regla DMU **solo enfierrado vigas** (detail con ``lap_detail_link_vigas_schema``).

    Por defecto equivale a ``compute_lap_segment_endpoints``. Sustituya la implementación
    aquí para lógica propia sin afectar shaft / borde losa (que usan la otra función).
    """
    return compute_lap_segment_endpoints(rebar_a, rebar_b, view, min_len_ft=min_len_ft)
