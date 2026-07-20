# -*- coding: utf-8 -*-
"""
Divide barras de armadura (Rebar) en varios tramos con traslape.

Flujo:
1. Seleccionar la barra original.
2. UI: elegir puntos de división en Revit (sobre el segmento mayor).
3. Extraer centerline, dividir el segmento mayor, conservar patas L en extremos.
4. Crear tramos replicando layout y posiciones del origen.

Limitaciones:
- Barras free-form: no soportadas.
- Conjuntos: cada tramo nuevo replica el layout original (cantidad y posiciones).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import os
import System
import datetime

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    Arc,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    IndependentTag,
    Line,
    Reference,
    StorageType,
    TagMode,
    TagOrientation,
    Transaction,
    UnitTypeId,
    UnitUtils,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarStyle,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from bimtools_rebar_hook_lengths import traslape_mm_from_nominal_diameter_mm
except Exception:
    traslape_mm_from_nominal_diameter_mm = None

_DIALOG_TITLE = u"Arainco: Dividir barra con traslape"
_TRANSACTION_NAME = u"Arainco: Dividir barra con traslape"

# Modos de solape respecto al punto de corte C (mm en el vano principal).
SPLICE_SYMMETRIC = u"symmetric"  # solape [C−L/2, C+L/2]
SPLICE_FORWARD = u"forward"  # solape [C, C+L] — T1 se alarga
SPLICE_BACKWARD = u"backward"  # solape [C−L, C] — T2 se alarga
SPLICE_MODES = (SPLICE_SYMMETRIC, SPLICE_FORWARD, SPLICE_BACKWARD)

SPLICE_MODE_LABELS = {
    SPLICE_SYMMETRIC: u"Simétrico (±½ traslape)",
    SPLICE_FORWARD: u"Hacia adelante (C → C+L)",
    SPLICE_BACKWARD: u"Hacia atrás (C−L → C)",
}
_MIN_MARGEN_MM = 50.0
_DIAG_DIALOG_LINES = 14


def _exception_text(ex):
    if ex is None:
        return u""
    try:
        return unicode(ex)
    except NameError:
        return str(ex)


class _DiagSession(object):
    """Instrumentación ligera: consola pyRevit + archivo en scripts/_diag_logs/."""

    def __init__(self):
        self._lines = []
        self._path = None
        self._open_log()

    def _open_log(self):
        try:
            base = os.path.dirname(os.path.abspath(__file__))
            log_dir = os.path.join(base, "_diag_logs")
            if not os.path.isdir(log_dir):
                os.makedirs(log_dir)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self._path = os.path.join(log_dir, u"dividir_barra_{0}.log".format(stamp))
            self.log(u"=== Dividir barra traslape — sesión de diagnóstico ===")
        except Exception as ex:
            self._path = None
            self._lines.append(u"[diag] No se pudo crear log: {0}".format(_exception_text(ex)))

    def log(self, msg):
        try:
            line = unicode(msg)
        except NameError:
            line = str(msg)
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        entry = u"[{0}] {1}".format(ts, line)
        self._lines.append(entry)
        try:
            print(entry)
        except Exception:
            pass
        if not self._path:
            return
        try:
            import codecs

            with codecs.open(self._path, "a", "utf-8") as fh:
                fh.write(entry + u"\n")
        except Exception:
            pass

    def step(self, name):
        self.log(u"--- {0} ---".format(name))

    def ex(self, label, ex):
        self.log(u"{0}: {1}".format(label, _exception_text(ex)))

    def path(self):
        return self._path

    def tail(self, n=None):
        n = n or _DIAG_DIALOG_LINES
        chunk = self._lines[-n:] if len(self._lines) > n else self._lines
        return u"\n".join(chunk)

    def failure_message(self, headline):
        parts = [headline]
        if self._path:
            parts.append(u"\nLog completo:\n{0}".format(self._path))
        parts.append(u"\nÚltimas líneas:\n{0}".format(self.tail()))
        return u"\n".join(parts)


def _describe_curve_chain(diag, doc, label, curves):
    if diag is None:
        return
    clist = list(curves or [])
    diag.log(
        u"Cadena '{0}': {1} tramo(s), L_total={2:.1f} mm".format(
            label,
            len(clist),
            _internal_to_mm(_chain_total_length(clist)),
        )
    )
    prev_end = None
    for i, c in enumerate(clist):
        try:
            tcrv = _clr_type_of(c)
            tname = tcrv.Name if tcrv is not None else u"?"
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            ln = _curve_length_safe(c)
            gap_mm = 0.0
            if prev_end is not None:
                gap_mm = _internal_to_mm(float(prev_end.DistanceTo(p0)))
            diag.log(
                u"  [{0}] {1} L={2:.1f} mm gap_prev={3:.2f} mm".format(
                    i, tname, _internal_to_mm(ln), gap_mm
                )
            )
            prev_end = p1
        except Exception as ex:
            diag.ex(u"  [{0}] curva".format(i), ex)


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _internal_to_mm(ft):
    return UnitUtils.ConvertFromInternalUnits(float(ft), UnitTypeId.Millimeters)


def _curve_clr_type():
    return clr.GetClrType(Line).BaseType


def _line_clr_type():
    return clr.GetClrType(Line)


def _arc_clr_type():
    return clr.GetClrType(Arc)


def _clr_type_of(obj):
    """Tipo CLR de una instancia (IronPython: no usar GetClrType sobre instancias)."""
    if obj is None:
        return None
    try:
        return obj.GetType()
    except Exception:
        try:
            return type(obj)
        except Exception:
            return None


def _curve_is_line(crv):
    if crv is None:
        return False
    try:
        if isinstance(crv, Line):
            return True
    except Exception:
        pass
    tcrv = _clr_type_of(crv)
    if tcrv is None:
        return False
    try:
        tname = tcrv.Name
        fn = tcrv.FullName or u""
        if _line_clr_type().Equals(tcrv) or _line_clr_type().IsAssignableFrom(tcrv):
            return True
        return tname == "Line" or fn.endswith(".Line")
    except Exception:
        return False


def _curve_is_arc(crv):
    if crv is None:
        return False
    try:
        if isinstance(crv, Arc):
            return True
    except Exception:
        pass
    tcrv = _clr_type_of(crv)
    if tcrv is None:
        return False
    try:
        tname = tcrv.Name
        fn = tcrv.FullName or u""
        if _arc_clr_type().Equals(tcrv) or _arc_clr_type().IsAssignableFrom(tcrv):
            return True
        return tname == "Arc" or fn.endswith(".Arc")
    except Exception:
        return False


def _short_curve_tol(doc):
    try:
        return float(doc.Application.ShortCurveTolerance)
    except Exception:
        return 1.0 / 304.8


def _curve_length_safe(crv):
    if crv is None:
        return 0.0
    try:
        return float(crv.Length)
    except Exception:
        return 0.0


def _ordered_trim_params(crv, pa, pb):
    pa = float(pa)
    pb = float(pb)
    if abs(pa - pb) < 1e-12:
        raise ValueError(u"Parámetros de recorte coinciden.")
    try:
        p0 = float(crv.GetEndParameter(0))
        p1 = float(crv.GetEndParameter(1))
        lo = min(p0, p1)
        hi = max(p0, p1)
        pa = max(lo, min(hi, pa))
        pb = max(lo, min(hi, pb))
    except Exception:
        pass
    if pa > pb:
        pa, pb = pb, pa
    if abs(pa - pb) < 1e-12:
        raise ValueError(u"Parámetros de recorte coinciden.")
    return pa, pb


def _line_between_points(p0, p1, min_len):
    if p0 is None or p1 is None:
        return None
    if p0.DistanceTo(p1) < min_len:
        return None
    try:
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def _anchor_line_start_to_point(line_crv, new_start, min_len):
    """Mueve el inicio de una línea a ``new_start`` conservando el extremo final."""
    if not _curve_is_line(line_crv) or new_start is None:
        return line_crv
    try:
        p1 = line_crv.GetEndPoint(1)
        return _line_between_points(new_start, p1, min_len) or line_crv
    except Exception:
        return line_crv


def _anchor_line_end_to_point(line_crv, new_end, min_len):
    """Mueve el final de una línea a ``new_end`` conservando el extremo inicial."""
    if not _curve_is_line(line_crv) or new_end is None:
        return line_crv
    try:
        p0 = line_crv.GetEndPoint(0)
        return _line_between_points(p0, new_end, min_len) or line_crv
    except Exception:
        return line_crv


def _trim_curve_segment(crv, pa, pb, min_len=None):
    if crv is None:
        raise ValueError(u"Curva nula.")
    pa, pb = _ordered_trim_params(crv, pa, pb)
    min_len = float(min_len) if min_len is not None else 1e-9
    p_a = crv.Evaluate(pa, False)
    p_b = crv.Evaluate(pb, False)
    if p_a.DistanceTo(p_b) < min_len:
        raise ValueError(u"Tramo recortado demasiado corto.")
    if _curve_is_line(crv):
        return Line.CreateBound(p_a, p_b)
    if _curve_is_arc(crv):
        tmid = 0.5 * (pa + pb)
        p_m = crv.Evaluate(tmid, False)
        try:
            arc = Arc.Create(p_a, p_m, p_b)
            if _curve_length_safe(arc) >= min_len:
                return arc
        except Exception:
            pass
        ln = _line_between_points(p_a, p_b, min_len)
        if ln is not None:
            return ln
        raise ValueError(u"No se pudo recortar el arco.")
    tcrv = _clr_type_of(crv)
    tname = tcrv.Name if tcrv is not None else u"?"
    raise ValueError(
        u"Tipo de curva no soportado para recorte: {} (solo línea y arco).".format(tname)
    )


def _clone_curve(crv):
    """Copia geométrica independiente (las curvas de GetCenterlineCurves no son creables tal cual)."""
    if crv is None:
        return None
    try:
        cloned = crv.Clone()
        if cloned is not None:
            return cloned
    except Exception:
        pass
    if _curve_is_line(crv):
        return _line_between_points(
            crv.GetEndPoint(0), crv.GetEndPoint(1), 1e-12
        )
    if _curve_is_arc(crv):
        try:
            pa = crv.GetEndParameter(0)
            pb = crv.GetEndParameter(1)
            return _trim_curve_segment(crv, pa, pb, 1e-12)
        except Exception:
            pass
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        try:
            pm = crv.Evaluate(0.5, True)
            arc = Arc.Create(p0, pm, p1)
            if arc is not None:
                return arc
        except Exception:
            pass
        return _line_between_points(p0, p1, 1e-12)
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        return _line_between_points(p0, p1, 1e-12)
    except Exception:
        return None


def _point_at_dist_on_curve(crv, dist):
    """Punto a distancia de arco ``dist`` desde el inicio de ``crv``."""
    cl = _curve_length_safe(crv)
    d = float(dist)
    if d <= 1e-12:
        return crv.GetEndPoint(0)
    if d >= cl - 1e-9:
        return crv.GetEndPoint(1)
    if _curve_is_line(crv):
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        u = p1.Subtract(p0)
        ln = float(u.GetLength())
        if ln < 1e-12:
            return p0
        return p0.Add(u.Multiply(d / ln))
    par = _param_at_dist_from_start(crv, d)
    return crv.Evaluate(par, False)


def _extract_piece_from_curve(c, local_a, local_b, cl, min_len, snap_start=None):
    la = float(local_a)
    lb = float(local_b)
    ml = float(min_len) if min_len is not None else 1e-9
    if lb <= la + 1e-9:
        return None
    full_eps = max(1e-9, float(cl) * 1e-6)
    if la < 1e-9 and abs(lb - float(cl)) < full_eps:
        return _clone_curve(c)
    if la < 1e-9 and snap_start is not None:
        p_a = snap_start
    else:
        p_a = _point_at_dist_on_curve(c, la)
    p_b = _point_at_dist_on_curve(c, lb)
    if _curve_is_line(c):
        return _line_between_points(p_a, p_b, ml)
    if _curve_is_arc(c):
        if la < 1e-9 and abs(lb - float(cl)) < full_eps:
            return _clone_curve(c)
        try:
            p0 = c.GetEndParameter(0)
            pa = p0 if la < 1e-9 else _param_at_dist_from_start(c, la)
            pb = _param_at_dist_from_start(c, lb)
            return _trim_curve_segment(c, pa, pb, ml)
        except Exception:
            return _line_between_points(p_a, p_b, ml)
    return _line_between_points(p_a, p_b, ml)


def _heal_chain_endpoints(curves, max_gap_ft, min_len, max_arc_heal_ft=None):
    """Cierra micro-huecos sin redibujar arcos de pata L con huecos grandes."""
    if not curves or len(curves) < 2:
        return list(curves or [])
    arc_heal = float(max_arc_heal_ft) if max_arc_heal_ft is not None else max(
        float(min_len), _mm_to_internal(1.0)
    )
    out = [curves[0]]
    for c in curves[1:]:
        if c is None:
            continue
        prev = out[-1]
        try:
            p_join = prev.GetEndPoint(1)
            p1 = c.GetEndPoint(1)
            gap = float(p_join.DistanceTo(c.GetEndPoint(0)))
        except Exception:
            out.append(c)
            continue
        if gap <= float(max_gap_ft):
            if _curve_is_arc(c):
                nc = c
            elif _curve_is_line(c):
                nc = _line_between_points(p_join, p1, min_len)
            else:
                nc = _line_between_points(p_join, p1, min_len)
            if nc is not None and _curve_length_safe(nc) >= min_len:
                out.append(nc)
            else:
                out.append(c)
        else:
            out.append(c)
    return out


def _snap_chain_micro_gaps(curves, doc):
    """Solo une huecos menores a 1 mm; no sustituye arcos de codos L."""
    if not curves or len(curves) < 2:
        return list(curves or [])
    ml = max(_short_curve_tol(doc), 1e-9)
    micro = max(ml, _mm_to_internal(1.0))
    return _heal_chain_endpoints(curves, micro, ml, max_arc_heal_ft=micro)


def _prepare_curves_for_rebar(doc, curves, diag=None, label=u"", aggressive_heal=False):
    """Filtra tramos cortos; curado suave por defecto (preserva arcos de pata L)."""
    tol = _short_curve_tol(doc)
    ml = max(tol, 1e-9)
    heal_gap = max(_mm_to_internal(25.0), ml * 50.0) if aggressive_heal else max(
        _mm_to_internal(1.0), ml
    )
    arc_heal = heal_gap if aggressive_heal else max(_mm_to_internal(1.0), ml)
    clist = []
    for c in curves or []:
        if c is None:
            continue
        try:
            if _curve_length_safe(c) >= ml:
                clist.append(c)
        except Exception:
            continue
    if not clist:
        return []
    healed = _snap_chain_micro_gaps(clist, doc)
    if diag and len(healed) != len(clist):
        diag.log(
            u"Cadena '{0}' curada: {1} -> {2} tramo(s)".format(
                label, len(clist), len(healed)
            )
        )
    for i in range(1, len(healed)):
        try:
            gap = float(
                healed[i - 1].GetEndPoint(1).DistanceTo(healed[i].GetEndPoint(0))
            )
            if gap > heal_gap and diag:
                diag.log(
                    u"Gap residual {0:.2f} mm en {1}[{2}]".format(
                        _internal_to_mm(gap), label, i
                    )
                )
        except Exception:
            pass
    return healed


def _normal_from_chain(curves, fallback):
    if not curves:
        return fallback
    try:
        p0 = curves[0].GetEndPoint(0)
        p1 = curves[0].GetEndPoint(1)
        t = p1.Subtract(p0)
        if t.GetLength() < 1e-12 and len(curves) > 1:
            p1 = curves[1].GetEndPoint(1)
            t = p1.Subtract(p0)
        if t.GetLength() < 1e-12:
            return fallback
        t = t.Normalize()
        p_ref = curves[1].GetEndPoint(1) if len(curves) > 1 else curves[0].GetEndPoint(1)
        v = p_ref.Subtract(p0)
        n = t.CrossProduct(v)
        if n is not None and n.GetLength() > 1e-9:
            return n.Normalize()
    except Exception:
        pass
    return fallback


def _xyz_fmt(p):
    if p is None:
        return u"None"
    try:
        return u"({0:.1f},{1:.1f},{2:.1f})".format(
            _internal_to_mm(float(p.X)),
            _internal_to_mm(float(p.Y)),
            _internal_to_mm(float(p.Z)),
        )
    except Exception:
        return u"?"


def _vec_fmt(v):
    if v is None:
        return u"None"
    try:
        return u"({0:.4f},{1:.4f},{2:.4f})".format(float(v.X), float(v.Y), float(v.Z))
    except Exception:
        return u"?"


def _dot_safe(a, b):
    try:
        return float(a.DotProduct(b))
    except Exception:
        return None


def _normals_same_hemisphere(n0, n1):
    d = _dot_safe(n0, n1)
    if d is None:
        return True
    return d >= 0.0


def _chain_sample_points(curves, n_samples=5):
    """Puntos a lo largo de la cadena (inicio…fin) para comparar pose."""
    if not curves:
        return []
    total = _chain_total_length(curves)
    if total <= 1e-12:
        try:
            return [curves[0].GetEndPoint(0)]
        except Exception:
            return []
    out = []
    n_samples = max(2, int(n_samples))
    for i in range(n_samples):
        target = total * (float(i) / float(n_samples - 1))
        accum = 0.0
        pt = None
        for c in curves:
            cl = float(c.Length)
            if target <= accum + cl + 1e-9:
                local = max(0.0, target - accum)
                try:
                    pt = c.Evaluate(_param_at_dist_from_start(c, local), False)
                except Exception:
                    try:
                        pt = c.GetEndPoint(0 if local < 0.5 * cl else 1)
                    except Exception:
                        pt = None
                break
            accum += cl
        if pt is None:
            try:
                pt = curves[-1].GetEndPoint(1)
            except Exception:
                continue
        out.append(pt)
    return out


def _mean_delta_xyz(points_a, points_b):
    """Delta medio A←B (b + delta ≈ a)."""
    if not points_a or not points_b or len(points_a) != len(points_b):
        return None, None
    sx = sy = sz = 0.0
    n = 0
    for pa, pb in zip(points_a, points_b):
        try:
            sx += float(pa.X) - float(pb.X)
            sy += float(pa.Y) - float(pb.Y)
            sz += float(pa.Z) - float(pb.Z)
            n += 1
        except Exception:
            continue
    if n < 1:
        return None, None
    delta = XYZ(sx / n, sy / n, sz / n)
    return delta, float(delta.GetLength())


def _log_pose_compare(diag, label, orig_rebar, new_rebar, expected_curves):
    """Instrumentación: normal, lado de layout y desfase espacial vs origen."""
    if diag is None:
        return
    try:
        a0 = _shape_driven_accessor(orig_rebar)
        a1 = _shape_driven_accessor(new_rebar)
        n0 = a0.Normal if a0 is not None else None
        n1 = a1.Normal if a1 is not None else None
        side0 = bool(a0.BarsOnNormalSide) if a0 is not None else None
        side1 = bool(a1.BarsOnNormalSide) if a1 is not None else None
        diag.log(
            u"Pose[{0}] normal_orig={1} normal_new={2} dot={3}".format(
                label, _vec_fmt(n0), _vec_fmt(n1), _dot_safe(n0, n1)
            )
        )
        diag.log(
            u"Pose[{0}] BarsOnNormalSide orig={1} new={2}".format(label, side0, side1)
        )
        n_pos = min(_cantidad_posiciones(orig_rebar), _cantidad_posiciones(new_rebar))
        # Solo comparar transforms relativos si ya no es Single.
        rule_new = _layout_rule_nombre(new_rebar) or u""
        if rule_new != u"Single" and n_pos >= 1:
            for i in range(max(1, n_pos)):
                ti = _get_bar_transform(orig_rebar, i)
                tn = _get_bar_transform(new_rebar, i)
                o_i = ti.Origin if ti is not None else None
                o_n = tn.Origin if tn is not None else None
                d = None
                if o_i is not None and o_n is not None:
                    try:
                        d = _internal_to_mm(float(o_i.DistanceTo(o_n)))
                    except Exception:
                        d = None
                diag.log(
                    u"Pose[{0}] bar{1} T_rel_orig={2} T_rel_new={3} d_rel={4} mm".format(
                        label, i, _xyz_fmt(o_i), _xyz_fmt(o_n), d
                    )
                )
                # Midpoints visuales (mundo)
                try:
                    m_o = _rebar_midpoint_xyz(orig_rebar, i)
                    m_n = _rebar_midpoint_xyz(new_rebar, i)
                    d_m = None
                    if m_o is not None and m_n is not None:
                        d_m = _internal_to_mm(float(m_o.DistanceTo(m_n)))
                    diag.log(
                        u"Pose[{0}] bar{1} mid_orig={2} mid_new={3} d_mid={4} mm".format(
                            label, i, _xyz_fmt(m_o), _xyz_fmt(m_n), d_m
                        )
                    )
                except Exception:
                    pass
        exp = list(expected_curves or [])
        act = _centerline_chain(new_rebar, 0)
        if exp and act:
            pe = _chain_sample_points(exp, 5)
            pa = _chain_sample_points(act, 5)
            delta, dist = _mean_delta_xyz(pe, pa)
            rms = 0.0
            nrms = 0
            if delta is not None and pe and pa:
                for ept, apt in zip(pe, pa):
                    try:
                        moved = apt.Add(delta)
                        rms += float(ept.DistanceTo(moved)) ** 2
                        nrms += 1
                    except Exception:
                        pass
                if nrms:
                    rms = (rms / nrms) ** 0.5
            diag.log(
                u"Pose[{0}] cadena_visual: mean_delta={1} mm rms_post={2:.2f} mm "
                u"exp0={3} act0={4}".format(
                    label,
                    _internal_to_mm(dist) if dist is not None else None,
                    _internal_to_mm(rms) if nrms else -1.0,
                    _xyz_fmt(pe[0]) if pe else u"?",
                    _xyz_fmt(pa[0]) if pa else u"?",
                )
            )
    except Exception as ex:
        diag.ex(u"PoseCompare[{0}]".format(label), ex)


def _rebind_curve_start(crv, new_start, min_len):
    if crv is None or new_start is None:
        return None
    try:
        p_end = crv.GetEndPoint(1)
    except Exception:
        return None
    if new_start.DistanceTo(p_end) < min_len:
        return None
    if _curve_is_line(crv):
        return _line_between_points(new_start, p_end, min_len)
    if _curve_is_arc(crv):
        try:
            p_mid = crv.Evaluate(0.5, True)
            arc = Arc.Create(new_start, p_mid, p_end)
            if _curve_length_safe(arc) >= min_len:
                return arc
        except Exception:
            pass
        return _line_between_points(new_start, p_end, min_len)
    return _line_between_points(new_start, p_end, min_len)


def _sanear_cadena_curvas(doc, curves):
    """
    Filtra tramos cortos y garantiza continuidad extremo a extremo para CreateFromCurves.
    """
    tol = _short_curve_tol(doc)
    min_len = max(tol, 1e-9)
    bridge_max = max(min_len * 200.0, _mm_to_internal(5.0))
    raw = []
    for c in curves or []:
        if c is None:
            continue
        try:
            if _curve_length_safe(c) >= min_len:
                raw.append(c)
        except Exception:
            continue
    if not raw:
        return []
    if len(raw) == 1:
        return raw

    out = [raw[0]]
    for c in raw[1:]:
        prev = out[-1]
        try:
            p_end = prev.GetEndPoint(1)
            p_start = c.GetEndPoint(0)
            gap = float(p_end.DistanceTo(p_start))
        except Exception:
            out.append(c)
            continue
        if gap <= tol:
            rebound = _rebind_curve_start(c, p_end, min_len)
            out.append(rebound if rebound is not None else c)
            continue
        if gap > bridge_max:
            rebound = _rebind_curve_start(c, p_end, min_len)
            out.append(rebound if rebound is not None else c)
            continue
        bridge = _line_between_points(p_end, p_start, min_len)
        if bridge is not None:
            out.append(bridge)
        out.append(c)

    cleaned = []
    for c in out:
        if c is not None and _curve_length_safe(c) >= min_len:
            cleaned.append(c)
    return cleaned


def _cadena_a_polilinea(doc, curves):
    """Último recurso: cadena de líneas entre vértices consecutivos."""
    tol = _short_curve_tol(doc)
    if not curves:
        return []
    pts = []
    try:
        pts.append(curves[0].GetEndPoint(0))
    except Exception:
        return []
    for c in curves:
        try:
            pts.append(c.GetEndPoint(1))
        except Exception:
            continue
    out = []
    for i in range(len(pts) - 1):
        ln = _line_between_points(pts[i], pts[i + 1], tol)
        if ln is not None:
            out.append(ln)
    return out


def _segment_length_between(crv, pa, pb):
    return _curve_length_safe(_trim_curve_segment(crv, pa, pb))


def _param_at_dist_from_start(curve, dist):
    p0 = curve.GetEndParameter(0)
    p1 = curve.GetEndParameter(1)
    if dist <= 1e-12:
        return p0
    if dist >= curve.Length - 1e-9:
        return p1
    prev_p = p0
    for k in range(1, 33):
        t = float(k) / 32.0
        pm = p0 + t * (p1 - p0)
        try:
            Lm = _segment_length_between(curve, p0, pm)
        except Exception:
            continue
        if Lm >= dist:
            lo_p, hi_p = prev_p, pm
            for _ in range(45):
                mid = 0.5 * (lo_p + hi_p)
                try:
                    Lmid = _segment_length_between(curve, p0, mid)
                except Exception:
                    hi_p = mid
                    continue
                if Lmid < dist:
                    lo_p = mid
                else:
                    hi_p = mid
            return 0.5 * (lo_p + hi_p)
        prev_p = pm
    return p1


def _chain_total_length(curves):
    return sum(float(c.Length) for c in curves if c is not None)


def _find_main_span_index(chain):
    """Índice del segmento más largo (vanos / tramo recto principal)."""
    best_i = 0
    best_len = -1.0
    for i, c in enumerate(chain or []):
        if c is None:
            continue
        ln = float(c.Length)
        if ln > best_len:
            best_len = ln
            best_i = i
    return best_i


def _decompose_bar_chain(chain):
    """Separa patas L / codos del vano principal (segmento más largo)."""
    clist = [c for c in (chain or []) if c is not None]
    if not clist:
        return {
            u"main_index": 0,
            u"prefix": [],
            u"main": [],
            u"suffix": [],
            u"prefix_len": 0.0,
            u"main_len": 0.0,
            u"suffix_len": 0.0,
        }
    idx = _find_main_span_index(clist)
    prefix = list(clist[:idx])
    main = [clist[idx]]
    suffix = list(clist[idx + 1 :])
    prefix_len = sum(float(c.Length) for c in prefix)
    main_len = float(main[0].Length) if main else 0.0
    suffix_len = sum(float(c.Length) for c in suffix)
    return {
        u"main_index": idx,
        u"prefix": prefix,
        u"main": main,
        u"suffix": suffix,
        u"prefix_len": prefix_len,
        u"main_len": main_len,
        u"suffix_len": suffix_len,
    }


def _clone_chain(curves):
    out = []
    for c in curves or []:
        cc = _clone_curve(c)
        if cc is not None:
            out.append(cc)
    return out


def _compose_chunk_curves(layout, d0_main, d1_main, document, diag=None, label=u""):
    """
    Construye la centerline de un tramo: patas L fijas solo en extremos del vano.
    ``d0_main`` / ``d1_main`` son distancias de arco sobre el segmento principal.
    """
    min_len = _short_curve_tol(document)
    main = layout.get(u"main") or []
    main_len = float(layout.get(u"main_len") or 0.0)
    if not main or main_len <= 1e-9:
        return []
    d0 = max(0.0, min(main_len, float(d0_main)))
    d1 = max(0.0, min(main_len, float(d1_main)))
    if d1 <= d0 + 1e-9:
        return []
    main_part = _extract_subchain_between_distances(main, d0, d1, min_len=min_len)
    if not main_part:
        return []
    prefix = _clone_chain(layout.get(u"prefix") or []) if d0 <= 1e-9 else []
    suffix = (
        _clone_chain(layout.get(u"suffix") or []) if d1 >= main_len - 1e-9 else []
    )
    if prefix and main_part:
        try:
            anchor = prefix[-1].GetEndPoint(1)
            if _curve_is_line(main_part[0]):
                main_part[0] = _anchor_line_start_to_point(
                    main_part[0], anchor, min_len
                )
        except Exception:
            pass
    if suffix and main_part:
        try:
            anchor = suffix[0].GetEndPoint(0)
            if _curve_is_line(main_part[-1]):
                main_part[-1] = _anchor_line_end_to_point(
                    main_part[-1], anchor, min_len
                )
        except Exception:
            pass
    parts = []
    parts.extend(prefix)
    parts.extend(main_part)
    parts.extend(suffix)
    return _stitch_chunk_parts(parts, document)


def _stitch_chunk_parts(parts, doc):
    """Une prefijo + vano + sufijo sin deformar arcos de pata L."""
    if not parts:
        return []
    ml = max(_short_curve_tol(doc), 1e-9)
    micro = max(ml, _mm_to_internal(2.0))
    out = [parts[0]]
    for c in parts[1:]:
        if c is None:
            continue
        prev = out[-1]
        try:
            p_join = prev.GetEndPoint(1)
            p_end = c.GetEndPoint(1)
            gap = float(p_join.DistanceTo(c.GetEndPoint(0)))
        except Exception:
            out.append(c)
            continue
        if gap < micro and _curve_is_line(c):
            nc = _line_between_points(p_join, p_end, ml)
            out.append(nc if nc is not None else c)
        else:
            out.append(c)
    return out


def _segment_kind_label(crv):
    return u"Arc" if _curve_is_arc(crv) else u"Line"


def build_bar_preview_model(chain, lap_mm, diameter_mm):
    """Modelo ligero (solo mm) para la UI de división."""
    layout = _decompose_bar_chain(chain)
    segments = []
    main_idx = int(layout[u"main_index"])
    for i, c in enumerate(chain or []):
        if c is None:
            continue
        if i < main_idx:
            role = u"prefix"
        elif i > main_idx:
            role = u"suffix"
        else:
            role = u"main"
        segments.append(
            {
                u"index": i,
                u"length_mm": _internal_to_mm(float(c.Length)),
                u"kind": _segment_kind_label(c),
                u"role": role,
            }
        )
    return {
        u"segments": segments,
        u"main_index": main_idx,
        u"main_length_mm": _internal_to_mm(float(layout[u"main_len"])),
        u"prefix_length_mm": _internal_to_mm(float(layout[u"prefix_len"])),
        u"suffix_length_mm": _internal_to_mm(float(layout[u"suffix_len"])),
        u"lap_mm": float(lap_mm),
        u"diameter_mm": diameter_mm,
        u"total_length_mm": _internal_to_mm(_chain_total_length(chain)),
    }


def normalize_splice_mode(mode):
    """Normaliza el modo de solape; por defecto simétrico."""
    try:
        key = unicode(mode or u"").strip().lower()
    except NameError:
        key = str(mode or u"").strip().lower()
    if key in SPLICE_MODES:
        return key
    return SPLICE_SYMMETRIC


def splice_overlap_zone_mm(cut_mm, lap_mm, splice_mode=None):
    """Zona de solape [a, b] en mm sobre el vano para un corte C."""
    mode = normalize_splice_mode(splice_mode)
    c = float(cut_mm)
    lap = float(lap_mm)
    if mode == SPLICE_FORWARD:
        return c, c + lap
    if mode == SPLICE_BACKWARD:
        return c - lap, c
    half = 0.5 * lap
    return c - half, c + half


def _piece_ranges_on_main(cuts_ft, main_len, lap_ft, splice_mode):
    """
    Rangos (d0, d1) de cada tramo sobre el vano principal (unidades internas).

    - symmetric: extremos ± half_lap alrededor de cada C
    - forward: tramos intermedios terminan en C+L; el siguiente empieza en C
    - backward: tramos intermedios empiezan en C−L; el anterior termina en C
    """
    mode = normalize_splice_mode(splice_mode)
    cuts = list(cuts_ft or [])
    n_pieces = len(cuts) + 1
    lap = float(lap_ft)
    main = float(main_len)
    out = []
    for i in range(n_pieces):
        if mode == SPLICE_FORWARD:
            if i == 0:
                d0, d1 = 0.0, cuts[0] + lap
            elif i == n_pieces - 1:
                d0, d1 = cuts[-1], main
            else:
                d0, d1 = cuts[i - 1], cuts[i] + lap
        elif mode == SPLICE_BACKWARD:
            if i == 0:
                d0, d1 = 0.0, cuts[0]
            elif i == n_pieces - 1:
                d0, d1 = cuts[-1] - lap, main
            else:
                d0, d1 = cuts[i - 1] - lap, cuts[i]
        else:
            half = 0.5 * lap
            if i == 0:
                d0, d1 = 0.0, cuts[0] + half
            elif i == n_pieces - 1:
                d0, d1 = cuts[-1] - half, main
            else:
                d0, d1 = cuts[i - 1] - half, cuts[i] + half
        d0 = max(0.0, min(main, float(d0)))
        d1 = max(0.0, min(main, float(d1)))
        out.append((d0, d1))
    return out


def _validate_cuts_on_main(cuts_mm, main_len_mm, lap_mm, splice_mode=None):
    """Valida lista de cortes (mm desde inicio del vano principal)."""
    if not cuts_mm:
        return False, u"Indique al menos un punto de división sobre el vano principal."
    try:
        cuts = sorted(set(float(c) for c in cuts_mm))
    except Exception:
        return False, u"Cortes no válidos."
    mode = normalize_splice_mode(splice_mode)
    lap = float(lap_mm)
    margen = float(_MIN_MARGEN_MM)
    main_len = float(main_len_mm)

    if mode == SPLICE_FORWARD:
        if main_len <= lap + 2.0 * margen:
            return (
                False,
                u"El vano principal ({:.0f} mm) es demasiado corto para el traslape.".format(
                    main_len
                ),
            )
        for c in cuts:
            if c < margen:
                return (
                    False,
                    u"Corte a {:.0f} mm demasiado cerca del inicio del vano.".format(c),
                )
            if c > main_len - lap - margen:
                return (
                    False,
                    u"Corte a {:.0f} mm demasiado cerca del final del vano "
                    u"(hace falta espacio para el traslape hacia adelante).".format(c),
                )
    elif mode == SPLICE_BACKWARD:
        if main_len <= lap + 2.0 * margen:
            return (
                False,
                u"El vano principal ({:.0f} mm) es demasiado corto para el traslape.".format(
                    main_len
                ),
            )
        for c in cuts:
            if c < lap + margen:
                return (
                    False,
                    u"Corte a {:.0f} mm demasiado cerca del inicio del vano "
                    u"(hace falta espacio para el traslape hacia atrás).".format(c),
                )
            if c > main_len - margen:
                return (
                    False,
                    u"Corte a {:.0f} mm demasiado cerca del final del vano.".format(c),
                )
    else:
        half = 0.5 * lap
        if main_len <= 2.0 * half + 2.0 * margen:
            return (
                False,
                u"El vano principal ({:.0f} mm) es demasiado corto para el traslape.".format(
                    main_len
                ),
            )
        for c in cuts:
            if c < half + margen:
                return (
                    False,
                    u"Corte a {:.0f} mm demasiado cerca del inicio del vano (traslape).".format(
                        c
                    ),
                )
            if c > main_len - half - margen:
                return (
                    False,
                    u"Corte a {:.0f} mm demasiado cerca del final del vano (traslape).".format(
                        c
                    ),
                )

    for i in range(len(cuts) - 1):
        if cuts[i + 1] - cuts[i] < lap + 2.0 * margen:
            return (
                False,
                u"Los cortes {:.0f} mm y {:.0f} mm están demasiado próximos.".format(
                    cuts[i], cuts[i + 1]
                ),
            )
    return True, None


def _main_span_cut_mm_from_full_chain_dist(layout, cut_dist_full):
    """Convierte distancia sobre cadena completa a mm sobre vano principal."""
    try:
        local = float(cut_dist_full) - float(layout[u"prefix_len"])
    except Exception:
        return None
    main_len = float(layout[u"main_len"])
    if local < -1e-6 or local > main_len + 1e-6:
        return None
    return _internal_to_mm(max(0.0, min(main_len, local)))


def _main_span_cut_mm_from_point(layout, point):
    main = layout.get(u"main") or []
    if not main or point is None:
        return None, None
    cut, perp = _closest_arc_distance_on_chain(main, point)
    if cut is None:
        return None, perp
    return _internal_to_mm(cut), perp


def project_point_to_main_span_mm(document, rebar, point, bar_index=0):
    """
    Proyecta un punto 3D sobre el segmento mayor (vano) de la centerline.

    Returns:
        (cut_mm, mensaje_error) — distancia mm desde el inicio del vano principal.
    """
    if point is None:
        return None, u"Punto no válido."
    try:
        bi = int(bar_index)
    except Exception:
        bi = 0
    chain = _centerline_chain(rebar, bi)
    if not chain:
        return None, u"No se pudo leer la centerline."
    layout = _decompose_bar_chain(chain)
    cut_mm, _perp = _main_span_cut_mm_from_point(layout, point)
    if cut_mm is None:
        return (
            None,
            u"El punto debe caer sobre el segmento mayor (vano), no sobre las patas L.",
        )
    return cut_mm, None


def revit_pick_main_span_cut_mm(uiapp, rebar):
    """
    Selección en vista activa → distancia mm sobre el vano principal.

    Returns:
        (cut_mm, mensaje_error)
    """
    if uiapp is None or rebar is None:
        return None, u"Sin documento o barra."
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None, u"No hay documento activo."
    point, bar_index = _pick_point_on_rebar(uidoc, rebar)
    if point is None:
        return None, None
    return project_point_to_main_span_mm(
        uidoc.Document, rebar, point, bar_index if bar_index is not None else 0
    )


def prepare_dividir_session(document, rebar):
    """
    Prepara la UI a partir de la barra original (centerline índice 0, vano mayor).

    Returns:
        (ok, mensaje_error, session_dict)
    """
    if not isinstance(rebar, Rebar):
        return False, u"No es un elemento Rebar.", None
    if _shape_driven_accessor(rebar) is None:
        return (
            False,
            u"Solo aplica a barras shape-driven (no free-form).",
            None,
        )

    chain = _centerline_chain(rebar, 0)
    if not chain:
        return False, u"No se pudo leer la línea media de la barra.", None

    layout = _decompose_bar_chain(chain)
    if layout[u"main_len"] <= 1e-9:
        return False, u"La barra no tiene segmento mayor utilizable.", None

    bar_type = document.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"No se pudo obtener RebarBarType.", None
    lap_mm, d_mm = _lap_mm_for_rebar(rebar, bar_type)
    if lap_mm is None:
        return False, u"No se pudo calcular el traslape según el diámetro.", None

    preview = build_bar_preview_model(chain, lap_mm, d_mm)
    n_pos = _cantidad_posiciones(rebar)

    session = {
        u"rebar_id": rebar.Id,
        u"n_pos": n_pos,
        u"layout_rule": _layout_rule_nombre(rebar),
        u"lap_mm": float(lap_mm),
        u"diameter_mm": d_mm,
        u"preview": preview,
        u"suggested_cuts_mm": [],
    }
    return True, None, session


def prepare_dividir_session_legacy(document, rebar, point=None, bar_index_hint=None):
    """Compatibilidad — usar ``prepare_dividir_session``."""
    ok, err, session = prepare_dividir_session(document, rebar)
    if not ok:
        return ok, err, session
    if point is not None:
        cut_mm, _ = project_point_to_main_span_mm(
            document, rebar, point, bar_index_hint or 0
        )
        if cut_mm is not None:
            main_len = float(session[u"preview"].get(u"main_length_mm") or 0)
            ok_cut, _ = _validate_cuts_on_main([cut_mm], main_len, session[u"lap_mm"])
            if ok_cut:
                session[u"suggested_cuts_mm"] = [cut_mm]
    return ok, err, session


def _extract_subchain_between_distances(curves, d0, d1, min_len=None):
    clist = [c for c in curves if c is not None]
    if not clist:
        return []
    total = _chain_total_length(clist)
    d0 = max(0.0, min(total, float(d0)))
    d1 = max(0.0, min(total, float(d1)))
    if d1 <= d0 + 1e-9:
        return []
    out = []
    accum = 0.0
    prev_end = None
    for c in clist:
        cl = float(c.Length)
        seg_a = accum
        seg_b = accum + cl
        ia = max(d0, seg_a)
        ib = min(d1, seg_b)
        if ib > ia + 1e-9:
            piece = _extract_piece_from_curve(
                c,
                ia - seg_a,
                ib - seg_a,
                cl,
                min_len,
                snap_start=prev_end,
            )
            if piece is not None and _curve_length_safe(piece) > 1e-9:
                out.append(piece)
                try:
                    prev_end = piece.GetEndPoint(1)
                except Exception:
                    prev_end = None
        accum = seg_b
    return out


def _closest_on_line_segment(line, point):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        v = p1.Subtract(p0)
        length = float(v.GetLength())
        if length < 1e-12:
            return 0.0, float(point.DistanceTo(p0))
        u = v.Normalize()
        t = float((point.Subtract(p0)).DotProduct(u))
        if t < 0.0:
            t = 0.0
        elif t > length:
            t = length
        closest = p0.Add(u.Multiply(t))
        return t, float(point.DistanceTo(closest))
    except Exception:
        return None, None


def _project_point_on_curve(curve, point):
    """
  Distancia de arco desde el inicio de ``curve`` hasta la proyección de ``point``.
  Returns (arc_dist, perp_dist) o (None, None).
    """
    if curve is None or point is None:
        return None, None
    try:
        p0 = curve.GetEndParameter(0)
        p1 = curve.GetEndParameter(1)
    except Exception:
        return None, None

    if _curve_is_line(curve):
        arc, perp = _closest_on_line_segment(curve, point)
        if arc is not None:
            return arc, perp

    try:
        ir = curve.Project(point)
        if ir is not None:
            pp = getattr(ir, "XYZPoint", None)
            if pp is None:
                pp = getattr(ir, "Point", None)
            par = ir.Parameter
            if pp is not None:
                par_lo = min(float(p0), float(p1))
                par_hi = max(float(p0), float(p1))
                par_use = max(par_lo, min(par_hi, float(par)))
                seg_len = _segment_length_between(curve, p0, par_use)
                q = curve.Evaluate(par_use, False)
                return seg_len, float(point.DistanceTo(q))
    except Exception:
        pass

    best_arc = None
    best_perp = None
    n_samples = 48
    for k in range(n_samples + 1):
        t = float(k) / float(n_samples)
        try:
            par = float(p0) + t * (float(p1) - float(p0))
            q = curve.Evaluate(par, False)
            perp = float(point.DistanceTo(q))
            arc = _segment_length_between(curve, p0, par)
            if best_perp is None or perp < best_perp:
                best_perp = perp
                best_arc = arc
        except Exception:
            continue
    return best_arc, best_perp


def _closest_arc_distance_on_chain(curves, point):
    clist = [c for c in curves if c is not None]
    if not clist or point is None:
        return None, None
    best_perp = None
    best_arc = None
    accum = 0.0
    for c in clist:
        local_arc, perp = _project_point_on_curve(c, point)
        if local_arc is not None:
            arc_d = accum + local_arc
            if best_perp is None or perp < best_perp:
                best_perp = perp
                best_arc = arc_d
        accum += float(c.Length)
    return best_arc, best_perp


def _centerline_chain_variants(rebar, bar_index=0):
    """Devuelve cadenas de línea media (distintas opciones API) para mayor robustez."""
    chains = []
    opts = (
        (False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves),
        (False, True, False, MultiplanarOption.IncludeAllMultiplanarCurves),
        (False, False, True, MultiplanarOption.IncludeAllMultiplanarCurves),
        (False, False, False, MultiplanarOption.IncludeOnlyPlanarCurves),
    )
    seen = set()
    for adj, hooks, bend, mp_opt in opts:
        try:
            curves = rebar.GetCenterlineCurves(
                adj, hooks, bend, mp_opt, int(bar_index)
            )
        except Exception:
            continue
        if curves is None or curves.Count == 0:
            continue
        chain = tuple(curves[i] for i in range(curves.Count))
        key = tuple(round(float(c.Length), 6) for c in chain)
        if key in seen:
            continue
        seen.add(key)
        chains.append(list(chain))
    if not chains:
        one = _centerline_chain(rebar, bar_index)
        if one:
            chains.append(one)
    return chains


def _resolver_barra_y_corte(rebar, point, bar_index_hint=None):
    """
    Determina índice de barra y distancia de corte a lo largo de la línea media.

    Returns:
        (bar_index, cut_dist, perp_dist, chain) o mensaje de error.
    """
    n = _cantidad_posiciones(rebar)
    if point is None:
        return None, None, None, None, u"Punto de división no válido."

    candidates = []
    if bar_index_hint is not None:
        try:
            hi = int(bar_index_hint)
            if 0 <= hi < n:
                candidates.append(hi)
        except Exception:
            pass
    for i in range(n):
        if i not in candidates:
            candidates.append(i)

    best = None
    for idx in candidates:
        try:
            if not rebar.IsBarIncluded(int(idx)):
                continue
        except Exception:
            pass
        for chain in _centerline_chain_variants(rebar, idx):
            if not chain:
                continue
            cut, perp = _closest_arc_distance_on_chain(chain, point)
            if cut is None:
                continue
            total = _chain_total_length(chain)
            if total <= 1e-9:
                continue
            item = (perp, idx, cut, chain)
            if best is None or item[0] < best[0]:
                best = item

    if best is None:
        return None, None, None, None, u"No se pudo proyectar el punto sobre la línea media."

    perp, idx, cut, chain = best
    max_perp_ft = _mm_to_internal(5000.0)
    if perp > max_perp_ft:
        return (
            None,
            None,
            None,
            None,
            u"El punto está demasiado lejos de la línea media ({:.0f} mm). "
            u"Clic directamente sobre la barra en la vista.".format(_internal_to_mm(perp)),
        )
    return idx, cut, perp, chain, None


def _cantidad_posiciones(rebar):
    best = 1
    for getter in (
        lambda: int(rebar.NumberOfBarPositions),
        lambda: int(rebar.GetNumberOfBarPositions()),
        lambda: int(rebar.Quantity),
    ):
        try:
            n = int(getter())
            if n > best:
                best = n
        except Exception:
            pass
    return best


def _layout_rule_nombre(rebar):
    try:
        r = rebar.LayoutRule
        if r is not None:
            s = r.ToString() or u""
            if s:
                return s
    except Exception:
        pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            r = acc.GetLayoutRule()
            if r is not None:
                s = r.ToString() or u""
                if s:
                    return s
    except Exception:
        pass
    return u""


def _get_bar_transform(rebar, bar_index):
    bi = int(bar_index)
    try:
        return rebar.GetBarPositionTransform(bi)
    except Exception:
        pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, "GetBarPositionTransform"):
            return acc.GetBarPositionTransform(bi)
    except Exception:
        pass
    return None


def _bar_included(rebar, bar_index):
    try:
        return bool(rebar.IsBarIncluded(int(bar_index)))
    except Exception:
        return True


def _layout_rule_name_accessor(rebar, acc):
    try:
        r = rebar.LayoutRule
        if r is not None:
            return r.ToString() or u""
    except Exception:
        pass
    if acc is not None:
        try:
            r = acc.GetLayoutRule()
            if r is not None:
                return r.ToString() or u""
        except Exception:
            pass
    return u""


def _spacing_internal(rebar):
    try:
        return float(rebar.MaxSpacing)
    except Exception:
        return 0.0


def _array_length_internal(acc):
    if acc is None:
        return 0.0
    try:
        return float(acc.ArrayLength)
    except Exception:
        try:
            return float(acc.GetArrayLength())
        except Exception:
            return 0.0


def _copy_layout_rebar_shape_driven(src, dst, diag=None):
    """Replica regla de layout, cantidad y posiciones del conjunto origen."""
    a0 = src.GetShapeDrivenAccessor()
    a1 = dst.GetShapeDrivenAccessor()
    if a0 is None or a1 is None:
        return False, u"ShapeDrivenAccessor nulo (layout no copiable)."

    rule_name = _layout_rule_name_accessor(src, a0)
    sp = _spacing_internal(src)
    alen = _array_length_internal(a0)
    b_side = bool(a0.BarsOnNormalSide)
    # Si CreateFromCurves dejó la normal invertida, invertir el lado del array.
    try:
        n0 = a0.Normal
        n1 = a1.Normal
        if n0 is not None and n1 is not None and not _normals_same_hemisphere(n0, n1):
            b_side = not b_side
            if diag:
                diag.log(
                    u"Layout: normal invertida vs origen → BarsOnNormalSide={0}".format(
                        b_side
                    )
                )
    except Exception:
        pass
    inc0 = bool(src.IncludeFirstBar)
    inc1 = bool(src.IncludeLastBar)
    nbars = _cantidad_posiciones(src)

    try:
        if rule_name == u"Single":
            a1.SetLayoutAsSingle()
        elif rule_name == u"MaximumSpacing":
            a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        elif rule_name in (u"Number", u"FixedNumber"):
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
        elif rule_name == u"NumberWithSpacing":
            a1.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_side, inc0, inc1)
        elif rule_name == u"MinimumClearSpacing":
            a1.SetLayoutAsMinimumClearSpacing(sp, alen, b_side, inc0, inc1)
        else:
            if rule_name:
                try:
                    a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
                except Exception:
                    a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
            else:
                a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        return True, u""
    except Exception as ex:
        if int(nbars) == 1:
            try:
                a1.SetLayoutAsSingle()
                return True, u""
            except Exception as ex2:
                return False, u"{0!s} | fallback Single: {1!s}".format(ex, ex2)
        return False, u"{0!s} (regla: «{1}»)".format(ex, rule_name or u"(vacía)")


def _sync_bars_included_from_original(src, dst):
    n = _cantidad_posiciones(src)
    for i in range(n):
        try:
            dst.SetBarIncluded(bool(src.IsBarIncluded(int(i))), int(i))
        except Exception:
            pass


def _chains_congruent(chain_a, chain_b, tol_ft=None):
    if not chain_a or not chain_b:
        return False
    if len(chain_a) != len(chain_b):
        return False
    tol = tol_ft if tol_ft is not None else _mm_to_internal(2.0)
    for ca, cb in zip(chain_a, chain_b):
        try:
            if abs(float(ca.Length) - float(cb.Length)) > tol:
                return False
        except Exception:
            return False
    return True


def _resolve_geometry_for_division(rebar, bar_index_cut, diag=None):
    """
    Determina layout de geometría (patas L + vano) y barra plantilla.

    - Cortes: se validan sobre la barra seleccionada (``bar_index_cut``).
    - Creación con conjunto (n>1): plantilla en índice 0 para que el layout
      Revit coincida con las posiciones del origen.
    """
    try:
        bi_cut = int(bar_index_cut)
    except Exception:
        bi_cut = 0

    n_pos = _cantidad_posiciones(rebar)
    chain_cut = _centerline_chain(rebar, bi_cut)
    if not chain_cut:
        return None, bi_cut, 0, u"Sin centerline en barra índice {0}.".format(bi_cut)

    layout_cut = _decompose_bar_chain(chain_cut)

    if n_pos > 1:
        chain0 = _centerline_chain(rebar, 0)
        if not chain0:
            return None, bi_cut, 0, u"Sin centerline en barra índice 0 del conjunto."
        if not _chains_congruent(chain0, chain_cut):
            return (
                None,
                bi_cut,
                0,
                u"Las posiciones del conjunto no son congruentes; "
                u"no se puede replicar layout y geometría.",
            )
        layout_tpl = _decompose_bar_chain(chain0)
        tpl_index = 0
        if diag:
            diag.log(
                u"Geometría plantilla: barra 0 (conjunto n={0}, corte ref. idx={1})".format(
                    n_pos, bi_cut
                )
            )
    else:
        layout_tpl = layout_cut
        tpl_index = bi_cut

    return layout_tpl, bi_cut, tpl_index, None


def _alinear_rebar_a_transform_barra(document, new_rebar, orig_rebar, bar_index, diag=None):
    """Traslada el conjunto nuevo para que la barra ``bar_index`` coincida con el origen."""
    try:
        bi = int(bar_index)
    except Exception:
        bi = 0
    ti = _get_bar_transform(orig_rebar, bi)
    tn = _get_bar_transform(new_rebar, bi)
    if ti is None or tn is None:
        if diag:
            diag.log(
                u"Alineado transform: sin GetBarPositionTransform "
                u"(orig={0} new={1} idx={2})".format(
                    ti is not None, tn is not None, bi
                )
            )
        return
    try:
        delta = ti.Origin.Subtract(tn.Origin)
        dist = float(delta.GetLength())
        if diag:
            diag.log(
                u"Alineado transform idx={0}: d={1:.2f} mm before={2} target={3}".format(
                    bi,
                    _internal_to_mm(dist),
                    _xyz_fmt(tn.Origin),
                    _xyz_fmt(ti.Origin),
                )
            )
        if dist < _mm_to_internal(0.5):
            return
        ElementTransformUtils.MoveElement(document, new_rebar.Id, delta)
        if diag:
            diag.log(
                u"Alineado conjunto {0:.2f} mm (barra idx={1})".format(
                    _internal_to_mm(dist), bi
                )
            )
    except Exception as ex:
        if diag:
            diag.ex(u"Alineado transform", ex)


def _try_fix_array_side_if_misaligned(document, orig, new_rb, curves_tpl, diag=None):
    """
    Si barra 0 cuadra pero barra 1 queda lejos, invertir BarsOnNormalSide y realinear.
    """
    n = _cantidad_posiciones(orig)
    if n < 2:
        return False
    ti0 = _get_bar_transform(orig, 0)
    tn0 = _get_bar_transform(new_rb, 0)
    ti1 = _get_bar_transform(orig, 1)
    tn1 = _get_bar_transform(new_rb, 1)
    if ti0 is None or tn0 is None or ti1 is None or tn1 is None:
        return False
    try:
        d0 = float(ti0.Origin.DistanceTo(tn0.Origin))
        d1 = float(ti1.Origin.DistanceTo(tn1.Origin))
    except Exception:
        return False
    # Solo actuar si bar0 está cerca y bar1 desviada claramente.
    if d0 > _mm_to_internal(15.0) or d1 < _mm_to_internal(25.0):
        return False
    if diag:
        diag.log(
            u"Pose: bar0 ok ({0:.1f} mm) pero bar1 desviada ({1:.1f} mm) "
            u"→ reintento con lado de array invertido".format(
                _internal_to_mm(d0), _internal_to_mm(d1)
            )
        )
    a0 = _shape_driven_accessor(orig)
    a1 = _shape_driven_accessor(new_rb)
    if a0 is None or a1 is None:
        return False
    try:
        rule_name = _layout_rule_name_accessor(orig, a0)
        sp = _spacing_internal(orig)
        alen = _array_length_internal(a0)
        b_side = not bool(a1.BarsOnNormalSide)
        inc0 = bool(orig.IncludeFirstBar)
        inc1 = bool(orig.IncludeLastBar)
        nbars = n
        if rule_name == u"MaximumSpacing":
            a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        elif rule_name in (u"Number", u"FixedNumber"):
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
        elif rule_name == u"NumberWithSpacing":
            a1.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_side, inc0, inc1)
        elif rule_name == u"MinimumClearSpacing":
            a1.SetLayoutAsMinimumClearSpacing(sp, alen, b_side, inc0, inc1)
        else:
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
        document.Regenerate()
        _alinear_rebar_a_transform_barra(document, new_rb, orig, 0, diag=diag)
        _alinear_rebar_a_cadena_esperada(document, new_rb, curves_tpl, diag=diag)
        return True
    except Exception as ex:
        if diag:
            diag.ex(u"Flip array side", ex)
        return False


def _finalize_new_rebar_set(document, orig, new_rb, curves_tpl, diag=None):
    """
    Replica solo layout/posición del origen; la geometría viene de ``curves_tpl``.
    """
    n = _cantidad_posiciones(orig)
    ok, err = True, u""

    if diag:
        _log_pose_compare(diag, u"pre-layout", orig, new_rb, curves_tpl)

    if n > 1:
        ok, err = _copy_layout_rebar_shape_driven(orig, new_rb, diag=diag)
        if diag:
            if ok:
                diag.log(
                    u"Layout copiado: {0} pos., regla {1}".format(
                        n, _layout_rule_nombre(orig) or u"?"
                    )
                )
            else:
                diag.log(u"Layout NO copiado: {0}".format(err))
        _sync_bars_included_from_original(orig, new_rb)
        try:
            document.Regenerate()
        except Exception:
            pass
        # Primero alinear el conjunto por transform de barra 0…
        _alinear_rebar_a_transform_barra(document, new_rb, orig, 0, diag=diag)

    # …y al final imponer la geometría del tramo (gana sobre el transform).
    _alinear_rebar_a_cadena_esperada(document, new_rb, curves_tpl, diag=diag)

    if n > 1:
        try:
            document.Regenerate()
        except Exception:
            pass
        _try_fix_array_side_if_misaligned(
            document, orig, new_rb, curves_tpl, diag=diag
        )
        _copy_moved_bar_transforms(orig, new_rb, diag=diag)

    try:
        document.Regenerate()
    except Exception:
        pass
    # Tras regen, constraints/shape pueden desplazar: reimponer pose visual.
    _alinear_rebar_a_cadena_esperada(document, new_rb, curves_tpl, diag=diag)

    if diag:
        _log_pose_compare(diag, u"post-align", orig, new_rb, curves_tpl)
    return ok, err


def _asegurar_layout_single(rebar):
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            acc.SetLayoutAsSingle()
    except Exception:
        pass


def _alinear_rebar_a_cadena_esperada(document, new_rebar, expected_curves, diag=None):
    """
    Corrige el desfase de CreateFromCurves moviendo la barra para que su
    centerline coincida (traslación rígida) con la cadena recortada.
    """
    if new_rebar is None or not expected_curves:
        return
    actual = _centerline_chain(new_rebar, 0)
    if not actual:
        return
    try:
        pe = _chain_sample_points(expected_curves, 5)
        pa = _chain_sample_points(actual, 5)
        delta, dist = _mean_delta_xyz(pe, pa)
        if delta is None or dist is None:
            # Fallback: solo extremo inicial
            exp_start = expected_curves[0].GetEndPoint(0)
            act_start = actual[0].GetEndPoint(0)
            delta = exp_start.Subtract(act_start)
            dist = float(delta.GetLength())
        if diag:
            diag.log(
                u"Alineado cadena: mean_delta={0:.2f} mm".format(_internal_to_mm(dist))
            )
        if dist < _mm_to_internal(0.5):
            return
        ElementTransformUtils.MoveElement(document, new_rebar.Id, delta)
        if diag:
            diag.log(
                u"Alineado inicio/cadena {0:.2f} mm (cadena esperada vs creada)".format(
                    _internal_to_mm(dist)
                )
            )
    except Exception as ex:
        if diag:
            diag.ex(u"Alineado cadena", ex)


def _shape_driven_accessor(rebar):
    try:
        return rebar.GetShapeDrivenAccessor()
    except Exception:
        return None


def _es_rebar_seleccionable(rebar):
    """Rebar shape-driven con al menos una posición incluida y línea media legible."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    if _shape_driven_accessor(rebar) is None:
        return False
    n = _cantidad_posiciones(rebar)
    for i in range(max(1, n)):
        try:
            if not rebar.IsBarIncluded(int(i)):
                continue
        except Exception:
            pass
        try:
            if _centerline_chain(rebar, i):
                return True
        except Exception:
            pass
        for chain in _centerline_chain_variants(rebar, i):
            if chain:
                return True
    return False


def _bar_index_desde_referencia(rebar, reference):
    if rebar is None or reference is None:
        return -1
    try:
        return int(rebar.GetBarIndexFromReference(reference))
    except Exception:
        return -1


def _bar_index_desde_punto(rebar, point):
    n = _cantidad_posiciones(rebar)
    if n <= 1:
        return 0
    if point is None:
        return 0
    best_i = 0
    best_perp = None
    for i in range(n):
        try:
            if not rebar.IsBarIncluded(int(i)):
                continue
        except Exception:
            pass
        chain = _centerline_chain(rebar, i)
        if not chain:
            continue
        _arc, perp = _closest_arc_distance_on_chain(chain, point)
        if perp is None:
            continue
        if best_perp is None or perp < best_perp:
            best_perp = perp
            best_i = i
    return best_i


def _excluir_barra_del_conjunto(rebar, bar_index, document, diag=None, allow_regenerate=False):
    try:
        idx = int(bar_index)
    except Exception:
        return False
    n = _cantidad_posiciones(rebar)
    if idx < 0 or idx >= n:
        if diag:
            diag.log(u"SetBarIncluded: índice {0} fuera de rango (n={1})".format(idx, n))
        return False
    try:
        if hasattr(rebar, "DoesBarExistAtPosition") and not rebar.DoesBarExistAtPosition(idx):
            if diag:
                diag.log(u"SetBarIncluded: no existe posición {0}".format(idx))
            return False
    except Exception:
        pass
    try:
        rebar.SetBarIncluded(False, idx)
        if diag:
            diag.log(u"SetBarIncluded(False, {0}) OK".format(idx))
    except Exception as ex:
        if diag:
            diag.ex(u"SetBarIncluded", ex)
        return False
    if not allow_regenerate:
        return True
    for _ in range(2):
        try:
            document.Regenerate()
        except Exception as ex:
            if diag:
                diag.ex(u"Regenerate (post-exclusión)", ex)
        if not _bar_included(rebar, idx):
            return True
    ok = not _bar_included(rebar, idx)
    if diag:
        diag.log(u"Verificación post-regen índice {0}: incluida={1}".format(idx, not ok))
    return ok


def _rebar_nominal_diameter_mm(bar_type):
    if not isinstance(bar_type, RebarBarType):
        return None
    try:
        d_mm = int(round(float(bar_type.BarNominalDiameter) * 304.8))
        return d_mm if d_mm > 0 else None
    except Exception:
        return None


def _hook_type(doc, hook_id):
    if hook_id is None or hook_id == ElementId.InvalidElementId:
        return None
    e = doc.GetElement(hook_id)
    return e if isinstance(e, RebarHookType) else None


def _rebar_normal(rebar):
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            n = acc.Normal
            if n is not None and n.GetLength() > 1e-12:
                return n.Normalize()
    except Exception:
        pass
    return XYZ.BasisZ


def _curves_to_array(curves_clean):
    ct = _curve_clr_type()
    n = len(curves_clean)
    arr = System.Array.CreateInstance(ct, n)
    for i in range(n):
        arr[i] = curves_clean[i]
    return arr


def _curves_to_list(curves_clean):
    lst = List[_curve_clr_type()]()
    for c in curves_clean:
        lst.Add(c)
    return lst


def _chain_has_bends(curves):
    n = 0
    for c in curves or []:
        if c is None:
            continue
        if _curve_is_arc(c):
            return True
        n += 1
        if n > 1:
            return True
    return False


def _rebar_shape_element(doc, rebar):
    if doc is None or rebar is None:
        return None
    try:
        sid = rebar.GetShapeId()
        if sid is not None and sid != ElementId.InvalidElementId:
            sh = doc.GetElement(sid)
            if sh is not None:
                return sh
    except Exception:
        pass
    return None


def _aplicar_ganchos_post_creacion(rebar, hook_start, hook_end):
    if rebar is None:
        return
    inv = ElementId.InvalidElementId
    for end_idx, hook in ((0, hook_start), (1, hook_end)):
        try:
            if hook is not None:
                rebar.SetHookTypeId(end_idx, hook.Id)
            else:
                rebar.SetHookTypeId(end_idx, inv)
        except Exception:
            pass


def _hook_create_variants(start_hook, end_hook):
    """Solo RebarHookType o None — IronPython/Revit 2024+ no acepta ElementId aquí."""
    if start_hook is not None and end_hook is not None:
        return ((start_hook, end_hook),)
    if start_hook is not None:
        return ((start_hook, None), (None, None))
    if end_hook is not None:
        return ((None, end_hook), (None, None))
    return ((None, None),)


def _orient_create_variants(start_orient, end_orient):
    return (
        (start_orient, end_orient),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )


def _create_rebar_chunk(
    doc,
    curves_list,
    host,
    norm,
    bar_type,
    style,
    start_hook,
    end_hook,
    start_orient,
    end_orient,
    diag=None,
    chunk_label=u"?",
    rebar_src=None,
):
    """
    Crea Rebar solo desde curvas recortadas — sin RebarShape ni largo del origen.

  ``useExistingShapeIfPossible=False`` y ``createNewShape=False`` evitan que Revit
    ajuste la geometría a una forma/largo catalogados.
    """
    _fail_logged = set()

    def _norm_variants(nvec, curves_try):
        """Prioriza la normal del Rebar origen; evita que un normal-from-chain
        (p.ej. desde la pata L) gane y deje el conjunto al lado opuesto."""
        out = []
        seen = []

        def _add(n):
            if n is None:
                return
            try:
                nn = n.Normalize()
            except Exception:
                nn = n
            for s in seen:
                d = _dot_safe(s, nn)
                if d is not None and abs(d) > 0.999:
                    return
            seen.append(nn)
            out.append(nn)

        _add(nvec)
        try:
            if nvec is not None:
                _add(nvec.Negate())
        except Exception:
            pass
        nc = _normal_from_chain(curves_try, None)
        _add(nc)
        try:
            if nc is not None:
                _add(nc.Negate())
        except Exception:
            pass
        return out if out else [XYZ.BasisZ]

    def _log_fail(key, via, use_exist, create_new, ex):
        if not diag or key in _fail_logged:
            return
        _fail_logged.add(key)
        diag.ex(
            u"CreateFromCurves FAIL [{0}] via={1} flags={2}/{3}".format(
                chunk_label, via, use_exist, create_new
            ),
            ex,
        )

    def _try_create(curves_clean, nvec, h0, h1, so, eo, use_exist, create_new, container, via):
        if not curves_clean:
            return None
        try:
            if container == "list":
                payload = _curves_to_list(curves_clean)
            else:
                payload = _curves_to_array(curves_clean)
            rb = Rebar.CreateFromCurves(
                doc,
                style,
                bar_type,
                h0,
                h1,
                host,
                nvec,
                payload,
                so,
                eo,
                bool(use_exist),
                bool(create_new),
            )
            if rb is not None and diag:
                diag.log(
                    u"CreateFromCurves OK [{0}] via={1} n={2} flags={3}/{4} normal={5}".format(
                        chunk_label,
                        via,
                        len(curves_clean),
                        use_exist,
                        create_new,
                        _vec_fmt(nvec),
                    )
                )
            return rb
        except Exception as ex:
            _log_fail(
                (via, use_exist, create_new, _exception_text(ex)[:80]),
                via,
                use_exist,
                create_new,
                ex,
            )
            return None

    curves_clean = _prepare_curves_for_rebar(
        doc, curves_list, diag=diag, label=chunk_label + u"/raw", aggressive_heal=False
    )
    if not curves_clean:
        curves_clean = list(curves_list or [])
    if diag:
        _describe_curve_chain(diag, doc, chunk_label + u"/prep", curves_clean)
        for i, c in enumerate(curves_clean):
            try:
                if _curve_is_arc(c):
                    diag.log(
                        u"  arco[{0}] L={1:.1f} mm".format(
                            i, _internal_to_mm(_curve_length_safe(c))
                        )
                    )
            except Exception:
                pass

    bent = _chain_has_bends(curves_clean)
    # Revit exige al menos un flag True; (False,True) = shape nueva solo desde curvas.
    if bent:
        flag_pairs = ((False, True), (True, False))
    else:
        flag_pairs = ((False, True), (True, False), (True, True))

    h0 = None if bent else start_hook
    h1 = None if bent else end_hook
    hook_variants = ((h0, h1),) if bent else _hook_create_variants(start_hook, end_hook)

    for h0_try, h1_try in hook_variants:
        for so, eo in _orient_create_variants(start_orient, end_orient):
            for nvec in _norm_variants(norm, curves_clean):
                for use_exist, create_new in flag_pairs:
                    for container in ("array", "list"):
                        via = u"curves/{0}".format(container)
                        rb = _try_create(
                            curves_clean,
                            nvec,
                            h0_try,
                            h1_try,
                            so,
                            eo,
                            use_exist,
                            create_new,
                            container,
                            via,
                        )
                        if rb is not None:
                            if not bent:
                                _aplicar_ganchos_post_creacion(rb, start_hook, end_hook)
                            return rb

    raise RuntimeError(
        u"CreateFromCurves [{0}]: ningún intento produjo Rebar válido.".format(chunk_label)
    )


def _centerline_chain(rebar, bar_index=0):
    """
    Centerline en posición visual (BarPosition + MovedBar).

    GetCenterlineCurves en shape-driven suele devolver la geometría en la
    posición «de manejo» de la barra 0; la posición real en modelo usa
    GetTransformedCenterlineCurves.
    """
    bi = int(bar_index)
    try:
        curves = rebar.GetTransformedCenterlineCurves(
            False,
            False,
            False,
            MultiplanarOption.IncludeAllMultiplanarCurves,
            bi,
        )
        if curves is not None and curves.Count > 0:
            return [curves[i] for i in range(curves.Count)]
    except Exception:
        pass
    curves = rebar.GetCenterlineCurves(
        False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, bi
    )
    if curves is None or curves.Count == 0:
        return []
    out = [curves[i] for i in range(curves.Count)]
    # Fallback BuildingCoder: aplicar solo BarPositionTransform.
    try:
        if bi != 0 and _shape_driven_accessor(rebar) is not None:
            tr = _get_bar_transform(rebar, bi)
            if tr is not None and not tr.IsIdentity:
                out = [c.CreateTransformed(tr) for c in out]
    except Exception:
        pass
    return out


def _centerline_chain_driving(rebar, bar_index=0):
    """Curvas de manejo (sin MovedBar); solo para diagnóstico."""
    bi = int(bar_index)
    curves = rebar.GetCenterlineCurves(
        False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, bi
    )
    if curves is None or curves.Count == 0:
        return []
    return [curves[i] for i in range(curves.Count)]


def _moved_bar_transform(rebar, bar_index):
    try:
        return rebar.GetMovedBarTransform(int(bar_index))
    except Exception:
        return None


def _log_moved_bar_transforms(diag, label, rebar):
    if diag is None or rebar is None:
        return
    n = _cantidad_posiciones(rebar)
    for i in range(max(1, n)):
        mt = _moved_bar_transform(rebar, i)
        if mt is None:
            diag.log(u"MovedBar[{0}] idx={1}: None".format(label, i))
            continue
        try:
            ident = bool(mt.IsIdentity)
        except Exception:
            ident = False
        origin = None
        try:
            origin = mt.Origin
        except Exception:
            pass
        diag.log(
            u"MovedBar[{0}] idx={1}: identity={2} origin={3}".format(
                label, i, ident, _xyz_fmt(origin)
            )
        )


def _copy_moved_bar_transforms(src, dst, diag=None):
    """Replica desplazamientos manuales por barra (MoveBarInSet)."""
    if src is None or dst is None:
        return
    n = min(_cantidad_posiciones(src), _cantidad_posiciones(dst))
    for i in range(n):
        mt = _moved_bar_transform(src, i)
        if mt is None:
            continue
        try:
            if bool(mt.IsIdentity):
                continue
        except Exception:
            pass
        try:
            dst.MoveBarInSet(int(i), mt)
            if diag:
                diag.log(u"MoveBarInSet idx={0} OK".format(i))
        except Exception as ex:
            if diag:
                diag.ex(u"MoveBarInSet idx={0}".format(i), ex)


def _log_driving_vs_transformed(diag, rebar):
    """Cuantifica desfase entre curvas de manejo y posición visual (barra 0)."""
    if diag is None or rebar is None:
        return
    try:
        drive = _centerline_chain_driving(rebar, 0)
        vis = _centerline_chain(rebar, 0)
        if not drive or not vis:
            return
        p0 = drive[0].GetEndPoint(0)
        p1 = vis[0].GetEndPoint(0)
        d = float(p0.DistanceTo(p1))
        diag.log(
            u"DrivingVsVisual bar0: d={0:.2f} mm drive0={1} vis0={2}".format(
                _internal_to_mm(d), _xyz_fmt(p0), _xyz_fmt(p1)
            )
        )
    except Exception as ex:
        diag.ex(u"DrivingVsVisual", ex)


def _rebar_midpoint_xyz(rebar, bar_index=0):
    chain = _centerline_chain(rebar, bar_index)
    if not chain:
        return None
    total = _chain_total_length(chain)
    if total <= 1e-12:
        return None
    half = 0.5 * total
    accum = 0.0
    for c in chain:
        cl = float(c.Length)
        if half <= accum + cl + 1e-9:
            local = half - accum
            try:
                par = _param_at_dist_from_start(c, local)
                return c.Evaluate(par, False)
            except Exception:
                return None
        accum += cl
    try:
        return chain[-1].GetEndPoint(1)
    except Exception:
        return None


def _copy_instance_parameters(src, dst, skip_names=None):
    skip = set(skip_names or ())
    skip.update(
        (
            u"Rebar Number",
            u"Bar Count",
            u"Quantity",
            u"Total Bar Length",
            u"Max Rebar Length",
            u"Bar Length",
            u"Actual Bar Length",
            u"Rebar Shape",
            u"Shape",
            u"Shape Image",
        )
    )
    for p in src.Parameters:
        if p is None or p.IsReadOnly:
            continue
        try:
            name = p.Definition.Name
        except Exception:
            continue
        if name in skip:
            continue
        dp = dst.LookupParameter(name)
        if dp is None or dp.IsReadOnly:
            continue
        try:
            st = p.StorageType
            if st == StorageType.String:
                dp.Set(p.AsString() or u"")
            elif st == StorageType.Integer:
                dp.Set(p.AsInteger())
            elif st == StorageType.Double:
                dp.Set(p.AsDouble())
            elif st == StorageType.ElementId:
                dp.Set(p.AsElementId())
        except Exception:
            pass


def _element_id_int(eid):
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        return int(eid.IntegerValue)
    except AttributeError:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _tag_rebar_int_if_match(tag, rebar_set, invalid):
    if tag is None:
        return None
    try:
        if getattr(tag, u"IsOrphaned", False):
            return None
    except Exception:
        pass
    for getter in (
        lambda: tag.GetTaggedLocalElementIds(),
        lambda: tag.GetTaggedElementIds(),
    ):
        try:
            ids = getter()
            if ids is None:
                continue
            for leid in ids:
                ti = _element_id_int(leid)
                if ti is not None and ti in rebar_set:
                    return ti
        except Exception:
            continue
    try:
        ref_ids = tag.GetTaggedReferences()
        if ref_ids:
            for r in ref_ids:
                ti = _element_id_int(r.ElementId)
                if ti is not None and ti in rebar_set:
                    return ti
    except Exception:
        pass
    try:
        rid = tag.TaggedLocalElementId
        ti = _element_id_int(rid)
        if ti is not None and ti in rebar_set:
            return ti
    except Exception:
        pass
    return None


def _capturar_etiquetas_rebar(doc, rebar_id):
    rid = _element_id_int(rebar_id)
    if rid is None:
        return []
    rebar_set = {rid}
    invalid = ElementId.InvalidElementId
    out = []
    try:
        coll = (
            FilteredElementCollector(doc)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return out
    for tag in coll:
        if _tag_rebar_int_if_match(tag, rebar_set, invalid) is None:
            continue
        info = {}
        try:
            info[u"type_id"] = tag.GetTypeId()
        except Exception:
            info[u"type_id"] = None
        try:
            info[u"view_id"] = tag.OwnerViewId
        except Exception:
            info[u"view_id"] = None
        try:
            info[u"head"] = tag.TagHeadPosition
        except Exception:
            info[u"head"] = None
        try:
            info[u"orient"] = tag.TagOrientation
        except Exception:
            info[u"orient"] = TagOrientation.Horizontal
        try:
            info[u"leader"] = bool(tag.HasLeader)
        except Exception:
            info[u"leader"] = True
        try:
            fam = doc.GetElement(info[u"type_id"])
            if fam is not None and hasattr(fam, u"Family"):
                sf = fam.Family
                info[u"family_name"] = sf.Name if sf is not None else u""
            else:
                info[u"family_name"] = u""
        except Exception:
            info[u"family_name"] = u""
        if info.get(u"view_id") is None or info.get(u"type_id") is None:
            continue
        out.append(info)
    return out


def _referencias_tag_rebar(doc, rebar, view):
    refs = []
    seen = set()

    def _add(r):
        if r is None:
            return
        try:
            key = r.ConvertToStableRepresentation(doc)
        except Exception:
            key = id(r)
        if key in seen:
            return
        seen.add(key)
        refs.append(r)

    try:
        subs = rebar.GetSubelements() if hasattr(rebar, "GetSubelements") else None
    except Exception:
        subs = None
    if subs:
        for sub in subs:
            if sub is None:
                continue
            try:
                sref = sub.GetReference() if hasattr(sub, "GetReference") else None
                if sref is not None:
                    _add(sref)
            except Exception:
                pass
    try:
        npos = int(rebar.NumberOfBarPositions)
    except Exception:
        npos = 0
    if npos > 0:
        for idx in (0, max(0, npos - 1)):
            try:
                if hasattr(rebar, "GetReferenceToBarPosition"):
                    rpos = rebar.GetReferenceToBarPosition(idx)
                elif hasattr(rebar, "GetReferenceForBarPosition"):
                    rpos = rebar.GetReferenceForBarPosition(idx)
                else:
                    rpos = None
                if rpos is not None:
                    _add(rpos)
            except Exception:
                pass
    try:
        _add(Reference(rebar))
    except Exception:
        pass
    return refs


def _crear_etiqueta_rebar(doc, view, rebar, type_id, head, orient, add_leader):
    if view is None or type_id is None or rebar is None:
        return None
    if head is None:
        head = _rebar_midpoint_xyz(rebar, 0)
    if head is None:
        return None
    try:
        sym = doc.GetElement(type_id)
        if sym is not None and not sym.IsActive:
            sym.Activate()
    except Exception:
        pass
    refs = _referencias_tag_rebar(doc, rebar, view)
    if not refs:
        return None
    for ref in refs:
        try:
            tag = IndependentTag.Create(
                doc, type_id, view.Id, ref, add_leader, orient, head
            )
            if tag is not None:
                return tag
        except Exception:
            pass
        try:
            tag = IndependentTag.Create(
                doc,
                view.Id,
                ref,
                add_leader,
                TagMode.TM_ADDBY_CATEGORY,
                orient,
                head,
            )
            if tag is not None:
                try:
                    tag.ChangeTypeId(type_id)
                except Exception:
                    try:
                        tag.SetTypeId(type_id)
                    except Exception:
                        pass
                return tag
        except Exception:
            pass
    return None


def _recrear_etiquetas(doc, tag_infos, new_rebars):
    creadas = 0
    for rb in new_rebars:
        if rb is None:
            continue
        head_rb = _rebar_midpoint_xyz(rb, 0)
        for info in tag_infos:
            view = doc.GetElement(info.get(u"view_id"))
            if view is None:
                continue
            head = head_rb or info.get(u"head")
            tag = _crear_etiqueta_rebar(
                doc,
                view,
                rb,
                info.get(u"type_id"),
                head,
                info.get(u"orient", TagOrientation.Horizontal),
                info.get(u"leader", True),
            )
            if tag is not None:
                creadas += 1
    return creadas


def _lap_mm_for_rebar(rebar, bar_type):
    d_mm = _rebar_nominal_diameter_mm(bar_type)
    if d_mm is None:
        return None, d_mm
    if traslape_mm_from_nominal_diameter_mm is not None:
        lap = traslape_mm_from_nominal_diameter_mm(d_mm)
        if lap is not None and lap > 0:
            return float(lap), d_mm
    return 40.0 * float(d_mm), d_mm


def dividir_rebar_en_cortes(
    document,
    rebar,
    cuts_mm_on_main,
    lap_mm=None,
    diag=None,
    splice_mode=None,
):
    """
    Divide la barra original en tramos según cortes (mm) sobre el segmento mayor.

    Pipeline:
    1. Centerline índice 0 → segmento mayor + patas L.
    2. Recorta el vano según cortes y traslape (modo de solape).
    3. Crea cada tramo y replica layout del origen.

    ``splice_mode``: symmetric | forward | backward (ver SPLICE_*).

    Returns:
        (ok, mensaje, ids_nuevos)
    """
    diag = diag or _DiagSession()
    mode = normalize_splice_mode(splice_mode)
    if not isinstance(rebar, Rebar):
        return False, u"No es un elemento Rebar.", []
    if _shape_driven_accessor(rebar) is None:
        return (
            False,
            u"Solo aplica a barras shape-driven (Structural Rebar con forma). "
            u"Las barras free-form no están soportadas.",
            [],
        )

    chain = _centerline_chain(rebar, 0)
    if not chain:
        return False, diag.failure_message(u"No se pudo leer la línea media."), []

    layout_geom = _decompose_bar_chain(chain)
    main_len_mm = _internal_to_mm(float(layout_geom[u"main_len"]))
    if main_len_mm <= 1e-6:
        return False, u"Sin vano principal para dividir.", []

    host = document.GetElement(rebar.GetHostId())
    if host is None:
        return False, u"La barra no tiene host válido.", []

    bar_type = document.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"No se pudo obtener RebarBarType.", []

    if lap_mm is None:
        lap_mm, d_mm = _lap_mm_for_rebar(rebar, bar_type)
        if lap_mm is None:
            return False, u"No se pudo calcular el traslape según el diámetro.", []
    else:
        d_mm = _rebar_nominal_diameter_mm(bar_type)

    ok_val, err_val = _validate_cuts_on_main(
        cuts_mm_on_main, main_len_mm, lap_mm, splice_mode=mode
    )
    if not ok_val:
        return False, diag.failure_message(err_val), []

    cuts_mm = sorted(set(float(c) for c in cuts_mm_on_main))
    cuts_ft = [_mm_to_internal(c) for c in cuts_mm]
    lap_ft = _mm_to_internal(lap_mm)
    main_len = float(layout_geom[u"main_len"])
    ranges = _piece_ranges_on_main(cuts_ft, main_len, lap_ft, mode)
    n_pieces = len(ranges)

    rid = _element_id_int(rebar.Id)
    n_pos = _cantidad_posiciones(rebar)
    diag.step(u"Entrada")
    diag.log(
        u"Rebar Id={0} n_pos={1} cortes_vano={2} solape={3}".format(
            rid, n_pos, [int(round(c)) for c in cuts_mm], mode
        )
    )
    _describe_curve_chain(diag, document, u"centerline/idx0", chain)
    diag.log(
        u"Segmento mayor: idx={0} L={1:.1f} mm (pref={2:.1f} suf={3:.1f})".format(
            layout_geom[u"main_index"],
            main_len_mm,
            _internal_to_mm(float(layout_geom[u"prefix_len"])),
            _internal_to_mm(float(layout_geom[u"suffix_len"])),
        )
    )

    chunk_specs = []
    for i, (d0, d1) in enumerate(ranges):
        h0 = i == 0
        h1 = i == n_pieces - 1
        label = u"T{0}".format(i + 1)
        if d1 <= d0 + 1e-9:
            return (
                False,
                diag.failure_message(u"Rango inválido en tramo {0}.".format(label)),
                [],
            )
        curves = _compose_chunk_curves(
            layout_geom, d0, d1, document, diag=diag, label=label
        )
        if not curves:
            return False, diag.failure_message(u"No se pudo generar tramo {0}.".format(label)), []
        chunk_specs.append((label, curves, h0, h1))

    diag.step(u"Tramos recortados")
    diag.log(
        u"traslape={0:.0f} mm Ø={1} mm modo={2}".format(
            float(lap_mm), d_mm or u"?", mode
        )
    )
    for label, curves, _h0, _h1 in chunk_specs:
        _describe_curve_chain(diag, document, label, curves)

    try:
        style = rebar.Style
    except Exception:
        style = RebarStyle.Standard
    norm = _rebar_normal(rebar)
    if diag:
        diag.log(u"Normal origen Rebar={0}".format(_vec_fmt(norm)))
        _log_driving_vs_transformed(diag, rebar)
        _log_moved_bar_transforms(diag, u"orig", rebar)
    hook_start = _hook_type(document, rebar.GetHookTypeId(0))
    hook_end = _hook_type(document, rebar.GetHookTypeId(1))
    try:
        so0 = rebar.GetHookOrientation(0)
        so1 = rebar.GetHookOrientation(1)
    except Exception:
        so0 = RebarHookOrientation.Right
        so1 = RebarHookOrientation.Left

    orig_has_l_geometry = (
        float(layout_geom[u"prefix_len"]) > 1e-9
        or float(layout_geom[u"suffix_len"]) > 1e-9
        or _chain_has_bends(chain)
    )
    if orig_has_l_geometry and diag:
        diag.log(u"Patas L en centerline: sin RebarShape ni ganchos Revit en tramos.")

    tag_infos = _capturar_etiquetas_rebar(document, rebar.Id)
    old_id = rebar.Id
    if n_pos > 1 or n_pieces != 2:
        tag_infos = []

    t = Transaction(document, _TRANSACTION_NAME)
    t.Start()
    nuevos = []
    try:
        for label, curves, has_start, has_end in chunk_specs:
            diag.step(u"Transacción — crear {0}".format(label))
            if orig_has_l_geometry:
                chunk_h0 = None
                chunk_h1 = None
            else:
                chunk_h0 = hook_start if has_start else None
                chunk_h1 = hook_end if has_end else None
            rb = _create_rebar_chunk(
                document,
                curves,
                host,
                norm,
                bar_type,
                style,
                chunk_h0,
                chunk_h1,
                so0 if has_start else RebarHookOrientation.Right,
                so1 if has_end else RebarHookOrientation.Left,
                diag=diag,
                chunk_label=label,
            )
            if rb is None:
                raise RuntimeError(u"CreateFromCurves devolvió None en {0}.".format(label))
            ok_lay, err_lay = _finalize_new_rebar_set(
                document, rebar, rb, curves, diag=diag
            )
            if not ok_lay and n_pos > 1:
                raise RuntimeError(
                    err_lay or u"No se pudo replicar el layout del conjunto."
                )
            _copy_instance_parameters(rebar, rb)
            try:
                document.Regenerate()
            except Exception:
                pass
            # Parámetros/constraints pueden desplazar tras copiar: reimponer pose.
            _alinear_rebar_a_cadena_esperada(document, rb, curves, diag=diag)
            if n_pos > 1:
                _copy_moved_bar_transforms(rebar, rb, diag=diag)
            nuevos.append(rb)
            if diag:
                _describe_curve_chain(
                    diag, document, label + u"/creado", _centerline_chain(rb, 0)
                )
                _log_pose_compare(diag, label + u"/final", rebar, rb, curves)
            diag.log(u"Creado {0} Id={1}".format(label, _element_id_int(rb.Id)))

        diag.step(u"Eliminar barra original")
        document.Delete(old_id)

        n_tags = 0
        if tag_infos and len(nuevos) == 2:
            diag.step(u"Recrear etiquetas")
            n_tags = _recrear_etiquetas(document, tag_infos, nuevos)

        diag.step(u"Commit")
        t.Commit()
    except Exception as ex:
        t.RollBack()
        diag.ex(u"ROLLBACK", ex)
        return False, diag.failure_message(_exception_text(ex) or u"Error en transacción."), []

    ids = [rb.Id for rb in nuevos]
    rule = _layout_rule_nombre(rebar)
    conjunto_txt = u""
    if n_pos > 1:
        conjunto_txt = (
            u" Cada tramo replica el conjunto original ({} pos., regla {}).".format(
                n_pos, rule or u"?"
            )
        )
    cuts_txt = u", ".join(u"{:.0f}".format(c) for c in cuts_mm)
    mode_lbl = SPLICE_MODE_LABELS.get(mode, mode)
    diag.step(u"OK")
    detalle = (
        u"{0} tramo(s) en vano principal (cortes: {1} mm; traslape {2:.0f} mm, Ø {3} mm; "
        u"solape: {4}).{5} Etiquetas recreadas: {6}."
    ).format(
        len(nuevos),
        cuts_txt,
        float(lap_mm),
        d_mm if d_mm is not None else u"?",
        mode_lbl,
        conjunto_txt,
        n_tags,
    )
    if diag.path():
        detalle += u"\n\nLog diagnóstico:\n{0}".format(diag.path())
    return True, detalle, ids


def dividir_rebar_en_punto(
    document, rebar, pick_point, bar_index=None, lap_mm=None, diag=None, splice_mode=None
):
    """
    Divide ``rebar`` en el punto ``pick_point`` con traslape (un solo corte).

    El corte se proyecta sobre el vano principal (segmento más largo).

    Returns:
        (ok, mensaje, ids_nuevos)
    """
    diag = diag or _DiagSession()
    if pick_point is None:
        return False, u"Punto de división no válido.", []

    bar_index, cut_dist, _perp, chain, err = _resolver_barra_y_corte(
        rebar, pick_point, bar_index_hint=bar_index
    )
    if err:
        return False, diag.failure_message(err), []
    if bar_index is None or cut_dist is None or not chain:
        return False, diag.failure_message(u"No se pudo proyectar el punto sobre la línea media."), []

    layout = _decompose_bar_chain(chain)
    cut_main_mm = _main_span_cut_mm_from_full_chain_dist(layout, cut_dist)
    if cut_main_mm is None:
        cut_main_mm, _ = _main_span_cut_mm_from_point(layout, pick_point)
    if cut_main_mm is None:
        return (
            False,
            diag.failure_message(
                u"El punto debe caer sobre el vano principal (tramo más largo), "
                u"no sobre las patas L."
            ),
            [],
        )

    return dividir_rebar_en_cortes(
        document,
        rebar,
        [cut_main_mm],
        lap_mm=lap_mm,
        diag=diag,
        splice_mode=splice_mode,
    )


class _FiltroRebarShapeDriven(ISelectionFilter):
    def AllowElement(self, elem):
        return _es_rebar_seleccionable(elem)

    def AllowReference(self, reference, position):
        return False


class _FiltroPuntoEnRebar(ISelectionFilter):
    def __init__(self, rebar_id):
        self._rebar_id = rebar_id

    def AllowElement(self, elem):
        if not isinstance(elem, Rebar):
            return False
        try:
            return elem.Id == self._rebar_id
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return True


def _object_type_point_on_element():
    try:
        return ObjectType.PointOnElement
    except Exception:
        return None


def _pick_rebar(uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroRebarShapeDriven(),
            u"1/1 — Selecciona la barra original (individual o conjunto con layout).",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    el = uidoc.Document.GetElement(ref.ElementId)
    if not _es_rebar_seleccionable(el):
        TaskDialog.Show(
            _DIALOG_TITLE,
            u"El elemento no es una barra shape-driven válida "
            u"(revisa free-form, geometría o posiciones del conjunto).",
        )
        return None
    return el


def _pick_point_on_rebar(uidoc, rebar):
    filt = _FiltroPuntoEnRebar(rebar.Id)
    prompt = (
        u"Clic sobre la barra (segmento mayor / vano) para indicar el punto de división."
    )
    bar_index = None
    point = None

    ot_sub = None
    try:
        ot_sub = ObjectType.Subelement
    except Exception:
        pass
    if ot_sub is not None:
        try:
            ref = uidoc.Selection.PickObject(ot_sub, filt, prompt)
            if ref is not None:
                idx = _bar_index_desde_referencia(rebar, ref)
                if idx >= 0:
                    bar_index = idx
                try:
                    point = ref.GlobalPoint
                except Exception:
                    pass
        except OperationCanceledException:
            return None, None
        except Exception:
            pass

    ot_poe = _object_type_point_on_element()
    if point is None and ot_poe is not None:
        try:
            ref = uidoc.Selection.PickObject(ot_poe, filt, prompt)
            if ref is not None:
                idx = _bar_index_desde_referencia(rebar, ref)
                if idx >= 0:
                    bar_index = idx
                try:
                    point = ref.GlobalPoint
                except Exception:
                    pass
        except OperationCanceledException:
            return None, None
        except Exception:
            pass

    if point is None:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                filt,
                prompt + u" (clic en la barra visible).",
            )
            if ref is not None:
                idx = _bar_index_desde_referencia(rebar, ref)
                if idx >= 0:
                    bar_index = idx
                try:
                    point = ref.GlobalPoint
                except Exception:
                    pass
        except OperationCanceledException:
            return None, None
        except Exception:
            pass

    if point is None:
        try:
            point = uidoc.Selection.PickPoint(
                u"2/2 — Clic cerca de la barra para indicar el punto de división."
            )
        except OperationCanceledException:
            return None, None
        except Exception:
            return None, None

    return point, bar_index


def run_pyrevit(__revit__):
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(_DIALOG_TITLE, u"No hay documento activo.")
        return

    rebar = _pick_rebar(uidoc)
    if rebar is None:
        return

    try:
        import dividir_barra_traslape_ui as _ui_mod
    except Exception:
        try:
            import imp
            import os

            _p = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                u"dividir_barra_traslape_ui.py",
            )
            _ui_mod = imp.load_source(u"dividir_barra_traslape_ui", _p)
        except Exception as ex:
            TaskDialog.Show(
                _DIALOG_TITLE,
                u"No se pudo cargar la interfaz:\n\n{0}".format(_exception_text(ex)),
            )
            return

    _ui_mod.show_dividir_barra_window(__revit__, rebar)
