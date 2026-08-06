# -*- coding: utf-8 -*-
"""
To Single — convierte un conjunto de Rebar (layout ≠ Single) en barras individuales.

Cada single hereda la geometría (copy del set + SetLayoutAsSingle + traslado a
su posición) y los parámetros de instancia ``Armadura_*`` del set original.

Si ``Armadura_Malla`` = Yes, se reaplican las etiquetas del set original en
**una sola** barra generada (cabeza en centroide del muro host).

Al pasar a Single la etiqueta puede dejar de reportar la separación: se crea un
TextNote ``@NNN`` (tipo ``2.5mm Arial_Arrow Filled 15 Degree``) **solo si** las
etiquetas de esa barra no muestran ya ``@`` numérico. Offset papel calibrado a
1:50; línea V. sube ~3 mm papel respecto a la cabeza compartida (evita solapar H.).

Revit 2024+ | IronPython | pyRevit
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    ElementId,
    ElementTransformUtils,
    Transaction,
    TransactionGroup,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

_TRANSACTION_NAME = u"Arainco: To Single"
_TX_GROUP_NAME = u"Arainco: To Single"
_TX_TAGS = u"Arainco: To Single — tags malla"
_ARMADURA_MALLA = u"Armadura_Malla"
_XYZ_ZERO = XYZ.Zero

# TextNote de separación (si la etiqueta ya no reporta @NNN).
# Prioridad: borde derecho del BoundingBox de la IndependentTag (mm papel vía escala
# de vista). Respaldo: media anchura del TagText (Arial 2.5 mm).
# Bloque malla (cabeza compartida): V. y H. en líneas distintas → dy sale del BB.
_SPACING_TEXT_TYPE_NAME = u"2.5mm Arial_Arrow Filled 15 Degree"
_SPACING_TEXT_HEIGHT_PAPER_MM = 2.5
_SPACING_CHAR_WIDTH_FACTOR = 0.45  # Arial condensado / media real < 0.52
_SPACING_TEXT_DX_PAPER_MM = 3.0  # respaldo sin TagText/BB
# Aire tras el diámetro (mm papel) — evita solape 0/@ y hueco excesivo.
_SPACING_TEXT_DX_GAP_PAPER_MM = 0.20
_SPACING_TEXT_DX_TIGHTEN = 0.88
_SPACING_BBOX_MAX_PAPER_MM = 22.0
_SPACING_BBOX_MIN_PAPER_MM = 1.2
# Extra en borde BB. Ligero a la izquierda del BB sin tapar el «0».
_SPACING_BBOX_DX_EXTRA_PAPER_MM = -0.20
# Nudge final (mm papel): menos izquierda = más a la derecha del @.
# 0.45 quedaba bien-ish; 0.70 lo corrió de más a la izq. Target: un pelo más a la der.
_SPACING_NUDGE_LEFT_PAPER_MM = 0.25
_SPACING_NUDGE_UP_PAPER_MM = 0.90
_SPACING_LINE_V_DY_PAPER_MM = 3.2
_SPACING_LINE_H_DY_PAPER_MM = 0.0
_SPACING_LINE_DEFAULT_DY_PAPER_MM = 0.0
_SPACING_REF_SCALE = 50

_PROG_BG = u"#071018"
_PROG_ACCENT = u"#5BC0DE"
_PROG_FG = u"#E8F4F8"


# ---------------------------------------------------------------------------
# Progress bar (pyRevit; no-op si no hay pyRevit)
# ---------------------------------------------------------------------------

def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


def _progress_palette():
    bg, accent, fg = _PROG_BG, _PROG_ACCENT, _PROG_FG
    try:
        from bimtools_ui_tokens import ACCENT_PRIMARY, BG_APP, FG_TITLE

        bg = BG_APP or bg
        accent = ACCENT_PRIMARY or accent
        fg = FG_TITLE or fg
    except Exception:
        pass
    return bg, accent, fg


def _color_from_hex(hex_str):
    from System.Windows.Media import Color

    s = (hex_str or u"").lstrip(u"#")
    if len(s) < 6:
        s = u"071018"
    return Color.FromRgb(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _apply_progress_theme(pb):
    if pb is None:
        return
    try:
        import System
        from System.Windows.Media import SolidColorBrush

        bg_h, ac_h, fg_h = _progress_palette()
        pb.Resources[u"pyRevitDarkBrush"] = SolidColorBrush(_color_from_hex(bg_h))
        pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(_color_from_hex(ac_h))
        pb.Resources[u"pyRevitAccentColor"] = _color_from_hex(ac_h)
        try:
            pb.Resources[System.Windows.SystemColors.WindowBrushKey] = (
                SolidColorBrush(_color_from_hex(fg_h))
            )
        except Exception:
            pass
    except Exception:
        pass


class ToSingleProgress(object):
    """Context manager de ProgressBar pyRevit (no-op si no disponible)."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._index = 0
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or u"Arainco: To Single"

    def __enter__(self):
        if not _pbar_enabled():
            return self
        try:
            from pyrevit import forms as _pyrevit_forms

            self._pb = _pyrevit_forms.ProgressBar(
                title=self._title(0),
                cancellable=False,
            )
            _apply_progress_theme(self._pb)
            self._pb.__enter__()
            self._open = True
        except Exception:
            self._pb = None
            self._open = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._open and self._pb is not None:
            try:
                self._pb.__exit__(exc_type, exc_val, exc_tb)
            except Exception:
                pass
        self._open = False
        self._pb = None
        return False

    def _title(self, index):
        return u"{0} {1}/{2}".format(
            self._title_prefix,
            min(int(index) + 1, self._total),
            int(self._total),
        )

    def step(self, phase_label=None):
        if self._pb is None:
            return
        i = int(self._index)
        if i >= self._total:
            i = self._total - 1
        self._index = i + 1
        label = phase_label or u""
        base = (
            u"{0} — {1}".format(self._title(i), label) if label else self._title(i)
        )
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(i + 1, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(i + 1, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass


def _estimate_progress_total(rebars):
    """Por set: 1 plan + N barras + 1 finalize."""
    total = 0
    if not rebars:
        return 1
    for rb in rebars:
        try:
            n = len(_indices_incluidos(rb))
        except Exception:
            n = 1
        total += 1 + max(1, int(n)) + 1
    return max(1, total)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _exception_text(ex):
    try:
        return _as_unicode(ex) if ex else u"Error desconocido."
    except Exception:
        return u"Error desconocido."


def _fmt_xyz(p):
    if p is None:
        return u"None"
    try:
        return u"({0:.4f}, {1:.4f}, {2:.4f})".format(
            float(p.X), float(p.Y), float(p.Z)
        )
    except Exception:
        return u"?"


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


class _RevitTransaction(object):
    """Context manager Transaction (Start / Commit|RollBack / Dispose)."""

    def __init__(self, document, name):
        self._doc = document
        self._name = name or _TRANSACTION_NAME
        self._t = None

    def __enter__(self):
        self._t = Transaction(self._doc, self._name)
        self._t.Start()
        return self._t

    def __exit__(self, exc_type, exc_val, exc_tb):
        t = self._t
        if t is None:
            return False
        try:
            try:
                started = bool(t.HasStarted())
                ended = bool(t.HasEnded())
            except Exception:
                started, ended = True, False
            if not started or ended:
                return False
            if exc_type is not None:
                try:
                    t.RollBack()
                except Exception:
                    pass
            else:
                try:
                    t.Commit()
                except Exception:
                    try:
                        t.RollBack()
                    except Exception:
                        pass
                    raise
        finally:
            try:
                t.Dispose()
            except Exception:
                pass
            self._t = None
        return False


class _RevitTransactionGroup(object):
    """Context manager TransactionGroup (Assimilate / RollBack / Dispose)."""

    def __init__(self, document, name):
        self._doc = document
        self._name = name or _TX_GROUP_NAME
        self._tg = None
        self._assimilate = True

    def __enter__(self):
        self._tg = TransactionGroup(self._doc, self._name)
        self._tg.Start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        tg = self._tg
        if tg is None:
            return False
        try:
            if exc_type is not None or not self._assimilate:
                try:
                    tg.RollBack()
                except Exception:
                    pass
            else:
                try:
                    tg.Assimilate()
                except Exception:
                    try:
                        tg.RollBack()
                    except Exception:
                        pass
        finally:
            try:
                tg.Dispose()
            except Exception:
                pass
            self._tg = None
        return False


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


def _es_layout_single(rebar):
    rule = _layout_rule_nombre(rebar)
    if not rule:
        return _cantidad_posiciones(rebar) <= 1
    return rule == u"Single" or rule.endswith(u".Single") or rule == u"Single"


def _bar_incluido(rebar, idx):
    i = int(idx)
    try:
        return bool(rebar.IsBarIncluded(i))
    except Exception:
        pass
    try:
        return bool(rebar.DoesBarExistAtPosition(i))
    except Exception:
        return True


def _indices_incluidos(rebar):
    n = max(0, _cantidad_posiciones(rebar))
    return [i for i in range(n) if _bar_incluido(rebar, i)]


def _get_bar_transform(rebar, bar_index):
    bi = int(bar_index)
    for getter in (
        lambda: rebar.GetBarPositionTransform(bi),
        lambda: rebar.GetMovedBarTransform(bi),
    ):
        try:
            t = getter()
            if t is not None:
                return t
        except Exception:
            continue
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, "GetBarPositionTransform"):
            return acc.GetBarPositionTransform(bi)
    except Exception:
        pass
    return None


def _iter_curve_list(curves):
    if curves is None:
        return []
    out = []
    try:
        n = int(curves.Count)
        for i in range(n):
            try:
                c = curves.get_Item(i)
            except Exception:
                try:
                    c = curves[i]
                except Exception:
                    c = None
            if c is not None:
                out.append(c)
        if out:
            return out
    except Exception:
        pass
    try:
        for c in curves:
            if c is not None:
                out.append(c)
    except Exception:
        pass
    return out


def _curve_midpoint(crv):
    if crv is None:
        return None
    try:
        return crv.Evaluate(0.5, True)
    except Exception:
        pass
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        return XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
    except Exception:
        return None


def _midpoint_from_curves(curves):
    clist = _iter_curve_list(curves)
    if not clist:
        return None
    best = None
    best_len = -1.0
    for c in clist:
        try:
            ln = float(c.Length)
        except Exception:
            ln = 0.0
        if ln >= best_len:
            best_len = ln
            best = c
    mid = _curve_midpoint(best)
    if mid is not None:
        return mid
    pts = []
    for c in clist:
        m = _curve_midpoint(c)
        if m is not None:
            pts.append(m)
    if not pts:
        return None
    sx = sy = sz = 0.0
    for p in pts:
        sx += float(p.X)
        sy += float(p.Y)
        sz += float(p.Z)
    n = float(len(pts))
    return XYZ(sx / n, sy / n, sz / n)


def _bar_world_midpoint(rebar, bar_index):
    if rebar is None:
        return None
    bi = int(bar_index)
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        mp = MultiplanarOption.IncludeAllMultiplanarCurves
    except Exception:
        mp = None

    if mp is not None and hasattr(rebar, "GetTransformedCenterlineCurves"):
        try:
            curves = rebar.GetTransformedCenterlineCurves(
                False, False, False, mp, bi
            )
            mid = _midpoint_from_curves(curves)
            if mid is not None:
                return mid
        except Exception:
            pass
        try:
            curves = rebar.GetTransformedCenterlineCurves(
                False, True, True, mp, bi
            )
            mid = _midpoint_from_curves(curves)
            if mid is not None:
                return mid
        except Exception:
            pass

    if mp is not None:
        try:
            curves = rebar.GetCenterlineCurves(False, False, False, mp, bi)
            mid = _midpoint_from_curves(curves)
            if mid is not None:
                tr = _get_bar_transform(rebar, bi)
                try:
                    if tr is not None and not bool(tr.IsIdentity):
                        return tr.OfPoint(mid)
                except Exception:
                    pass
                return mid
        except Exception:
            pass

    tr = _get_bar_transform(rebar, bi)
    if tr is not None:
        try:
            o = tr.Origin
            if o is not None and (
                abs(float(o.X)) + abs(float(o.Y)) + abs(float(o.Z))
            ) > 1e-6:
                return o
        except Exception:
            pass
    return None


def _layout_step_and_normal(rebar):
    acc = None
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return None, 0.0, 1

    normal = None
    try:
        normal = acc.Normal
    except Exception:
        normal = None
    if normal is None:
        return None, 0.0, 1
    try:
        if float(normal.GetLength()) < 1e-12:
            return None, 0.0, 1
        normal = normal.Normalize()
    except Exception:
        return None, 0.0, 1

    try:
        if not bool(acc.BarsOnNormalSide):
            normal = XYZ(-float(normal.X), -float(normal.Y), -float(normal.Z))
    except Exception:
        pass

    n = max(1, _cantidad_posiciones(rebar))
    alen = 0.0
    try:
        alen = float(acc.ArrayLength)
    except Exception:
        try:
            alen = float(acc.GetArrayLength())
        except Exception:
            alen = 0.0
    sp = 0.0
    try:
        sp = float(rebar.MaxSpacing)
    except Exception:
        sp = 0.0

    step = 0.0
    if n > 1 and alen > 1e-9:
        step = alen / float(n - 1)
    elif sp > 1e-9:
        step = sp
    return normal, step, n


def _bar_offset_from_layout(rebar, bar_index, normal, step):
    if normal is None or step is None or step < 1e-12:
        return None
    try:
        k = float(int(bar_index)) * float(step)
        return XYZ(
            float(normal.X) * k,
            float(normal.Y) * k,
            float(normal.Z) * k,
        )
    except Exception:
        return None


def _xyz_sub(a, b):
    if a is None or b is None:
        return None
    try:
        return XYZ(
            float(a.X) - float(b.X),
            float(a.Y) - float(b.Y),
            float(a.Z) - float(b.Z),
        )
    except Exception:
        return None


def _compute_move_delta(src_rebar, bar_index, new_single=None, positions_cache=None):
    """
    Vector de traslado para la pose de ``bar_index``.

    Con ``new_single=None`` (plan precomputado): mid_tgt − mid_src0 / layout /
    transform (misma lógica que al no poder leer mid del single nuevo).
    """
    bi = int(bar_index)
    method = u"none"
    delta = None

    mid_tgt = None
    if positions_cache is not None and bi in positions_cache:
        mid_tgt = positions_cache.get(bi)
    if mid_tgt is None:
        mid_tgt = _bar_world_midpoint(src_rebar, bi)

    mid_new = (
        _bar_world_midpoint(new_single, 0) if new_single is not None else None
    )

    if mid_tgt is not None and mid_new is not None:
        delta = _xyz_sub(mid_tgt, mid_new)
        method = u"mid_tgt - mid_single"
    elif mid_tgt is not None:
        mid0 = None
        if positions_cache is not None and 0 in positions_cache:
            mid0 = positions_cache.get(0)
        if mid0 is None:
            mid0 = _bar_world_midpoint(src_rebar, 0)
        if mid0 is not None:
            delta = _xyz_sub(mid_tgt, mid0)
            method = u"mid_tgt - mid_src0"
        else:
            nrm, step, _n = _layout_step_and_normal(src_rebar)
            off = _bar_offset_from_layout(src_rebar, bi, nrm, step)
            if off is not None:
                delta = off
                method = u"layout_offset_only"
    else:
        nrm, step, _n = _layout_step_and_normal(src_rebar)
        off_i = _bar_offset_from_layout(src_rebar, bi, nrm, step)
        off_0 = _bar_offset_from_layout(src_rebar, 0, nrm, step)
        if off_i is not None:
            if off_0 is not None:
                delta = _xyz_sub(off_i, off_0)
                if delta is None:
                    delta = off_i
            else:
                delta = off_i
            method = u"layout Normal*step*idx"

    if delta is None or (delta is not None and float(delta.GetLength()) < 1e-12):
        t0 = _get_bar_transform(src_rebar, 0)
        ti = _get_bar_transform(src_rebar, bi)
        if t0 is not None and ti is not None:
            try:
                d2 = ti.Origin - t0.Origin
                if float(d2.GetLength()) > 1e-9:
                    delta = d2
                    method = u"BarPositionTransform.Origin"
            except Exception:
                pass

    dlen = 0.0
    try:
        if delta is not None:
            dlen = float(delta.GetLength())
    except Exception:
        dlen = 0.0
    return delta, method, dlen


def _precompute_bar_positions(rebar, indices):
    cache = {}
    nrm, step, npos = _layout_step_and_normal(rebar)

    t0 = _get_bar_transform(rebar, 0)
    all_tr_zero = True
    if t0 is not None:
        try:
            for bi in indices[: min(5, len(indices))]:
                ti = _get_bar_transform(rebar, bi)
                if ti is None:
                    continue
                if float((ti.Origin - t0.Origin).GetLength()) > 1e-6:
                    all_tr_zero = False
                    break
        except Exception:
            pass

    distinct = 0
    prev = None
    for bi in sorted(set([0] + list(indices))):
        mid = _bar_world_midpoint(rebar, bi)
        if mid is None and nrm is not None and step > 1e-12:
            off = _bar_offset_from_layout(rebar, bi, nrm, step)
            mid0 = _bar_world_midpoint(rebar, 0)
            if mid0 is not None and off is not None:
                mid = XYZ(
                    float(mid0.X) + float(off.X),
                    float(mid0.Y) + float(off.Y),
                    float(mid0.Z) + float(off.Z),
                )
        if mid is not None:
            cache[bi] = mid
            if prev is not None:
                try:
                    if float((mid - prev).GetLength()) > 1e-4:
                        distinct += 1
                except Exception:
                    pass
            prev = mid

    n_filled = sum(1 for v in cache.values() if v is not None)
    if n_filled >= 2 and distinct == 0 and nrm is not None and step > 1e-12:
        anchor = None
        for bi in sorted(cache.keys()):
            if cache.get(bi) is not None:
                anchor = cache[bi]
                break
        if anchor is None:
            anchor = XYZ(0, 0, 0)
        for bi in sorted(set([0] + list(indices))):
            off = _bar_offset_from_layout(rebar, bi, nrm, step)
            if off is None:
                continue
            cache[bi] = XYZ(
                float(anchor.X) + float(off.X),
                float(anchor.Y) + float(off.Y),
                float(anchor.Z) + float(off.Z),
            )
    return cache


def _is_in_group(rebar):
    try:
        gid = rebar.GroupId
        if gid is not None and gid != ElementId.InvalidElementId:
            return True
    except Exception:
        pass
    return False


def _es_malla(rebar):
    if rebar is None:
        return False
    p = None
    try:
        p = rebar.LookupParameter(_ARMADURA_MALLA)
    except Exception:
        p = None
    if p is None:
        try:
            for sp in rebar.Parameters:
                try:
                    name = sp.Definition.Name if sp.Definition is not None else u""
                except Exception:
                    name = u""
                if _as_unicode(name).strip().lower() == _ARMADURA_MALLA.lower():
                    p = sp
                    break
        except Exception:
            p = None
    if p is None:
        return False
    try:
        if int(p.AsInteger()) == 1:
            return True
    except Exception:
        pass
    try:
        vs = _as_unicode(p.AsValueString()).strip().upper()
        if vs in (u"YES", u"Y", u"SI", u"SÍ", u"1", u"TRUE", u"VERDADERO"):
            return True
    except Exception:
        pass
    return False


def _es_rebar_separable(rebar):
    if rebar is None or not isinstance(rebar, Rebar):
        return False, u"No es un elemento Rebar."
    if _is_in_group(rebar):
        return False, u"La barra pertenece a un Group; no se puede separar."
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False, u"Solo aplica a barras shape-driven (no free-form)."
    if _es_layout_single(rebar):
        return False, u"El layout ya es Single; no hay conjunto que separar."
    idxs = _indices_incluidos(rebar)
    if len(idxs) < 2:
        return (
            False,
            u"El conjunto debe tener al menos dos barras incluidas "
            u"(layout distinto de Single).",
        )
    return True, u""


# ---------------------------------------------------------------------------
# Core
# ---------------------------------------------------------------------------

def _precompute_move_plan(src_rebar, idxs, positions_cache):
    plan = {}
    for bi in idxs:
        delta, method, dlen = _compute_move_delta(
            src_rebar, bi, new_single=None, positions_cache=positions_cache
        )
        plan[int(bi)] = {
            u"delta": delta,
            u"method": method,
            u"dlen": dlen,
        }
    return plan


def _spacing_mm_from_rebar(rebar):
    """Separación entre posiciones del set en mm (MaxSpacing o paso del array)."""
    if rebar is None:
        return None
    try:
        sp = float(rebar.MaxSpacing)
        if sp > 1e-9:
            return float(sp) * 304.8
    except Exception:
        pass
    try:
        normal, step, n = _layout_step_and_normal(rebar)
        if step is not None and float(step) > 1e-9 and n is not None and int(n) > 1:
            return float(step) * 304.8
    except Exception:
        pass
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, u"GetSpacing"):
            sp2 = float(acc.GetSpacing())
            if sp2 > 1e-9:
                return float(sp2) * 304.8
    except Exception:
        pass
    return None


def _texto_separacion_barras(spacing_mm):
    try:
        n = int(round(float(spacing_mm)))
    except Exception:
        return None
    if n <= 0:
        return None
    return u"@{0}".format(n)


def _malla_orient_rebar(rebar):
    """``vertical`` | ``horizontal`` | None."""
    try:
        from armado_muros_malla_rebar_tags import _orient_rebar_malla_desde_param

        o = _orient_rebar_malla_desde_param(rebar)
        if o in (u"vertical", u"horizontal"):
            return o
    except Exception:
        pass
    try:
        from armado_muros_rebar_params import get_armadura_malla_orientacion

        o = get_armadura_malla_orientacion(rebar)
        if o in (u"vertical", u"horizontal"):
            return o
    except Exception:
        pass
    return None


def _paper_mm_to_model_ft(paper_mm, view):
    """mm sobre papel → ft modelo según ``View.Scale`` (1:50 → Scale=50)."""
    try:
        scale = int(view.Scale)
    except Exception:
        scale = _SPACING_REF_SCALE
    if scale <= 0:
        scale = _SPACING_REF_SCALE
    return float(paper_mm) * float(scale) / 304.8


def _normalize_xyz(v):
    if v is None:
        return None
    try:
        L = float(v.GetLength())
        if L < 1e-12:
            return None
        return XYZ(float(v.X) / L, float(v.Y) / L, float(v.Z) / L)
    except Exception:
        return None


def _tag_text_raw(tag):
    if tag is None:
        return u""
    try:
        t = tag.TagText
        if t:
            return u"{}".format(t)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import BuiltInParameter

        p = tag.get_Parameter(BuiltInParameter.TAG_LABEL)
        if p is not None and p.HasValue:
            s = p.AsString()
            if s:
                return u"{}".format(s)
    except Exception:
        pass
    return u""


def _tag_text_shows_spacing(text):
    """True si el texto de etiqueta ya incluye sufijo @numérico (p. ej. @200)."""
    if not text:
        return False
    try:
        s = u"{}".format(text)
    except Exception:
        return False
    # Compactar espacios; buscar @digits (también tras salto de línea).
    i = 0
    n = len(s)
    while i < n:
        if s[i] == u"@" or s[i] == u"＠":
            j = i + 1
            while j < n and s[j] in (u" ", u"\t", u"\xa0"):
                j += 1
            if j < n and s[j].isdigit():
                return True
        i += 1
    return False


def _independent_tags_for_rebar(document, rebar_id, view=None):
    """IndependentTags que etiquetan ``rebar_id`` (opcionalmente en una vista)."""
    rid = _element_id_int(rebar_id)
    if rid is None or document is None:
        return []
    rebar_set = {rid}
    out = []
    try:
        from Autodesk.Revit.DB import FilteredElementCollector, IndependentTag

        invalid = ElementId.InvalidElementId
        coll = (
            FilteredElementCollector(document)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return out

    try:
        from dividir_rebar_punto_tags import tag_rebar_int_if_match
    except Exception:
        tag_rebar_int_if_match = None

    for tag in coll:
        if view is not None:
            try:
                if tag.OwnerViewId != view.Id:
                    continue
            except Exception:
                pass
        matched = None
        if tag_rebar_int_if_match is not None:
            try:
                matched = tag_rebar_int_if_match(tag, rebar_set, invalid)
            except Exception:
                matched = None
        if matched is None:
            # Respaldo: GetTaggedLocalElementIds / GetTaggedElementIds
            try:
                ids = tag.GetTaggedLocalElementIds()
                for eid in ids or []:
                    if _element_id_int(eid) == rid:
                        matched = rid
                        break
            except Exception:
                pass
            if matched is None:
                try:
                    ids = tag.GetTaggedElementIds()
                    # LinkElementId list in newer APIs
                    for le in ids or []:
                        try:
                            eid = le.HostElementId
                        except Exception:
                            eid = le
                        if _element_id_int(eid) == rid:
                            matched = rid
                            break
                except Exception:
                    pass
        if matched is None:
            continue
        out.append(tag)
    return out


def _rebar_tags_already_show_spacing(document, rebar_id, view):
    for tag in _independent_tags_for_rebar(document, rebar_id, view=view):
        if _tag_text_shows_spacing(_tag_text_raw(tag)):
            return True
    return False


def _paper_text_width_mm(s):
    """Ancho estimado de texto (mm papel) a altura 2.5 mm Arial."""
    if not s:
        return 0.0
    n = 0
    try:
        for ch in u"{}".format(s):
            if ch in (u"\r", u"\n", u"\t"):
                continue
            n += 1
    except Exception:
        try:
            n = len(str(s))
        except Exception:
            n = 0
    return float(n) * float(_SPACING_CHAR_WIDTH_FACTOR) * float(
        _SPACING_TEXT_HEIGHT_PAPER_MM
    )


def _tag_lines(tag_text):
    if not tag_text:
        return []
    out = []
    try:
        raw = u"{}".format(tag_text).replace(u"\r\n", u"\n").replace(u"\r", u"\n")
    except Exception:
        return []
    for ln in raw.split(u"\n"):
        t = ln.strip()
        if t:
            out.append(t)
    return out


def _diameter_line_from_tag_text(tag_text, malla_orient):
    """
    Extrae la línea de diámetro (sin @sep) relevante a V./H.
    Quita prefijos D.M. en la misma línea.
    """
    lines = _tag_lines(tag_text)
    if not lines:
        return None

    def _is_v(ln):
        u = ln.upper().replace(u" ", u"")
        return u.startswith(u"V.") or u.startswith(u"V=") or u"V.=" in u or u"V=" in u

    def _is_h(ln):
        u = ln.upper().replace(u" ", u"")
        return u.startswith(u"H.") or u.startswith(u"H=") or u"H.=" in u or u"H=" in u

    def _is_mha(ln):
        u = ln.upper().replace(u" ", u"")
        return u.startswith(u"M.H.A") or u.startswith(u"MHA")

    target = None
    if malla_orient == u"vertical":
        for ln in lines:
            if _is_v(ln):
                target = ln
                break
    elif malla_orient == u"horizontal":
        for ln in lines:
            if _is_h(ln):
                target = ln
                break
    if target is None:
        for ln in reversed(lines):
            if _is_mha(ln):
                continue
            if ln.upper().replace(u" ", u"") in (u"D.M.", u"DM.", u"D.M"):
                continue
            target = ln
            break
    if target is None:
        target = lines[-1]

    # Quitar D.M. al inicio de la línea
    t = target
    for pref in (u"D.M.", u"D.M", u"DM.", u"DM"):
        if t.upper().startswith(pref.upper()):
            t = t[len(pref) :].lstrip()
            break

    # Solo tramo de diámetro (hasta @)
    at = t.find(u"@")
    if at < 0:
        at = t.find(u"＠")
    if at >= 0:
        t = t[:at].rstrip()
    return t if t else None


def _dx_paper_mm_after_diameter(tag_text, malla_orient):
    """
    Offset horizontal (mm papel) desde la cabeza hasta el inicio de @ (respaldo).
    Usa TagText; se multiplica por ``_SPACING_TEXT_DX_TIGHTEN`` para cerrar hueco.
    """
    gap = float(_SPACING_TEXT_DX_GAP_PAPER_MM)
    diam = _diameter_line_from_tag_text(tag_text, malla_orient)
    lines = _tag_lines(tag_text)

    if diam is None:
        if malla_orient == u"horizontal":
            diam = u"H.=\u00f810"
        else:
            diam = u"V.=\u00f810"

    w_diam = _paper_text_width_mm(diam)
    if w_diam < 0.5:
        return float(_SPACING_TEXT_DX_PAPER_MM)

    content_lines = []
    for ln in lines:
        uu = ln.upper().replace(u" ", u"")
        if uu.startswith(u"M.H.A") or uu in (u"D.M.", u"DM.", u"D.M"):
            continue
        content_lines.append(ln)

    tight = float(_SPACING_TEXT_DX_TIGHTEN)
    if len(content_lines) <= 1 and len(lines) <= 1:
        return tight * (0.5 * w_diam) + gap

    w_max = 0.0
    for ln in lines:
        w_max = max(w_max, _paper_text_width_mm(ln))
    if w_max < 0.5:
        w_max = w_diam

    dx = -0.5 * w_max + w_diam
    dx = tight * dx + gap
    if dx < 0.3:
        dx = tight * (0.5 * w_diam) + gap
    if dx > w_diam + gap:
        dx = tight * (0.5 * w_diam) + gap
    return dx


def _line_dy_paper_mm(malla_orient):
    if malla_orient == u"vertical":
        return float(_SPACING_LINE_V_DY_PAPER_MM)
    if malla_orient == u"horizontal":
        return float(_SPACING_LINE_H_DY_PAPER_MM)
    return float(_SPACING_LINE_DEFAULT_DY_PAPER_MM)


def _view_scale_den(view):
    try:
        s = int(view.Scale)
        if s > 0:
            return s
    except Exception:
        pass
    return int(_SPACING_REF_SCALE)


def _model_ft_to_paper_mm(ft, view):
    scale = _view_scale_den(view)
    return float(ft) * 304.8 / float(scale)


def _tag_bb_corners(bb):
    if bb is None or bb.Min is None or bb.Max is None:
        return []
    try:
        x0, y0, z0 = float(bb.Min.X), float(bb.Min.Y), float(bb.Min.Z)
        x1, y1, z1 = float(bb.Max.X), float(bb.Max.Y), float(bb.Max.Z)
    except Exception:
        return []
    out = []
    for x in (x0, x1):
        for y in (y0, y1):
            for z in (z0, z1):
                out.append(XYZ(x, y, z))
    return out


def _tag_bbox_extents_from_head(tag, view, head, along, across):
    """
    Proyecciones del BB de la etiqueta respecto a ``head``:
    (max_along ft, mid_across ft) o (None, None).
    """
    if tag is None or view is None or head is None:
        return None, None
    if along is None or across is None:
        return None, None
    bb = None
    try:
        bb = tag.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = tag.get_BoundingBox(None)
        except Exception:
            bb = None
    corners = _tag_bb_corners(bb)
    if not corners:
        return None, None

    dots_a = []
    dots_c = []
    try:
        hx, hy, hz = float(head.X), float(head.Y), float(head.Z)
        ax, ay, az = float(along.X), float(along.Y), float(along.Z)
        cx, cy, cz = float(across.X), float(across.Y), float(across.Z)
    except Exception:
        return None, None

    for p in corners:
        try:
            dx = float(p.X) - hx
            dy = float(p.Y) - hy
            dz = float(p.Z) - hz
            dots_a.append(dx * ax + dy * ay + dz * az)
            dots_c.append(dx * cx + dy * cy + dz * cz)
        except Exception:
            continue
    if not dots_a:
        return None, None

    max_along = max(dots_a)
    mid_across = 0.5 * (min(dots_c) + max(dots_c))

    # Filtrar BB absurdos (p. ej. con leader o catálogo completo).
    max_paper = _model_ft_to_paper_mm(max_along, view)
    if max_paper > float(_SPACING_BBOX_MAX_PAPER_MM):
        return None, None
    if max_paper < float(_SPACING_BBOX_MIN_PAPER_MM) and abs(mid_across) < 1e-9:
        return None, None
    return max_along, mid_across


def _reading_axes(view, tag_orient=None):
    right = _normalize_xyz(getattr(view, u"RightDirection", None))
    up = _normalize_xyz(getattr(view, u"UpDirection", None))
    if right is None:
        right = XYZ(1.0, 0.0, 0.0)
    if up is None:
        up = XYZ(0.0, 0.0, 1.0)
    is_vert = False
    try:
        from Autodesk.Revit.DB import TagOrientation

        if tag_orient is not None and tag_orient == TagOrientation.Vertical:
            is_vert = True
    except Exception:
        is_vert = False
    # Malla fuerza TagOrientation.Horizontal; along = Right de la vista.
    along = up if is_vert else right
    across = right if is_vert else up
    return along, across


def _origin_texto_separacion(
    view,
    head,
    tag_orient=None,
    malla_orient=None,
    dx_paper_mm=None,
    dy_paper_mm=None,
    dx_ft=None,
    dy_ft=None,
):
    """
    Origen del TextNote ``@NNN`` justo tras el tramo de diámetro.

    Preferir ``dx_ft``/``dy_ft`` (desde BB); si no, offsets en mm papel.
    """
    if view is None or head is None:
        return head

    along, across = _reading_axes(view, tag_orient)

    if dx_ft is None:
        if dx_paper_mm is None:
            dx_paper_mm = float(_SPACING_TEXT_DX_PAPER_MM)
        dx_ft = _paper_mm_to_model_ft(float(dx_paper_mm), view)
    if dy_ft is None:
        if dy_paper_mm is None:
            dy_paper_mm = _line_dy_paper_mm(malla_orient)
        dy_ft = _paper_mm_to_model_ft(float(dy_paper_mm), view)

    try:
        origin = XYZ(
            float(head.X) + float(along.X) * float(dx_ft) + float(across.X) * float(dy_ft),
            float(head.Y) + float(along.Y) * float(dx_ft) + float(across.Y) * float(dy_ft),
            float(head.Z) + float(along.Z) * float(dx_ft) + float(across.Z) * float(dy_ft),
        )
    except Exception:
        return head
    try:
        from armado_muros_etiqueta_malla import _proyectar_punto_plano_vista

        origin = _proyectar_punto_plano_vista(origin, view) or origin
    except Exception:
        pass
    return origin


def _apply_spacing_nudge(view, dx_ft, dy_ft):
    """Desplaza el origen: izquierda (menos dx) y arriba (más dy / Up)."""
    left = _paper_mm_to_model_ft(float(_SPACING_NUDGE_LEFT_PAPER_MM), view)
    up = _paper_mm_to_model_ft(float(_SPACING_NUDGE_UP_PAPER_MM), view)
    try:
        return float(dx_ft) - float(left), float(dy_ft) + float(up)
    except Exception:
        return dx_ft, dy_ft


def _resolve_spacing_text_origin(
    document, view, rebar_id, head, malla_orient, tag_orient=None
):
    """
    Origen final del @: 1) BB de IndependentTag de la barra, 2) TagText estimado.
    Luego nudge izquierda/arriba calibrado en 1:50.
    """
    along, across = _reading_axes(view, tag_orient)
    gap_ft = _paper_mm_to_model_ft(float(_SPACING_TEXT_DX_GAP_PAPER_MM), view)

    tags = _independent_tags_for_rebar(document, rebar_id, view=view)
    pick = None
    tag_text = u""
    for tag in tags:
        if pick is None:
            pick = tag
        if not tag_text:
            tag_text = _tag_text_raw(tag)
        try:
            th = tag.TagHeadPosition
            if th is not None:
                head = th
                pick = tag
        except Exception:
            pass

    # 1) BoundingBox real de la etiqueta (borde derecho del diametro)
    if pick is not None and head is not None:
        max_along, mid_across = _tag_bbox_extents_from_head(
            pick, view, head, along, across
        )
        if max_along is not None:
            extra = _paper_mm_to_model_ft(
                float(_SPACING_BBOX_DX_EXTRA_PAPER_MM), view
            )
            dx_ft = float(max_along) + float(gap_ft) + float(extra)
            dy_ft = float(mid_across) if mid_across is not None else 0.0
            est_paper = _dx_paper_mm_after_diameter(tag_text, malla_orient)
            est_ft = _paper_mm_to_model_ft(est_paper, view)
            if est_ft > 1e-9 and dx_ft > 1.35 * est_ft:
                dx_ft = 0.5 * (dx_ft + est_ft)
            if est_ft > 1e-9 and dx_ft < 0.55 * est_ft:
                dx_ft = 0.5 * (dx_ft + est_ft)
            dx_ft, dy_ft = _apply_spacing_nudge(view, dx_ft, dy_ft)
            return _origin_texto_separacion(
                view,
                head,
                tag_orient=tag_orient,
                malla_orient=malla_orient,
                dx_ft=dx_ft,
                dy_ft=dy_ft,
            ), head, tag_text

    # 2) Estimación por TagText + dy fijo por orientación
    dx_paper = _dx_paper_mm_after_diameter(tag_text, malla_orient)
    dy_paper = _line_dy_paper_mm(malla_orient)
    dx_ft = _paper_mm_to_model_ft(dx_paper, view)
    dy_ft = _paper_mm_to_model_ft(dy_paper, view)
    dx_ft, dy_ft = _apply_spacing_nudge(view, dx_ft, dy_ft)
    return (
        _origin_texto_separacion(
            view,
            head,
            tag_orient=tag_orient,
            malla_orient=malla_orient,
            dx_ft=dx_ft,
            dy_ft=dy_ft,
        ),
        head,
        tag_text,
    )


def _find_spacing_text_type(document):
    try:
        from armado_muros_etiqueta_malla import (
            _find_text_note_type_named,
            _text_note_type_obligatorio,
        )

        t = _text_note_type_obligatorio(document)
        if t is not None:
            return t
        return _find_text_note_type_named(document, _SPACING_TEXT_TYPE_NAME)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import FilteredElementCollector, TextNoteType

        want = _SPACING_TEXT_TYPE_NAME.lower().replace(u" ", u"")
        for tnt in FilteredElementCollector(document).OfClass(TextNoteType):
            try:
                nm = (tnt.Name or u"").lower().replace(u" ", u"")
            except Exception:
                continue
            if nm == want or (
                nm.startswith(u"2.5")
                and u"arial" in nm
                and u"arrow" in nm
                and u"filled" in nm
                and u"15" in nm
            ):
                return tnt
    except Exception:
        pass
    return None


def _crear_text_note_separacion(document, view, origin, texto, text_type):
    """Crea TextNote; Left+Middle para que @NNN quede a la derecha del diámetro."""
    if document is None or view is None or origin is None or not texto or text_type is None:
        return None
    try:
        from Autodesk.Revit.DB import (
            HorizontalTextAlignment,
            TextNote,
            TextNoteOptions,
            VerticalTextAlignment,
        )
    except Exception:
        return None

    def _opts(con_alin):
        try:
            o = TextNoteOptions(text_type.Id)
        except Exception:
            o = TextNoteOptions()
            try:
                o.TypeId = text_type.Id
            except Exception:
                pass
        if con_alin:
            try:
                o.HorizontalAlignment = HorizontalTextAlignment.Left
            except Exception:
                pass
            try:
                o.VerticalAlignment = VerticalTextAlignment.Middle
            except Exception:
                pass
        return o

    for con_alin in (True, False):
        try:
            tn = TextNote.Create(document, view.Id, origin, texto, _opts(con_alin))
            return tn
        except Exception:
            continue
    return None


def _crear_textos_separacion(
    document, plan, rebar_id, spacing_mm, tag_infos_fallback=None
):
    """
    TextNote ``@NNN`` solo si las etiquetas de **esta** barra no muestran ya @sep.
    Posición por orientación V/H de malla (evita solapar la línea H cuando se
    singlea V. y H. aún reporta @200).
    """
    texto = _texto_separacion_barras(spacing_mm)
    if not texto:
        return 0
    text_type = _find_spacing_text_type(document)
    if text_type is None:
        return 0

    rb = None
    try:
        rb = document.GetElement(rebar_id) if rebar_id is not None else None
    except Exception:
        rb = None

    malla_orient = plan.get(u"malla_orient")
    if malla_orient is None and rb is not None:
        malla_orient = _malla_orient_rebar(rb)

    tag_infos = list(tag_infos_fallback or plan.get(u"tag_infos") or [])
    if not tag_infos:
        avid = plan.get(u"active_view_id")
        if avid is not None:
            tag_infos = [{u"view_id": avid, u"head": None, u"orient": None}]

    n = 0
    seen_views = set()

    for info in tag_infos:
        vid = info.get(u"view_id")
        try:
            vkey = int(vid.IntegerValue)
        except Exception:
            vkey = None
        if vkey is not None and vkey in seen_views:
            continue
        if vkey is not None:
            seen_views.add(vkey)
        view = document.GetElement(vid) if vid is not None else None
        if view is None:
            continue

        # No duplicar si la etiqueta de esta barra ya escribe @NNN.
        if _rebar_tags_already_show_spacing(document, rebar_id, view):
            continue

        head = None
        if rb is not None:
            head, _wall = _tag_head_centroide_host(document, rb, view)
        if head is None:
            head = info.get(u"head")
        if head is None and rb is not None:
            try:
                head = _bar_world_midpoint(rb, 0)
            except Exception:
                head = None
            if head is not None:
                try:
                    from armado_muros_etiqueta_malla import _proyectar_punto_plano_vista

                    head = _proyectar_punto_plano_vista(head, view) or head
                except Exception:
                    pass
        if head is None:
            continue

        origin, head, _tag_text = _resolve_spacing_text_origin(
            document,
            view,
            rebar_id,
            head,
            malla_orient,
            tag_orient=info.get(u"orient"),
        )
        if origin is None:
            continue
        tn = _crear_text_note_separacion(document, view, origin, texto, text_type)
        if tn is not None:
            n += 1

    return n


def _build_explode_plan(document, rebar):
    plan = {
        u"ok": False,
        u"message": u"",
        u"idxs": [],
        u"positions_cache": {},
        u"moves": {},
        u"is_malla": False,
        u"tag_infos": [],
        u"spacing_mm": None,
        u"malla_orient": None,
        u"old_id": None,
        u"old_id_int": None,
        u"src_rebar": rebar,
    }

    ok_pre, msg_pre = _es_rebar_separable(rebar)
    if not ok_pre:
        plan[u"message"] = msg_pre
        return plan

    idxs = _indices_incluidos(rebar)
    plan[u"idxs"] = idxs
    positions_cache = _precompute_bar_positions(rebar, idxs)
    plan[u"positions_cache"] = positions_cache
    plan[u"moves"] = _precompute_move_plan(rebar, idxs, positions_cache)
    plan[u"spacing_mm"] = _spacing_mm_from_rebar(rebar)
    plan[u"malla_orient"] = _malla_orient_rebar(rebar)

    is_malla = _es_malla(rebar)
    plan[u"is_malla"] = is_malla

    tag_infos = []
    try:
        from dividir_rebar_punto_tags import capture_rebar_tag_infos

        tag_infos = capture_rebar_tag_infos(document, rebar.Id) or []
    except Exception:
        tag_infos = []
    plan[u"tag_infos"] = tag_infos
    plan[u"old_id"] = rebar.Id
    plan[u"old_id_int"] = _element_id_int(rebar.Id)
    plan[u"ok"] = True
    return plan


def _crear_single_en_posicion(
    document, rebar, bar_index, move_item, copy_armadura_fn
):
    bi = int(bar_index)
    info = {
        u"bar_index": bi,
        u"new_id": None,
        u"delta": None,
        u"n_armadura": 0,
        u"moved": False,
        u"move_method": u"",
    }

    try:
        new_ids = ElementTransformUtils.CopyElement(document, rebar.Id, _XYZ_ZERO)
    except Exception as ex:
        return None, u"CopyElement falló (índice {}): {}".format(
            bi, _exception_text(ex)
        ), info

    if new_ids is None or len(new_ids) < 1:
        return None, u"CopyElement no devolvió elementos (índice {}).".format(bi), info

    rb = document.GetElement(new_ids[0])
    if rb is None or not isinstance(rb, Rebar):
        return None, u"La copia no es Rebar (índice {}).".format(bi), info

    info[u"new_id"] = _element_id_int(rb.Id)

    try:
        acc = rb.GetShapeDrivenAccessor()
        if acc is None:
            return None, u"Copia sin ShapeDrivenAccessor (índice {}).".format(bi), info
        acc.SetLayoutAsSingle()
    except Exception as ex:
        try:
            document.Delete(rb.Id)
        except Exception:
            pass
        return None, u"SetLayoutAsSingle falló (índice {}): {}".format(
            bi, _exception_text(ex)
        ), info

    delta = None
    method = u"precomputed"
    dlen = 0.0
    if move_item is not None:
        delta = move_item.get(u"delta")
        method = move_item.get(u"method") or method
        try:
            dlen = float(move_item.get(u"dlen") or 0.0)
        except Exception:
            dlen = 0.0
        if delta is not None and dlen < 1e-12:
            try:
                dlen = float(delta.GetLength())
            except Exception:
                dlen = 0.0
    info[u"delta"] = dlen
    info[u"move_method"] = method

    if delta is not None and dlen > 1e-9:
        try:
            ElementTransformUtils.MoveElement(document, rb.Id, delta)
            info[u"moved"] = True
        except Exception as ex:
            try:
                document.Delete(rb.Id)
            except Exception:
                pass
            return None, u"No se pudo trasladar la barra {}: {}".format(
                bi, _exception_text(ex)
            ), info

    if copy_armadura_fn is not None:
        try:
            info[u"n_armadura"] = int(copy_armadura_fn(rebar, rb) or 0)
        except Exception:
            pass

    return rb, u"", info


def _apply_explode_plan_in_open_txn(document, plan, progress=None):
    rebar = plan.get(u"src_rebar")
    idxs = plan.get(u"idxs") or []
    moves = plan.get(u"moves") or {}
    old_id = plan.get(u"old_id")
    old_id_int = plan.get(u"old_id_int")
    is_malla = bool(plan.get(u"is_malla"))
    tag_infos = plan.get(u"tag_infos") or []

    copy_armadura_fn = None
    try:
        from dividir_rebar_punto_core import copy_armadura_instance_parameters

        copy_armadura_fn = copy_armadura_instance_parameters
    except Exception:
        pass

    new_rebars = []
    created_info = []
    n_idx = len(idxs)

    for k, bi in enumerate(idxs):
        if progress is not None:
            progress.step(
                u"Id {0}: barra {1}/{2} (idx {3})".format(
                    old_id_int if old_id_int is not None else u"?",
                    k + 1,
                    n_idx,
                    bi,
                )
            )
        rb, err, info = _crear_single_en_posicion(
            document, rebar, bi, moves.get(int(bi)), copy_armadura_fn
        )
        created_info.append(info)
        if rb is None:
            raise RuntimeError(
                err or u"Fallo al crear single en índice {}.".format(bi)
            )
        new_rebars.append(rb)

    if not new_rebars:
        raise RuntimeError(u"No se generó ninguna barra single.")

    if progress is not None:
        progress.step(
            u"Id {0}: finalize (regen / delete / tags)".format(
                old_id_int if old_id_int is not None else u"?"
            )
        )

    try:
        document.Regenerate()
    except Exception:
        pass

    try:
        document.Delete(old_id)
    except Exception as ex:
        raise RuntimeError(
            u"Singles creados pero no se pudo eliminar el set original: {}".format(
                _exception_text(ex)
            )
        )

    new_ids = []
    for rb in new_rebars:
        try:
            if rb.Id is not None:
                new_ids.append(rb.Id)
        except Exception:
            pass

    n_tags = 0
    if is_malla and tag_infos and new_ids:
        n_tags = _etiquetar_una_barra_malla(
            document, tag_infos, new_ids[0], open_transaction=True
        )
        try:
            document.Regenerate()
        except Exception:
            pass

    n_spacing_text = 0
    spacing_mm = plan.get(u"spacing_mm")
    if new_ids and spacing_mm is not None:
        try:
            n_spacing_text = _crear_textos_separacion(
                document,
                plan,
                new_ids[0],
                spacing_mm,
                tag_infos_fallback=tag_infos,
            )
        except Exception:
            n_spacing_text = 0

    return {
        u"new_ids": new_ids,
        u"n_singles": len(new_ids),
        u"n_tags": n_tags,
        u"n_spacing_text": n_spacing_text,
        u"created_info": created_info,
    }


def separar_rebar_set_a_singles(
    document, rebar, open_transaction=False, progress=None, active_view=None
):
    """
    Separa ``rebar`` en N singles y elimina el set original.

    1) Plan en memoria (sin Transaction)
    2) Una Transaction: creates + 1×Regenerate + Delete + tags malla + TextNote @sep
    """
    result = {
        u"ok": False,
        u"message": u"",
        u"new_ids": [],
        u"n_singles": 0,
        u"is_malla": False,
        u"n_tags": 0,
        u"n_spacing_text": 0,
        u"tag_infos": [],
    }

    rid = _element_id_int(rebar.Id) if rebar is not None else None
    if progress is not None:
        progress.step(
            u"Planificar set id={0}".format(rid if rid is not None else u"?")
        )

    plan = _build_explode_plan(document, rebar)
    if active_view is not None:
        try:
            plan[u"active_view_id"] = active_view.Id
        except Exception:
            plan[u"active_view_id"] = None
    else:
        plan[u"active_view_id"] = None
    result[u"is_malla"] = bool(plan.get(u"is_malla"))
    result[u"tag_infos"] = plan.get(u"tag_infos") or []
    if not plan.get(u"ok"):
        result[u"message"] = plan.get(u"message") or u"Plan no válido."
        if progress is not None:
            try:
                n_idxs = max(1, len(_indices_incluidos(rebar)) if rebar else 1)
            except Exception:
                n_idxs = 1
            for _ in range(n_idxs + 1):
                progress.step(u"Omitido (set no válido)")
        return result

    def _run():
        return _apply_explode_plan_in_open_txn(document, plan, progress=progress)

    try:
        if open_transaction:
            applied = _run()
        else:
            with _RevitTransaction(document, _TRANSACTION_NAME):
                applied = _run()
    except Exception as ex:
        result[u"message"] = _exception_text(ex)
        return result

    new_ids = applied.get(u"new_ids") or []
    valid_ids = []
    for eid in new_ids:
        try:
            if document.GetElement(eid) is not None:
                valid_ids.append(eid)
        except Exception:
            pass
    result[u"new_ids"] = valid_ids
    result[u"n_singles"] = len(valid_ids)
    result[u"n_tags"] = int(applied.get(u"n_tags") or 0)
    result[u"n_spacing_text"] = int(applied.get(u"n_spacing_text") or 0)

    partes = [u"Se generaron {} barra(s) Single.".format(len(valid_ids))]
    if plan.get(u"is_malla"):
        partes.append(
            u"Malla (Armadura_Malla=Yes): {} etiqueta(s) reaplicada(s) "
            u"en una sola barra.".format(result[u"n_tags"])
        )
    if result[u"n_spacing_text"]:
        partes.append(
            u"TextNote separación: {}.".format(result[u"n_spacing_text"])
        )
    result[u"ok"] = True
    result[u"message"] = u"\n".join(partes)
    return result


# ---------------------------------------------------------------------------
# Tags (malla)
# ---------------------------------------------------------------------------

def _host_wall_from_rebar(document, rebar):
    if document is None or rebar is None:
        return None
    try:
        from Autodesk.Revit.DB import Wall
    except Exception:
        Wall = None
    host = None
    try:
        hid = rebar.GetHostId()
        if hid is not None and hid != ElementId.InvalidElementId:
            host = document.GetElement(hid)
    except Exception:
        host = None
    if host is None:
        return None
    if Wall is not None:
        try:
            if isinstance(host, Wall):
                return host
        except Exception:
            pass
    try:
        cat = host.Category
        if cat is not None and int(cat.Id.IntegerValue) == -2000011:
            return host
    except Exception:
        pass
    try:
        if hasattr(host, u"WallType") or getattr(host, u"Width", None) is not None:
            return host
    except Exception:
        pass
    return None


def _tag_head_centroide_host(document, rebar, view):
    wall = _host_wall_from_rebar(document, rebar)
    if wall is None:
        return None, None
    try:
        from armado_muros_malla_rebar_tags import _head_pos_centroide_muro

        head = _head_pos_centroide_muro(wall, view)
        if head is not None:
            return head, wall
    except Exception:
        pass
    try:
        from armado_muros_etiqueta_malla import (
            _proyectar_punto_plano_vista,
            centroide_geometria_muro,
        )

        c = centroide_geometria_muro(wall, view)
        if c is not None:
            head = _proyectar_punto_plano_vista(c, view)
            if head is None:
                head = c
            return head, wall
    except Exception:
        pass
    try:
        bb = wall.get_BoundingBox(view)
        if bb is None:
            bb = wall.get_BoundingBox(None)
        if bb is not None and bb.Min is not None and bb.Max is not None:
            head = XYZ(
                0.5 * (float(bb.Min.X) + float(bb.Max.X)),
                0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
                0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
            )
            return head, wall
    except Exception:
        pass
    return None, wall


def _etiquetar_una_barra_malla(
    document, tag_infos, rebar_id, open_transaction=False
):
    """Etiqueta una barra; cabeza en centroide del muro host."""
    if document is None or not tag_infos or rebar_id is None:
        return 0
    rb = document.GetElement(rebar_id)
    if rb is None or not isinstance(rb, Rebar):
        return 0

    try:
        from dividir_rebar_punto_tags import (
            create_rebar_independent_tag,
            resolve_tag_type_id_for_rebar,
        )
        from Autodesk.Revit.DB import TagOrientation
    except Exception:
        return 0

    def _head_fallback(rebar, info, view):
        try:
            from dividir_rebar_punto_tags import _head_for_divided_rebar

            return _head_for_divided_rebar(rebar, info, view, 0)
        except Exception:
            return info.get(u"head") if info else None

    def _do_create():
        creadas = 0
        cache = {}
        seen_views = set()
        for info in tag_infos:
            vid = info.get(u"view_id")
            try:
                vkey = int(vid.IntegerValue)
            except Exception:
                vkey = None
            if vkey is not None and vkey in seen_views:
                continue
            if vkey is not None:
                seen_views.add(vkey)
            view = document.GetElement(vid) if vid is not None else None
            if view is None:
                continue
            type_id = resolve_tag_type_id_for_rebar(document, rb, info, cache)
            if type_id is None:
                continue
            head, _wall = _tag_head_centroide_host(document, rb, view)
            if head is None:
                head = _head_fallback(rb, info, view)
            tag = create_rebar_independent_tag(
                document,
                view,
                rb,
                type_id,
                head,
                info.get(u"orient", TagOrientation.Horizontal),
                info.get(u"leader", True),
                rotation=info.get(u"rotation"),
                bar_index=0,
            )
            if tag is not None:
                if head is not None:
                    try:
                        tag.TagHeadPosition = head
                    except Exception:
                        pass
                creadas += 1
        return creadas

    if open_transaction:
        try:
            return _do_create()
        except Exception:
            return 0
    try:
        with _RevitTransaction(document, _TX_TAGS):
            return _do_create()
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# Selection / UI
# ---------------------------------------------------------------------------

class _FiltroRebarSetNoSingle(ISelectionFilter):
    def AllowElement(self, elem):
        if not isinstance(elem, Rebar):
            return False
        ok, _ = _es_rebar_separable(elem)
        return ok

    def AllowReference(self, reference, position):
        return True


def _rebars_desde_seleccion_actual(uidoc):
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count < 1:
        return []
    out = []
    for eid in ids:
        el = doc.GetElement(eid)
        if not isinstance(el, Rebar):
            continue
        ok, _ = _es_rebar_separable(el)
        if ok:
            out.append(el)
    return out


def _references_from_rebars(rebars):
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import Reference

    out = List[Reference]()
    if not rebars:
        return out
    for rb in rebars:
        if rb is None:
            continue
        try:
            out.Add(Reference(rb))
        except Exception:
            try:
                out.Add(Reference(rb.Id))
            except Exception:
                pass
    return out


def _rebars_from_pick_refs(doc, refs):
    out = []
    seen = set()
    if refs is None:
        return out
    for ref in refs:
        if ref is None:
            continue
        el = None
        try:
            el = doc.GetElement(ref)
        except Exception:
            try:
                el = doc.GetElement(ref.ElementId)
            except Exception:
                el = None
        if not isinstance(el, Rebar):
            continue
        ok, _ = _es_rebar_separable(el)
        if not ok:
            continue
        rid = _element_id_int(el.Id)
        if rid is not None and rid in seen:
            continue
        if rid is not None:
            seen.add(rid)
        out.append(el)
    return out


def _pick_rebar_sets(uidoc, preselected=None):
    """
    Multiselección PickObjects. preselected se siembra (Ctrl+clic deselecciona).
    Returns list[Rebar] | None (None = cancelado).
    """
    if uidoc is None:
        return None

    pre = list(preselected or [])
    filt = _FiltroRebarSetNoSingle()
    prompt = (
        u"Seleccione uno o más sets de Rebar (layout ≠ Single). "
        u"Ctrl+clic deselecciona. Finalizar confirma; Esc cancela."
    )
    if pre:
        prompt = (
            u"{0} set(s) preseleccionado(s). Añada o quite sets "
            u"(Ctrl+clic deselecciona). Finalizar confirma; Esc cancela."
        ).format(len(pre))

    refs = None
    used_seed = False
    if pre:
        pre_refs = _references_from_rebars(pre)
        if pre_refs is not None and pre_refs.Count > 0:
            try:
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element, filt, prompt, pre_refs
                )
                used_seed = True
            except OperationCanceledException:
                return None
            except Exception:
                refs = None
                used_seed = False

    if refs is None and not used_seed:
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, filt, prompt)
        except OperationCanceledException:
            return None
        except Exception:
            return None

    rebars = _rebars_from_pick_refs(uidoc.Document, refs)
    try:
        from System.Collections.Generic import List

        id_list = List[ElementId]()
        for rb in rebars:
            try:
                id_list.Add(rb.Id)
            except Exception:
                pass
        uidoc.Selection.SetElementIds(id_list)
    except Exception:
        pass
    return rebars


def run_pyrevit(__revit__):
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        return

    # Directo a selección (sin diálogo de instrucciones ni resumen final).
    pre = _rebars_desde_seleccion_actual(uidoc)
    rebars = _pick_rebar_sets(uidoc, preselected=pre)
    if rebars is None:
        return
    if not rebars:
        return

    doc = uidoc.Document

    def _process_lote(progress):
        counters = {u"ok": 0, u"singles": 0, u"tags": 0}
        errs = []
        n_sets = len(rebars)
        for i, rebar in enumerate(rebars):
            try:
                rebar = doc.GetElement(rebar.Id)
            except Exception:
                pass
            if rebar is None or not isinstance(rebar, Rebar):
                errs.append(u"Set {0}: elemento ya no existe.".format(i + 1))
                if progress is not None:
                    progress.step(
                        u"Set {0}/{1}: omitido".format(i + 1, n_sets)
                    )
                continue
            try:
                res = separar_rebar_set_a_singles(
                    doc,
                    rebar,
                    progress=progress,
                    active_view=uidoc.ActiveView,
                )
            except Exception as ex:
                rid = _element_id_int(rebar.Id) if rebar is not None else None
                errs.append(
                    u"Id {}: excepción — {}".format(
                        rid if rid is not None else u"?",
                        _exception_text(ex),
                    )
                )
                continue

            if res.get(u"ok"):
                counters[u"ok"] += 1
                counters[u"singles"] += int(res.get(u"n_singles") or 0)
                counters[u"tags"] += int(res.get(u"n_tags") or 0)
            else:
                rid = _element_id_int(rebar.Id) if rebar is not None else None
                emsg = res.get(u"message") or u""
                errs.append(
                    u"Id {}: {}".format(
                        rid if rid is not None else u"?", emsg
                    )
                )
        return counters, errs

    total_steps = _estimate_progress_total(rebars)

    def _run_with_progress():
        with ToSingleProgress(total_steps, title_prefix=u"Arainco: To Single") as pb:
            return _process_lote(pb)

    try:
        if len(rebars) > 1:
            with _RevitTransactionGroup(doc, _TX_GROUP_NAME):
                _run_with_progress()
        else:
            _run_with_progress()
    except Exception:
        pass


def run(__revit__):
    run_pyrevit(__revit__)


def main_rps():
    try:
        run_pyrevit(__revit__)  # noqa: F821
    except NameError:
        pass
