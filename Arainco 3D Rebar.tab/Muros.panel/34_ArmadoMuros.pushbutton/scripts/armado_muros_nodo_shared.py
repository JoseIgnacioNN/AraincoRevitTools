# -*- coding: utf-8 -*-
"""
Versión portable de utilidades rebar-set para Armado Muros Cabezal.

Contiene solo las funciones de manipulación de layout/extremos de Rebar
(sin dependencias de pyrevit, area_reinforcement_losa, etc.).
"""

from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB.Structure import Rebar

try:
    from bimtools_runtime import use_rebar_setlayout_inclusion
except Exception:
    def use_rebar_setlayout_inclusion(doc=None):
        return False


def _rebar_layout_rule_nombre(rebar, acc):
    try:
        r = rebar.LayoutRule
        if r is not None:
            s = r.ToString() or u""
            if s:
                return s
    except Exception:
        pass
    if acc is not None:
        try:
            r = acc.GetLayoutRule()
            if r is not None:
                s = r.ToString() or u""
                if s:
                    return s
        except Exception:
            pass
    return u""


def _rebar_max_spacing_internal(rebar):
    try:
        return float(rebar.MaxSpacing)
    except Exception:
        return 0.0


def _rebar_array_length_internal(acc):
    if acc is None:
        return 0.0
    try:
        return float(acc.ArrayLength)
    except Exception:
        try:
            return float(acc.GetArrayLength())
        except Exception:
            return 0.0


def _rebar_cantidad_posiciones(rebar):
    best = 1
    for n in (
        _safe_int(getattr(rebar, u"NumberOfBarPositions", None)),
        _safe_int_call(rebar, u"GetNumberOfBarPositions"),
        _safe_int(getattr(rebar, u"Quantity", None)),
        _safe_int(getattr(rebar, u"TotalBarCount", None)),
        _safe_distribution_path_count(rebar),
    ):
        if n is not None and n > best:
            best = n
    return best


def _safe_int(value):
    if value is None:
        return None
    try:
        return int(value)
    except Exception:
        return None


def _safe_int_call(obj, method_name):
    if obj is None:
        return None
    try:
        fn = getattr(obj, method_name, None)
        if fn is None or not callable(fn):
            return None
        return int(fn())
    except Exception:
        return None


def _safe_distribution_path_count(rebar):
    if rebar is None:
        return None
    try:
        path = rebar.GetDistributionPath()
        if path is None:
            return None
        return int(path.Count)
    except Exception:
        return None


def _rebar_bar_included(rebar, idx):
    try:
        return bool(rebar.IsBarIncluded(int(idx)))
    except Exception:
        return True


def _excluir_barras_por_indices(rebar, indices, doc=None):
    """Excluye posiciones concretas de un set (``SetBarIncluded``), con verificación."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    n = _rebar_cantidad_posiciones(rebar)
    if n < 1:
        return False
    seen = set()
    targets = []
    for raw in indices or []:
        try:
            idx = int(raw)
        except Exception:
            continue
        if idx in seen or idx < 0 or idx >= n:
            continue
        seen.add(idx)
        targets.append(idx)
    if not targets:
        return False
    for idx in sorted(targets, reverse=True):
        try:
            rebar.SetBarIncluded(False, idx)
        except Exception:
            pass
    if doc is not None:
        try:
            doc.Regenerate()
        except Exception:
            pass
    pending = [i for i in targets if _rebar_bar_included(rebar, i)]
    if pending:
        for idx in sorted(pending, reverse=True):
            try:
                rebar.SetBarIncluded(False, idx)
            except Exception:
                pass
        if doc is not None:
            try:
                doc.Regenerate()
            except Exception:
                pass
    still = [i for i in targets if _rebar_bar_included(rebar, i)]
    return len(still) < len(targets)


def ajustar_inclusion_extremos_rebar_set(rebar, document, include_first=True, include_last=True):
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    if include_first and include_last:
        return False
    if not use_rebar_setlayout_inclusion(document):
        return False

    n = _rebar_cantidad_posiciones(rebar)
    if not include_first and not include_last:
        if n < 3:
            return False
    elif n < 2:
        return False

    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False
    rule = _rebar_layout_rule_nombre(rebar, acc)
    if rule == u"Single" or u"Single" in rule:
        return False
    sp = _rebar_max_spacing_internal(rebar)
    alen = _rebar_array_length_internal(acc)
    if alen < 1e-12:
        return False
    try:
        b_side = bool(acc.BarsOnNormalSide)
    except Exception:
        b_side = True
    inc0, inc1 = bool(include_first), bool(include_last)
    nbars = n

    def _aplicar(b_side_):
        if rule == u"MaximumSpacing":
            acc.SetLayoutAsMaximumSpacing(sp, alen, b_side_, inc0, inc1)
        elif rule in (u"Number", u"FixedNumber"):
            acc.SetLayoutAsFixedNumber(nbars, alen, b_side_, inc0, inc1)
        elif rule == u"NumberWithSpacing":
            acc.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_side_, inc0, inc1)
        elif rule == u"MinimumClearSpacing":
            acc.SetLayoutAsMinimumClearSpacing(sp, alen, b_side_, inc0, inc1)
        else:
            if rule:
                try:
                    acc.SetLayoutAsFixedNumber(nbars, alen, b_side_, inc0, inc1)
                except Exception:
                    acc.SetLayoutAsMaximumSpacing(sp, alen, b_side_, inc0, inc1)
            else:
                acc.SetLayoutAsMaximumSpacing(sp, alen, b_side_, inc0, inc1)

    for b_try in (b_side, not b_side):
        try:
            _aplicar(b_try)
            if document is not None:
                try:
                    document.Regenerate()
                except Exception:
                    pass
            return True
        except Exception:
            continue
    try:
        acc.FlipRebarSet()
    except Exception:
        return False
    for b_try in (b_side, not b_side):
        try:
            _aplicar(b_try)
            if document is not None:
                try:
                    document.Regenerate()
                except Exception:
                    pass
            return True
        except Exception:
            continue
    return False


def _excluir_barras_extremos_por_indice(rebar, include_first, include_last):
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    n = _rebar_cantidad_posiciones(rebar)
    if n < 2:
        return False
    if not include_first and not include_last and n < 3:
        return False
    ok = False
    if not include_first:
        try:
            rebar.SetBarIncluded(False, 0)
            ok = True
        except Exception:
            pass
    if not include_last:
        try:
            rebar.SetBarIncluded(False, int(n) - 1)
            ok = True
        except Exception:
            pass
    return ok


def ajustar_inclusion_extremos_rebar_set_con_fallback(
    rebar, document, include_first=True, include_last=True,
):
    if not use_rebar_setlayout_inclusion(document):
        if include_first and include_last:
            return False
        return _excluir_barras_extremos_por_indice(rebar, include_first, include_last)
    if ajustar_inclusion_extremos_rebar_set(
        rebar, document, include_first, include_last,
    ):
        return True
    if include_first and include_last:
        return False
    return _excluir_barras_extremos_por_indice(rebar, include_first, include_last)


def desactivar_extremos_rebar_set(rebar, document):
    return ajustar_inclusion_extremos_rebar_set_con_fallback(rebar, document, False, False)
