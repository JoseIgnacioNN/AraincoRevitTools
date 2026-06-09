# -*- coding: utf-8 -*-
"""
Divide un conjunto de barras (Rebar) con regla de trazado *Maximum Spacing* en dos
subconjuntos en la barra seleccionada.

Flujo:
1. Seleccionar el Rebar (shape-driven, layout Maximum Spacing, más de una barra).
2. Seleccionar una barra del conjunto (subelemento).
3. Dividir el set en el índice de esa barra.

Implementación:
- Revit 2027+ con ``SplitRebar``: API nativa.
- Revit 2024–2026: recrea dos conjuntos copiando el Rebar, ajustando longitud de
  distribución con ``GetBarPositionTransform`` y ``SetLayoutAsMaximumSpacing``.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import ElementId, ElementTransformUtils, Transaction, XYZ
from Autodesk.Revit.DB.Structure import Rebar
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.DB.Structure import RebarShapeDrivenLayoutRule
except Exception:
    RebarShapeDrivenLayoutRule = None

_DIALOG_TITLE = u"Arainco: Dividir rebar set Maximum Spacing"
_TRANSACTION_NAME = u"Arainco: Dividir rebar set Maximum Spacing"


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


def _es_layout_maximum_spacing(rebar):
    if RebarShapeDrivenLayoutRule is not None:
        try:
            return rebar.LayoutRule == RebarShapeDrivenLayoutRule.MaximumSpacing
        except Exception:
            pass
    rule = _layout_rule_nombre(rebar)
    return rule == u"MaximumSpacing" or u"MaximumSpacing" in rule


def _es_rebar_divisible(rebar):
    if rebar is None or not isinstance(rebar, Rebar):
        return False, u"No es un elemento Rebar."
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False, u"Solo aplica a barras shape-driven (no free-form)."
    if not _es_layout_maximum_spacing(rebar):
        return False, u"La regla de trazado debe ser Maximum Spacing (separación máxima)."
    n = _cantidad_posiciones(rebar)
    if n < 2:
        return False, u"El conjunto debe tener al menos dos posiciones de barra."
    return True, u""


def _bar_index_desde_referencia(rebar, reference):
    if rebar is None or reference is None:
        return -1
    try:
        idx = int(rebar.GetBarIndexFromReference(reference))
    except Exception:
        idx = -1
    return idx


def _validar_indice_division(rebar, bar_index):
    n = _cantidad_posiciones(rebar)
    idx = int(bar_index)
    if idx < 0 or idx >= n:
        return False, u"Índice de barra fuera de rango (0–{}).".format(max(0, n - 1))
    max_split = n - 2
    if idx > max_split:
        return False, (
            u"No se puede dividir en la última barra del conjunto "
            u"(selecciona una barra anterior)."
        )
    if _split_rebar_api_disponible(rebar):
        try:
            if hasattr(rebar, "AreBarIndicesValidForSplit"):
                indices = List[int]()
                indices.Add(idx)
                if not rebar.AreBarIndicesValidForSplit(indices):
                    return False, u"Revit rechaza la división en el índice {}.".format(idx)
        except Exception:
            pass
    return True, u""


def _split_rebar_api_disponible(rebar=None):
    try:
        if rebar is not None and getattr(rebar, "SplitRebar", None) is not None:
            return True
    except Exception:
        pass
    try:
        import System.Reflection as SR

        for mi in Rebar.GetMethods(SR.BindingFlags.Public | SR.BindingFlags.Instance):
            if mi.Name == "SplitRebar":
                return True
    except Exception:
        pass
    return False


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


def _layout_params(rebar):
    acc = rebar.GetShapeDrivenAccessor()
    sp = 0.0
    try:
        sp = float(rebar.MaxSpacing)
    except Exception:
        pass
    alen = 0.0
    if acc is not None:
        try:
            alen = float(acc.ArrayLength)
        except Exception:
            try:
                alen = float(acc.GetArrayLength())
            except Exception:
                pass
    b_side = True
    if acc is not None:
        try:
            b_side = bool(acc.BarsOnNormalSide)
        except Exception:
            pass
    try:
        inc0 = bool(rebar.IncludeFirstBar)
    except Exception:
        inc0 = True
    try:
        inc1 = bool(rebar.IncludeLastBar)
    except Exception:
        inc1 = True
    return acc, sp, alen, b_side, inc0, inc1


def _aplicar_layout_segmento(acc, spacing, array_len, b_side, inc_first, inc_last):
    if acc is None:
        return False
    if array_len < 1e-9 or spacing >= array_len - 1e-9:
        try:
            acc.SetLayoutAsSingle()
            return True
        except Exception:
            return False
    combos = (
        (bool(b_side), bool(inc_first), bool(inc_last)),
        (not bool(b_side), bool(inc_first), bool(inc_last)),
    )
    for b_try, i0, i1 in combos:
        try:
            acc.SetLayoutAsMaximumSpacing(float(spacing), float(array_len), b_try, i0, i1)
            return True
        except Exception:
            continue
    return False


def _distribucion_desde_rebar(rebar, n):
    t0 = _get_bar_transform(rebar, 0)
    t_last = _get_bar_transform(rebar, n - 1)
    if t0 is None or t_last is None:
        return None, None
    delta = t_last.Origin - t0.Origin
    if delta.GetLength() < 1e-9:
        return None, t0
    return delta.Normalize(), t0


def _posicion_escalar(rebar, bar_index, direction, t0):
    t_bar = _get_bar_transform(rebar, bar_index)
    if t_bar is None or t0 is None or direction is None:
        return None
    try:
        return float((t_bar.Origin - t0.Origin).DotProduct(direction))
    except Exception:
        return None


def _dividir_con_split_rebar_api(document, rebar, bar_index):
    indices = List[int]()
    indices.Add(int(bar_index))
    rebar.SplitRebar(indices)
    document.Regenerate()


def _dividir_manual_max_spacing(document, rebar, bar_index):
    idx = int(bar_index)
    n = _cantidad_posiciones(rebar)
    acc0, spacing, _alen_total, b_side, inc0, inc1 = _layout_params(rebar)
    if acc0 is None:
        return False, u"GetShapeDrivenAccessor no disponible."

    direction, t0 = _distribucion_desde_rebar(rebar, n)
    if direction is None or t0 is None:
        return False, u"No se pudieron leer posiciones de barras (GetBarPositionTransform)."

    pos_idx = _posicion_escalar(rebar, idx, direction, t0)
    pos_next = _posicion_escalar(rebar, idx + 1, direction, t0)
    pos_last = _posicion_escalar(rebar, n - 1, direction, t0)
    if pos_idx is None or pos_next is None or pos_last is None:
        return False, u"No se pudo calcular la posición de corte."

    len1 = max(0.0, float(pos_idx))
    len2 = max(0.0, float(pos_last - pos_next))

    t_next = _get_bar_transform(rebar, idx + 1)
    if t_next is None:
        return False, u"No se pudo leer la transformación de la barra {}.".format(idx + 1)
    delta_move = t_next.Origin - t0.Origin

    try:
        new_ids = ElementTransformUtils.CopyElement(document, rebar.Id, XYZ.Zero)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        return False, u"No se pudo copiar el Rebar: {}".format(msg)

    if new_ids is None or len(new_ids) < 1:
        return False, u"CopyElement no devolvió elementos."

    rb2 = document.GetElement(new_ids[0])
    if rb2 is None:
        return False, u"No se pudo obtener la copia del Rebar."

    if not _aplicar_layout_segmento(acc0, spacing, len1, b_side, inc0, True):
        return False, u"No se pudo aplicar Maximum Spacing al primer subconjunto."

    acc2 = rb2.GetShapeDrivenAccessor()
    if acc2 is None:
        return False, u"La copia no tiene ShapeDrivenAccessor."

    if delta_move.GetLength() > 1e-9:
        try:
            ElementTransformUtils.MoveElement(document, rb2.Id, delta_move)
        except Exception as ex:
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            return False, u"No se pudo trasladar el segundo subconjunto: {}".format(msg)

    if not _aplicar_layout_segmento(acc2, spacing, len2, b_side, True, inc1):
        return False, u"No se pudo aplicar Maximum Spacing al segundo subconjunto."

    document.Regenerate()
    return True, u""


def dividir_rebar_set_en_indice(document, rebar, bar_index):
    """
    Divide ``rebar`` en el índice indicado.

    Returns:
        (ok: bool, mensaje: unicode, ids_resultantes: list)
    """
    ok_pre, msg_pre = _es_rebar_divisible(rebar)
    if not ok_pre:
        return False, msg_pre, []

    ok_idx, msg_idx = _validar_indice_division(rebar, bar_index)
    if not ok_idx:
        return False, msg_idx, []

    idx = int(bar_index)
    n = _cantidad_posiciones(rebar)
    metodo = u"manual"

    t = Transaction(document, _TRANSACTION_NAME)
    t.Start()
    try:
        if _split_rebar_api_disponible(rebar):
            try:
                _dividir_con_split_rebar_api(document, rebar, idx)
                metodo = u"SplitRebar"
            except Exception:
                ok_m, msg_m = _dividir_manual_max_spacing(document, rebar, idx)
                if not ok_m:
                    raise RuntimeError(msg_m)
        else:
            ok_m, msg_m = _dividir_manual_max_spacing(document, rebar, idx)
            if not ok_m:
                raise RuntimeError(msg_m)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        try:
            msg = unicode(ex) if ex else u"Error al dividir el conjunto."
        except NameError:
            msg = str(ex) if ex else u"Error al dividir el conjunto."
        return False, msg, []

    detalle = (
        u"Corte tras la barra índice {} ({}): subconjunto 1 (barras 0–{}), "
        u"subconjunto 2 (barras {}–{})."
    ).format(idx, metodo, idx, idx + 1, max(idx + 1, n - 1))
    return True, detalle, []


class _FiltroRebarMaxSpacing(ISelectionFilter):
    def AllowElement(self, elem):
        if not isinstance(elem, Rebar):
            return False
        ok, _ = _es_rebar_divisible(elem)
        return ok

    def AllowReference(self, reference, position):
        return False


class _FiltroBarraDeRebar(ISelectionFilter):
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


def _object_type_subelement():
    try:
        return ObjectType.Subelement
    except Exception:
        return None


def _pick_rebar_max_spacing(uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroRebarMaxSpacing(),
            u"1/2 — Selecciona un Rebar con layout Maximum Spacing (más de una barra).",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    return uidoc.Document.GetElement(ref.ElementId)


def _pick_bar_index(uidoc, rebar):
    ot_sub = _object_type_subelement()
    prompt = u"2/2 — Selecciona la barra donde dividir el conjunto."
    if ot_sub is not None:
        try:
            ref = uidoc.Selection.PickObject(
                ot_sub,
                _FiltroBarraDeRebar(rebar.Id),
                prompt,
            )
        except OperationCanceledException:
            return None
        except Exception:
            ref = None
        if ref is not None:
            idx = _bar_index_desde_referencia(rebar, ref)
            if idx >= 0:
                return idx

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FiltroBarraDeRebar(rebar.Id),
            prompt + u" (clic en el conjunto; se usará la barra central si no hay subelemento).",
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None

    if ref is None:
        return None

    idx = _bar_index_desde_referencia(rebar, ref)
    if idx >= 0:
        return idx

    n = _cantidad_posiciones(rebar)
    if n >= 2:
        return min(n // 2, n - 2)
    return None


def _rebar_desde_seleccion_actual(uidoc):
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count != 1:
        return None
    el = doc.GetElement(ids[0])
    if not isinstance(el, Rebar):
        return None
    ok, _ = _es_rebar_divisible(el)
    return el if ok else None


def run_pyrevit(__revit__):
    from dividir_rebar_set_instruction_dialog import (
        show_info_dialog,
        show_selection_instructions,
    )

    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        show_info_dialog(_DIALOG_TITLE, u"No hay documento activo.", uiapp=__revit__)
        return

    if not show_selection_instructions(__revit__):
        return

    rebar = _rebar_desde_seleccion_actual(uidoc)
    if rebar is None:
        rebar = _pick_rebar_max_spacing(uidoc)
    if rebar is None:
        return

    bar_index = _pick_bar_index(uidoc, rebar)
    if bar_index is None:
        return

    ok_idx, msg_idx = _validar_indice_division(rebar, bar_index)
    if not ok_idx:
        show_info_dialog(_DIALOG_TITLE, msg_idx, uiapp=__revit__)
        return

    doc = uidoc.Document
    ok, msg, _ = dividir_rebar_set_en_indice(doc, rebar, bar_index)
    if ok:
        show_info_dialog(
            _DIALOG_TITLE,
            u"Conjunto dividido correctamente.\n\n{}\nÍndice de corte: {}".format(
                msg, bar_index
            ),
            uiapp=__revit__,
        )
    else:
        show_info_dialog(
            _DIALOG_TITLE,
            u"No se pudo dividir:\n\n{}".format(msg),
            uiapp=__revit__,
        )


def main_rps():
    """RevitPythonShell: ejecuta con un Rebar Maximum Spacing preseleccionado o flujo interactivo."""
    try:
        run_pyrevit(__revit__)  # noqa: F821
    except NameError:
        TaskDialog.Show(_DIALOG_TITLE, u"Ejecuta en pyRevit o RPS con __revit__ disponible.")
