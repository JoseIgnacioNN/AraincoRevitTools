# -*- coding: utf-8 -*-
"""
Divide un conjunto de barras (Rebar) shape-driven en dos subconjuntos.

UI (pushbutton): layout Maximum Spacing + selección interactiva.
Pipeline (mallas desacople): también NumberWithSpacing / FixedNumber vía
``dividir_rebar_set_en_indice_en_tx`` (sin Transaction propia).

Implementación:
- Revit 2027+ con ``SplitRebar``: API nativa (si disponible).
- Revit 2024–2026: Copy + Move por ``GetBarPositionTransform`` + SetLayout
  con ArrayLength parcial (anclado en barra 0 / barra al corte).
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
_PARAM_CONJUNTO_GUID = u"Armadura_Conjunto_GUID"
_PARAM_MALLA_ORIENTACION = u"Armadura_Malla_Orientacion"
_PARAM_ARMADURA_NIVEL = u"Armadura_Nivel"


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def mostrar_aviso(uiapp, instruction, content=u""):
    """Aviso WPF BIMTools; respaldo a TaskDialog."""
    instruction = _as_unicode(instruction)
    content = _as_unicode(content)
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = None
        try:
            hwnd = revit_main_hwnd(uiapp)
        except Exception:
            hwnd = None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=instruction,
            content=content,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        msg = instruction
        if content:
            msg = u"{0}\n\n{1}".format(instruction, content)
        TaskDialog.Show(_DIALOG_TITLE, msg)
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


def _find_string_param(element, param_name):
    if element is None or not param_name:
        return None
    p = None
    try:
        p = element.LookupParameter(param_name)
    except Exception:
        pass
    if p is None:
        try:
            for pr in element.Parameters:
                try:
                    defn = pr.Definition
                    if defn is not None and defn.Name == param_name:
                        p = pr
                        break
                except Exception:
                    continue
        except Exception:
            pass
    if p is None:
        return None
    val = None
    try:
        val = p.AsString()
    except Exception:
        pass
    if not val:
        try:
            val = p.AsValueString()
        except Exception:
            pass
    if not val:
        return None
    try:
        t = unicode(val).strip()
    except Exception:
        try:
            t = str(val or u"").strip()
        except Exception:
            return None
    return t or None


def _get_conjunto_guid(element):
    try:
        from armado_muros_rebar_params import get_armadura_conjunto_guid

        return get_armadura_conjunto_guid(element)
    except Exception:
        return _find_string_param(element, _PARAM_CONJUNTO_GUID)


def _get_malla_orientacion(rebar):
    try:
        from armado_muros_rebar_params import get_armadura_malla_orientacion

        return get_armadura_malla_orientacion(rebar)
    except Exception:
        val = _find_string_param(rebar, _PARAM_MALLA_ORIENTACION)
        if not val:
            return None
        try:
            tl = unicode(val).strip().lower()
        except NameError:
            tl = str(val or u"").strip().lower()
        if tl.startswith(u"v"):
            return u"vertical"
        if tl.startswith(u"h"):
            return u"horizontal"
        return None


def _get_armadura_nivel(element):
    """Lee ``Armadura_Nivel`` (nombre de nivel) o ``None``."""
    return _find_string_param(element, _PARAM_ARMADURA_NIVEL)


def _collect_rebar_ids_conjunto_guid(doc, conjunto_guid):
    if doc is None or not conjunto_guid:
        return []
    try:
        from armado_muros_rebar_params import collect_rebars_por_conjunto_guid

        return list(collect_rebars_por_conjunto_guid(doc, conjunto_guid))
    except Exception:
        from Autodesk.Revit.DB import FilteredElementCollector

        try:
            target = unicode(conjunto_guid).strip()
        except NameError:
            target = str(conjunto_guid or u"").strip()
        if not target:
            return []
        ids = []
        try:
            rebars = (
                FilteredElementCollector(doc)
                .OfClass(Rebar)
                .WhereElementIsNotElementType()
            )
        except Exception:
            return []
        for rb in rebars:
            gid = _get_conjunto_guid(rb)
            if gid == target:
                try:
                    ids.append(rb.Id)
                except Exception:
                    pass
        return ids


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


def _layout_rule_kind(rebar):
    """
    ``maximum`` | ``number_spacing`` | ``fixed`` | ``other``.
    """
    rule = _layout_rule_nombre(rebar) or u""
    if rule == u"MaximumSpacing" or u"MaximumSpacing" in rule:
        return u"maximum"
    if rule == u"NumberWithSpacing" or u"NumberWithSpacing" in rule:
        return u"number_spacing"
    if rule in (u"Number", u"FixedNumber") or u"FixedNumber" in rule:
        return u"fixed"
    if RebarShapeDrivenLayoutRule is not None:
        try:
            lr = rebar.LayoutRule
            name = lr.ToString() if lr is not None else u""
            if u"MaximumSpacing" in name:
                return u"maximum"
            if u"NumberWithSpacing" in name:
                return u"number_spacing"
            if u"FixedNumber" in name or name == u"Number":
                return u"fixed"
        except Exception:
            pass
    return u"other"


def _es_rebar_divisible(rebar):
    """UI: solo Maximum Spacing."""
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


def _es_rebar_divisible_pipeline(rebar):
    """Pipeline desacople: MaxSpacing / NumberWithSpacing / FixedNumber."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False, u"No es un elemento Rebar."
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False, u"Solo aplica a barras shape-driven (no free-form)."
    kind = _layout_rule_kind(rebar)
    if kind not in (u"maximum", u"number_spacing", u"fixed"):
        return False, u"Layout no soportado para dividir: {0}.".format(
            _layout_rule_nombre(rebar) or kind,
        )
    n = _cantidad_posiciones(rebar)
    if n < 2:
        return False, u"El conjunto debe tener al menos dos posiciones de barra."
    return True, u""


def _peers_misma_orientacion_y_guid(doc, seed_rebar):
    """
    Otras rebars con el mismo ``Armadura_Conjunto_GUID``,
    ``Armadura_Malla_Orientacion`` y ``Armadura_Nivel``
    (cara opuesta u host distinto en el mismo nivel).

    Mallas de otros niveles se excluyen aunque compartan el GUID de creación.
    """
    guid = _get_conjunto_guid(seed_rebar)
    orient = _get_malla_orientacion(seed_rebar)
    if not guid or not orient:
        return []
    seed_nivel = _get_armadura_nivel(seed_rebar)
    seed_id = _element_id_int(seed_rebar.Id)
    peers = []
    for eid in _collect_rebar_ids_conjunto_guid(doc, guid):
        try:
            rb = doc.GetElement(eid)
        except Exception:
            continue
        if rb is None or not isinstance(rb, Rebar):
            continue
        if _element_id_int(rb.Id) == seed_id:
            continue
        if _get_malla_orientacion(rb) != orient:
            continue
        if _get_armadura_nivel(rb) != seed_nivel:
            continue
        ok, _ = _es_rebar_divisible(rb)
        if not ok:
            continue
        peers.append(rb)
    return peers


def _targets_dividir_ui(doc, seed_rebar):
    return [seed_rebar] + _peers_misma_orientacion_y_guid(doc, seed_rebar)


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


def _ajustar_indice_ultima_barra(rebar, bar_index):
    """Si el índice es la última barra, cortar antes para aislarla (n-2)."""
    n = _cantidad_posiciones(rebar)
    idx = int(bar_index)
    if n >= 2 and idx >= n - 1:
        return n - 2
    return idx


def _indice_equivalente(rebar_src, idx_src, rebar_dst):
    """
    Índice de corte en ``rebar_dst`` equivalente al de ``rebar_src``.

    Misma cantidad → mismo índice (si es válido). Distinta cantidad → mapeo
    proporcional en 0..n-1, acotado a [0, n-2] para no dejar un set vacío.
    """
    n_src = _cantidad_posiciones(rebar_src)
    n_dst = _cantidad_posiciones(rebar_dst)
    max_split = n_dst - 2
    if max_split < 0:
        return None
    idx_src = int(idx_src)
    if n_dst == n_src:
        if 0 <= idx_src <= max_split:
            return idx_src
        if idx_src >= n_src - 1:
            return max_split
        return max(0, min(idx_src, max_split))
    if n_src <= 1:
        return max(0, min(idx_src, max_split))
    mapped = int(round(float(idx_src) * float(n_dst - 1) / float(n_src - 1)))
    return max(0, min(mapped, max_split))


def _plan_division_ui(document, seed, bar_index):
    """
    Plan de corte por Rebar (semilla + peers GUID/orientación/nivel).

    Returns:
        ``(ok, mensaje_error, plan, avisos)``
        ``plan`` = lista de ``(rebar, idx)``; si la semilla no es válida, ``ok`` es False.
    """
    idx_seed = _ajustar_indice_ultima_barra(seed, bar_index)
    seed_id = _element_id_int(seed.Id)
    plan = []
    avisos = []
    for rb in _targets_dividir_ui(document, seed):
        rid = _element_id_int(rb.Id)
        ok_pre, msg_pre = _es_rebar_divisible(rb)
        if not ok_pre:
            if rid == seed_id:
                return False, u"Rebar {0}: {1}".format(rid, msg_pre), [], []
            avisos.append(u"Rebar {0}: {1}".format(rid, msg_pre))
            continue
        if rid == seed_id:
            idx_rb = idx_seed
        else:
            idx_rb = _indice_equivalente(seed, idx_seed, rb)
            if idx_rb is None:
                avisos.append(
                    u"Rebar {0}: no se pudo calcular un índice de corte equivalente.".format(
                        rid
                    )
                )
                continue
        ok_idx, msg_idx = _validar_indice_division(rb, idx_rb)
        if not ok_idx:
            if rid == seed_id:
                return False, u"Rebar {0}: {1}".format(rid, msg_idx), [], []
            avisos.append(u"Rebar {0}: {1}".format(rid, msg_idx))
            continue
        plan.append((rb, idx_rb))
    if not plan:
        return False, u"Ningún conjunto se pudo dividir.", [], avisos
    return True, u"", plan, avisos


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


def _aplicar_layout_segmento(
    acc, spacing, array_len, b_side, inc_first, inc_last,
    rule_kind=u"maximum", n_bars=None,
):
    """Aplica layout al segmento. ``n_bars`` obligatorio si no es MaximumSpacing."""
    if acc is None:
        return False
    kind = rule_kind or u"maximum"
    try:
        n_seg = int(n_bars) if n_bars is not None else 0
    except Exception:
        n_seg = 0

    if n_seg == 1 or array_len < 1e-9:
        try:
            acc.SetLayoutAsSingle()
            return True
        except Exception:
            return False

    if kind == u"maximum" and spacing > 1e-12 and spacing >= array_len - 1e-9:
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
            if kind == u"number_spacing":
                if n_seg < 1 or spacing < 1e-12:
                    continue
                acc.SetLayoutAsNumberWithSpacing(
                    int(n_seg), float(spacing), float(array_len), b_try, i0, i1,
                )
                return True
            if kind == u"fixed":
                if n_seg < 1:
                    continue
                acc.SetLayoutAsFixedNumber(
                    int(n_seg), float(array_len), b_try, i0, i1,
                )
                return True
            # maximum (default)
            acc.SetLayoutAsMaximumSpacing(
                float(spacing), float(array_len), b_try, i0, i1,
            )
            return True
        except Exception:
            continue
    # Último recurso: FixedNumber / Single
    if n_seg >= 2:
        for b_try, i0, i1 in combos:
            try:
                acc.SetLayoutAsFixedNumber(
                    int(n_seg), float(array_len), b_try, i0, i1,
                )
                return True
            except Exception:
                continue
    if n_seg == 1 or array_len < 1e-9:
        try:
            acc.SetLayoutAsSingle()
            return True
        except Exception:
            pass
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


def _distribucion_fallback(rebar, n, spacing, alen):
    """
    Si los transforms no distinguen posiciones, estima eje de reparto con
    Normal del shape-driven × eje vertical (mallas verticales).
    """
    t0 = _get_bar_transform(rebar, 0)
    if t0 is None:
        return None, None, None
    direction = None
    try:
        acc = rebar.GetShapeDrivenAccessor()
        normal = acc.Normal if acc is not None else None
        if normal is not None:
            for up in (XYZ(0.0, 0.0, 1.0), XYZ(1.0, 0.0, 0.0), XYZ(0.0, 1.0, 0.0)):
                cross = normal.CrossProduct(up)
                if cross is not None and cross.GetLength() > 1e-6:
                    direction = cross.Normalize()
                    break
    except Exception:
        direction = None
    if direction is None:
        return None, None, None

    positions = []
    if spacing is not None and float(spacing) > 1e-12 and n > 1:
        for i in range(n):
            positions.append(float(i) * float(spacing))
    elif alen is not None and float(alen) > 1e-12 and n > 1:
        for i in range(n):
            positions.append(float(i) / float(n - 1) * float(alen))
    else:
        return None, None, None
    return direction, t0, positions


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


def _rebar_ids_in_document(document):
    out = set()
    try:
        from Autodesk.Revit.DB import FilteredElementCollector

        for e in FilteredElementCollector(document).OfClass(Rebar):
            iid = _element_id_int(getattr(e, "Id", None))
            if iid is not None:
                out.add(int(iid))
    except Exception:
        pass
    return out


def _dividir_manual(document, rebar, bar_index):
    """
    Copy + Move + SetLayout. Devuelve ``(ok, msg, rb_left, rb_right)``.
    ``rb_left`` = original acortado (barras 0..idx); ``rb_right`` = copia.
    """
    idx = int(bar_index)
    n = _cantidad_posiciones(rebar)
    rule_kind = _layout_rule_kind(rebar)
    acc0, spacing, alen_total, b_side, inc0, inc1 = _layout_params(rebar)
    if acc0 is None:
        return False, u"GetShapeDrivenAccessor no disponible.", None, None

    direction, t0 = _distribucion_desde_rebar(rebar, n)
    positions_fb = None
    if direction is None:
        direction, t0, positions_fb = _distribucion_fallback(
            rebar, n, spacing, alen_total,
        )
        if direction is None or t0 is None or not positions_fb:
            return False, u"No se pudieron leer posiciones de barras (GetBarPositionTransform).", None, None

    if positions_fb is not None:
        pos_idx = float(positions_fb[idx])
        pos_next = float(positions_fb[idx + 1])
        pos_last = float(positions_fb[n - 1])
        delta_move = XYZ(
            float(direction.X) * pos_next,
            float(direction.Y) * pos_next,
            float(direction.Z) * pos_next,
        )
    else:
        pos_idx = _posicion_escalar(rebar, idx, direction, t0)
        pos_next = _posicion_escalar(rebar, idx + 1, direction, t0)
        pos_last = _posicion_escalar(rebar, n - 1, direction, t0)
        if pos_idx is None or pos_next is None or pos_last is None:
            return False, u"No se pudo calcular la posición de corte.", None, None
        t_next = _get_bar_transform(rebar, idx + 1)
        if t_next is None:
            return False, u"No se pudo leer la transformación de la barra {}.".format(idx + 1), None, None
        delta_move = t_next.Origin - t0.Origin

    len1 = max(0.0, float(pos_idx))
    len2 = max(0.0, float(pos_last - pos_next))
    n1 = int(idx) + 1
    n2 = int(n) - (int(idx) + 1)
    if n1 < 1 or n2 < 1:
        return False, u"Segmentos con n<1 (n1={0} n2={1}).".format(n1, n2), None, None

    try:
        new_ids = ElementTransformUtils.CopyElement(document, rebar.Id, XYZ.Zero)
    except Exception as ex:
        try:
            msg = unicode(ex)
        except NameError:
            msg = str(ex)
        return False, u"No se pudo copiar el Rebar: {}".format(msg), None, None

    if new_ids is None or len(new_ids) < 1:
        return False, u"CopyElement no devolvió elementos.", None, None

    rb2 = document.GetElement(new_ids[0])
    if rb2 is None:
        return False, u"No se pudo obtener la copia del Rebar.", None, None

    if not _aplicar_layout_segmento(
        acc0, spacing, len1, b_side, inc0, True,
        rule_kind=rule_kind, n_bars=n1,
    ):
        try:
            document.Delete(rb2.Id)
        except Exception:
            pass
        return False, u"No se pudo aplicar layout al primer subconjunto.", None, None

    acc2 = rb2.GetShapeDrivenAccessor()
    if acc2 is None:
        try:
            document.Delete(rb2.Id)
        except Exception:
            pass
        return False, u"La copia no tiene ShapeDrivenAccessor.", None, None

    if delta_move.GetLength() > 1e-9:
        try:
            ElementTransformUtils.MoveElement(document, rb2.Id, delta_move)
        except Exception as ex:
            try:
                document.Delete(rb2.Id)
            except Exception:
                pass
            try:
                msg = unicode(ex)
            except NameError:
                msg = str(ex)
            return False, u"No se pudo trasladar el segundo subconjunto: {}".format(msg), None, None

    if not _aplicar_layout_segmento(
        acc2, spacing, len2, b_side, True, inc1,
        rule_kind=rule_kind, n_bars=n2,
    ):
        try:
            document.Delete(rb2.Id)
        except Exception:
            pass
        return False, u"No se pudo aplicar layout al segundo subconjunto.", None, None

    try:
        document.Regenerate()
    except Exception:
        pass
    return True, u"", rebar, rb2


def _dividir_manual_max_spacing(document, rebar, bar_index):
    """Compat UI: ``(ok, msg)``."""
    ok, msg, _a, _b = _dividir_manual(document, rebar, bar_index)
    return ok, msg


def dividir_rebar_set_en_indice_en_tx(document, rebar, bar_index):
    """
    Divide ``rebar`` en el índice indicado **sin** abrir Transaction.

    Prefiere el camino manual (Copy+Move+SetLayout) porque devuelve ambos
    Rebar de forma fiable. ``SplitRebar`` solo si el manual falla.

    Returns:
        ``(ok, mensaje, [rb_left, rb_right])`` — elementos Rebar o lista vacía.
    """
    ok_pre, msg_pre = _es_rebar_divisible_pipeline(rebar)
    if not ok_pre:
        return False, msg_pre, []

    ok_idx, msg_idx = _validar_indice_division(rebar, bar_index)
    if not ok_idx:
        return False, msg_idx, []

    idx = int(bar_index)
    n = _cantidad_posiciones(rebar)

    ok_m, msg_m, rb_l, rb_r = _dividir_manual(document, rebar, idx)
    if ok_m and rb_l is not None and rb_r is not None:
        detalle = (
            u"Corte tras índice {} (manual): left=0–{}, right={}-{}."
        ).format(idx, idx, idx + 1, max(idx + 1, n - 1))
        return True, detalle, [rb_l, rb_r]

    if _split_rebar_api_disponible(rebar):
        ids_before = _rebar_ids_in_document(document)
        id0 = _element_id_int(getattr(rebar, "Id", None))
        try:
            _dividir_con_split_rebar_api(document, rebar, idx)
            ids_after = _rebar_ids_in_document(document)
            new_ids = [i for i in ids_after if i not in ids_before]
            rb_left = None
            rb_right = None
            try:
                rb_left = document.GetElement(rebar.Id)
            except Exception:
                rb_left = None
            if rb_left is None or not isinstance(rb_left, Rebar):
                rb_left = None
            for iid in new_ids:
                try:
                    el = document.GetElement(ElementId(iid))
                except Exception:
                    el = None
                if el is not None and isinstance(el, Rebar):
                    if rb_left is None:
                        rb_left = el
                    elif rb_right is None and el.Id != rb_left.Id:
                        rb_right = el
            if rb_left is not None and rb_right is None and id0 is not None:
                # Buscar cualquier otro rebar nuevo
                for iid in ids_after:
                    if iid == id0:
                        continue
                    try:
                        el = document.GetElement(ElementId(iid))
                    except Exception:
                        el = None
                    if (
                        el is not None
                        and isinstance(el, Rebar)
                        and (rb_left is None or el.Id != rb_left.Id)
                    ):
                        if rb_left is None:
                            rb_left = el
                        else:
                            rb_right = el
                            break
            if rb_left is not None and rb_right is not None:
                detalle = (
                    u"Corte tras índice {} (SplitRebar): left=0–{}, right={}-{}."
                ).format(idx, idx, idx + 1, max(idx + 1, n - 1))
                return True, detalle, [rb_left, rb_right]
        except Exception as ex_sp:
            try:
                msg_m = u"manual: {0}; SplitRebar: {1}".format(msg_m, ex_sp)
            except Exception:
                pass

    return False, msg_m or u"división falló", []


def dividir_rebar_set_en_indice(document, rebar, bar_index):
    """
    Divide ``rebar`` en el índice indicado (abre Transaction).

    Si la barra tiene ``Armadura_Conjunto_GUID`` y ``Armadura_Malla_Orientacion``,
    divide también las demás rebars del mismo GUID, misma orientación y
    mismo ``Armadura_Nivel`` (p. ej. cara interior y exterior de una malla
    en el mismo nivel). No se dividen mallas de otros niveles aunque
    compartan el GUID de creación.

    Returns:
        (ok: bool, mensaje: unicode, ids_resultantes: list)
    """
    ok_plan, msg_plan, plan, avisos = _plan_division_ui(document, rebar, bar_index)
    if not ok_plan:
        extra = u""
        if avisos:
            extra = u"\n" + u"\n".join(avisos)
        return False, (msg_plan or u"Ningún conjunto se pudo dividir.") + extra, []

    t = Transaction(document, _TRANSACTION_NAME)
    t.Start()
    msg = u""
    try:
        for rb, idx_rb in plan:
            ok, msg_rb, _ = dividir_rebar_set_en_indice_en_tx(document, rb, idx_rb)
            if not ok:
                rid = _element_id_int(rb.Id)
                raise RuntimeError(
                    u"Rebar {0}: {1}".format(rid, msg_rb or u"Error al dividir el conjunto."),
                )
            if not msg:
                msg = msg_rb or u""
        t.Commit()
    except Exception as ex:
        t.RollBack()
        try:
            msg = unicode(ex) if ex else u"Error al dividir el conjunto."
        except NameError:
            msg = str(ex) if ex else u"Error al dividir el conjunto."
        return False, msg, []

    n_extra = len(plan) - 1
    if n_extra > 0:
        msg = (
            u"{0}\n\nConjuntos adicionales divididos (misma GUID, orientación y nivel): {1}."
        ).format(msg, n_extra)
    if avisos:
        msg = (
            u"{0}\n\nConjuntos no divididos (índice equivalente o layout):\n{1}"
        ).format(msg, u"\n".join(avisos))

    return True, msg, []


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
    prompt = u"2/2 — Selecciona la barra de corte (última del primer set; la última aísla esa barra)."
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
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        mostrar_aviso(__revit__, u"No hay documento activo.")
        return

    rebar = _rebar_desde_seleccion_actual(uidoc)
    if rebar is None:
        rebar = _pick_rebar_max_spacing(uidoc)
    if rebar is None:
        return

    bar_index = _pick_bar_index(uidoc, rebar)
    if bar_index is None:
        return

    n_pick = _cantidad_posiciones(rebar)
    nota_ultima = u""
    if n_pick >= 2 and int(bar_index) >= n_pick - 1:
        bar_index = n_pick - 2
        nota_ultima = u"\nLa última barra quedó aislada como segundo subconjunto."

    ok_idx, msg_idx = _validar_indice_division(rebar, bar_index)
    if not ok_idx:
        mostrar_aviso(__revit__, msg_idx)
        return

    doc = uidoc.Document
    ok, msg, _ = dividir_rebar_set_en_indice(doc, rebar, bar_index)
    if ok:
        mostrar_aviso(
            __revit__,
            u"Conjunto dividido correctamente.",
            u"{0}\nÍndice de corte: {1}{2}".format(msg, bar_index, nota_ultima),
        )
    else:
        mostrar_aviso(__revit__, u"No se pudo dividir.", msg)


def run(__revit__):
    run_pyrevit(__revit__)


def main_rps():
    """RevitPythonShell: ejecuta con un Rebar Maximum Spacing preseleccionado o flujo interactivo."""
    try:
        run_pyrevit(__revit__)  # noqa: F821
    except NameError:
        TaskDialog.Show(_DIALOG_TITLE, u"Ejecuta en pyRevit o RPS con __revit__ disponible.")
