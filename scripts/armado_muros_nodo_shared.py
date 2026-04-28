# -*- coding: utf-8 -*-
"""
Utilidades compartidas para **Armado Muros Nodo** (BIMTools).
Lógica equivalente a `armado_muros_area` en BrassDev: combos de barra, espaciado, vista 3D, diálogo de uniones.
"""

from System.Windows.Controls import ComboBoxItem

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    SpecTypeId,
    UnitUtils,
    UnitTypeId,
    View,
    View3D,
    Wall,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    Rebar,
    RebarInSystem,
    StructuralWallUsage,
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

from pyrevit import forms

from area_reinforcement_losa import _get_rebar_bar_types
from bimtools_joined_geometry import get_joined_element_ids
from wall_node_boolean_section_rps import get_otros_element_ids_boceto_nudo


def _aplicar_solido_visible_en_vista(refuerzo, doc, view):
    """
    ``SetUnobscuredInView`` + ``SetSolidInView`` para ``AreaReinforcement``, ``Rebar`` o
    ``RebarInSystem`` (sin ocultar por el muro, forma sólida en vista 3D/planta/sección
    cuando la API lo permite).
    """
    if refuerzo is None or doc is None or view is None:
        return
    if not isinstance(view, View):
        return
    if not isinstance(
        refuerzo, (AreaReinforcement, Rebar, RebarInSystem)
    ):
        return
    try:
        refuerzo.SetUnobscuredInView(view, True)
    except Exception:
        pass
    try:
        refuerzo.SetSolidInView(view, True)
    except Exception:
        try:
            fn = getattr(refuerzo, "SetSolidInView", None)
            if fn is not None:
                fn(view, True)
        except Exception:
            pass
    if isinstance(refuerzo, (Rebar, RebarInSystem)):
        return
    if not isinstance(refuerzo, AreaReinforcement):
        return
    try:
        ids = refuerzo.GetRebarInSystemIds()
    except Exception:
        ids = None
    if not ids:
        return
    for rid in ids:
        el = doc.GetElement(rid)
        if el is None:
            continue
        _aplicar_solido_visible_en_vista(el, doc, view)


def _aplicar_solid_en_vista_3d(area_rein, doc, view):
    """
    Misma lógica que :func:`_aplicar_solido_visible_en_vista` (cualquier ``View``, no solo 3D).
    Mantiene el nombre para compatibilidad.
    """
    _aplicar_solido_visible_en_vista(area_rein, doc, view)


def _element_id_int(eid):
    """Id numérico de ElementId: Revit 2024+ usa ``Value``; versiones antiguas ``IntegerValue``."""
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        return int(eid.IntegerValue)
    except AttributeError:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _fill_bar_combo(combo, doc):
    """Rellena ComboBox con RebarBarType del proyecto."""
    combo.Items.Clear()
    pairs = _get_rebar_bar_types(doc)
    for name, bar_type in pairs:
        it = ComboBoxItem()
        it.Content = name
        it.Tag = bar_type.Id
        combo.Items.Add(it)
    if combo.Items.Count > 0:
        combo.SelectedIndex = 0


def _is_structural_wall(wall):
    """True si el uso estructural está definido y no es NonBearing."""
    if wall is None or not isinstance(wall, Wall):
        return False
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM)
        if p is None or not p.HasValue:
            return False
        u = p.AsInteger()
        return u != int(StructuralWallUsage.NonBearing)
    except Exception:
        return False


def _describe_elemento_linea(doc, eid):
    """Una línea legible: Id | categoría | nombre o tipo."""
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    el = doc.GetElement(eid)
    if el is None:
        return "ID {}: (elemento no encontrado)".format(_element_id_int(eid))
    cat = "?"
    try:
        if el.Category is not None:
            cat = el.Category.Name or cat
    except Exception:
        pass
    nombre = ""
    try:
        nombre = el.Name
    except Exception:
        pass
    if not nombre:
        try:
            nombre = type(el).__name__
        except Exception:
            nombre = ""
    return "ID {} | {} | {}".format(_element_id_int(eid), cat, nombre)


def _taskdialog_elementos_unidos_seleccion(doc, wall_ids):
    """
    ``TaskDialog`` con los elementos en *Join Geometry* respecto a cada muro de la selección inicial.
    """
    if doc is None or not wall_ids:
        return
    lineas = []
    try:
        n_muros = len(wall_ids)
        lineas.append("Muros estructurales en la selección: {}.".format(n_muros))
        lineas.append("")
        for wid in wall_ids:
            w = doc.GetElement(wid)
            if w is None:
                lineas.append("— Muro ID ? — (no encontrado)")
                lineas.append("")
                continue
            w_id_str = str(_element_id_int(wid)) if _element_id_int(wid) is not None else "?"
            lineas.append("— Muro ID {} —".format(w_id_str))
            # 1) Solo lo que Revit expone en *Unir geometría* (puede omitir muros en L)
            unidos = []
            try:
                raw_ids = get_joined_element_ids(doc, w)
            except Exception as ex:
                lineas.append(
                    u"  [Unir geometría API] Error al leer: {}".format(str(ex))
                )
                raw_ids = []
            if raw_ids:
                wid_int = _element_id_int(wid)
                try:
                    for eid in raw_ids:
                        if eid is None or eid == ElementId.InvalidElementId:
                            continue
                        if wid_int is not None and _element_id_int(eid) == wid_int:
                            continue
                        unidos.append(eid)
                except Exception:
                    pass
            lineas.append(u"  Unir geometría (GetJoinedElements):")
            if not unidos:
                lineas.append(u"    (ningún otro elemento; el muro puede no estar unido aún a sus vecinos)")
            else:
                for eid in unidos:
                    ln = _describe_elemento_linea(doc, eid)
                    if ln:
                        lineas.append(u"    • " + ln)
            # 2) Mismo criterio que el boceto del nudo: intersección, extremos, bbox, suelos…
            lineas.append(u"  Candidatos boceto nudo (BIMTools — puede incl. muro en L aunque no en API):")
            try:
                boceto_ids = get_otros_element_ids_boceto_nudo(doc, w)
            except Exception as ex2:
                lineas.append(u"    (error: {})".format(str(ex2)))
                boceto_ids = []
            if not boceto_ids:
                lineas.append(u"    (ninguno detectado; revisar LocationCurve y extensión del muro)")
            else:
                set_join = set()
                for j in unidos:
                    ji = _element_id_int(j)
                    if ji is not None:
                        set_join.add(ji)
                for eid in boceto_ids:
                    ln = _describe_elemento_linea(doc, eid)
                    if not ln:
                        continue
                    ei = _element_id_int(eid)
                    suf = (
                        u" [también en Unir geom.]"
                        if ei in set_join
                        else u" [heurística boceto: p. ej. muro en L sin Unir geom. en API]"
                    )
                    lineas.append(u"    • {}{}".format(ln, suf))
            # 3) Misma selección actual (contexto)
            _wid_i = _element_id_int(wid)
            if _wid_i is not None:
                otros_misma_sel = [
                    o
                    for o in wall_ids
                    if o is not None
                    and _element_id_int(o) is not None
                    and _element_id_int(o) != _wid_i
                ]
                if otros_misma_sel:
                    lineas.append(
                        u"  Otros muros en la misma selección (complementario):"
                    )
                    for oid in otros_misma_sel:
                        ln2 = _describe_elemento_linea(doc, oid)
                        if ln2:
                            lineas.append(u"    · " + ln2)
            lineas.append("")
    except Exception as ex:
        lineas = ["Error al listar uniones: {}".format(str(ex))]

    texto = "\n".join(lineas).strip()
    max_len = 12000
    if len(texto) > max_len:
        texto = texto[: max_len - 40] + "\n… (texto truncado)"

    try:
        td = TaskDialog("BIMTools — Elementos unidos (Armado Muros Nodo)")
        td.MainInstruction = (
            u"Por cada muro: (1) Unir geometría según API de Revit; (2) candidatos del boceto nudo (BIMTools)."
        )
        td.MainContent = texto if texto else "(sin contenido)"
        td.CommonButtons = TaskDialogCommonButtons.Ok
        td.DefaultButton = TaskDialogResult.Ok
        td.Show()
    except Exception:
        try:
            forms.alert(
                texto if texto else "(sin datos)",
                title="BIMTools — Elementos unidos",
            )
        except Exception:
            pass


def _resolver_vista_3d(uidoc, controller):
    """
    Vista 3D fiable: con ventana modal, ``ActiveView`` a veces deja de ser la 3D.
    Prioriza ActiveView si es View3D; si no, la vista al instanciar el controlador (antes de ShowDialog).
    """
    try:
        v = uidoc.ActiveView
        if v is not None and isinstance(v, View3D):
            return v
    except Exception:
        pass
    try:
        v = getattr(controller, "_active_view_when_form_opened", None)
        if v is not None and isinstance(v, View3D):
            return v
    except Exception:
        pass
    return None


def _vistas_refuerzo_para_visibilidad(uidoc, controller):
    """
    Vistas a las que aplicar visibilidad: 3D resuelta, vista al abrir el formulario, vista activa.
    Sin duplicados; puede incluir planta, alzado o 3D.
    """
    if uidoc is None:
        return []
    out = []
    seen = set()
    for v in (
        _resolver_vista_3d(uidoc, controller),
        getattr(controller, "_active_view_when_form_opened", None),
        uidoc.ActiveView,
    ):
        if v is None or id(v) in seen:
            continue
        if not isinstance(v, View):
            continue
        seen.add(id(v))
        out.append(v)
    return out


def _selected_bar_id(combo):
    it = combo.SelectedItem
    if it is None:
        return None
    try:
        tid = it.Tag
        if tid is not None and tid != ElementId.InvalidElementId:
            return tid
    except Exception:
        pass
    return None


def _spacing_text_to_internal(doc, text):
    """Convierte texto numérico a unidades internas según unidades de longitud del proyecto."""
    t = str(text).strip().replace(",", ".")
    val = float(t)
    try:
        opts = doc.GetUnits().GetFormatOptions(SpecTypeId.Length)
        uid = opts.GetUnitTypeId()
        return UnitUtils.ConvertToInternalUnits(val, uid)
    except Exception:
        return UnitUtils.ConvertToInternalUnits(val, UnitTypeId.Millimeters)


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
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 1


def desactivar_extremos_rebar_set(rebar, document):
    """
    En un ``Rebar`` con reparto (más de una posición), desactiva la primera y la última
    barra del conjunto (``includeFirstBar`` / ``includeLastBar`` = False) manteniendo la
    regla de layout, separación y longitud de conjunto actuales.

    No aplica a conjunto de una sola posición, a regla *Single* ni si hay menos de 3 posiciones
    (desactivar ambos extremos requiere al menos 3 posiciones iniciales).

    :returns: ``True`` si se aplicó un cambio de layout; ``False`` si no aplica o falló.
    """
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    n = _rebar_cantidad_posiciones(rebar)
    if n < 3:
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
    inc0, inc1 = False, False
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
