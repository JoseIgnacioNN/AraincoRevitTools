# -*- coding: utf-8 -*-
"""
Filtro Armadura Eje — crea/aplica filtros de vista por ``Armadura_Eje``.

Revit 2024+ | pyRevit | IronPython 2.7 / 3.4

Flujo:
  1. Recorrer todas las vistas Building Section del documento (no plantillas).
  2. Conservar solo las que tienen ``Armadura_Eje`` con valor.
  3. Por cada valor de eje, crear o actualizar un ``ParameterFilterElement``
     (regla Not Equals sobre Structural Rebar).
  4. Aplicar el filtro a cada vista candidata con visibilidad apagada y quitar
     de esa vista otros filtros ``Armadura_Eje …`` previos.
  5. Si la plantilla de vista incluye «V/G Overrides Filters», desmarcarlo
     para poder aplicar un filtro distinto por vista.
  6. Heredar en cada vista los filtros que ya tenía la plantilla (visibilidad,
     activación y overrides) y, además, aplicar el filtro ``Armadura_Eje``.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    ElementParameterFilter,
    FilteredElementCollector,
    ParameterFilterElement,
    ParameterFilterRuleFactory,
    ParameterFilterUtilities,
    OverrideGraphicSettings,
    SharedParameterElement,
    StorageType,
    Transaction,
    View,
    ViewFamily,
    ViewType,
)
from Autodesk.Revit.UI import TaskDialog

try:
    from Autodesk.Revit.DB import SpecTypeId
except Exception:
    SpecTypeId = None

TOOL_TITLE = u"Arainco: Filtro Armadura Eje"
PARAM_NAME = u"Armadura_Eje"
TX_NAME = u"Arainco: Filtro Armadura Eje"

_VISTA_DETALLE_MARKERS = (
    u"detail",
    u"detalle",
    u"callout",
    u"recuadro",
    u"detailed",
)
_BUILDING_SECTION_MARKERS = (
    u"building section",
    u"sección de edificio",
    u"seccion de edificio",
)


def _as_unicode(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)


def _show_message(uiapp, instruction, content=u""):
    """Diálogo WPF BIMTools; respaldo TaskDialog."""
    try:
        from bimtools_instruction_dialog import show_message_dialog
        from revit_wpf_window_position import revit_main_hwnd

        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            TOOL_TITLE,
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
            msg = instruction + u"\n\n" + content
        TaskDialog.Show(TOOL_TITLE, msg)
    except Exception:
        print(instruction)
        if content:
            print(content)


def _canon_key(text):
    return _as_unicode(text).strip().lower()


def _view_type_suffix(view):
    if view is None:
        return u""
    try:
        vt = view.ViewType
        try:
            s = vt.ToString()
        except Exception:
            s = str(vt)
    except Exception:
        return u""
    s = (s or u"").strip()
    if u"." in s:
        s = s.split(u".")[-1]
    return s


def _enum_equals(valor, enum_obj):
    if valor is None or enum_obj is None:
        return False
    try:
        if valor == enum_obj:
            return True
    except Exception:
        pass
    try:
        if int(valor) == int(enum_obj):
            return True
    except Exception:
        pass
    try:
        a = _canon_key(valor.ToString() if hasattr(valor, u"ToString") else valor)
        b = _canon_key(
            enum_obj.ToString() if hasattr(enum_obj, u"ToString") else enum_obj
        )
        if a and b and a.split(u".")[-1] == b.split(u".")[-1]:
            return True
    except Exception:
        pass
    return False


def _parametro_texto(element, *builtins):
    if element is None:
        return u""
    for bip in builtins:
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            s = p.AsValueString()
            if s:
                return _as_unicode(s).strip()
        except Exception:
            pass
        try:
            p = element.get_Parameter(bip)
            if p is None:
                continue
            s = p.AsString()
            if s:
                return _as_unicode(s).strip()
        except Exception:
            pass
    return u""


def _view_family_type_element(view):
    if view is None:
        return None
    try:
        doc = view.Document
    except Exception:
        doc = None
    if doc is None:
        return None
    try:
        tid = view.GetTypeId()
        if tid is not None and tid != ElementId.InvalidElementId:
            vft = doc.GetElement(tid)
            if vft is not None and hasattr(vft, u"ViewFamily"):
                return vft
    except Exception:
        pass
    return None


def _view_family_type_name(view):
    vft = _view_family_type_element(view)
    if vft is not None:
        try:
            nm = vft.Name or u""
            if nm:
                return _as_unicode(nm)
        except Exception:
            pass
    try:
        raw = _parametro_texto(
            view,
            BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
            BuiltInParameter.SYMBOL_NAME_PARAM,
        )
        if u":" in raw:
            raw = raw.split(u":", 1)[1].strip()
        if raw:
            return raw
    except Exception:
        pass
    return u""


def _nombre_es_building_section(name):
    n = _canon_key(name or u"")
    if not n:
        return False
    for bad in _VISTA_DETALLE_MARKERS:
        if bad in n:
            return False
    for ok in _BUILDING_SECTION_MARKERS:
        if ok in n:
            return True
    return False


def _vft_es_familia_section(vft):
    if vft is None:
        return False
    try:
        return _enum_equals(vft.ViewFamily, ViewFamily.Section)
    except Exception:
        pass
    try:
        vf = vft.ViewFamily
        s = vf.ToString() if hasattr(vf, u"ToString") else str(vf)
        return u"Section" in (s or u"")
    except Exception:
        return False


def es_vista_building_section(view):
    """True si la vista es una sección de edificio (Building Section)."""
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    if _view_type_suffix(view) == u"Detail":
        return False
    try:
        if _enum_equals(view.ViewType, ViewType.Detail):
            return False
    except Exception:
        pass
    if _view_type_suffix(view) != u"Section":
        try:
            if not _enum_equals(view.ViewType, ViewType.Section):
                return False
        except Exception:
            return False

    vft = _view_family_type_element(view)
    if vft is not None:
        if not _vft_es_familia_section(vft):
            return False
        try:
            if _nombre_es_building_section(vft.Name):
                return True
        except Exception:
            pass

    nombre_tipo = _view_family_type_name(view)
    if _nombre_es_building_section(nombre_tipo):
        return True

    if vft is not None and _vft_es_familia_section(vft):
        n = _canon_key(nombre_tipo or u"")
        if n:
            for bad in _VISTA_DETALLE_MARKERS:
                if bad in n:
                    return False
        return True

    return False


def _param_value_as_text(param):
    """Extrae texto usable de un parámetro de vista (string / int / value string)."""
    if param is None:
        return None
    try:
        storage = param.StorageType
    except Exception:
        storage = None

    if storage == StorageType.String:
        try:
            text = param.AsString()
        except Exception:
            text = None
        if text is None:
            try:
                text = param.AsValueString()
            except Exception:
                text = None
    elif storage == StorageType.Integer:
        try:
            if param.HasValue:
                text = _as_unicode(param.AsInteger())
            else:
                text = None
        except Exception:
            text = None
        if not text:
            try:
                text = param.AsValueString()
            except Exception:
                text = None
    else:
        try:
            text = param.AsValueString()
        except Exception:
            text = None
        if text is None:
            try:
                text = param.AsString()
            except Exception:
                text = None

    if text is None:
        return None
    text = _as_unicode(text).strip()
    return text or None


def _leer_armadura_eje_vista(view):
    if view is None:
        return None
    try:
        param = view.LookupParameter(PARAM_NAME)
    except Exception:
        param = None
    return _param_value_as_text(param)


def _vista_admite_filtros(view):
    if view is None or not isinstance(view, View):
        return False
    try:
        if view.AreGraphicsOverridesAllowed():
            return True
    except Exception:
        pass
    try:
        view.GetFilters()
        return True
    except Exception:
        return False


def _view_display_name(view):
    try:
        return _as_unicode(view.Name)
    except Exception:
        return u"(sin nombre)"


def _categoria_rebar_ids():
    cat_list = List[ElementId]()
    cat_list.Add(ElementId(BuiltInCategory.OST_Rebar))
    return cat_list


def _find_shared_param_id(doc, param_name):
    try:
        for spe in FilteredElementCollector(doc).OfClass(SharedParameterElement):
            try:
                if spe and spe.Name == param_name:
                    return spe.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def _find_filterable_param_id(doc, cat_list, param_name):
    try:
        for p_id in ParameterFilterUtilities.GetFilterableParametersInCommon(
            doc, cat_list
        ):
            p_elem = doc.GetElement(p_id)
            if p_elem is not None and getattr(p_elem, "Name", None) == param_name:
                return p_id
    except Exception:
        pass
    return None


def _find_param_id_from_sample_rebar(doc, param_name):
    rebar = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rebar)
        .WhereElementIsNotElementType()
        .FirstElement()
    )
    if rebar is None:
        return None, None
    try:
        param = rebar.LookupParameter(param_name)
    except Exception:
        param = None
    if param is None:
        return None, None
    return param.Id, param.StorageType


def _infer_storage_from_param_element(doc, param_id):
    try:
        p_elem = doc.GetElement(param_id)
    except Exception:
        p_elem = None
    if p_elem is None:
        return None

    definition = None
    try:
        definition = p_elem.GetDefinition()
    except Exception:
        definition = None
    if definition is None:
        return None

    if SpecTypeId is not None:
        try:
            data_type = definition.GetDataType()
            if data_type == SpecTypeId.String.Text:
                return StorageType.String
            if data_type == SpecTypeId.Int.Integer:
                return StorageType.Integer
            if data_type == SpecTypeId.Number:
                return StorageType.Double
            if data_type == SpecTypeId.Boolean.YesNo:
                return StorageType.Integer
        except Exception:
            pass

    try:
        ptype = definition.ParameterType
        name = _as_unicode(ptype)
        if u"Text" in name or u"String" in name:
            return StorageType.String
        if u"Integer" in name or u"YesNo" in name:
            return StorageType.Integer
        if u"Number" in name or u"Length" in name:
            return StorageType.Double
    except Exception:
        pass
    return None


def _resolve_target_param(doc, cat_list, param_name):
    """
    Resuelve ElementId y StorageType del parámetro objetivo.

    Prioridad: SharedParameterElement → parámetros filtrables → muestra Rebar.
    """
    storage = None
    param_id = _find_shared_param_id(doc, param_name)
    if param_id is None:
        param_id = _find_filterable_param_id(doc, cat_list, param_name)

    sample_id, sample_storage = _find_param_id_from_sample_rebar(doc, param_name)
    if param_id is None:
        param_id = sample_id
    if sample_storage is not None:
        storage = sample_storage

    if storage is None and param_id is not None:
        storage = _infer_storage_from_param_element(doc, param_id)

    if storage is None:
        storage = StorageType.String

    return param_id, storage


def _create_not_equals_rule(param_id, storage_type, value_text):
    if storage_type == StorageType.String:
        return ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, value_text)

    if storage_type == StorageType.Integer:
        try:
            val_int = int(value_text)
        except (TypeError, ValueError):
            raise ValueError(
                u"El parámetro espera un entero pero la vista tiene '{0}'.".format(
                    value_text
                )
            )
        return ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, val_int)

    if storage_type == StorageType.Double:
        try:
            val_dbl = float(value_text.replace(u",", u"."))
        except (TypeError, ValueError, AttributeError):
            raise ValueError(
                u"El parámetro espera un número pero la vista tiene '{0}'.".format(
                    value_text
                )
            )
        return ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, val_dbl)

    raise ValueError(
        u"StorageType no soportado para este filtro (debe ser Texto, Entero o Número)."
    )


def _filter_name_prefix():
    return PARAM_NAME + u" "


def _is_managed_filter_name(name):
    if not name:
        return False
    return _as_unicode(name).startswith(_filter_name_prefix())


def _find_parameter_filter_by_name(doc, filter_name):
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        try:
            if f.Name == filter_name:
                return f
        except Exception:
            continue
    return None


def _remove_other_armadura_eje_filters_from_view(view, keep_filter_id):
    removed = 0
    try:
        current_ids = list(view.GetFilters())
    except Exception:
        return 0

    doc = view.Document
    keep_int = None
    try:
        keep_int = int(keep_filter_id.IntegerValue)
    except Exception:
        pass

    for fid in current_ids:
        try:
            fid_int = int(fid.IntegerValue)
        except Exception:
            fid_int = None
        if keep_int is not None and fid_int == keep_int:
            continue
        try:
            fe = doc.GetElement(fid)
        except Exception:
            fe = None
        if fe is None:
            continue
        try:
            fname = fe.Name
        except Exception:
            continue
        if not _is_managed_filter_name(fname):
            continue
        try:
            view.RemoveFilter(fid)
            removed += 1
        except Exception:
            pass
    return removed


def _apply_filter_to_view(view, filter_elem):
    fid = filter_elem.Id
    already = False
    try:
        already = view.GetFilters().Contains(fid)
    except Exception:
        try:
            already = fid in list(view.GetFilters())
        except Exception:
            already = False

    if not already:
        view.AddFilter(fid)

    try:
        view.SetIsFilterEnabled(fid, True)
    except Exception:
        pass
    view.SetFilterVisibility(fid, False)
    return already


def _eid_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    try:
        return int(eid.Value)
    except Exception:
        return None


def _is_invalid_eid(eid):
    if eid is None:
        return True
    try:
        if eid == ElementId.InvalidElementId:
            return True
    except Exception:
        pass
    n = _eid_int(eid)
    return n is None or n == -1


def _vg_filters_bip_id():
    bip = getattr(BuiltInParameter, u"VIS_GRAPHICS_FILTERS", None)
    if bip is None:
        return None
    try:
        return ElementId(bip)
    except Exception:
        pass
    try:
        return ElementId(int(bip))
    except Exception:
        return None


def _param_name_es_vg_filters(name):
    n = _canon_key(name)
    if not n:
        return False
    tiene_filtro = (u"filter" in n) or (u"filtro" in n)
    if not tiene_filtro:
        return False
    return (
        (u"v/g" in n)
        or (u"vg " in n)
        or n.startswith(u"vg")
        or (u"overrides" in n)
        or (u"visib" in n)
        or (u"gráfico" in n)
        or (u"grafico" in n)
        or (u"reemplazo" in n)
        or (u"sustituc" in n)
    )


def _resolve_vg_filters_param_id(template):
    """ElementId del parámetro de plantilla «V/G Overrides Filters»."""
    known = _vg_filters_bip_id()
    known_int = _eid_int(known) if known is not None else None

    param_ids = []
    try:
        param_ids = list(template.GetTemplateParameterIds())
    except Exception:
        param_ids = []

    if known_int is not None:
        for eid in param_ids:
            if _eid_int(eid) == known_int:
                return eid
        return known

    by_int = {}
    try:
        for p in template.Parameters:
            try:
                pid = p.Id
            except Exception:
                continue
            nint = _eid_int(pid)
            if nint is None:
                continue
            try:
                pname = p.Definition.Name
            except Exception:
                pname = u""
            by_int[nint] = (pid, pname)
    except Exception:
        pass

    for eid in param_ids:
        nint = _eid_int(eid)
        hit = by_int.get(nint)
        if hit is None:
            continue
        _pid, pname = hit
        if _param_name_es_vg_filters(pname):
            return eid
    return None


def _get_view_template(view):
    if view is None:
        return None
    try:
        tid = view.ViewTemplateId
    except Exception:
        return None
    if _is_invalid_eid(tid):
        return None
    try:
        el = view.Document.GetElement(tid)
    except Exception:
        el = None
    if el is None or not isinstance(el, View):
        return None
    try:
        if not el.IsTemplate:
            return None
    except Exception:
        pass
    return el


def _non_controlled_ints(template):
    out = set()
    try:
        for eid in template.GetNonControlledTemplateParameterIds():
            n = _eid_int(eid)
            if n is not None:
                out.add(n)
    except Exception:
        pass
    return out


def _template_parameter_ints(template):
    out = set()
    try:
        for eid in template.GetTemplateParameterIds():
            n = _eid_int(eid)
            if n is not None:
                out.add(n)
    except Exception:
        pass
    return out


def plantilla_controla_filtros_vg(template):
    """
    True si la plantilla incluye (rige) «V/G Overrides Filters».

    Un parámetro de plantilla está controlado si aparece en
    ``GetTemplateParameterIds`` y no en ``GetNonControlledTemplateParameterIds``.
    """
    if template is None:
        return False
    vg_id = _resolve_vg_filters_param_id(template)
    if vg_id is None:
        return False
    vg_int = _eid_int(vg_id)
    if vg_int is None:
        return False
    if vg_int in _non_controlled_ints(template):
        return False
    tmpl_ids = _template_parameter_ints(template)
    if tmpl_ids:
        return vg_int in tmpl_ids
    return True


def vista_filtros_controlados_por_plantilla(view):
    """True si la vista no puede recibir filtros propios por su plantilla."""
    return plantilla_controla_filtros_vg(_get_view_template(view))


def desbloquear_filtros_en_plantilla(template):
    """
    Desmarca «V/G Overrides Filters» en la plantilla para permitir
    filtros distintos por vista.

    Debe llamarse dentro de una ``Transaction`` abierta.

    Returns:
        (changed: bool, error: unicode or None)
    """
    if template is None:
        return False, None
    if not plantilla_controla_filtros_vg(template):
        return False, None

    vg_id = _resolve_vg_filters_param_id(template)
    if vg_id is None:
        return False, (
            u"No se encontró el parámetro «V/G Overrides Filters» "
            u"en la plantilla de vista."
        )

    vg_int = _eid_int(vg_id)
    ids = List[ElementId]()
    seen = set()
    try:
        for eid in template.GetNonControlledTemplateParameterIds():
            n = _eid_int(eid)
            if n is None or n in seen:
                continue
            seen.add(n)
            ids.Add(eid)
    except Exception as ex:
        return False, _as_unicode(ex)

    if vg_int not in seen:
        ids.Add(vg_id)

    try:
        template.SetNonControlledTemplateParameterIds(ids)
    except Exception:
        try:
            from System.Collections.Generic import HashSet

            hs = HashSet[ElementId]()
            for eid in ids:
                hs.Add(eid)
            template.SetNonControlledTemplateParameterIds(hs)
        except Exception as ex:
            return False, _as_unicode(ex)

    if plantilla_controla_filtros_vg(template):
        return False, (
            u"La plantilla sigue rigiendo los filtros tras intentar "
            u"desmarcar «V/G Overrides Filters»."
        )
    return True, None


def ensure_vista_permite_filtros_propios(view):
    """
    Si la plantilla de la vista rige los filtros, la edita para no incluirlos.

    Returns:
        dict: ok, unlocked, template_name, error
    """
    result = {
        u"ok": True,
        u"unlocked": False,
        u"template_name": None,
        u"error": None,
    }
    template = _get_view_template(view)
    if template is None:
        return result
    try:
        result[u"template_name"] = _as_unicode(template.Name)
    except Exception:
        result[u"template_name"] = u"(plantilla)"

    if not plantilla_controla_filtros_vg(template):
        return result

    changed, err = desbloquear_filtros_en_plantilla(template)
    if err:
        result[u"ok"] = False
        result[u"error"] = err
        return result
    result[u"unlocked"] = bool(changed)
    return result


def _view_has_filter(view, fid):
    want = _eid_int(fid)
    if view is None or want is None:
        return False
    try:
        for existing in view.GetFilters():
            if _eid_int(existing) == want:
                return True
    except Exception:
        pass
    return False


def _filter_ids_ordered(view):
    if view is None:
        return []
    try:
        return list(view.GetOrderedFilters())
    except Exception:
        pass
    try:
        return list(view.GetFilters())
    except Exception:
        return []


def _copy_filter_overrides(source_view, fid):
    if source_view is None or _is_invalid_eid(fid):
        return None
    try:
        ovr = source_view.GetFilterOverrides(fid)
    except Exception:
        return None
    if ovr is None:
        return None
    try:
        return OverrideGraphicSettings(ovr)
    except Exception:
        return ovr


def snapshot_filters_from_view(source_view):
    """
    Copia estado de filtros (orden, visibilidad, activación, overrides).

    No incluye filtros gestionados ``Armadura_Eje …``: esos se aplican
    por vista según el valor del eje.
    """
    entries = []
    if source_view is None:
        return entries
    try:
        doc = source_view.Document
    except Exception:
        doc = None
    for fid in _filter_ids_ordered(source_view):
        if _is_invalid_eid(fid):
            continue
        fname = u""
        if doc is not None:
            try:
                fe = doc.GetElement(fid)
                if fe is not None:
                    fname = _as_unicode(fe.Name)
            except Exception:
                fname = u""
        if _is_managed_filter_name(fname):
            continue
        rec = {
            u"id": fid,
            u"visible": True,
            u"enabled": True,
            u"overrides": None,
        }
        try:
            rec[u"visible"] = bool(source_view.GetFilterVisibility(fid))
        except Exception:
            pass
        try:
            rec[u"enabled"] = bool(source_view.GetIsFilterEnabled(fid))
        except Exception:
            pass
        rec[u"overrides"] = _copy_filter_overrides(source_view, fid)
        entries.append(rec)
    return entries


def inherit_filter_snapshot_to_view(view, entries):
    """
    Aplica en ``view`` los filtros heredados de la plantilla.

    Returns:
        int: filtros añadidos o actualizados.
    """
    applied = 0
    if view is None:
        return 0
    for rec in entries or []:
        fid = rec.get(u"id")
        if _is_invalid_eid(fid):
            continue
        try:
            if not _view_has_filter(view, fid):
                view.AddFilter(fid)
            try:
                view.SetIsFilterEnabled(fid, bool(rec.get(u"enabled", True)))
            except Exception:
                pass
            try:
                view.SetFilterVisibility(fid, bool(rec.get(u"visible", True)))
            except Exception:
                pass
            ovr = rec.get(u"overrides")
            if ovr is not None:
                try:
                    view.SetFilterOverrides(fid, OverrideGraphicSettings(ovr))
                except Exception:
                    try:
                        view.SetFilterOverrides(fid, ovr)
                    except Exception:
                        pass
            applied += 1
        except Exception:
            continue
    return applied


def _iter_views_using_template(doc, template):
    if doc is None or template is None:
        return
    tid_int = _eid_int(template.Id)
    if tid_int is None:
        return
    try:
        views = FilteredElementCollector(doc).OfClass(View).ToElements()
    except Exception:
        return
    for view in views:
        if view is None:
            continue
        try:
            if view.IsTemplate:
                continue
        except Exception:
            continue
        try:
            if _eid_int(view.ViewTemplateId) == tid_int:
                yield view
        except Exception:
            continue


def inherit_template_filters_to_views(doc, template, snapshot, views=None):
    """
    Replica los filtros de la plantilla en las vistas indicadas
    (o en todas las que usan esa plantilla).
    """
    n = 0
    targets = views
    if targets is None:
        targets = list(_iter_views_using_template(doc, template))
    for view in targets or []:
        n += inherit_filter_snapshot_to_view(view, snapshot)
    return n


def _looks_like_template_filter_lock(ex):
    s = _canon_key(_as_unicode(ex))
    if not s:
        return False
    return (
        (u"template" in s)
        or (u"plantilla" in s)
        or (u"controlled" in s)
        or (u"cannot be modified" in s)
        or (u"no se puede modificar" in s)
        or (u"no se pueden modificar" in s)
    )


def prepare_armadura_eje_filter_context(doc):
    """
    Resuelve parámetro y categorías para crear filtros (solo lectura).

    Returns:
        (ctx, instruction_error, content_error)
        ctx es dict o None.
    """
    cat_list = _categoria_rebar_ids()
    param_id, storage_type = _resolve_target_param(doc, cat_list, PARAM_NAME)
    if param_id is None:
        return (
            None,
            u"No se encontró el parámetro «{0}» en Structural Rebar.".format(
                PARAM_NAME
            ),
            u"Compruebe que el parámetro compartido exista en el proyecto "
            u"y esté asignado a la categoría Structural Rebar.",
        )
    ctx = {
        u"cat_list": cat_list,
        u"param_id": param_id,
        u"storage_type": storage_type,
        u"cache": {},
        u"template_names_unlocked": [],
        u"templates_inherited": set(),
    }
    return ctx, None, None


def apply_armadura_eje_filter_to_view(doc, view, eje_valor, ctx):
    """
    Crea/actualiza el ``ParameterFilterElement`` del eje y lo aplica a la vista
    (visibilidad apagada). Hereda los filtros de la plantilla. Si la plantilla
    rige los filtros, los desmarca para permitir el filtro por eje.

    Debe llamarse dentro de una ``Transaction`` abierta.

    Returns:
        (ok: bool, info: dict)
        info: created_new, template_unlocked, template_name,
        inherited_count, error
    """
    info = {
        u"created_new": False,
        u"template_unlocked": False,
        u"template_name": None,
        u"inherited_count": 0,
        u"error": None,
    }
    if doc is None or view is None or ctx is None:
        info[u"error"] = u"Datos incompletos para el filtro."
        return False, info

    eje_text = _as_unicode(eje_valor).strip()
    if not eje_text:
        info[u"error"] = u"Sin valor de «{0}».".format(PARAM_NAME)
        return False, info

    try:
        if not bool(view.AreGraphicsOverridesAllowed()):
            info[u"error"] = u"La vista no admite filtros de visibilidad."
            return False, info
    except Exception:
        pass

    template = _get_view_template(view)
    snapshot = []
    if template is not None:
        snapshot = snapshot_filters_from_view(template)

    unlock = ensure_vista_permite_filtros_propios(view)
    if unlock.get(u"unlocked"):
        info[u"template_unlocked"] = True
        info[u"template_name"] = unlock.get(u"template_name")
        name = info[u"template_name"]
        names = ctx.get(u"template_names_unlocked")
        if name and isinstance(names, list) and name not in names:
            names.append(name)
    if not unlock.get(u"ok"):
        info[u"template_name"] = unlock.get(u"template_name")
        info[u"error"] = unlock.get(u"error") or (
            u"La plantilla de vista impide crear filtros por vista."
        )
        return False, info

    def _inherit_template_filters():
        inherited_on = ctx.get(u"templates_inherited")
        if not isinstance(inherited_on, set):
            inherited_on = set()
            ctx[u"templates_inherited"] = inherited_on
        tid = _eid_int(template.Id) if template is not None else None
        if (
            template is not None
            and tid is not None
            and info.get(u"template_unlocked")
            and tid not in inherited_on
        ):
            inherit_template_filters_to_views(doc, template, snapshot)
            inherited_on.add(tid)
        info[u"inherited_count"] = inherit_filter_snapshot_to_view(view, snapshot)

    def _do_apply():
        _inherit_template_filters()
        filter_elem, created_new = _ensure_filter_for_eje(
            doc,
            ctx[u"cat_list"],
            ctx[u"param_id"],
            ctx[u"storage_type"],
            eje_text,
            ctx[u"cache"],
        )
        _remove_other_armadura_eje_filters_from_view(view, filter_elem.Id)
        _apply_filter_to_view(view, filter_elem)
        return created_new

    try:
        info[u"created_new"] = _do_apply()
    except Exception as ex:
        if not _looks_like_template_filter_lock(ex):
            info[u"error"] = _as_unicode(ex)
            return False, info
        unlock2 = ensure_vista_permite_filtros_propios(view)
        if unlock2.get(u"unlocked"):
            info[u"template_unlocked"] = True
            info[u"template_name"] = unlock2.get(u"template_name")
            name = info[u"template_name"]
            names = ctx.get(u"template_names_unlocked")
            if name and isinstance(names, list) and name not in names:
                names.append(name)
            if not snapshot and template is None:
                template = _get_view_template(view)
                snapshot[:] = snapshot_filters_from_view(template) if template else []
        if not unlock2.get(u"ok"):
            info[u"error"] = unlock2.get(u"error") or _as_unicode(ex)
            return False, info
        try:
            info[u"created_new"] = _do_apply()
        except Exception as ex2:
            info[u"error"] = _as_unicode(ex2)
            return False, info

    return True, info


def collect_building_sections_with_eje(doc):
    """
    Building Sections no plantilla con ``Armadura_Eje`` no vacío.

    Returns:
        list of (view, eje_valor)
    """
    result = []
    try:
        views = FilteredElementCollector(doc).OfClass(View).ToElements()
    except Exception:
        return result

    for view in views:
        if view is None:
            continue
        try:
            if view.IsTemplate:
                continue
        except Exception:
            pass
        if not es_vista_building_section(view):
            continue
        if not _vista_admite_filtros(view):
            continue
        eje = _leer_armadura_eje_vista(view)
        if not eje:
            continue
        result.append((view, eje))

    try:
        result.sort(key=lambda item: (_as_unicode(item[1]), _view_display_name(item[0])))
    except Exception:
        pass
    return result


def _ensure_filter_for_eje(doc, cat_list, param_id, storage_type, eje_valor, cache):
    """
    Obtiene o crea el ParameterFilterElement para un valor de eje.

    ``cache``: dict eje_valor -> (filter_elem, created_new: bool)
    """
    if eje_valor in cache:
        # Ya resuelto en esta corrida: no contar de nuevo como «creado».
        return cache[eje_valor][0], False

    filter_name = u"{0} {1}".format(PARAM_NAME, eje_valor)
    rule = _create_not_equals_rule(param_id, storage_type, eje_valor)
    elem_filter = ElementParameterFilter(rule)

    filter_elem = _find_parameter_filter_by_name(doc, filter_name)
    created_new = False
    if filter_elem is None:
        filter_elem = ParameterFilterElement.Create(
            doc, filter_name, cat_list, elem_filter
        )
        created_new = True
    else:
        filter_elem.SetElementFilter(elem_filter)
        filter_elem.SetCategories(cat_list)

    cache[eje_valor] = (filter_elem, created_new)
    return filter_elem, created_new


def apply_filters_to_building_sections(doc):
    """
    Procesa todas las Building Section con ``Armadura_Eje``.

    Returns:
        (ok: bool, instruction: unicode, content: unicode)
    """
    targets = collect_building_sections_with_eje(doc)
    if not targets:
        return (
            False,
            u"No hay vistas Building Section con «{0}» definido.".format(PARAM_NAME),
            u"Asigne un valor a «{0}» en las secciones de edificio que "
            u"deban filtrar armadura por eje.".format(PARAM_NAME),
        )

    ctx, err_i, err_c = prepare_armadura_eje_filter_context(doc)
    if ctx is None:
        return False, err_i or u"No se pudo preparar el filtro.", err_c or u""

    applied = []
    errors = []
    filters_created = 0

    t = Transaction(doc, TX_NAME)
    t.Start()
    try:
        for view, eje_valor in targets:
            view_name = _view_display_name(view)
            try:
                ok_f, finfo = apply_armadura_eje_filter_to_view(
                    doc, view, eje_valor, ctx
                )
                if ok_f:
                    if (finfo or {}).get(u"created_new"):
                        filters_created += 1
                    applied.append((view_name, eje_valor))
                else:
                    errors.append(
                        u"{0}: {1}".format(
                            view_name,
                            (finfo or {}).get(u"error")
                            or u"no se pudo aplicar el filtro",
                        )
                    )
            except Exception as ex:
                errors.append(u"{0}: {1}".format(view_name, _as_unicode(ex)))

        if not applied:
            t.RollBack()
            detail = u"\n".join(errors) if errors else u""
            return (
                False,
                u"No se pudo aplicar el filtro a ninguna vista.",
                detail,
            )

        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        return (
            False,
            u"No se pudo crear o aplicar los filtros.",
            _as_unicode(ex),
        )

    # Resumen por eje
    by_eje = {}
    for view_name, eje_valor in applied:
        by_eje.setdefault(eje_valor, []).append(view_name)

    lines = []
    for eje_valor in sorted(by_eje.keys(), key=_as_unicode):
        names = by_eje[eje_valor]
        lines.append(
            u"• Eje {0}: {1} vista(s) — {2}".format(
                eje_valor,
                len(names),
                u", ".join(names),
            )
        )

    n_filters = len(ctx.get(u"cache") or {})
    instruction = (
        u"Filtro aplicado en {0} Building Section(s) "
        u"({1} valor(es) de «{2}»)."
    ).format(len(applied), n_filters, PARAM_NAME)

    content_parts = [u"\n".join(lines)]
    if filters_created:
        content_parts.append(
            u"Filtros de proyecto creados: {0}.".format(filters_created)
        )
    unlocked = ctx.get(u"template_names_unlocked") or []
    if unlocked:
        content_parts.append(
            u"Plantilla(s) editada(s) para permitir filtros por vista: {0}.".format(
                u", ".join(unlocked)
            )
        )
    if errors:
        content_parts.append(
            u"Omitidas / error:\n{0}".format(u"\n".join(errors))
        )

    return True, instruction, u"\n\n".join(content_parts)


def _name_contains_armadura_eje(name):
    token = _as_unicode(PARAM_NAME).strip().lower()
    if not token:
        return False
    return token in _as_unicode(name).strip().lower()


def _host_allows_filter_edit(host):
    if host is None:
        return False
    try:
        if not bool(host.AreGraphicsOverridesAllowed()):
            return False
    except Exception:
        pass
    return True


def _read_filter_enabled(host, fid):
    try:
        return bool(host.GetIsFilterEnabled(fid))
    except Exception:
        return True


def _write_filter_enabled(host, fid, enabled):
    if host is None or _is_invalid_eid(fid):
        return False
    try:
        host.SetIsFilterEnabled(fid, bool(enabled))
        return True
    except Exception:
        return False


def _resolve_filter_edit_host(view):
    """Vista activa o su plantilla si esta rige los filtros V/G."""
    if vista_filtros_controlados_por_plantilla(view):
        template = _get_view_template(view)
        if template is not None and _host_allows_filter_edit(template):
            return template
    if view is not None and _host_allows_filter_edit(view):
        return view
    return None


def collect_applied_armadura_eje_filters(view):
    """
    Filtros aplicados en ``view`` cuyo nombre contiene ``Armadura_Eje``.

    Returns:
        list[dict]: fid, name, host, was_enabled
    """
    found = []
    if view is None:
        return found
    try:
        doc = view.Document
    except Exception:
        return found

    for fid in _filter_ids_ordered(view):
        if _is_invalid_eid(fid):
            continue
        try:
            fe = doc.GetElement(fid)
        except Exception:
            fe = None
        if fe is None:
            continue
        try:
            fname = _as_unicode(fe.Name)
        except Exception:
            continue
        if not _name_contains_armadura_eje(fname):
            continue
        host = _resolve_filter_edit_host(view)
        if host is None:
            continue
        found.append(
            {
                u"fid": fid,
                u"name": fname,
                u"host": host,
                u"was_enabled": _read_filter_enabled(host, fid),
            }
        )
    return found


def _commit_filter_enabled_changes(doc, items, enable, tx_name, uidoc=None):
    """
    ``items``: dicts con fid, host, was_enabled.
    Si ``enable`` es False, desactiva los que estaban activos.
    Si ``enable`` es True, reactiva solo los que ``was_enabled`` era True.
    """
    pending = []
    for rec in items or []:
        fid = rec.get(u"fid")
        host = rec.get(u"host")
        if host is None or _is_invalid_eid(fid):
            continue
        if enable:
            if not bool(rec.get(u"was_enabled", True)):
                continue
            want = True
        else:
            if not bool(rec.get(u"was_enabled", True)):
                continue
            want = False
        if _read_filter_enabled(host, fid) == want:
            continue
        pending.append((host, fid, want))

    if not pending or doc is None:
        return 0

    tx = Transaction(doc, tx_name)
    tx.Start()
    changed = 0
    try:
        for host, fid, want in pending:
            if _write_filter_enabled(host, fid, want):
                changed += 1
        if changed:
            tx.Commit()
        else:
            tx.RollBack()
    except Exception:
        try:
            if tx.HasStarted():
                tx.RollBack()
        except Exception:
            pass
        raise

    if changed and uidoc is not None:
        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass
    return changed


def suspend_armadura_eje_filters_in_active_view(uidoc):
    """
    Desactiva temporalmente los filtros aplicados cuyo nombre contiene
    ``Armadura_Eje`` en la vista activa.

    Returns:
        dict: token para ``restore_armadura_eje_filters`` (puede estar vacío).
    """
    token = {u"uidoc": uidoc, u"doc": None, u"items": []}
    if uidoc is None:
        return token
    try:
        view = uidoc.ActiveView
    except Exception:
        view = None
    if view is None:
        return token
    try:
        token[u"doc"] = view.Document
    except Exception:
        token[u"doc"] = None
    items = collect_applied_armadura_eje_filters(view)
    token[u"items"] = items
    if not items:
        return token
    _commit_filter_enabled_changes(
        token[u"doc"],
        items,
        enable=False,
        tx_name=u"Arainco: Desactivar filtro Armadura_Eje",
        uidoc=uidoc,
    )
    return token


def restore_armadura_eje_filters(token):
    """Reactiva los filtros ``Armadura_Eje`` que estaban activos al suspender."""
    if not token:
        return 0
    items = token.get(u"items") or []
    if not items:
        return 0
    return _commit_filter_enabled_changes(
        token.get(u"doc"),
        items,
        enable=True,
        tx_name=u"Arainco: Activar filtro Armadura_Eje",
        uidoc=token.get(u"uidoc"),
    )


def run(revit_app):
    """Punto de entrada pyRevit: ``run(__revit__)``."""
    uidoc = None
    try:
        uidoc = revit_app.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _show_message(revit_app, u"No hay documento activo.")
        return

    doc = uidoc.Document
    if doc is None:
        _show_message(revit_app, u"No hay documento activo.")
        return

    ok, instruction, content = apply_filters_to_building_sections(doc)
    _show_message(revit_app, instruction, content)
