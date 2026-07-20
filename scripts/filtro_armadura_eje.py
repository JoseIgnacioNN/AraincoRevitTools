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

    cat_list = _categoria_rebar_ids()
    param_id, storage_type = _resolve_target_param(doc, cat_list, PARAM_NAME)
    if param_id is None:
        return (
            False,
            u"No se encontró el parámetro «{0}» en Structural Rebar.".format(
                PARAM_NAME
            ),
            u"Compruebe que el parámetro compartido exista en el proyecto "
            u"y esté asignado a la categoría Structural Rebar.",
        )

    filter_cache = {}
    applied = []
    errors = []
    filters_created = 0

    t = Transaction(doc, TX_NAME)
    t.Start()
    try:
        for view, eje_valor in targets:
            view_name = _view_display_name(view)
            try:
                filter_elem, created_new = _ensure_filter_for_eje(
                    doc, cat_list, param_id, storage_type, eje_valor, filter_cache
                )
                if created_new:
                    filters_created += 1

                _remove_other_armadura_eje_filters_from_view(view, filter_elem.Id)
                _apply_filter_to_view(view, filter_elem)
                applied.append((view_name, eje_valor))
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

    n_filters = len(filter_cache)
    instruction = (
        u"Filtro aplicado en {0} Building Section(s) "
        u"({1} valor(es) de «{2}»)."
    ).format(len(applied), n_filters, PARAM_NAME)

    content_parts = [u"\n".join(lines)]
    if filters_created:
        content_parts.append(
            u"Filtros de proyecto creados: {0}.".format(filters_created)
        )
    if errors:
        content_parts.append(
            u"Omitidas / error:\n{0}".format(u"\n".join(errors))
        )

    return True, instruction, u"\n\n".join(content_parts)


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
