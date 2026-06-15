# -*- coding: utf-8 -*-
"""Parámetros compartidos en Rebar creados por Armado muros Cabezal."""

import clr

clr.AddReference("System")
from System import AppDomain

from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB.Structure import Rebar

ARMADURA_ARAINCO_PARAM = u"Armadura_Arainco"
ARMADURA_UBICACION_PARAM = u"Armadura_Ubicacion"
ARMADURA_MALLA_TIPO_PARAM = u"Armadura_Malla_Tipo"
ARMADURA_MALLA_ORIENTACION_PARAM = u"Armadura_Malla_Orientacion"
ARMADURA_CONJUNTO_GUID_PARAM = u"Armadura_Conjunto_GUID"
ARMADURA_MALLA_TIPO_DM = u"D.M."
ARMADURA_MALLA_ORIENT_V = u"V."
ARMADURA_MALLA_ORIENT_H = u"H."
_APPDOMAIN_CONJUNTO_GUID_KEY = u"Arainco_ArmadoMuros_Conjunto_GUID"


def _norm_param_def_name(name):
    if name is None:
        return u""
    try:
        t = unicode(name).replace(u"\u00A0", u" ").strip()
    except Exception:
        try:
            t = str(name).strip()
        except Exception:
            return u""
    return t


def _iter_element_parameters(element):
    if element is None:
        return
    try:
        for p in element.Parameters:
            yield p
    except Exception:
        pass


def _iter_rebar_parameters(rebar):
    return _iter_element_parameters(rebar)


def _find_element_parameter(element, param_name):
    """
    ``LookupParameter`` y, si falla, barrido por ``Parameters``
    (algunos parámetros compartidos solo aparecen al iterar).
    """
    if element is None or not param_name:
        return None
    target = _norm_param_def_name(param_name).lower()
    if not target:
        return None
    try:
        p = element.LookupParameter(param_name)
        if p is not None:
            return p
    except Exception:
        pass
    try:
        for p in _iter_element_parameters(element):
            if p is None:
                continue
            try:
                dn = _norm_param_def_name(p.Definition.Name).lower()
            except Exception:
                continue
            if dn == target:
                return p
    except Exception:
        pass
    return None


def _find_rebar_parameter(rebar, param_name):
    return _find_element_parameter(rebar, param_name)


def generar_armadura_conjunto_guid():
    """UUID de corrida (único por ejecución unificada)."""
    try:
        import uuid
        return unicode(uuid.uuid4())
    except Exception:
        pass
    try:
        import System
        return unicode(System.Guid.NewGuid())
    except Exception:
        return None


def iniciar_armadura_conjunto_guid_ejecucion(conjunto_guid=None):
    """Activa el GUID de corrida en AppDomain para estampar rebars de la ejecución."""
    gid = conjunto_guid or generar_armadura_conjunto_guid()
    if not gid:
        return None
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONJUNTO_GUID_KEY, gid)
    except Exception:
        pass
    return gid


def obtener_armadura_conjunto_guid_actual():
    """GUID de corrida activo o ``None``."""
    try:
        gid = AppDomain.CurrentDomain.GetData(_APPDOMAIN_CONJUNTO_GUID_KEY)
    except Exception:
        gid = None
    if not gid:
        return None
    try:
        t = unicode(gid).strip()
    except Exception:
        try:
            t = str(gid or u"").strip()
        except Exception:
            return None
    return t or None


def finalizar_armadura_conjunto_guid_ejecucion():
    """Limpia el GUID de corrida al terminar la ejecución."""
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONJUNTO_GUID_KEY, None)
    except Exception:
        pass


def _set_element_string_param(element, param_name, value):
    """Escribe un parámetro de instancia tipo texto en un elemento."""
    if element is None or not param_name:
        return False
    try:
        valor = unicode(value)
    except Exception:
        try:
            valor = str(value)
        except Exception:
            return False
    p = _find_element_parameter(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    try:
        st = p.StorageType
        if st == StorageType.String:
            p.Set(valor)
            return True
    except Exception:
        pass
    try:
        p.SetValueString(valor)
        return True
    except Exception:
        pass
    try:
        p.Set(valor)
        return True
    except Exception:
        return False


def _set_rebar_string_param(rebar, param_name, value):
    """Escribe un parámetro de instancia tipo texto en un ``Rebar``."""
    if rebar is None or not isinstance(rebar, Rebar) or not param_name:
        return False
    return _set_element_string_param(rebar, param_name, value)


def stamp_armadura_conjunto_guid(element, conjunto_guid=None):
    """Escribe ``Armadura_Conjunto_GUID`` en un elemento (Rebar, Detail Item, etc.)."""
    if element is None:
        return False
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return False
    return _set_element_string_param(element, ARMADURA_CONJUNTO_GUID_PARAM, gid)


def stamp_armadura_conjunto_guid_por_ids(doc, element_ids, conjunto_guid=None):
    """Escribe ``Armadura_Conjunto_GUID`` en una lista de ``ElementId``."""
    if doc is None or not element_ids:
        return 0
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return 0
    n = 0
    for eid in element_ids:
        try:
            el = doc.GetElement(eid)
        except Exception:
            continue
        if stamp_armadura_conjunto_guid(el, conjunto_guid=gid):
            n += 1
    return n


def get_armadura_conjunto_guid(element):
    """Lee ``Armadura_Conjunto_GUID`` de un elemento o ``None``."""
    if element is None:
        return None
    p = _find_element_parameter(element, ARMADURA_CONJUNTO_GUID_PARAM)
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


def _normalize_conjunto_guid_target(conjunto_guid):
    if not conjunto_guid:
        return None
    try:
        target = unicode(conjunto_guid).strip()
    except Exception:
        try:
            target = str(conjunto_guid or u"").strip()
        except Exception:
            return None
    return target or None


def collect_rebars_por_conjunto_guid(doc, conjunto_guid):
    """
    Devuelve ``ElementId`` de rebars con el mismo ``Armadura_Conjunto_GUID``.

    ``conjunto_guid`` debe coincidir exactamente con el valor leído por
    ``get_armadura_conjunto_guid`` (texto ya normalizado).
    """
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    from Autodesk.Revit.DB import FilteredElementCollector

    ids = []
    try:
        rebars = (
            FilteredElementCollector(doc)
            .OfClass(Rebar)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for rebar in rebars:
        try:
            gid = get_armadura_conjunto_guid(rebar)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(rebar.Id)
            except Exception:
                pass
    return ids


def collect_empalmes_por_conjunto_guid(doc, conjunto_guid):
    """
    Devuelve ``ElementId`` de representaciones de empalme (Detail Items)
    con el mismo ``Armadura_Conjunto_GUID``.
    """
    target = _normalize_conjunto_guid_target(conjunto_guid)
    if doc is None or not target:
        return []

    from Autodesk.Revit.DB import BuiltInCategory, FamilyInstance, FilteredElementCollector

    ids = []
    try:
        details = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_DetailComponents)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []

    for el in details:
        if not isinstance(el, FamilyInstance):
            continue
        try:
            gid = get_armadura_conjunto_guid(el)
        except Exception:
            continue
        if gid == target:
            try:
                ids.append(el.Id)
            except Exception:
                pass
    return ids


def collect_corrida_por_conjunto_guid(doc, conjunto_guid):
    """
    Barras + representaciones de empalme con el mismo GUID de corrida.

    Retorna ``dict`` con claves ``rebar_ids``, ``empalme_ids``, ``all_ids``.
    """
    rebar_ids = collect_rebars_por_conjunto_guid(doc, conjunto_guid)
    empalme_ids = collect_empalmes_por_conjunto_guid(doc, conjunto_guid)
    all_ids = list(rebar_ids)
    seen = set()
    for eid in rebar_ids:
        try:
            seen.add(int(eid.IntegerValue))
        except Exception:
            pass
    for eid in empalme_ids:
        try:
            key = int(eid.IntegerValue)
        except Exception:
            key = None
        if key is not None and key not in seen:
            all_ids.append(eid)
            seen.add(key)
    return {
        u"rebar_ids": rebar_ids,
        u"empalme_ids": empalme_ids,
        u"all_ids": all_ids,
    }


def armadura_ubicacion_valor_desde_capa(layer_index):
    """Capa índice 0 → ``1``, índice 1 → ``2``, etc."""
    try:
        li = int(layer_index)
    except Exception:
        li = 0
    return unicode(max(0, li) + 1)


def set_armadura_ubicacion_desde_capa(rebar, layer_index):
    """Escribe ``Armadura_Ubicacion`` según índice de capa (1-based)."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    valor_txt = armadura_ubicacion_valor_desde_capa(layer_index)
    try:
        valor_int = int(valor_txt)
    except Exception:
        valor_int = 1
    try:
        p = _find_rebar_parameter(rebar, ARMADURA_UBICACION_PARAM)
        if p is None or p.IsReadOnly:
            return False
        st = p.StorageType
        if st == StorageType.Integer:
            p.Set(valor_int)
        elif st == StorageType.String:
            p.Set(valor_txt)
        else:
            p.Set(valor_txt)
        return True
    except Exception:
        return False


def stamp_cabezal_longitudinal_rebar(rebar, layer_index=0):
    """``Armadura_Arainco`` + ``Armadura_Ubicacion`` en barras longitudinales de cabezal."""
    if rebar is None:
        return rebar
    activar_armadura_arainco(rebar)
    set_armadura_ubicacion_desde_capa(rebar, layer_index)
    return rebar


def activar_armadura_arainco(rebar):
    """Activa el parámetro booleano ``Armadura_Arainco`` en un ``Rebar``."""
    if rebar is None or not isinstance(rebar, Rebar):
        return False
    ok = False
    try:
        p = _find_rebar_parameter(rebar, ARMADURA_ARAINCO_PARAM)
        if p is None or p.IsReadOnly:
            stamp_armadura_conjunto_guid(rebar)
            return False
        st = p.StorageType
        if st == StorageType.Integer:
            p.Set(1)
            ok = True
        elif st == StorageType.String:
            p.Set(u"1")
            ok = True
        else:
            try:
                p.Set(1)
                ok = True
            except Exception:
                pass
            if not ok:
                try:
                    p.Set(True)
                    ok = True
                except Exception:
                    pass
            if not ok:
                try:
                    p.SetValueString(u"1")
                    ok = True
                except Exception:
                    pass
            if not ok:
                try:
                    p.SetValueString(u"Yes")
                    ok = True
                except Exception:
                    pass
    except Exception:
        stamp_armadura_conjunto_guid(rebar)
        return False
    stamp_armadura_conjunto_guid(rebar)
    return ok


def activar_armadura_arainco_por_ids(doc, rebar_ids):
    """Activa ``Armadura_Arainco`` en una lista de ``ElementId``."""
    if doc is None or not rebar_ids:
        return 0
    n = 0
    for eid in rebar_ids:
        try:
            el = doc.GetElement(eid)
            if activar_armadura_arainco(el):
                n += 1
        except Exception:
            pass
    return n


def stamp_malla_vertical_rebar(rebar):
    """Barras verticales de malla: ``Armadura_Malla_Tipo`` = D.M., orientación V."""
    if rebar is None:
        return rebar
    _set_rebar_string_param(rebar, ARMADURA_MALLA_TIPO_PARAM, ARMADURA_MALLA_TIPO_DM)
    _set_rebar_string_param(
        rebar, ARMADURA_MALLA_ORIENTACION_PARAM, ARMADURA_MALLA_ORIENT_V,
    )
    stamp_armadura_conjunto_guid(rebar)
    return rebar


def stamp_malla_horizontal_rebar(rebar):
    """Barras horizontales de malla: ``Armadura_Malla_Orientacion`` = H."""
    if rebar is None:
        return rebar
    _set_rebar_string_param(
        rebar, ARMADURA_MALLA_ORIENTACION_PARAM, ARMADURA_MALLA_ORIENT_H,
    )
    stamp_armadura_conjunto_guid(rebar)
    return rebar


def get_armadura_malla_orientacion(rebar):
    """Lee ``Armadura_Malla_Orientacion`` (``V.`` / ``H.``) o ``None``."""
    if rebar is None:
        return None
    p = _find_rebar_parameter(rebar, ARMADURA_MALLA_ORIENTACION_PARAM)
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
        try:
            if p.HasValue and p.StorageType == StorageType.String:
                val = p.AsString()
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
    tl = t.lower().replace(u"\u00a0", u" ")
    if tl in (ARMADURA_MALLA_ORIENT_V.lower(), u"v", u"v."):
        return u"vertical"
    if tl in (ARMADURA_MALLA_ORIENT_H.lower(), u"h", u"h."):
        return u"horizontal"
    if tl.startswith(u"v"):
        return u"vertical"
    if tl.startswith(u"h"):
        return u"horizontal"
    return None
