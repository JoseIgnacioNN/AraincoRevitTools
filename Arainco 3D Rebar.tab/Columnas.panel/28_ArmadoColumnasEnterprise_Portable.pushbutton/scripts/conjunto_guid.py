# -*- coding: utf-8 -*-
"""
GUID de corrida (``Armadura_Conjunto_GUID``) — compartido por herramientas de armadura.

Copia mínima de ``armado_muros_rebar_params`` / FloorGeometryCanvas ``conjunto_guid``
sin depender de otros pushbuttons.
"""

import clr

clr.AddReference("System")
from System import AppDomain

from Autodesk.Revit.DB import StorageType

ARMADURA_CONJUNTO_GUID_PARAM = u"Armadura_Conjunto_GUID"
_APPDOMAIN_CONJUNTO_GUID_KEY = u"Arainco_ArmadoColumnas_Conjunto_GUID"


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


def _find_element_parameter(element, param_name):
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


def generar_armadura_conjunto_guid():
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
    gid = conjunto_guid or generar_armadura_conjunto_guid()
    if not gid:
        return None
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONJUNTO_GUID_KEY, gid)
    except Exception:
        pass
    return gid


def obtener_armadura_conjunto_guid_actual():
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
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CONJUNTO_GUID_KEY, None)
    except Exception:
        pass


def _set_element_string_param(element, param_name, value):
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


def stamp_armadura_conjunto_guid(element, conjunto_guid=None):
    if element is None:
        return False
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return False
    return _set_element_string_param(element, ARMADURA_CONJUNTO_GUID_PARAM, gid)


def stamp_armadura_conjunto_guid_en_rebars(rebars, conjunto_guid=None):
    if not rebars:
        return 0
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return 0
    n = 0
    for rb in rebars:
        try:
            if stamp_armadura_conjunto_guid(rb, conjunto_guid=gid):
                n += 1
        except Exception:
            pass
    return n
