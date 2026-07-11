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
ARMADURA_MALLA_PARAM = u"Armadura_Malla"
ARMADURA_ARAINCO_PARAM = u"Armadura_Arainco"
ARMADURA_UBICACION_PARAM = u"Armadura_Ubicacion"
ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
ARMADURA_UBICACION_INFERIOR = u"F"
ARMADURA_UBICACION_SUPERIOR = u"F'"
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


def inspect_conjunto_guid_param(element):
    """
    Diagnóstico del parámetro ``Armadura_Conjunto_GUID`` en un elemento.

    Devuelve dict con claves: found, read_only, storage_type, current_value, category.
    """
    out = {
        u"found": False,
        u"read_only": None,
        u"storage_type": None,
        u"current_value": None,
        u"category": None,
        u"element_class": None,
    }
    if element is None:
        return out
    try:
        out[u"element_class"] = unicode(element.GetType().Name)
    except Exception:
        pass
    try:
        if element.Category is not None:
            out[u"category"] = unicode(element.Category.Name)
    except Exception:
        pass
    p = _find_element_parameter(element, ARMADURA_CONJUNTO_GUID_PARAM)
    if p is None:
        return out
    out[u"found"] = True
    try:
        out[u"read_only"] = bool(p.IsReadOnly)
    except Exception:
        pass
    try:
        out[u"storage_type"] = unicode(p.StorageType)
    except Exception:
        pass
    try:
        out[u"current_value"] = p.AsString()
    except Exception:
        try:
            out[u"current_value"] = p.AsValueString()
        except Exception:
            pass
    return out


def stamp_armadura_conjunto_guid_with_reason(element, conjunto_guid=None):
    """
    Igual que ``stamp_armadura_conjunto_guid`` pero devuelve ``(ok, motivo)``.
    """
    if element is None:
        return False, u"elemento None"
    gid = conjunto_guid or obtener_armadura_conjunto_guid_actual()
    if not gid:
        return False, u"sin GUID de corrida (no generado ni en AppDomain)"
    info = inspect_conjunto_guid_param(element)
    if not info.get(u"found"):
        return False, (
            u"parámetro «{0}» no encontrado en {1} ({2})".format(
                ARMADURA_CONJUNTO_GUID_PARAM,
                info.get(u"element_class") or u"?",
                info.get(u"category") or u"sin categoría",
            )
        )
    if info.get(u"read_only"):
        return False, u"parámetro «{0}» es solo lectura".format(ARMADURA_CONJUNTO_GUID_PARAM)
    if _set_element_string_param(element, ARMADURA_CONJUNTO_GUID_PARAM, gid):
        return True, u"ok"
    return False, (
        u"Set falló (storage={0}, valor previo={1})".format(
            info.get(u"storage_type") or u"?",
            info.get(u"current_value") or u"(vacío)",
        )
    )


def _set_element_yes_no_param(element, param_name, yes=True):
    """Escribe un parámetro Yes/No (Integer 0/1, bool o texto Yes/No)."""
    if element is None or not param_name:
        return False
    p = _find_element_parameter(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    if yes:
        candidates = (1, True, u"1", u"Yes", u"yes", u"Sí", u"SI")
    else:
        candidates = (0, False, u"0", u"No", u"no")
    try:
        st = p.StorageType
        if st == StorageType.Integer:
            p.Set(1 if yes else 0)
            return True
    except Exception:
        pass
    for val in candidates:
        try:
            p.Set(val)
            return True
        except Exception:
            continue
    try:
        p.SetValueString(u"Yes" if yes else u"No")
        return True
    except Exception:
        return False


def stamp_armadura_malla(element, yes=True):
    """Marca ``Armadura_Malla`` = Yes en barras de malla (Rebar / RebarInSystem)."""
    if element is None:
        return False
    return _set_element_yes_no_param(element, ARMADURA_MALLA_PARAM, yes=yes)


def stamp_armadura_arainco(element, yes=True):
    """Marca ``Armadura_Arainco`` = Yes en barras creadas por herramientas Arainco."""
    if element is None:
        return False
    return _set_element_yes_no_param(element, ARMADURA_ARAINCO_PARAM, yes=yes)


def stamp_armadura_ubicacion(element, valor):
    """Escribe ``Armadura_Ubicacion`` (p. ej. ``F`` / ``F'``) si el parámetro es editable."""
    if element is None or not valor:
        return False
    return _set_element_string_param(element, ARMADURA_UBICACION_PARAM, valor)


def stamp_armadura_nivel(element, valor):
    """Escribe ``Armadura_Nivel`` (nombre del nivel del host) como string."""
    if element is None or not valor:
        return False
    return _set_element_string_param(element, ARMADURA_NIVEL_PARAM, valor)
