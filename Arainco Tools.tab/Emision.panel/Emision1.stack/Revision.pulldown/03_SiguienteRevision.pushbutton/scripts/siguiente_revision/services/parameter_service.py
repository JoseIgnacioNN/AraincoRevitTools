# -*- coding: utf-8 -*-
"""
ParameterService — acceso seguro a parámetros de láminas y cajetines.

Encapsula la estrategia de búsqueda: busca primero en el ViewSheet y luego en
cada instancia de cajetín (TitleBlock) colocada en la lámina. Soporta escritura
con conversión automática según StorageType.
"""

from __future__ import print_function

try:
    unicode
except NameError:
    unicode = str

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, StorageType


def title_block_instances(doc, sheet):
    """Devuelve las instancias de cajetín colocadas en la lámina."""
    bic = getattr(BuiltInCategory, "OST_TitleBlocks", None)
    if bic is None or sheet is None or doc is None:
        return []
    try:
        return list(
            FilteredElementCollector(doc, sheet.Id)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return []


def lookup(sheet, doc, param_name):
    """
    Busca un parámetro por nombre, primero en la lámina y luego en cada cajetín.

    Returns:
        (host_element, Parameter) o (None, None) si no se encuentra.
    """
    if not param_name or sheet is None or doc is None:
        return None, None
    try:
        p = sheet.LookupParameter(param_name)
        if p is not None:
            return sheet, p
    except Exception:
        pass
    for tb in title_block_instances(doc, sheet):
        try:
            p = tb.LookupParameter(param_name)
            if p is not None:
                return tb, p
        except Exception:
            continue
    return None, None


def set_value(param, value):
    """
    Asigna un valor al parámetro según su StorageType.

    Acepta valores string, int o float; convierte automáticamente.

    Returns:
        True si se asignó con éxito, False en caso contrario.
    """
    try:
        uval = unicode(value) if value is not None else u""
    except Exception:
        uval = u""
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType == StorageType.String:
            param.Set(uval)
            return True
        if param.StorageType == StorageType.Integer:
            try:
                param.Set(int(round(float(uval.strip() or u"0"))))
            except Exception:
                return False
            return True
        if param.StorageType == StorageType.Double:
            try:
                param.Set(float(uval.strip().replace(u",", u".")))
            except Exception:
                return False
            return True
    except Exception:
        return False
    return False


def set_named(sheet, doc, param_name, value):
    """
    Busca el parámetro por nombre y asigna el valor.

    Returns:
        True si encontró y asignó, False en caso contrario.
    """
    _, p = lookup(sheet, doc, param_name)
    if p is None:
        return False
    return set_value(p, value)


def get_text(sheet, doc, param_name):
    """
    Devuelve el valor textual de un parámetro (strip), o None si no existe.
    """
    _, p = lookup(sheet, doc, param_name)
    if p is None:
        return None
    try:
        if p.StorageType == StorageType.String:
            return (unicode(p.AsString()) or u"").strip()
        return (unicode(p.AsValueString() or u"")).strip()
    except Exception:
        try:
            return (unicode(p.AsValueString() or u"")).strip()
        except Exception:
            return u""


def is_empty(sheet, doc, param_name):
    """
    True si el parámetro no tiene valor o su texto está vacío.
    False si el parámetro no existe (no se puede emitir).
    """
    _, p = lookup(sheet, doc, param_name)
    if p is None:
        return False
    try:
        if not p.HasValue:
            return True
    except Exception:
        pass
    txt = get_text(sheet, doc, param_name)
    if txt is None:
        return False
    return not txt
