# -*- coding: utf-8 -*-
"""
Servicio dominio Revit para Area Reinforcement en muros — sin UI.

Usado por:
- FloorGeometryCanvas (espesor de muro, ganchos)

Solo References RevitAPI (no PresentationFramework).

Revit 2024+ · IronPython para pyRevit.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    UnitUtils,
    UnitTypeId,
    XYZ,
)
from area_reinforcement_losa import (
    _buscar_hook_por_largo,
    _crear_hook_desde_largo,
)


# Resta al espesor de muro para el largo del gancho (alineado con botón 09).
_HOOK_RESTA_MURO_MM = 40


def obtener_espesor_muro_mm(wall):
    """
    Obtiene el espesor del muro en mm.
    Prioriza WALL_ATTR_WIDTH_PARAM y parámetros de tipo conocidos.
    """
    if wall is None:
        return None
    try:
        param = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    try:
        param = wall.LookupParameter("Default Thickness")
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    try:
        type_id = wall.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            wall_type = wall.Document.GetElement(type_id)
            if wall_type:
                for pname in ("Default Thickness", "Thickness", "Espesor", "Width"):
                    param = wall_type.LookupParameter(pname)
                    if param and param.HasValue:
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def obtener_o_crear_hook_desde_espesor_muro(document, wall, en_transaccion=True):
    """
    RebarHookType con Hook Length = espesor_muro - 40 mm.
    Debe ejecutarse dentro de Transaction (normalmente como SubTransaction).
    """
    espesor_mm = obtener_espesor_muro_mm(wall)
    if espesor_mm is None:
        return ElementId.InvalidElementId
    largo_target = espesor_mm - _HOOK_RESTA_MURO_MM
    if largo_target <= 0:
        return ElementId.InvalidElementId
    hook = _buscar_hook_por_largo(document, largo_target)
    if hook:
        return hook.Id
    nuevo = _crear_hook_desde_largo(document, largo_target, en_transaccion=en_transaccion)
    return nuevo.Id if nuevo else ElementId.InvalidElementId


def obtener_direccion_principal_muro(wall):
    """Dirección de trazado principal (eje del muro) desde Location.Curve."""
    try:
        loc = wall.Location
        if loc is None:
            return XYZ(1, 0, 0)
        curve = getattr(loc, "Curve", None)
        if curve is None:
            return XYZ(1, 0, 0)
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        dz = p1.Z - p0.Z
        length = (dx * dx + dy * dy + dz * dz) ** 0.5
        if length > 1e-6:
            return XYZ(dx / length, dy / length, dz / length)
    except Exception:
        pass
    return XYZ(1, 0, 0)


# Alias con prefijo `_` (compatibilidad histórica).
_obtener_o_crear_hook_desde_espesor_muro = obtener_o_crear_hook_desde_espesor_muro
_obtener_direccion_principal_muro = obtener_direccion_principal_muro
