# -*- coding: utf-8 -*-
"""
Script para editar parámetros de un elemento Area Reinforcement existente.
Ejecutable en RevitPythonShell (RPS) o pyRevit — Revit 2024+ | IronPython 3.4

Requisito: Tener seleccionado un elemento Area Reinforcement antes de ejecutar.
"""

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement

# ── Boilerplate RPS / pyRevit ───────────────────────────────────────────────
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# BuiltInParameter para robustez multiidioma (fallback por LookupParameter)
_BIP_BAR_TYPES = [
    "REBAR_SYSTEM_BAR_TYPE_TOP_DIR_1_GENERIC",      # Top Major
    "REBAR_SYSTEM_BAR_TYPE_TOP_DIR_2_GENERIC",      # Top Minor
    "REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_1_GENERIC",   # Bottom Major
    "REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_2_GENERIC",   # Bottom Minor
]
_BIP_SPACING = [
    "REBAR_SYSTEM_SPACING_TOP_DIR_1_GENERIC",       # Top Major
    "REBAR_SYSTEM_SPACING_TOP_DIR_2_GENERIC",       # Top Minor
    "REBAR_SYSTEM_SPACING_BOTTOM_DIR_1_GENERIC",    # Bottom Major
    "REBAR_SYSTEM_SPACING_BOTTOM_DIR_2_GENERIC",    # Bottom Minor
]

# Nombres LookupParameter (inglés) como fallback
_LOOKUP_BAR_NAMES = [
    [u"Top Major Bar Type", u"Exterior Major Bar Type", u"Exterior Major Rebar Type"],
    [u"Top Minor Bar Type", u"Exterior Minor Bar Type", u"Exterior Minor Rebar Type"],
    [u"Bottom Major Bar Type", u"Interior Major Bar Type", u"Interior Major Rebar Type"],
    [u"Bottom Minor Bar Type", u"Interior Minor Bar Type", u"Interior Minor Rebar Type"],
]
_LOOKUP_SPACING_NAMES = [
    [u"Top Major Spacing", u"Exterior Major Spacing"],
    [u"Top Minor Spacing", u"Exterior Minor Spacing"],
    [u"Bottom Major Spacing", u"Interior Major Spacing"],
    [u"Bottom Minor Spacing", u"Interior Minor Spacing"],
]


def _set_param_safe(param, value, value_type="element"):
    """Intenta asignar valor al parámetro si existe y no es ReadOnly."""
    if param is None or param.IsReadOnly:
        return False
    if value_type == "element":
        if value is None or value == ElementId.InvalidElementId:
            return False
        if param.StorageType != StorageType.ElementId:
            return False
    elif value_type == "double":
        if value is None:
            return False
        if param.StorageType != StorageType.Double:
            return False
    try:
        param.Set(value)
        return True
    except Exception:
        return False


def editar_parametros_area_reinforcement(
    area_rein,
    top_major_bar_type_id=None,
    top_major_spacing=None,
    top_minor_bar_type_id=None,
    top_minor_spacing=None,
    bottom_major_bar_type_id=None,
    bottom_major_spacing=None,
    bottom_minor_bar_type_id=None,
    bottom_minor_spacing=None,
):
    """
    Edita los parámetros de Bar Type y Spacing de un AreaReinforcement.

    Parámetros:
        area_rein: Elemento AreaReinforcement (o ElementId)
        top_major_bar_type_id, top_minor_bar_type_id: ElementId del RebarBarType (o None para no cambiar)
        bottom_major_bar_type_id, bottom_minor_bar_type_id: idem
        top_major_spacing, top_minor_spacing: Double en unidades internas de Revit (o None para no cambiar)
        bottom_major_spacing, bottom_minor_spacing: idem

    Los valores None se ignoran (no se modifica ese parámetro).
    """
    if area_rein is None:
        return False
    if isinstance(area_rein, ElementId):
        area_rein = doc.GetElement(area_rein)
    if not area_rein or not isinstance(area_rein, AreaReinforcement):
        return False

    bar_ids = [
        top_major_bar_type_id,
        top_minor_bar_type_id,
        bottom_major_bar_type_id,
        bottom_minor_bar_type_id,
    ]
    spacings = [
        top_major_spacing,
        top_minor_spacing,
        bottom_major_spacing,
        bottom_minor_spacing,
    ]

    modified = 0

    # 1) Intentar por BuiltInParameter
    for i, (bip_bar, bip_spc) in enumerate(zip(_BIP_BAR_TYPES, _BIP_SPACING)):
        bip_bar_enum = getattr(BuiltInParameter, bip_bar, None)
        bip_spc_enum = getattr(BuiltInParameter, bip_spc, None)
        if bip_bar_enum is not None and bar_ids[i] is not None:
            try:
                p = area_rein.get_Parameter(bip_bar_enum)
                if _set_param_safe(p, bar_ids[i], "element"):
                    modified += 1
            except Exception:
                pass
        if bip_spc_enum is not None and spacings[i] is not None:
            try:
                p = area_rein.get_Parameter(bip_spc_enum)
                if _set_param_safe(p, spacings[i], "double"):
                    modified += 1
            except Exception:
                pass

    # 2) Fallback por LookupParameter
    for i in range(4):
        if bar_ids[i] is not None:
            for name in _LOOKUP_BAR_NAMES[i]:
                try:
                    p = area_rein.LookupParameter(name)
                    if _set_param_safe(p, bar_ids[i], "element"):
                        modified += 1
                        break
                except Exception:
                    continue
        if spacings[i] is not None:
            for name in _LOOKUP_SPACING_NAMES[i]:
                try:
                    p = area_rein.LookupParameter(name)
                    if _set_param_safe(p, spacings[i], "double"):
                        modified += 1
                        break
                except Exception:
                    continue

    return modified > 0


def mm_a_internas(valor_mm):
    """Convierte mm a unidades internas de Revit (para espaciado)."""
    return UnitUtils.ConvertToInternalUnits(float(valor_mm), UnitTypeId.Millimeters)


# ── Ejecución principal (RPS/pyRevit) ─────────────────────────────────────────
try:
    elem_ids = list(uidoc.Selection.GetElementIds())
    if not elem_ids:
        print("Error: No hay elemento seleccionado. Selecciona un Area Reinforcement.")
    else:
        elem = doc.GetElement(elem_ids[0])
        if elem is None:
            print("Error: No se pudo obtener el elemento seleccionado.")
        elif not isinstance(elem, AreaReinforcement):
            print("Error: El elemento seleccionado no es un Area Reinforcement. Tipo: {}".format(
                type(elem).__name__
            ))
        else:
            trans = Transaction(doc, "Editar Refuerzo de Área")
            trans.Start()
            try:
                # Ejemplo: editar espaciado a 150 mm (usa mm_a_internas para conversión)
                spacing_internal = mm_a_internas(150.0)
                ok = editar_parametros_area_reinforcement(
                    elem,
                    top_major_spacing=spacing_internal,
                    top_minor_spacing=spacing_internal,
                    bottom_major_spacing=spacing_internal,
                    bottom_minor_spacing=spacing_internal,
                    # top_major_bar_type_id=...,  # pasar ElementId si se desea cambiar tipo de barra
                )
                trans.Commit()
                if ok:
                    print("Parámetros del Area Reinforcement (ID: {}) actualizados correctamente.".format(
                        elem.Id.IntegerValue
                    ))
                else:
                    print("No se modificaron parámetros. Revisa que los valores sean válidos.")
            except Exception as ex:
                trans.RollBack()
                print("Error al editar: {}".format(str(ex)))
except Exception as ex:
    print("Error: {}".format(str(ex)))
