# -*- coding: utf-8 -*-
"""
Numeración Fundación — misma lógica que FundacionAislada / SeleccionarPorNumeracion.
"""

from Autodesk.Revit.DB import FilteredElementCollector, StorageType


def _leer_valor_param(p):
    """Extrae el valor del parámetro como string."""
    if not p or not p.HasValue:
        return None
    try:
        st = p.StorageType
        if st == StorageType.String:
            s = p.AsString()
            return str(s).strip() if s and str(s).strip() else None
        if st == StorageType.Integer:
            return str(p.AsInteger())
        if st == StorageType.Double:
            return str(int(round(p.AsDouble())))
    except Exception:
        pass
    try:
        s = p.AsString()
        if s is not None and str(s).strip():
            return str(s).strip()
        vs = p.AsValueString()
        if vs is not None and str(vs).strip():
            return str(vs).strip()
    except Exception:
        pass
    return None


def leer_numeracion_fundacion(element):
    """
    Lee el parámetro 'Numeracion Fundacion' del elemento (instancia).
    Retorna el valor como string o None.
    """
    if not element:
        return None
    param_names = (
        "Numeracion Fundacion",
        "Numeracion fundacion",
        "Numeracion",
        "Foundation Numbering",
    )
    try:
        for param_name in param_names:
            p = element.LookupParameter(param_name)
            if p is not None:
                val = _leer_valor_param(p)
                if val is not None:
                    return val
        if hasattr(element, "GetOrderedParameters"):
            for p in element.GetOrderedParameters():
                if p is None:
                    continue
                try:
                    def_name = (p.Definition.Name or "").lower()
                    if (
                        ("numeracion" in def_name and "fundacion" in def_name)
                        or "foundation numbering" in def_name
                        or def_name == "numeracion"
                    ):
                        val = _leer_valor_param(p)
                        if val is not None:
                            return val
                except Exception:
                    continue
        if hasattr(element, "Parameters"):
            for p in element.Parameters:
                if p is None:
                    continue
                try:
                    def_name = (p.Definition.Name or "").lower()
                    if (
                        ("numeracion" in def_name and "fundacion" in def_name)
                        or "foundation numbering" in def_name
                        or def_name == "numeracion"
                    ):
                        val = _leer_valor_param(p)
                        if val is not None:
                            return val
                except Exception:
                    continue
    except Exception:
        pass
    return None


def buscar_fundaciones_por_numeracion(document, numeracion_val):
    """
    Elementos con Numeracion Fundacion igual a ``numeracion_val``.
    Retorna lista de ``ElementId``.
    """
    if numeracion_val is None:
        return []
    result = []
    try:
        collector = FilteredElementCollector(document).WhereElementIsNotElementType()
        for elem in collector:
            if elem is None:
                continue
            valor = leer_numeracion_fundacion(elem)
            if valor is not None:
                v1 = str(valor).strip()
                v2 = str(numeracion_val).strip()
                if v1 == v2:
                    result.append(elem.Id)
    except Exception:
        pass
    return result
