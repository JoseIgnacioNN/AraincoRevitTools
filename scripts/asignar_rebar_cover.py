# -*- coding: utf-8 -*-
"""
Asignar Rebar Cover a elementos del proyecto.
Ejecutar en Revit Python Shell (RPS).
Variables doc y uidoc deben estar disponibles en el contexto global.
"""

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    Transaction,
)
from Autodesk.Revit.DB.Structure import RebarCoverType, RebarHostData
from Autodesk.Revit.UI import TaskDialog

# Conversión: 1 mm = 0.00328084 pies (API de Revit usa pies internamente)
MM_TO_FEET = 0.00328084

# Categorías y recubrimientos (mm): (BuiltInCategory, valor_mm)
CATEGORIAS_COVER = {
    "Muros (Walls)": (BuiltInCategory.OST_Walls, 25),
    "Suelos (Floors)": (BuiltInCategory.OST_Floors, 20),
    "Cimentaciones (Structural Foundations)": (BuiltInCategory.OST_StructuralFoundation, 50),
    "Armazón Estructural (Structural Framing)": (BuiltInCategory.OST_StructuralFraming, 25),
}


def mm_to_feet(mm):
    """Convierte milímetros a pies."""
    return mm * MM_TO_FEET


def get_or_create_rebar_cover_type(document, mm_value):
    """
    Obtiene un RebarCoverType existente con el valor dado (en mm) o lo crea.
    Retorna el RebarCoverType o None si falla.
    """
    cover_feet = mm_to_feet(mm_value)
    name = "{}mm".format(int(mm_value))

    # Buscar RebarCoverType existente por nombre o valor
    try:
        rct_type = clr.GetClrType(RebarCoverType)
    except Exception:
        rct_type = RebarCoverType
    collector = FilteredElementCollector(document).OfClass(rct_type)
    for elem in collector:
        if elem.Name == name:
            return elem
        # Tolerancia para comparar CoverDistance (evitar errores de punto flotante)
        if abs(elem.CoverDistance - cover_feet) < 1e-6:
            return elem

    # Crear nuevo RebarCoverType si no existe
    try:
        return RebarCoverType.Create(document, name, cover_feet)
    except Exception:
        return None


def asignar_cover_a_elementos(document, categoria_bic, cover_type):
    """
    Asigna el RebarCoverType a todos los elementos válidos de la categoría.
    Retorna (procesados, omitidos).
    """
    procesados = 0
    omitidos = 0

    elementos = (
        FilteredElementCollector(document)
        .OfCategory(categoria_bic)
        .WhereElementIsNotElementType()
    )

    for elem in elementos:
        host_data = RebarHostData.GetRebarHostData(elem)
        if host_data is None:
            omitidos += 1
            continue
        try:
            if not host_data.IsValidHost():
                omitidos += 1
                continue
            host_data.SetCommonCoverType(cover_type)
            procesados += 1
        except Exception:
            omitidos += 1
        finally:
            if host_data is not None:
                try:
                    host_data.Dispose()
                except Exception:
                    pass

    return procesados, omitidos


def main():
    # doc y uidoc disponibles en RPS. En pyRevit: doc = __revit__.ActiveUIDocument.Document
    resumen_lineas = []
    total_procesados = 0
    total_omitidos = 0

    with Transaction(doc, "Asignar Rebar Cover") as txn:
        txn.Start()

        for nombre_cat, (bic, mm_val) in CATEGORIAS_COVER.items():
            cover_type = get_or_create_rebar_cover_type(doc, mm_val)
            if cover_type is None:
                resumen_lineas.append("{}: Error al obtener/crear tipo {}mm".format(nombre_cat, mm_val))
                continue

            procesados, omitidos = asignar_cover_a_elementos(doc, bic, cover_type)
            total_procesados += procesados
            total_omitidos += omitidos
            resumen_lineas.append(
                "{}: {} procesados, {} omitidos (cover {}mm)".format(
                    nombre_cat, procesados, omitidos, mm_val
                )
            )

        txn.Commit()

    mensaje = "Recubrimientos de armadura asignados:\n\n"
    mensaje += "\n".join(resumen_lineas)
    mensaje += "\n\nTotal procesados: {}".format(total_procesados)
    mensaje += "\nTotal omitidos: {}".format(total_omitidos)

    TaskDialog.Show("Rebar Cover - Resumen", mensaje)


main()
