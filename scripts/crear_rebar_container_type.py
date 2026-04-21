# -*- coding: utf-8 -*-
"""
Crea/asegura un tipo de RebarContainerType para poder crear nuevos Rebar Containers.

Uso (Revit Python Shell - RPS):
- Deben existir variables globales: `doc` y `uidoc` (habitualmente RPS las provee).

Si el documento no trae el tipo requerido, este script lo crea con un nombre configurable.
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Transaction, FilteredElementCollector
from Autodesk.Revit.DB.Structure import RebarContainerType, RebarContainer
from Autodesk.Revit.UI import TaskDialog


# Nombre del tipo a crear/recuperar.
# Puedes cambiarlo según tu convención.
TYPE_NAME = "BIMTools_RebarContainerType"


def get_existing_type_id(document, type_name):
    """Devuelve ElementId del tipo si ya existe; si no, devuelve None."""
    # En IronPython a veces conviene pasar el CLR type explícito.
    try:
        rct_type = clr.GetClrType(RebarContainerType)
    except Exception:
        rct_type = RebarContainerType
    collector = FilteredElementCollector(document).OfClass(rct_type)
    for t in collector:
        try:
            if t.Name == type_name:
                return t.Id
        except Exception:
            pass
    return None


def get_type_id(document, type_name):
    """
    Intenta crear/recuperar el tipo usando la API (preferido).
    Si falla, usa el "default" como fallback.
    """
    # GetOrCreate requiere Transaction; se llama desde main.
    type_id = RebarContainerType.GetOrCreateRebarContainerType(document, type_name)
    if type_id is not None and type_id.IntegerValue >= 0:
        return type_id

    # Fallback defensivo: crear el tipo default.
    return RebarContainerType.CreateDefaultRebarContainerType(document)


def get_any_rebar_container_type_id(document):
    """Si existen instancias de RebarContainer, devuelve el TypeId de alguna de ellas."""
    try:
        rc_type = clr.GetClrType(RebarContainer)
    except Exception:
        rc_type = RebarContainer

    collector = FilteredElementCollector(document).OfClass(rc_type)
    for rc in collector:
        try:
            type_id = rc.GetTypeId()
            if type_id is not None:
                return type_id
        except Exception:
            pass
    return None


def main():
    # En RPS, normalmente existe `doc`. Si no, intentamos obtenerlo desde __revit__.
    try:
        document = doc
    except NameError:
        document = __revit__.ActiveUIDocument.Document

    type_name = (TYPE_NAME or "").strip()
    if not type_name:
        raise Exception("TYPE_NAME no puede estar vacío.")

    # 1) Verificar si ya existen instancias de RebarContainer.
    container_type_id = get_any_rebar_container_type_id(document)
    if container_type_id is not None:
        type_id = container_type_id
    else:
        # 2) Si no hay instancias, validar si ya existe el tipo requerido.
        existing_type_id = get_existing_type_id(document, type_name)
        if existing_type_id is not None:
            type_id = existing_type_id
        else:
            # 3) Si no hay instancias y no existe el tipo, crearlo.
            with Transaction(document, "Crear RebarContainerType") as txn:
                txn.Start()
                type_id = get_type_id(document, type_name)

    type_elem = document.GetElement(type_id)
    real_name = type_elem.Name if type_elem else str(type_id)
    msg = (
        "RebarContainerType listo.\n"
        "Nombre: {}\n"
        "Id: {}".format(real_name, type_id.IntegerValue if type_id else "N/A")
    )

    try:
        TaskDialog.Show("BIMTools - RebarContainerType", msg)
    except Exception:
        # Por si RPS está en modo que no permite TaskDialog.
        print(msg)


main()

