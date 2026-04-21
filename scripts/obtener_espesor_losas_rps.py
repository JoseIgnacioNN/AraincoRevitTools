# -*- coding: utf-8 -*-
"""
Script RPS: Obtener espesor de losas seleccionadas.
Ejecutable en Revit Python Shell (RPS) — Revit 2024+ | IronPython 3.4

Requisito: Seleccionar losas (Floor) en el modelo antes de ejecutar.
Imprime nombre, ID y espesor en mm para cada losa.
"""

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import BuiltInParameter, Floor

# ── Boilerplate RPS ─────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None

# Factor de conversión: pies (unidades internas Revit) a milímetros
PIES_A_MM = 304.8


def main():
    if doc is None or uidoc is None:
        print("Error: Ejecuta este script dentro de Revit Python Shell (RPS).")
        return

    elem_ids = list(uidoc.Selection.GetElementIds())
    if not elem_ids:
        print("No hay elementos seleccionados. Selecciona losas (Floor) y vuelve a ejecutar.")
        return

    losas = []
    for elem_id in elem_ids:
        elem = doc.GetElement(elem_id)
        if elem is not None and isinstance(elem, Floor):
            losas.append(elem)

    if not losas:
        print("Ninguno de los elementos seleccionados es una losa (Floor).")
        return

    print("-" * 60)
    print("Espesor de losas seleccionadas")
    print("-" * 60)

    for losa in losas:
        nombre = losa.Name if losa.Name else "(sin nombre)"
        elem_id = losa.Id.IntegerValue

        param = losa.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        if param is None or not param.HasValue:
            espesor_mm = None
        else:
            espesor_pies = param.AsDouble()
            espesor_mm = espesor_pies * PIES_A_MM

        if espesor_mm is not None:
            print("Nombre: {} | ID: {} | Espesor: {:.2f} mm".format(
                nombre, elem_id, espesor_mm))
        else:
            print("Nombre: {} | ID: {} | Espesor: (no disponible)".format(
                nombre, elem_id))

    print("-" * 60)
    print("Total: {} losa(s) procesada(s).".format(len(losas)))


if __name__ == "__main__" or (doc is not None and uidoc is not None):
    main()
