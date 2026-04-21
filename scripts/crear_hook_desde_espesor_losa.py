# -*- coding: utf-8 -*-
"""
Script RPS: Crear RebarHookType en función del espesor de la losa seleccionada.
Ejecutable en Revit Python Shell (RPS) — Revit 2024+ | IronPython 3.4

Requisito: Seleccionar una o más losas (Floor) antes de ejecutar.
El Hook Length = espesor_losa_mm - 40 mm.
Nombre: "Rebar Hook - [grados]º - [largo] mm"

Ejecutar: File > Run Script... y seleccionar este archivo .py
O en la consola: exec(open(r'RUTA\crear_hook_desde_espesor_losa.py', encoding='utf-8').read())
"""

import math
import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import BuiltInParameter, Floor, Transaction, FilteredElementCollector, UnitUtils, UnitTypeId
from Autodesk.Revit.DB.Structure import RebarHookType, RebarBarType

# ── Boilerplate RPS ─────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None

PIES_A_MM = 304.8


def obtener_espesor_losa_mm(losa):
    """Obtiene el espesor de una losa en mm."""
    param = losa.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
    if param is None or not param.HasValue:
        return None
    return param.AsDouble() * PIES_A_MM


def obtener_losas_seleccionadas():
    """Obtiene las losas (Floor) actualmente seleccionadas."""
    if doc is None or uidoc is None:
        return []
    elem_ids = list(uidoc.Selection.GetElementIds())
    losas = []
    for elem_id in elem_ids:
        elem = doc.GetElement(elem_id)
        if elem is not None and isinstance(elem, Floor):
            losas.append(elem)
    return losas


def obtener_rebar_bar_types(documento):
    """Obtiene todos los RebarBarType del documento."""
    collector = FilteredElementCollector(documento)
    return list(collector.OfClass(RebarBarType))


def obtener_primer_rebar_bar_type(documento):
    """Obtiene el primer RebarBarType del documento."""
    bar_types = obtener_rebar_bar_types(documento)
    return bar_types[0] if bar_types else None


def obtener_nombres_hook_existentes(documento):
    """Obtiene los nombres de todos los RebarHookType existentes."""
    collector = FilteredElementCollector(documento)
    hook_types = list(collector.OfClass(RebarHookType))
    return [ht.Name for ht in hook_types]


def crear_hook_type_en_doc(documento, nombre, angulo_grados, multiplicador_extension, largo_hook_mm, rebar_bar_type=None, en_transaccion=True):
    """
    Crea un nuevo RebarHookType con el Hook Length indicado.
    """
    bar_type = rebar_bar_type or obtener_primer_rebar_bar_type(documento)
    if not bar_type:
        raise Exception("No hay RebarBarType en el documento. Crea al menos un tipo de barra de armadura.")

    nombres_existentes = obtener_nombres_hook_existentes(documento)
    angulo_str = str(int(angulo_grados)) if angulo_grados == int(angulo_grados) else str(angulo_grados)

    t = Transaction(documento, "Crear tipo de gancho desde espesor losa") if en_transaccion else None
    if t:
        t.Start()

    try:
        angulo_radianes = math.radians(angulo_grados)
        hook_type = RebarHookType.Create(documento, angulo_radianes, multiplicador_extension)

        largo_interno = UnitUtils.ConvertToInternalUnits(largo_hook_mm, UnitTypeId.Millimeters)
        for bt in obtener_rebar_bar_types(documento):
            bt.SetAutoCalcHookLengths(hook_type.Id, False)
            bt.SetHookLength(hook_type.Id, largo_interno)

        largo_str = "{} mm".format(int(round(largo_hook_mm)))
        nombre_base = "{} - {}º - {}".format(nombre, angulo_str, largo_str)
        nombre_final = nombre_base
        if nombre_base in nombres_existentes:
            contador = 1
            while nombre_final in nombres_existentes:
                nombre_final = "{} ({})".format(nombre_base, contador)
                contador += 1
            print("Nombre '{}' ya existe. Usando '{}'.".format(nombre_base, nombre_final))

        hook_type.Name = nombre_final

        if t:
            t.Commit()
        return hook_type

    except Exception as ex:
        if t:
            t.RollBack()
        raise


RESTA_MM = 40  # Se resta al espesor de la losa para obtener el Hook Length


def main(
    nombre="Rebar Hook",
    angulo_grados=90.0,
    multiplicador_extension=12.0,
    usar_espesor_minimo=True
):
    """
    Crea un RebarHookType con Hook Length = espesor_losa_mm - 40 mm.
    Nombre: "Rebar Hook - [grados]º - [largo] mm"

    Args:
        nombre: Nombre base del tipo de gancho (default "Rebar Hook")
        angulo_grados: Ángulo del gancho en grados (default 90)
        multiplicador_extension: Multiplicador de extensión (default 12.0)
        usar_espesor_minimo: Si True y hay varias losas, usa el espesor mínimo; si False, el máximo
    """
    if doc is None or uidoc is None:
        print("Error: Ejecuta este script dentro de Revit Python Shell (RPS).")
        return None

    losas = obtener_losas_seleccionadas()
    if not losas:
        print("No hay losas seleccionadas. Selecciona una o más losas (Floor) y vuelve a ejecutar.")
        return None

    espesores = []
    for losa in losas:
        e = obtener_espesor_losa_mm(losa)
        if e is not None:
            espesores.append(e)

    if not espesores:
        print("No se pudo obtener el espesor de ninguna losa seleccionada.")
        return None

    espesor_ref = min(espesores) if usar_espesor_minimo else max(espesores)
    largo_hook_mm = max(0, espesor_ref - RESTA_MM)

    print("-" * 60)
    print("Crear Hook desde espesor de losa")
    print("-" * 60)
    print("Losa(s) seleccionada(s): {}".format(len(losas)))
    print("Espesor de referencia: {:.2f} mm".format(espesor_ref))
    print("Resta aplicada: {} mm".format(RESTA_MM))
    print("Hook Length calculado: {:.2f} mm".format(largo_hook_mm))
    print("-" * 60)

    hook_type = crear_hook_type_en_doc(
        doc,
        nombre=nombre,
        angulo_grados=angulo_grados,
        multiplicador_extension=multiplicador_extension,
        largo_hook_mm=largo_hook_mm,
        en_transaccion=True
    )

    print("Tipo de gancho creado: '{}' (Hook Length: {:.2f} mm)".format(hook_type.Name, largo_hook_mm))
    return hook_type


# ── Ejecución ───────────────────────────────────────────────────────────────
if __name__ == "__main__" or (doc is not None and uidoc is not None):
    # Hook Length = espesor_losa - 40 mm. Nombre: "Rebar Hook - 90º - 140 mm"
    main(
        nombre="Rebar Hook",
        angulo_grados=90.0,
        multiplicador_extension=12.0,
        usar_espesor_minimo=True
    )
