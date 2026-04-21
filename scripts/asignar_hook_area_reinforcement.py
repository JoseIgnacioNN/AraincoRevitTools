"""
Script para Revit Python Shell (RPS) / IronPython 3.4
Revit 2024 - Asigna un Rebar Hook por defecto al inicio y final de todos los Area Reinforcements seleccionados.

Tarea: Asignar un Rebar Hook (el primero encontrado en el proyecto) al inicio y final de todas las
       Area Reinforcements seleccionadas.

Ejecutar: File > Run Script... y seleccionar este archivo .py
O en la consola: exec(open(r'RUTA_COMPLETA\asignar_hook_area_reinforcement.py', encoding='utf-8').read())

Variables globales RPS: doc (Active Document), uidoc (Active UI Document)
"""

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    ElementId,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement, RebarHookType


def obtener_rebar_hook_types(documento):
    """
    Obtiene todos los RebarHookType del proyecto (igual que get_rebar_hook_types).
    Returns: lista de RebarHookType.
    """
    collector = FilteredElementCollector(documento)
    return list(collector.OfClass(RebarHookType))


def obtener_rebar_hook_por_defecto(documento):
    """
    Devuelve el primer RebarHookType encontrado en el proyecto.
    Returns: RebarHookType o None si no hay ninguno.
    """
    hook_types = obtener_rebar_hook_types(documento)
    return hook_types[0] if hook_types else None


def obtener_area_reinforcements_seleccionados(uidocumento):
    """
    Obtiene los Area Reinforcements de la selección actual.
    Returns: lista de AreaReinforcement.
    """
    selection = uidocumento.Selection
    elem_ids = selection.GetElementIds()
    area_reins = []
    for eid in elem_ids:
        elem = doc.GetElement(eid)
        if elem and isinstance(elem, AreaReinforcement):
            area_reins.append(elem)
    return area_reins


def asignar_hook_a_area_reinforcement(area_rein, hook_type_id):
    """
    Asigna el RebarHookType (ElementId) al inicio y final de todas las capas
    del Area Reinforcement. Usa LookupParameter y .Set() con ElementId.
    """
    if not hook_type_id or hook_type_id == ElementId.InvalidElementId:
        return False

    # Parámetros de gancho por capa (inicio=Exterior/Top, final=Interior/Bottom)
    # Referencia: script AreaReinforcementLosa - parámetros Hook Type
    hook_param_names = [
        u"Exterior Major Hook Type", u"Top Major Hook Type",
        u"Exterior Minor Hook Type", u"Top Minor Hook Type",
        u"Interior Major Hook Type", u"Bottom Major Hook Type",
        u"Interior Minor Hook Type", u"Bottom Minor Hook Type",
    ]
    modificado = False
    for pname in hook_param_names:
        try:
            p = area_rein.LookupParameter(pname)
            if p and not p.IsReadOnly:
                p.Set(hook_type_id)
                modificado = True
        except Exception:
            continue
    return modificado


def ejecutar():
    """Flujo principal: obtener hook por defecto, selección, asignar en Transaction."""
    # 1. Obtener primer RebarHookType del proyecto
    hook_type = obtener_rebar_hook_por_defecto(doc)
    if not hook_type:
        print("Error: No hay RebarHookType en el proyecto.")
        print("Crea al menos un tipo de gancho (Structure > Rebar > Hook Types) o usa crear_hook_type.py.")
        return

    # 2. Obtener Area Reinforcements seleccionados
    area_reins = obtener_area_reinforcements_seleccionados(uidoc)
    if not area_reins:
        print("Error: No hay Area Reinforcements en la selección.")
        print("Selecciona uno o más Area Reinforcements y vuelve a ejecutar.")
        return

    # 3. Modificar en Transaction
    t = Transaction(doc, "Asignar Rebar Hook a Area Reinforcements")
    t.Start()
    try:
        modificados = 0
        for ar in area_reins:
            if asignar_hook_a_area_reinforcement(ar, hook_type.Id):
                modificados += 1
        t.Commit()
        nombre_hook = hook_type.Name if hook_type.Name else "(sin nombre)"
        print("OK: Hook '{}' asignado a {} de {} Area Reinforcements.".format(
            nombre_hook, modificados, len(area_reins)))
    except Exception as ex:
        t.RollBack()
        print("Error al asignar ganchos: {}".format(str(ex)))
        raise


# Ejecutar si doc y uidoc están definidos (entorno RPS)
if "doc" in dir() and "uidoc" in dir():
    ejecutar()
else:
    print("Este script debe ejecutarse dentro de Revit Python Shell (doc y uidoc predefinidos).")
