# -*- coding: utf-8 -*-
"""
Script RPS (RevitPythonShell / IronPython 3.4)
Lista:
1) Todos los RebarHookType disponibles en el proyecto actual.
2) Qué RebarHookType está asignado en cada slot de AreaReinforcement:
   - Exterior Major Hook Type
   - Top Major Hook Type
   - Exterior Minor Hook Type
   - Top Minor Hook Type
   - Interior Major Hook Type
   - Bottom Major Hook Type
   - Interior Minor Hook Type
   - Bottom Minor Hook Type

Uso:
File > Run Script... y seleccionar este .py

Notas:
- El “Hook Length” puede variar según RebarBarType; aquí calculamos la(s)
  longitud(es) únicas (mm) usando todos los RebarBarType del proyecto.
"""

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
    StorageType,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement, RebarBarType, RebarHookType
from Autodesk.Revit.UI import TaskDialog


def _get_doc():
    # En RPS normalmente ya existen doc/uidoc; en pyRevit puede existir __revit__.
    if "doc" in globals() and doc is not None:
        return doc
    try:
        return __revit__.ActiveUIDocument.Document
    except Exception:
        return None


doc = _get_doc()
if doc is None:
    TaskDialog.Show("Error", "No hay documento activo. Ejecuta este script dentro de Revit (RPS).")
    raise Exception("No hay documento activo.")


HOOK_SLOTS = [
    u"Exterior Major Hook Type",
    u"Top Major Hook Type",
    u"Exterior Minor Hook Type",
    u"Top Minor Hook Type",
    u"Interior Major Hook Type",
    u"Bottom Major Hook Type",
    u"Interior Minor Hook Type",
    u"Bottom Minor Hook Type",
]


def get_all_hook_types(document):
    collector = FilteredElementCollector(document)
    return list(collector.OfClass(RebarHookType))


def get_all_bar_types(document):
    collector = FilteredElementCollector(document)
    return list(collector.OfClass(RebarBarType))


def get_hook_lengths_mm_for_type(bar_types, hook_type, document):
    """
    Devuelve un set de longitudes únicas (mm) según todos los RebarBarType.
    """
    lengths = set()
    try:
        for bt in bar_types:
            if bt is None or hook_type is None:
                continue
            try:
                largo_interno = bt.GetHookLength(hook_type.Id)
                largo_mm = UnitUtils.ConvertFromInternalUnits(largo_interno, UnitTypeId.Millimeters)
                # Redondeo para agrupar valores muy cercanos.
                lengths.add(round(largo_mm, 2))
            except Exception:
                continue
    except Exception:
        pass
    return lengths


def collect_slot_assignments(document):
    """
    Recorre todos los AreaReinforcement y acumula qué RebarHookType está asignado
    en cada slot (ElementId leído desde los parámetros).
    """
    slot_to_hook_ids = {}
    for slot in HOOK_SLOTS:
        slot_to_hook_ids[slot] = set()

    hook_types = get_all_hook_types(document)
    hook_id_to_name = {}
    for ht in hook_types:
        if ht:
            hook_id_to_name[ht.Id.IntegerValue] = ht.Name

    area_reins = list(FilteredElementCollector(document).OfClass(AreaReinforcement))

    for ar in area_reins:
        if ar is None:
            continue
        for slot in HOOK_SLOTS:
            try:
                p = ar.LookupParameter(slot)
                if p is None or p.IsReadOnly:
                    continue
                if p.StorageType != StorageType.ElementId:
                    continue
                eid = p.AsElementId()
                if eid is None or eid == ElementId.InvalidElementId:
                    continue
                slot_to_hook_ids[slot].add(eid.IntegerValue)
            except Exception:
                continue

    return slot_to_hook_ids, hook_id_to_name


def main():
    hook_types = get_all_hook_types(doc)
    bar_types = get_all_bar_types(doc)

    # Index para encontrar nombre por ID al imprimir.
    hook_id_to_name = {}
    for ht in hook_types:
        try:
            hook_id_to_name[ht.Id.IntegerValue] = ht.Name
        except Exception:
            pass

    # 1) Lista completa de hook types
    print("=" * 80)
    print("REBARHOOK TYPES DISPONIBLES (modelo actual)")
    print("Total RebarHookType:", len(hook_types))
    print("-" * 80)

    for ht in sorted(hook_types, key=lambda x: x.Name or ""):
        try:
            lengths = get_hook_lengths_mm_for_type(bar_types, ht, doc)
            if lengths:
                lengths_str = ", ".join([str(x) for x in sorted(list(lengths))])
            else:
                lengths_str = "(sin longitudes calculables)"
            name = ht.Name if ht.Name else "(sin nombre)"
            print("- {0} | Id={1} | HookLengths(mm)={2}".format(name, ht.Id.IntegerValue, lengths_str))
        except Exception as ex:
            try:
                print("- (error imprimiendo hook) Id={0} | {1}".format(ht.Id.IntegerValue, str(ex)))
            except Exception:
                print("- (error imprimiendo hook)")

    # 2) Asignaciones por slot en AreaReinforcement
    print("=" * 80)
    print("ASIGNACIONES POR SLOT EN AreaReinforcement")
    slot_to_hook_ids, hook_id_to_name2 = collect_slot_assignments(doc)

    total_area_reins = len(list(FilteredElementCollector(doc).OfClass(AreaReinforcement)))
    print("AreaReinforcement encontrados:", total_area_reins)
    print("-" * 80)

    for slot in HOOK_SLOTS:
        ids = slot_to_hook_ids.get(slot, set())
        if not ids:
            print("{0}: (sin asignación)".format(slot))
            continue

        # Ordenar por nombre para legibilidad
        names = []
        for hid in ids:
            names.append(hook_id_to_name2.get(hid, hook_id_to_name.get(hid, "Id={0}".format(hid))))
        names_sorted = sorted(list(set(names)))
        print("{0}: {1} tipo(s) -> {2}".format(slot, len(names_sorted), ", ".join(names_sorted)))

    # Resumen para TaskDialog (corto)
    resumen = "RebarHookType: {0}\nAreaReinforcement: {1}\nRevisar consola para el detalle.".format(
        len(hook_types), total_area_reins
    )
    TaskDialog.Show("Lista SLOTS Rebar Hooks", resumen)


main()

