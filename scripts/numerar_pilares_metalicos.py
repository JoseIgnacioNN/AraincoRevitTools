# -*- coding: utf-8 -*-
"""
Numerar pilares metálicos (Type Mark).

- Revit 2024–2026 | IronPython (pyRevit).
- Filtra ``OST_StructuralColumns`` cuyo material estructural sea **Steel**
  (Material for model behavior = Steel).
- **Type Mark** en el tipo de familia (solo ``LookupParameter``):
  los tipos que **ya tienen** marca (cualquier texto no vacío) **no se modifican**.
  Los que están **vacíos** reciben números **nuevos al final** de la serie: ``max(marcas
  numéricas existentes entre pilares de acero) + 1, +2, …`` ordenados por nombre de tipo.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Transaction,
)
from Autodesk.Revit.DB.Structure import StructuralMaterialType
from Autodesk.Revit.UI import TaskDialog

from System.Collections.Generic import List


def _es_pilar_acero(fi):
    """True si el FamilyInstance es pilar estructural con material Steel."""
    if fi is None or not isinstance(fi, FamilyInstance):
        return False
    try:
        if fi.Category is None:
            return False
    except Exception:
        return False
    try:
        sm = fi.StructuralMaterialType
        return sm == StructuralMaterialType.Steel
    except Exception:
        return False


def _type_mark_param(elem_type):
    """
    Type Mark solo vía LookupParameter (sin BuiltInParameter ni iterar Parameters:
    evita caché de módulos antiguos y posibles fallos del binding CLR).
    """
    if elem_type is None:
        return None
    for name in (
        u"Type Mark",
        u"Marca de tipo",
        u"Marca tipo",
        u"Marca de Tipo",
        u"TYPE_MARK",
    ):
        try:
            p = elem_type.LookupParameter(name)
            if p is not None:
                return p
        except Exception:
            pass
    return None


def _nombre_tipo(doc, tid):
    t = doc.GetElement(tid)
    if t is None:
        return u""
    try:
        return (t.Name or u"").strip()
    except Exception:
        return u""


def _texto_marca_actual(p):
    if p is None:
        return None
    try:
        if not p.HasValue:
            return u""
    except Exception:
        pass
    try:
        s = p.AsString()
        if s is None:
            return u""
        return s.strip()
    except Exception:
        return None


def _marca_vacia(s):
    if s is None:
        return True
    return len(s) == 0


def _entero_desde_marca(s):
    """Si ``s`` es solo un entero (ej. 12, 01), devuelve el int; si no, None."""
    if s is None:
        return None
    t = s.strip()
    if not t:
        return None
    try:
        if t.isdigit():
            return int(t)
        if t.startswith(u"-") and len(t) > 1 and t[1:].isdigit():
            return int(t)
    except Exception:
        pass
    return None


def run(revit):
    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(
            u"BIMTools — Numerar pilares metálicos",
            u"No hay documento activo.",
        )
        return

    doc = uidoc.Document

    instancias = []
    for el in (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralColumns)
        .WhereElementIsNotElementType()
    ):
        if _es_pilar_acero(el):
            instancias.append(el)

    if not instancias:
        TaskDialog.Show(
            u"BIMTools — Numerar pilares metálicos",
            u"No se encontraron pilares estructurales con material Acero (Steel) en el modelo.",
        )
        return

    type_ids = set()
    for inst in instancias:
        try:
            tid = inst.GetTypeId()
            if tid is not None and tid != ElementId.InvalidElementId:
                type_ids.add(tid)
        except Exception:
            pass

    if not type_ids:
        TaskDialog.Show(
            u"BIMTools — Numerar pilares metálicos",
            u"No se pudieron obtener tipos de familia de los pilares seleccionados.",
        )
        return

    ordenados = sorted(type_ids, key=lambda x: _nombre_tipo(doc, x).lower())

    # Máximo numérico entre marcas ya existentes (solo enteros en el texto de la marca).
    max_num = 0
    for tid in ordenados:
        sym = doc.GetElement(tid)
        p = _type_mark_param(sym)
        if p is None:
            continue
        txt = _texto_marca_actual(p)
        if _marca_vacia(txt):
            continue
        n0 = _entero_desde_marca(txt)
        if n0 is not None:
            if n0 > max_num:
                max_num = n0

    conservados = 0
    ok = 0
    skip = 0
    errores = []

    tx = Transaction(doc, u"BIMTools: Type Mark pilares acero")
    tx.Start()
    try:
        siguiente = max_num + 1
        for tid in ordenados:
            sym = doc.GetElement(tid)
            p = _type_mark_param(sym)
            if p is None:
                skip += 1
                errores.append(_nombre_tipo(doc, tid) or str(tid.IntegerValue))
                continue
            if p.IsReadOnly:
                skip += 1
                errores.append(_nombre_tipo(doc, tid) or str(tid.IntegerValue))
                continue
            txt = _texto_marca_actual(p)
            if not _marca_vacia(txt):
                conservados += 1
                continue
            try:
                s = str(siguiente)
                if p.Set(s):
                    ok += 1
                    siguiente += 1
                else:
                    skip += 1
                    errores.append(_nombre_tipo(doc, tid) or str(tid.IntegerValue))
            except Exception:
                skip += 1
                errores.append(_nombre_tipo(doc, tid) or str(tid.IntegerValue))
        tx.Commit()
    except Exception as ex:
        tx.RollBack()
        TaskDialog.Show(
            u"BIMTools — Numerar pilares metálicos",
            u"Error en la transacción: {0}".format(ex),
        )
        return

    # Resaltar ejemplares considerados
    try:
        ids = List[ElementId]()
        for inst in instancias:
            ids.Add(inst.Id)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass

    msg = (
        u"Pilares de acero (ejemplares): {0}\n"
        u"Tipos distintos: {1}\n"
        u"Tipos con marca ya existente (sin cambiar): {2}\n"
        u"Nuevas marcas asignadas (al final de la serie): {3}\n"
        u"Omitidos (sin parámetro o solo lectura): {4}"
    ).format(len(instancias), len(ordenados), conservados, ok, skip)
    if errores:
        muestra = errores[:8]
        msg += u"\n\nTipos con incidencia:\n" + u"\n".join(muestra)
        if len(errores) > 8:
            msg += u"\n…"

    TaskDialog.Show(u"BIMTools — Numerar pilares metálicos", msg)
