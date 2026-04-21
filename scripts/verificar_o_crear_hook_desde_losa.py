# -*- coding: utf-8 -*-
"""
Script RPS: Verifica si existe un RebarHookType con Hook Length = espesor_losa - 40 mm.
Si no existe, crea uno nuevo. Asigna el gancho a los Area Reinforcements de la losa.

Ejecutable en Revit Python Shell (RPS) — Revit 2024+ | IronPython 3.4

Requisito: Seleccionar una losa (Floor) antes de ejecutar.
- Extrae el espesor de la losa (FLOOR_ATTR_THICKNESS_PARAM) y resta 40 mm
- Busca RebarHookType con ese Hook Length en TODOS los RebarBarType
- Si no encuentra coincidencia, crea un nuevo RebarHookType
- Asigna el gancho a los Area Reinforcements hospedados en la losa
"""

import math
import clr

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Floor,
    Transaction,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement, RebarBarType, RebarHookType

# ── Boilerplate RPS ─────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None

RESTA_MM = 40
TOLERANCIA_MM = 0.5  # Tolerancia para comparar Hook Length


def obtener_espesor_losa_mm(losa):
    """
    Obtiene el espesor de una losa en mm.
    Prioriza FLOOR_ATTR_THICKNESS_PARAM (instancia) y LookupParameter('Default Thickness')
    en instancia y tipo como fallback (el espesor suele estar en el FloorType).
    """
    if losa is None:
        return None
    # 1) BuiltInParameter en la instancia
    try:
        param = losa.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 2) LookupParameter "Default Thickness" en la instancia
    try:
        param = losa.LookupParameter("Default Thickness")
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 3) LookupParameter en el tipo (común en losas estructurales)
    try:
        type_id = losa.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            floor_type = losa.Document.GetElement(type_id)
            if floor_type:
                for pname in ("Default Thickness", "Thickness", "Espesor"):
                    param = floor_type.LookupParameter(pname)
                    if param and param.HasValue:
                        return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def obtener_losa_seleccionada():
    """Obtiene la primera losa (Floor) seleccionada."""
    if doc is None or uidoc is None:
        return None
    elem_ids = list(uidoc.Selection.GetElementIds())
    for elem_id in elem_ids:
        elem = doc.GetElement(elem_id)
        if elem is not None and isinstance(elem, Floor):
            return elem
    return None


def obtener_rebar_bar_types(documento):
    """Obtiene todos los RebarBarType del documento."""
    collector = FilteredElementCollector(documento)
    return list(collector.OfClass(RebarBarType))


def obtener_primer_rebar_bar_type(documento):
    """Obtiene el primer RebarBarType del documento."""
    bar_types = obtener_rebar_bar_types(documento)
    return bar_types[0] if bar_types else None


def obtener_hook_length_mm(bar_type, hook_type):
    """Obtiene el Hook Length en mm desde la tabla Hook Lengths del RebarBarType."""
    try:
        largo_interno = bar_type.GetHookLength(hook_type.Id)
        return UnitUtils.ConvertFromInternalUnits(largo_interno, UnitTypeId.Millimeters)
    except Exception:
        return None


def buscar_hook_por_largo(documento, largo_target_mm):
    """
    Busca un RebarHookType cuyo Hook Length sea igual a largo_target_mm
    en TODOS los RebarBarType. Solo retorna un gancho si todos los bar types
    tienen el largo correcto para ese hook.
    """
    bar_types = obtener_rebar_bar_types(documento)
    if not bar_types:
        return None
    for ht in FilteredElementCollector(documento).OfClass(RebarHookType):
        if ht is None:
            continue
        todos_coinciden = True
        for bar_type in bar_types:
            try:
                largo_mm = obtener_hook_length_mm(bar_type, ht)
                if largo_mm is None or abs(largo_mm - largo_target_mm) > TOLERANCIA_MM:
                    todos_coinciden = False
                    break
            except Exception:
                todos_coinciden = False
                break
        if todos_coinciden:
            return ht
    return None


def crear_hook_type_en_doc(documento, largo_hook_mm, nombre="Rebar Hook", angulo_grados=90.0,
                           multiplicador_extension=12.0, en_transaccion=True):
    """
    Crea un nuevo RebarHookType con el Hook Length indicado.
    Adaptado de crear_hook_desde_espesor_losa.py
    """
    bar_type = obtener_primer_rebar_bar_type(documento)
    if not bar_type:
        raise Exception("No hay RebarBarType en el documento. Crea al menos un tipo de barra de armadura.")

    nombres_existentes = [ht.Name for ht in FilteredElementCollector(documento).OfClass(RebarHookType) if ht and ht.Name]
    angulo_str = str(int(angulo_grados)) if angulo_grados == int(angulo_grados) else str(angulo_grados)

    t = Transaction(documento, "Crear Rebar Hook desde espesor losa") if en_transaccion else None
    if t:
        t.Start()

    try:
        angulo_rad = math.radians(angulo_grados)
        hook_type = RebarHookType.Create(documento, angulo_rad, multiplicador_extension)
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


def obtener_area_reinforcements_de_losa(documento, losa_id):
    """Obtiene los Area Reinforcements hospedados en la losa."""
    resultados = []
    for ar in FilteredElementCollector(documento).OfClass(AreaReinforcement):
        try:
            if ar.GetHostId() == losa_id:
                resultados.append(ar)
        except Exception:
            continue
    return resultados


def asignar_hook_a_area_reinforcement(area_rein, hook_type_id):
    """Asigna el RebarHookType al Area Reinforcement."""
    if not hook_type_id or hook_type_id == ElementId.InvalidElementId:
        return False
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


def main():
    """Flujo principal: verifica o crea RebarHookType según espesor de losa - 40 mm."""
    if doc is None or uidoc is None:
        print("Error: Ejecuta este script dentro de Revit Python Shell (RPS).")
        return None

    losa = obtener_losa_seleccionada()
    if losa is None:
        print("No hay losa seleccionada. Selecciona una losa (Floor) y vuelve a ejecutar.")
        return None

    espesor_mm = obtener_espesor_losa_mm(losa)
    if espesor_mm is None:
        print("No se pudo obtener el espesor de la losa seleccionada.")
        return None

    largo_hook_mm = max(0, espesor_mm - RESTA_MM)

    print("-" * 60)
    print("Verificar o crear Rebar Hook desde espesor de losa")
    print("-" * 60)
    print("Losa seleccionada: {} (ID: {})".format(losa.Name or "(sin nombre)", losa.Id.IntegerValue))
    print("Espesor de losa: {:.2f} mm".format(espesor_mm))
    print("Resta aplicada: {} mm".format(RESTA_MM))
    print("Hook Length objetivo: {:.2f} mm".format(largo_hook_mm))
    print("-" * 60)

    # Buscar hook existente con ese largo (en TODOS los RebarBarType)
    hook = buscar_hook_por_largo(doc, largo_hook_mm)

    if not hook:
        # No existe: crear uno nuevo
        print("No se encontró RebarHookType con Hook Length = {:.2f} mm. Creando uno nuevo...".format(largo_hook_mm))
        hook = crear_hook_type_en_doc(
            doc,
            largo_hook_mm=largo_hook_mm,
            nombre="Rebar Hook",
            angulo_grados=90.0,
            multiplicador_extension=12.0,
            en_transaccion=True
        )
        print("Tipo de gancho creado: '{}' (Hook Length: {:.2f} mm)".format(hook.Name, largo_hook_mm))
    else:
        print("OK: Se encontró RebarHookType existente: '{}' (Hook Length: {:.2f} mm)".format(
            hook.Name, largo_hook_mm))

    # Asignar el gancho a los Area Reinforcements de la losa
    area_reins = obtener_area_reinforcements_de_losa(doc, losa.Id)
    if area_reins:
        trans = Transaction(doc, "Asignar Hook a Area Reinforcements")
        trans.Start()
        try:
            asignados = 0
            for ar in area_reins:
                if asignar_hook_a_area_reinforcement(ar, hook.Id):
                    asignados += 1
            trans.Commit()
            print("Hook asignado a {} Area Reinforcement(s) de la losa.".format(asignados))
        except Exception as ex:
            trans.RollBack()
            print("Error al asignar gancho: {}".format(str(ex)))
    else:
        print("No hay Area Reinforcements en esta losa. El gancho está listo para usarse.")

    return hook


# ── Ejecución ───────────────────────────────────────────────────────────────
if __name__ == "__main__" or (doc is not None and uidoc is not None):
    main()
