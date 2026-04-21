# -*- coding: utf-8 -*-
"""
Script RPS/pyRevit: Asigna Rebar Hooks a Area Reinforcements según el espesor del host.
Los ganchos quedan contenidos en el espesor: Hook Length = Default Thickness - 40 mm.

Requisito: Seleccionar Area Reinforcements, o losas (Floor) / muros (Wall).
- Si se seleccionan Area Reinforcements: se obtiene el host de cada uno y su espesor.
- Si se seleccionan Floor/Wall: se buscan los Area Reinforcements hospedados y se les asigna el gancho.

Busca un RebarHookType existente con Hook Length = espesor - 40 mm.
Si no existe, crea uno nuevo (usa crear_hook_desde_espesor_losa.crear_hook_type_en_doc).
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
    Wall,
)
from Autodesk.Revit.DB.Structure import AreaReinforcement, RebarBarType, RebarHookType

try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None

RESTA_MM = 40
TOLERANCIA_MM = 0.5  # Tolerancia para comparar Hook Length


def obtener_espesor_host_mm(host, documento):
    """
    Obtiene el espesor del elemento host en mm.
    Prioriza BuiltInParameter (espesor real de instancia) y LookupParameter('Default Thickness') como fallback.
    """
    if host is None:
        return None
    # 1) BuiltInParameter: espesor real de la instancia (igual que crear_hook_desde_espesor_losa)
    try:
        if isinstance(host, Floor):
            p = host.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        elif isinstance(host, Wall):
            p = host.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        else:
            p = None
        if p and p.HasValue:
            return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 2) LookupParameter "Default Thickness" en el elemento
    try:
        p = host.LookupParameter("Default Thickness")
        if p and p.HasValue:
            return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    # 3) LookupParameter "Default Thickness" en el tipo
    try:
        type_id = host.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            host_type = documento.GetElement(type_id)
            if host_type:
                p = host_type.LookupParameter("Default Thickness")
                if p and p.HasValue:
                    return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
    except Exception:
        pass
    return None


def obtener_rebar_bar_types(documento):
    """Obtiene todos los RebarBarType del documento."""
    return list(FilteredElementCollector(documento).OfClass(RebarBarType))


def obtener_primer_rebar_bar_type(documento):
    """Obtiene el primer RebarBarType."""
    bar_types = obtener_rebar_bar_types(documento)
    return bar_types[0] if bar_types else None


def obtener_hook_length_desde_parametro_mm(bar_type, hook_type, documento):
    """
    Obtiene el Hook Length en mm desde el parámetro 'Hook Lengths' del RebarBarType.
    GetHookLength lee de la tabla Hook Lengths del RebarBarType.
    """
    try:
        largo_interno = bar_type.GetHookLength(hook_type.Id)
        return UnitUtils.ConvertFromInternalUnits(largo_interno, UnitTypeId.Millimeters)
    except Exception:
        return None


def buscar_hook_por_largo(documento, largo_target_mm):
    """
    Busca un RebarHookType cuyo Hook Length (parámetro Hook Lengths) sea igual a largo_target_mm
    en TODOS los RebarBarType. El Hook Length es por bar type; si solo un bar type tiene 140mm
    pero otro tiene 50mm, al asignar el gancho el AR podría usar el bar type con 50mm.
    Solo retornamos un gancho si TODOS los bar types tienen el largo correcto.
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
                largo_mm = obtener_hook_length_desde_parametro_mm(bar_type, ht, documento)
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
        raise Exception("No hay RebarBarType en el documento.")
    nombres_existentes = [ht.Name for ht in FilteredElementCollector(documento).OfClass(RebarHookType) if ht and ht.Name]
    angulo_str = str(int(angulo_grados)) if angulo_grados == int(angulo_grados) else str(angulo_grados)
    t = Transaction(documento, "Crear Rebar Hook desde espesor host") if en_transaccion else None
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
        hook_type.Name = nombre_final
        if t:
            t.Commit()
        return hook_type
    except Exception as ex:
        if t:
            t.RollBack()
        raise


def obtener_o_crear_hook(documento, largo_hook_mm, en_transaccion=True):
    """
    Busca un RebarHookType con Hook Length = largo_hook_mm.
    Si no existe, crea uno nuevo.
    Retorna RebarHookType o None.
    """
    hook = buscar_hook_por_largo(documento, largo_hook_mm)
    if hook:
        return hook
    return crear_hook_type_en_doc(documento, largo_hook_mm, en_transaccion=en_transaccion)


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


def obtener_area_reinforcements_y_hosts(uidocumento, documento):
    """
    Obtiene (AreaReinforcement, host) para cada elemento seleccionado.
    Si se selecciona AreaReinforcement: (ar, host).
    Si se selecciona Floor/Wall: busca Area Reinforcements con ese host.
    """
    elem_ids = list(uidocumento.Selection.GetElementIds())
    if not elem_ids:
        return []
    resultados = []
    for eid in elem_ids:
        elem = documento.GetElement(eid)
        if elem is None:
            continue
        if isinstance(elem, AreaReinforcement):
            try:
                host_id = elem.GetHostId()
                host = documento.GetElement(host_id) if host_id else None
                if host:
                    resultados.append((elem, host))
            except Exception:
                pass
        elif isinstance(elem, (Floor, Wall)):
            host = elem
            for ar in FilteredElementCollector(documento).OfClass(AreaReinforcement):
                try:
                    if ar.GetHostId() == host.Id:
                        resultados.append((ar, host))
                except Exception:
                    continue
    return resultados


def ejecutar():
    """Flujo principal."""
    if doc is None or uidoc is None:
        print("Error: Ejecuta en Revit Python Shell o pyRevit.")
        return
    pares = obtener_area_reinforcements_y_hosts(uidoc, doc)
    if not pares:
        print("Error: Selecciona Area Reinforcements o losas (Floor) / muros (Wall) y vuelve a ejecutar.")
        return
    trans = Transaction(doc, "Asignar Hook desde espesor host")
    trans.Start()
    try:
        asignados = 0
        for ar, host in pares:
            espesor_mm = obtener_espesor_host_mm(host, doc)
            if espesor_mm is None:
                print("Advertencia: No se pudo obtener espesor del host (ID: {}). Se omite.".format(host.Id.IntegerValue))
                continue
            target_hook_mm = max(0, espesor_mm - RESTA_MM)
            hook = obtener_o_crear_hook(doc, target_hook_mm, en_transaccion=False)
            if hook and asignar_hook_a_area_reinforcement(ar, hook.Id):
                asignados += 1
        trans.Commit()
        print("OK: Hook asignado a {} Area Reinforcement(s). Espesor - {} mm = Hook Length.".format(
            asignados, RESTA_MM))
    except Exception as ex:
        trans.RollBack()
        print("Error: {}".format(str(ex)))
        raise


if __name__ == "__main__" or (doc is not None and uidoc is not None):
    ejecutar()
