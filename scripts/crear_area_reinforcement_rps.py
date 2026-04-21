# -*- coding: utf-8 -*-
"""
Script RPS: Crear AreaReinforcement con Create(doc, host, curves, layoutDirection, areaTypeId, barTypeId, hookTypeId)
Ejecutable en RevitPythonShell (RPS) — Revit 2024+ | IronPython 3.4

Requisito: Tener seleccionada una losa (Floor) o muro (Wall) antes de ejecutar.
"""

import math
import os
import sys
import clr

_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    CurveLoop,
    ElementId,
    ElementTypeGroup,
    FilteredElementCollector,
    Floor,
    Line,
    Sketch,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    AreaReinforcementType,
    RebarBarType,
    RebarHookType,
)
from Autodesk.Revit.UI import TaskDialog

# ── Boilerplate RPS ─────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


def aplicar_offset_recubrimiento(curves, offset_mm, document):
    """Aplica un offset hacia adentro (recubrimiento) a las curvas. offset_mm en mm."""
    if not curves or len(curves) < 2:
        return curves
    try:
        offset_internal = UnitUtils.ConvertToInternalUnits(offset_mm, UnitTypeId.Millimeters)
        curve_list = List[Curve](curves)
        loop = CurveLoop.Create(curve_list)
        if loop is None:
            return curves
        if not loop.HasPlane():
            return curves

        normal = XYZ(0, 0, 1)
        # CreateViaOffset es sensible al sentido del CurveLoop (CW/CCW) y a la normal.
        # Para garantizar "offset hacia adentro", probamos ambos signos y escogemos el
        # que reduce el perímetro (longitud total) respecto al original.
        try:
            base_len = loop.GetExactLength()
        except Exception:
            base_len = None

        offset_candidates = []
        for sign in (1.0, -1.0):
            try:
                cl = CurveLoop.CreateViaOffset(loop, sign * offset_internal, normal)
                if cl is None:
                    continue
                try:
                    cand_len = cl.GetExactLength()
                except Exception:
                    cand_len = None
                offset_candidates.append((cl, cand_len))
            except Exception:
                continue

        if not offset_candidates:
            return curves

        # Preferir el candidato con menor longitud (más "hacia adentro").
        # Si no podemos comparar longitudes, usar el primero válido.
        if base_len is not None:
            best = None
            best_len = None
            for cl, cand_len in offset_candidates:
                if cand_len is None:
                    continue
                if best is None or cand_len < best_len:
                    best = cl
                    best_len = cand_len
            if best is None:
                best = offset_candidates[0][0]
        else:
            best = offset_candidates[0][0]

        return [c for c in best]
    except Exception:
        return curves


def obtener_curvas_floor(floor, document):
    """Obtiene las curvas del perímetro exterior del sketch de la losa."""
    try:
        sketch_id = floor.SketchId
        if sketch_id is None or sketch_id == ElementId.InvalidElementId:
            return None
        sketch = document.GetElement(sketch_id)
        if sketch is None or not isinstance(sketch, Sketch):
            return None
        profile = sketch.Profile
        if profile is None:
            return None
        n_loops = profile.Size
        if n_loops < 1:
            return None
        curve_array = profile.get_Item(0)
        if curve_array is None:
            return None
        curves = []
        n_curves = curve_array.Size
        for j in range(n_curves):
            c = curve_array.get_Item(j)
            if c is not None:
                curves.append(c)
        if not curves:
            return None
        return aplicar_offset_recubrimiento(curves, 20.0, document) or curves
    except Exception:
        return None


def obtener_curvas_wall(wall, document):
    """Obtiene las curvas del contorno en planta del muro (bounding box base)."""
    try:
        bbox = wall.get_BoundingBox(None)
        if bbox is None:
            return None
        min_pt = bbox.Min
        max_pt = bbox.Max
        z = min_pt.Z
        p0 = XYZ(min_pt.X, min_pt.Y, z)
        p1 = XYZ(max_pt.X, min_pt.Y, z)
        p2 = XYZ(max_pt.X, max_pt.Y, z)
        p3 = XYZ(min_pt.X, max_pt.Y, z)
        curves = [
            Line.CreateBound(p0, p1),
            Line.CreateBound(p1, p2),
            Line.CreateBound(p2, p3),
            Line.CreateBound(p3, p0),
        ]
        return aplicar_offset_recubrimiento(curves, 20.0, document) or curves
    except Exception:
        return None


def obtener_direccion(curves):
    """Calcula XYZ de dirección principal desde la primera curva."""
    try:
        if curves and len(curves) > 0:
            first = curves[0]
            p0 = first.GetEndPoint(0)
            p1 = first.GetEndPoint(1)
            dx = p1.X - p0.X
            dy = p1.Y - p0.Y
            dz = p1.Z - p0.Z
            length = (dx * dx + dy * dy + dz * dz) ** 0.5
            if length > 1e-6:
                return XYZ(dx / length, dy / length, dz / length)
        return XYZ(1, 0, 0)
    except Exception:
        return XYZ(1, 0, 0)


def get_area_reinforcement_type_id(document):
    """Obtiene el primer AreaReinforcementType del proyecto."""
    try:
        default_id = document.GetDefaultElementTypeId(ElementTypeGroup.AreaReinforcementType)
        if default_id and default_id != ElementId.InvalidElementId:
            return default_id
    except Exception:
        pass
    try:
        for elem in FilteredElementCollector(document).OfClass(AreaReinforcementType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return None


def get_first_rebar_bar_type_id(document):
    """Obtiene el ID del primer RebarBarType."""
    try:
        for elem in FilteredElementCollector(document).OfClass(RebarBarType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return None


def get_first_rebar_hook_type_id(document):
    """Obtiene el ID del primer RebarHookType (o InvalidElementId si no hay)."""
    try:
        for elem in FilteredElementCollector(document).OfClass(RebarHookType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return ElementId.InvalidElementId


def asignar_hook_a_area_reinforcement(area_rein, hook_type_id):
    """
    Asigna el RebarHookType al inicio y final de todas las capas del Area Reinforcement.
    El Create(7 params) acepta hookTypeId pero Revit NO lo aplica a las capas;
    hay que asignar explícitamente estos parámetros.
    Usa BuiltInParameter (invariante al idioma) + LookupParameter + fallback por iteración.
    """
    if not area_rein or not hook_type_id or hook_type_id == ElementId.InvalidElementId:
        return
    asignados = 0
    # 1) BuiltInParameter (invariante al idioma)
    bip_names = [
        "REBAR_SYSTEM_HOOK_TYPE_MAJOR_TOP", "REBAR_SYSTEM_HOOK_TYPE_MAJOR_BOTTOM",
        "REBAR_SYSTEM_HOOK_TYPE_MINOR_TOP", "REBAR_SYSTEM_HOOK_TYPE_MINOR_BOTTOM",
        "REBAR_SYSTEM_HOOK_TYPE_EXTERIOR_MAJOR", "REBAR_SYSTEM_HOOK_TYPE_EXTERIOR_MINOR",
        "REBAR_SYSTEM_HOOK_TYPE_INTERIOR_MAJOR", "REBAR_SYSTEM_HOOK_TYPE_INTERIOR_MINOR",
        "REBAR_SYSTEM_HOOK_TYPE_TOP_DIR_1", "REBAR_SYSTEM_HOOK_TYPE_TOP_DIR_2",
        "REBAR_SYSTEM_HOOK_TYPE_BOTTOM_DIR_1", "REBAR_SYSTEM_HOOK_TYPE_BOTTOM_DIR_2",
    ]
    for name in bip_names:
        try:
            bip = getattr(BuiltInParameter, name, None)
            if bip is not None:
                p = area_rein.get_Parameter(bip)
                if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                    p.Set(hook_type_id)
                    asignados += 1
        except Exception:
            continue
    # 2) LookupParameter por nombre (inglés)
    if asignados == 0:
        hook_param_names = [
            u"Exterior Major Hook Type", u"Top Major Hook Type",
            u"Exterior Minor Hook Type", u"Top Minor Hook Type",
            u"Interior Major Hook Type", u"Bottom Major Hook Type",
            u"Interior Minor Hook Type", u"Bottom Minor Hook Type",
        ]
        for pname in hook_param_names:
            try:
                p = area_rein.LookupParameter(pname)
                if p and not p.IsReadOnly and p.StorageType == StorageType.ElementId:
                    p.Set(hook_type_id)
                    asignados += 1
            except Exception:
                continue
    # 3) Fallback: iterar parámetros y asignar los que sean tipo ElementId y contengan "ook" (Hook/Gancho)
    if asignados == 0:
        try:
            for p in area_rein.Parameters:
                if p is None or p.IsReadOnly or p.StorageType != StorageType.ElementId:
                    continue
                try:
                    nombre = p.Definition.Name if p.Definition else ""
                    if "ook" in nombre.lower() or "gancho" in nombre.lower():
                        p.Set(hook_type_id)
                        asignados += 1
                except Exception:
                    continue
        except Exception:
            pass


def crear_gancho_por_defecto(document):
    """
    Crea un RebarHookType por defecto (90°, extensión 12× diámetro) si no existe ninguno.
    Debe ejecutarse dentro de una Transaction ya iniciada.
    Retorna ElementId del gancho creado.
    """
    angulo_rad = math.radians(90.0)
    multiplicador = 12.0
    hook_type = RebarHookType.Create(document, angulo_rad, multiplicador)
    largo_mm = 50.0
    largo_interno = UnitUtils.ConvertToInternalUnits(largo_mm, UnitTypeId.Millimeters)
    try:
        for bt in FilteredElementCollector(document).OfClass(RebarBarType):
            bt.SetAutoCalcHookLengths(hook_type.Id, False)
            bt.SetHookLength(hook_type.Id, largo_interno)
    except Exception:
        pass
    try:
        hook_type.Name = u"Rebar Hook - 90º - 50.0 mm (por defecto)"
    except Exception:
        pass
    return hook_type.Id


def _etiquetar_area_rein_en_planta_si_aplica(document, uidocument, area_rein):
    """
    Si la vista activa es planta (FloorPlan) y admite etiquetas, crea IndependentTag
    por categoría (misma lógica que area_reinforcement_losa / etiquetar_area_reinforcement_rps).
    """
    try:
        from area_reinforcement_losa import (
            _crear_etiqueta_area_reinforcement,
            _es_vista_planta_area_reinforcement,
            _vista_valida_etiqueta_area_reinforcement,
        )
    except ImportError:
        return
    view = uidocument.ActiveView
    if not _es_vista_planta_area_reinforcement(view):
        return
    ok, _ = _vista_valida_etiqueta_area_reinforcement(view)
    if not ok:
        return
    try:
        document.Regenerate()
    except Exception:
        pass
    t = Transaction(document, u"Etiquetar Area Reinforcement (vista planta)")
    try:
        t.Start()
        _crear_etiqueta_area_reinforcement(document, view, area_rein)
        t.Commit()
        print(u"Etiqueta de Area Reinforcement creada en la vista de planta activa.")
    except Exception as ex:
        if t.HasStarted():
            try:
                t.RollBack()
            except Exception:
                pass
        print(u"Aviso: no se pudo crear la etiqueta en planta: {}".format(str(ex)))


def run(document, uidocument):
    """Ejecuta la lógica principal. Recibe document y uidocument para compatibilidad con pyRevit."""
    doc = document
    uidoc = uidocument
    try:
        TaskDialog.Show("Crear Area Reinforcement", "Seleccionar losa a Enfierrar.")
        elem_ids = list(uidoc.Selection.GetElementIds())
        if not elem_ids:
            print("Error: No hay ningún elemento seleccionado. Selecciona una losa (Floor) o muro (Wall) y vuelve a ejecutar.")
        else:
            host = doc.GetElement(elem_ids[0])
            if host is None:
                print("Error: No se pudo obtener el elemento seleccionado.")
            elif isinstance(host, Floor):
                curves = obtener_curvas_floor(host, doc)
                if not curves or len(curves) == 0:
                    print("Error: La losa no tiene sketch válido o no se pudieron obtener las curvas.")
                else:
                    layout_dir = obtener_direccion(curves)
                    area_type_id = get_area_reinforcement_type_id(doc)
                    bar_type_id = get_first_rebar_bar_type_id(doc)
                    hook_type_id = get_first_rebar_hook_type_id(doc)
                    if not area_type_id or not bar_type_id:
                        print("Error: No hay AreaReinforcementType o RebarBarType en el proyecto.")
                    else:
                        curve_list = List[Curve](curves)
                        trans = Transaction(doc, "Crear AreaReinforcement (7 params)")
                        trans.Start()
                        try:
                            if not hook_type_id or hook_type_id == ElementId.InvalidElementId:
                                hook_type_id = crear_gancho_por_defecto(doc)
                            ar = AreaReinforcement.Create(
                                doc, host, curve_list, layout_dir,
                                area_type_id, bar_type_id, hook_type_id
                            )
                            asignar_hook_a_area_reinforcement(ar, hook_type_id)
                            trans.Commit()
                            print("AreaReinforcement creado correctamente (ID: {})".format(ar.Id.IntegerValue))
                            _etiquetar_area_rein_en_planta_si_aplica(doc, uidoc, ar)
                        except Exception as ex:
                            trans.RollBack()
                            print("Error al crear AreaReinforcement: {}".format(str(ex)))
            elif isinstance(host, Wall):
                curves = obtener_curvas_wall(host, doc)
                if not curves or len(curves) == 0:
                    print("Error: No se pudieron obtener las curvas del muro.")
                else:
                    layout_dir = obtener_direccion(curves)
                    area_type_id = get_area_reinforcement_type_id(doc)
                    bar_type_id = get_first_rebar_bar_type_id(doc)
                    hook_type_id = get_first_rebar_hook_type_id(doc)
                    if not area_type_id or not bar_type_id:
                        print("Error: No hay AreaReinforcementType o RebarBarType en el proyecto.")
                    else:
                        curve_list = List[Curve](curves)
                        trans = Transaction(doc, "Crear AreaReinforcement (7 params)")
                        trans.Start()
                        try:
                            if not hook_type_id or hook_type_id == ElementId.InvalidElementId:
                                hook_type_id = crear_gancho_por_defecto(doc)
                            ar = AreaReinforcement.Create(
                                doc, host, curve_list, layout_dir,
                                area_type_id, bar_type_id, hook_type_id
                            )
                            asignar_hook_a_area_reinforcement(ar, hook_type_id)
                            trans.Commit()
                            print("AreaReinforcement creado correctamente (ID: {})".format(ar.Id.IntegerValue))
                            _etiquetar_area_rein_en_planta_si_aplica(doc, uidoc, ar)
                        except Exception as ex:
                            trans.RollBack()
                            print("Error al crear AreaReinforcement: {}".format(str(ex)))
            else:
                print("Error: El elemento seleccionado no es una losa (Floor) ni un muro (Wall). Tipo: {}".format(type(host).__name__))
    except Exception as ex:
        print("Error: {}".format(str(ex)))


# ── Ejecución al cargar (RPS directo o pyRevit) ──────────────────────────────
if doc is not None and uidoc is not None:
    run(doc, uidoc)
