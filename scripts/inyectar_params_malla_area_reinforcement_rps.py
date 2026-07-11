# -*- coding: utf-8 -*-
"""
RevitPythonShell (RPS) — Revit 2025 | IronPython 2.7 / 3.x

Selecciona todos los Area Reinforcement del proyecto e inyecta en sus barras
(Rebar / RebarInSystem) los parámetros de malla:

  - Armadura_Malla     = Yes
  - Armadura_Ubicacion = F' (malla superior / Top) · F (malla inferior / Bottom)
  - Armadura_Nivel     = nombre del nivel de la losa host

La lógica de ubicación y nivel replica la de
BIMTools.tab/Armadura.panel/08_CrearAreaReinforcementRPS.pushbutton
(``area_reinforcement_losa``).

IMPORTANTE: en RPS usar File > Run script (o execfile / #load).
No pegar el archivo en la consola interactiva (>>>): las lineas en blanco
cierran los bloques y producen "expected an indented block".
"""

import os
import sys

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Floor,
    GeometryInstance,
    Options,
    Solid,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    MultiplanarOption,
    Rebar,
    RebarInSystem,
)
from Autodesk.Revit.UI import TaskDialog

# ── Documento activo (RPS / pyRevit) ─────────────────────────────────────────
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

ARMADURA_MALLA_PARAM = u"Armadura_Malla"
ARMADURA_UBICACION_PARAM = u"Armadura_Ubicacion"
ARMADURA_NIVEL_PARAM = u"Armadura_Nivel"
ARMADURA_UBICACION_INFERIOR = u"F"
ARMADURA_UBICACION_SUPERIOR = u"F'"
_TX_NAME = u"Arainco: Inyectar params malla Area Reinforcement"


# ── sys.path → scripts/ (para importar conjunto_guid si está disponible) ─────
def _script_containing_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return None


def _ensure_scripts_on_path():
    candidates = []
    d0 = _script_containing_dir()
    if d0:
        candidates.append(d0)
    try:
        home = os.path.expanduser(u"~")
        guess = os.path.join(
            home, u"CustomRevitExtensions", u"BIMTools.extension", u"scripts",
        )
        if os.path.isdir(guess):
            candidates.append(guess)
    except Exception:
        pass
    for p in candidates:
        if p and p not in sys.path and os.path.isdir(p):
            sys.path.insert(0, p)


_ensure_scripts_on_path()

try:
    from conjunto_guid import (
        stamp_armadura_malla as _stamp_malla_ext,
        stamp_armadura_ubicacion as _stamp_ubic_ext,
        stamp_armadura_nivel as _stamp_nivel_ext,
    )
except Exception:
    _stamp_malla_ext = None
    _stamp_ubic_ext = None
    _stamp_nivel_ext = None


# ── Helpers de parámetros ────────────────────────────────────────────────────
def _as_text(value):
    """str/unicode compatible IronPython 2.7 y 3.4."""
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def _norm_param_def_name(name):
    if name is None:
        return u""
    try:
        return _as_text(name).replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def _find_element_parameter(element, param_name):
    if element is None or not param_name:
        return None
    target = _norm_param_def_name(param_name).lower()
    try:
        p = element.LookupParameter(param_name)
        if p is not None:
            return p
    except Exception:
        pass
    try:
        for p in element.Parameters:
            if p is None:
                continue
            try:
                dn = _norm_param_def_name(p.Definition.Name).lower()
            except Exception:
                continue
            if dn == target:
                return p
    except Exception:
        pass
    return None


def _set_yes_no(element, param_name, yes=True):
    p = _find_element_parameter(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    if yes:
        candidates = (1, True, u"1", u"Yes", u"yes", u"Si", u"SI")
    else:
        candidates = (0, False, u"0", u"No", u"no")
    try:
        if p.StorageType == StorageType.Integer:
            p.Set(1 if yes else 0)
            return True
    except Exception:
        pass
    for val in candidates:
        try:
            p.Set(val)
            return True
        except Exception:
            continue
    try:
        p.SetValueString(u"Yes" if yes else u"No")
        return True
    except Exception:
        return False


def _set_string(element, param_name, valor):
    if not valor:
        return False
    p = _find_element_parameter(element, param_name)
    if p is None or p.IsReadOnly:
        return False
    try:
        if p.StorageType == StorageType.String:
            p.Set(valor)
            return True
    except Exception:
        pass
    try:
        p.Set(valor)
        return True
    except Exception:
        pass
    try:
        p.SetValueString(valor)
        return True
    except Exception:
        return False


def stamp_armadura_malla(element, yes=True):
    if _stamp_malla_ext is not None:
        try:
            return bool(_stamp_malla_ext(element, yes=yes))
        except Exception:
            pass
    return _set_yes_no(element, ARMADURA_MALLA_PARAM, yes=yes)


def stamp_armadura_ubicacion(element, valor):
    if _stamp_ubic_ext is not None:
        try:
            return bool(_stamp_ubic_ext(element, valor))
        except Exception:
            pass
    return _set_string(element, ARMADURA_UBICACION_PARAM, valor)


def stamp_armadura_nivel(element, valor):
    if _stamp_nivel_ext is not None:
        try:
            return bool(_stamp_nivel_ext(element, valor))
        except Exception:
            pass
    return _set_string(element, ARMADURA_NIVEL_PARAM, valor)


# ── IDs / geometría ──────────────────────────────────────────────────────────
def _element_id_int(eid):
    if eid is None:
        return None
    try:
        v = getattr(eid, "Value", None)
        if v is not None:
            return int(v)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        pass
    return None


def _iter_solids(element):
    if element is None:
        return
    opts = Options()
    opts.ComputeReferences = False
    try:
        geom = element.get_Geometry(opts)
    except Exception:
        return
    if geom is None:
        return
    for obj in geom:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Faces.Size > 0:
            yield obj
            continue
        if isinstance(obj, GeometryInstance):
            try:
                inst = obj.GetInstanceGeometry()
                if inst:
                    for g in inst:
                        if isinstance(g, Solid) and g.Faces.Size > 0:
                            yield g
            except Exception:
                pass


def _espesor_losa_mm(floor):
    if floor is None:
        return None
    try:
        param = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(
                param.AsDouble(), UnitTypeId.Millimeters,
            )
    except Exception:
        pass
    try:
        param = floor.LookupParameter(u"Default Thickness")
        if param and param.HasValue:
            return UnitUtils.ConvertFromInternalUnits(
                param.AsDouble(), UnitTypeId.Millimeters,
            )
    except Exception:
        pass
    try:
        type_id = floor.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            ftype = doc.GetElement(type_id)
            if ftype is not None:
                param = ftype.LookupParameter(u"Default Thickness")
                if param and param.HasValue:
                    return UnitUtils.ConvertFromInternalUnits(
                        param.AsDouble(), UnitTypeId.Millimeters,
                    )
    except Exception:
        pass
    return None


def _obtener_z_caras_losa(floor):
    """Z superior / inferior de la losa (misma lógica que area_reinforcement_losa)."""
    if floor is None or not isinstance(floor, Floor):
        return None, None
    z_top = None
    z_bottom = None
    try:
        from Autodesk.Revit.DB import PlanarFace
    except Exception:
        PlanarFace = type(None)
    for solid in _iter_solids(floor):
        try:
            n_faces = int(solid.Faces.Size)
        except Exception:
            continue
        for fi in range(n_faces):
            try:
                face = solid.Faces.get_Item(fi)
            except Exception:
                try:
                    face = solid.Faces[fi]
                except Exception:
                    continue
            if not isinstance(face, PlanarFace):
                continue
            try:
                nz = float(face.FaceNormal.Z)
            except Exception:
                continue
            try:
                z_face = float(face.Origin.Z)
            except Exception:
                try:
                    fbb = face.GetBoundingBox()
                    z_face = (float(fbb.Min.Z) + float(fbb.Max.Z)) * 0.5
                except Exception:
                    continue
            if nz >= 0.9:
                if z_top is None or z_face > z_top:
                    z_top = z_face
            elif nz <= -0.9:
                if z_bottom is None or z_face < z_bottom:
                    z_bottom = z_face
    if z_top is not None and z_bottom is not None:
        return z_top, z_bottom
    esp_mm = _espesor_losa_mm(floor)
    if z_top is not None and esp_mm is not None:
        esp_ft = UnitUtils.ConvertToInternalUnits(float(esp_mm), UnitTypeId.Millimeters)
        return z_top, z_top - esp_ft
    try:
        bb = floor.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Max.Z), float(bb.Min.Z)
    except Exception:
        pass
    return None, None


def _host_losa(area_rein):
    if area_rein is None:
        return None
    try:
        hid = area_rein.GetHostId()
    except Exception:
        return None
    if hid is None or hid == ElementId.InvalidElementId:
        return None
    try:
        host = doc.GetElement(hid)
    except Exception:
        return None
    return host if isinstance(host, Floor) else None


def _nivel_losa_como_string(floor):
    """Nombre del nivel de la losa host (misma lógica que area_reinforcement_losa)."""
    if floor is None:
        return None
    lid = None
    try:
        lid = floor.LevelId
        if lid is None or lid == ElementId.InvalidElementId:
            lid = None
    except Exception:
        lid = None
    if lid is None:
        for bip_name in (
            u"INSTANCE_REFERENCE_LEVEL_PARAM",
            u"LEVEL_PARAM",
            u"SCHEDULE_LEVEL_PARAM",
        ):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                p = floor.get_Parameter(bip)
                if p is None or not p.HasValue or p.StorageType != StorageType.ElementId:
                    continue
                eid = p.AsElementId()
                if eid is not None and eid != ElementId.InvalidElementId:
                    lid = eid
                    break
            except Exception:
                pass
    if lid is None:
        return None
    try:
        level = doc.GetElement(lid)
        if level is None or level.Name is None:
            return None
        return _as_text(level.Name)
    except Exception:
        return None


def _z_promedio_barra(barra):
    if barra is None:
        return None
    zs = []
    for include_all in (True, False):
        try:
            opt = (
                MultiplanarOption.IncludeAllMultiplanarCurves
                if include_all
                else MultiplanarOption.IncludeOnlyPlanarCurves
            )
            curves = barra.GetCenterlineCurves(False, False, False, opt, 0)
            if curves is None:
                continue
            try:
                n = int(curves.Count)
            except Exception:
                n = 0
            for i in range(n):
                try:
                    c = curves[i]
                    zs.append(float(c.GetEndPoint(0).Z))
                    zs.append(float(c.GetEndPoint(1).Z))
                except Exception:
                    continue
            if zs:
                break
        except Exception:
            continue
    if not zs:
        try:
            bb = barra.get_BoundingBox(None)
            if bb is not None:
                return (float(bb.Max.Z) + float(bb.Min.Z)) * 0.5
        except Exception:
            pass
        return None
    return sum(zs) / float(len(zs))


def _ubicacion_por_z(z_barra, z_top, z_bottom):
    if z_barra is None:
        return None
    if z_top is not None and z_bottom is not None:
        z_mid = (float(z_top) + float(z_bottom)) * 0.5
        if float(z_barra) >= z_mid:
            return ARMADURA_UBICACION_SUPERIOR
        return ARMADURA_UBICACION_INFERIOR
    if z_top is not None:
        if float(z_barra) >= float(z_top) - 1e-9:
            return ARMADURA_UBICACION_SUPERIOR
        return ARMADURA_UBICACION_INFERIOR
    if z_bottom is not None:
        if float(z_barra) <= float(z_bottom) + 1e-9:
            return ARMADURA_UBICACION_INFERIOR
        return ARMADURA_UBICACION_SUPERIOR
    return None


def _ubicacion_por_geometria_cara_losa(barra, z_top, z_bottom):
    """
    F' si la barra está en la mitad superior de la losa; F si está en la inferior.
    Misma heurística que ``_ubicacion_por_geometria_cara_losa`` en area_reinforcement_losa.
    """
    if z_top is None or z_bottom is None:
        return None
    z_mid = (float(z_top) + float(z_bottom)) * 0.5
    try:
        bb = barra.get_BoundingBox(None)
        if bb is not None:
            z_max = float(bb.Max.Z)
            z_min = float(bb.Min.Z)
            if z_min >= z_mid:
                return ARMADURA_UBICACION_SUPERIOR
            if z_max <= z_mid:
                return ARMADURA_UBICACION_INFERIOR
            z_centro = (z_max + z_min) * 0.5
            return _ubicacion_por_z(z_centro, z_top, z_bottom)
    except Exception:
        pass
    return _ubicacion_por_z(_z_promedio_barra(barra), z_top, z_bottom)


# ── Colección de AreaReinforcement y barras ──────────────────────────────────
def collect_all_area_reinforcements(document):
    return list(
        FilteredElementCollector(document)
        .OfClass(AreaReinforcement)
        .WhereElementIsNotElementType()
    )


def _collect_rebars(document, area_rein):
    out = []
    if area_rein is None or document is None:
        return out
    try:
        from Autodesk.Revit.DB import ElementCategoryFilter

        flt = ElementCategoryFilter(BuiltInCategory.OST_Rebar)
        dep = area_rein.GetDependentElements(flt)
        if dep is None:
            return out
        for i in range(int(dep.Count)):
            try:
                el = document.GetElement(dep[i])
                if isinstance(el, Rebar):
                    out.append(el)
            except Exception:
                continue
    except Exception:
        pass
    return out


def _collect_rebar_in_system(document, area_rein):
    barras = []
    if area_rein is None or document is None:
        return barras
    seen = set()
    try:
        sys_ids = area_rein.GetRebarInSystemIds()
    except Exception:
        sys_ids = None
    if sys_ids is not None:
        try:
            nd = int(sys_ids.Count)
        except Exception:
            nd = 0
        for i in range(nd):
            try:
                eid = sys_ids[i]
                eid_int = _element_id_int(eid)
                if eid_int is None or eid_int in seen:
                    continue
                el = document.GetElement(eid)
                if isinstance(el, RebarInSystem):
                    barras.append(el)
                    seen.add(eid_int)
            except Exception:
                continue
    if barras:
        return barras
    try:
        from Autodesk.Revit.DB import ElementCategoryFilter

        flt = ElementCategoryFilter(BuiltInCategory.OST_RebarInSystem)
        dep = area_rein.GetDependentElements(flt)
        if dep is None:
            return barras
        for i in range(int(dep.Count)):
            try:
                eid_int = _element_id_int(dep[i])
                if eid_int is None or eid_int in seen:
                    continue
                el = document.GetElement(dep[i])
                if isinstance(el, RebarInSystem):
                    barras.append(el)
                    seen.add(eid_int)
            except Exception:
                continue
    except Exception:
        pass
    return barras


def collect_barras_de_area(document, area_rein):
    """Lista de (barra, tipo_str) sin duplicados."""
    barras = []
    seen = set()
    for rb in _collect_rebars(document, area_rein):
        eid = _element_id_int(rb.Id)
        if eid is None or eid in seen:
            continue
        barras.append((rb, u"Rebar"))
        seen.add(eid)
    for rb in _collect_rebar_in_system(document, area_rein):
        eid = _element_id_int(rb.Id)
        if eid is None or eid in seen:
            continue
        barras.append((rb, u"RebarInSystem"))
        seen.add(eid)
    return barras


def seleccionar_elementos(elementos):
    ids = List[ElementId]()
    for el in elementos or []:
        if el is None:
            continue
        try:
            ids.Add(el.Id)
        except Exception:
            pass
    if ids.Count > 0:
        uidoc.Selection.SetElementIds(ids)
    return int(ids.Count)


# ── Proceso principal ────────────────────────────────────────────────────────
def inyectar_params_malla_en_proyecto(document):
    areas = collect_all_area_reinforcements(document)
    stats = {
        u"n_areas": len(areas),
        u"n_barras": 0,
        u"n_malla_ok": 0,
        u"n_ubicacion_ok": 0,
        u"n_nivel_ok": 0,
        u"n_sin_barras": 0,
        u"n_sin_nivel": 0,
        u"n_sin_ubicacion": 0,
        u"areas_detalle": [],
    }
    if not areas:
        return stats

    t = Transaction(document, _TX_NAME)
    t.Start()
    try:
        document.Regenerate()
        for ar in areas:
            if ar is None:
                continue
            ar_id = _element_id_int(ar.Id)
            floor = _host_losa(ar)
            nivel_valor = _nivel_losa_como_string(floor)
            z_top, z_bottom = _obtener_z_caras_losa(floor)
            barras = collect_barras_de_area(document, ar)
            if not barras:
                stats[u"n_sin_barras"] += 1
            if not nivel_valor:
                stats[u"n_sin_nivel"] += 1

            n_m = n_u = n_n = 0
            for barra, _tipo in barras:
                stats[u"n_barras"] += 1
                ubicacion = _ubicacion_por_geometria_cara_losa(barra, z_top, z_bottom)
                if ubicacion is None:
                    stats[u"n_sin_ubicacion"] += 1

                if stamp_armadura_malla(barra, yes=True):
                    n_m += 1
                    stats[u"n_malla_ok"] += 1
                if ubicacion and stamp_armadura_ubicacion(barra, ubicacion):
                    n_u += 1
                    stats[u"n_ubicacion_ok"] += 1
                if nivel_valor and stamp_armadura_nivel(barra, nivel_valor):
                    n_n += 1
                    stats[u"n_nivel_ok"] += 1

            stats[u"areas_detalle"].append({
                u"area_id": ar_id,
                u"n_barras": len(barras),
                u"nivel": nivel_valor,
                u"malla": n_m,
                u"ubicacion": n_u,
                u"nivel_ok": n_n,
            })
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
        raise
    return stats


def _format_resumen(stats):
    lines = [
        u"Area Reinforcement: {0}".format(stats.get(u"n_areas") or 0),
        u"Barras procesadas: {0}".format(stats.get(u"n_barras") or 0),
        u"Armadura_Malla=Yes: {0}".format(stats.get(u"n_malla_ok") or 0),
        u"Armadura_Ubicacion (F/F'): {0}".format(stats.get(u"n_ubicacion_ok") or 0),
        u"Armadura_Nivel: {0}".format(stats.get(u"n_nivel_ok") or 0),
    ]
    if stats.get(u"n_sin_barras"):
        lines.append(
            u"Áreas sin barras: {0}".format(stats[u"n_sin_barras"]),
        )
    if stats.get(u"n_sin_nivel"):
        lines.append(
            u"Áreas sin nivel de losa: {0}".format(stats[u"n_sin_nivel"]),
        )
    if stats.get(u"n_sin_ubicacion"):
        lines.append(
            u"Barras sin ubicación resoluble: {0}".format(stats[u"n_sin_ubicacion"]),
        )
    for det in (stats.get(u"areas_detalle") or [])[:12]:
        lines.append(
            u"  · Area {0}: {1} barra(s), Nivel={2}, Malla={3}, Ubic={4}, NivelOK={5}".format(
                det.get(u"area_id"),
                det.get(u"n_barras"),
                det.get(u"nivel") or u"?",
                det.get(u"malla"),
                det.get(u"ubicacion"),
                det.get(u"nivel_ok"),
            )
        )
    n_det = len(stats.get(u"areas_detalle") or [])
    if n_det > 12:
        lines.append(u"  … y {0} área(s) más.".format(n_det - 12))
    return u"\n".join(lines)


def main():
    areas = collect_all_area_reinforcements(doc)
    if not areas:
        TaskDialog.Show(
            u"Arainco: Params malla Area Reinforcement",
            u"No hay Area Reinforcement en el proyecto.",
        )
        return

    n_sel = seleccionar_elementos(areas)
    stats = inyectar_params_malla_en_proyecto(doc)
    # Mantener selección de todos los Area Reinforcement tras la transacción.
    seleccionar_elementos(areas)

    resumen = _format_resumen(stats)
    print(resumen)
    TaskDialog.Show(
        u"Arainco: Params malla Area Reinforcement",
        u"Seleccionados: {0} Area Reinforcement.\n\n{1}".format(n_sel, resumen),
    )


if __name__ == "__main__":
    main()
