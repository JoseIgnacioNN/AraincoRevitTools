# -*- coding: utf-8 -*-
# Ejecutar en RPS: File > Run script (no pegar línea a línea).
"""
Detect shaft openings that cut the selected floor (or footprint roof).

Evaluation: "Does a shaft cut this floor?" — we list shafts that intersect/cut
the selected losa, not the other way around.

Uses Revit API database-level filters:
  - ElementIntersectsElementFilter(floor): keeps shafts whose bbox intersects the floor.
  - Optionally ElementIntersectsSolidFilter(floor solid): refinement by solid geometry.

Compatible with pyRevit and Revit Python Shell (RPS).
Creates model lines for each detected shaft boundary (Transaction).

Usage: Select exactly one Floor or FootPrintRoof, then run the script.
"""

import math
import clr
clr.AddReference("RevitAPI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Arc,
    Curve,
    ElementId,
    BuiltInCategory,
    BuiltInParameter,
    BoundingBoxIntersectsFilter,
    ElementIntersectsElementFilter,
    ElementIntersectsSolidFilter,
    ElementTypeGroup,
    FilteredElementCollector,
    Floor,
    FootPrintRoof,
    GeometryInstance,
    Line,
    Options,
    Outline,
    PlanarFace,
    Plane,
    Sketch,
    SketchPlane,
    Solid,
    StorageType,
    Transaction,
    UnitUtils,
    UnitTypeId,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    AreaReinforcementType,
    RebarBarType,
    RebarHookType,
)

# ── Document / UIDocument (pyRevit usa __revit__; RPS puede predefinir doc/uidoc) ─
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    # RPS: no sobrescribir doc/uidoc si ya están definidos por el shell
    try:
        doc
    except NameError:
        doc = None
    try:
        uidoc
    except NameError:
        uidoc = None


def _get_shaft_opening_category():
    """
    Returns the BuiltInCategory for shaft openings.
    Some Revit versions use OST_ShaftOpening, others OST_Opening.
    """
    for attr in ("OST_ShaftOpening", "OST_Opening"):
        try:
            cat = getattr(BuiltInCategory, attr, None)
            if cat is not None:
                return cat
        except Exception:
            continue
    return None


def _get_floor_solids(element, options=None):
    """
    Extracts all Solid geometry from a floor/roof element.
    Handles GeometryElement with direct Solid or GeometryInstance.
    Returns list of Solid (may be empty).
    """
    if element is None:
        return []
    if options is None:
        options = Options()
        options.ComputeReferences = False
    try:
        geom_elem = element.get_Geometry(options)
    except Exception:
        return []
    if geom_elem is None:
        return []
    solids = []
    for obj in geom_elem:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Volume > 0:
            solids.append(obj)
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                if inst_geom is not None:
                    for g in inst_geom:
                        if isinstance(g, Solid) and g.Volume > 0:
                            solids.append(g)
            except Exception:
                pass
    return solids


def _shafts_by_bbox_intersects_filter(document, floor_element, shaft_cat):
    """
    Shafts whose bounding box intersects the floor's bounding box.
    Uses BoundingBoxIntersectsFilter(Outline(floor_bbox)) — same approach as
    pick_face_losa_superior; often more reliable than ElementIntersectsElementFilter.
    """
    try:
        bbox = floor_element.get_BoundingBox(None)
    except Exception:
        return []
    if bbox is None:
        return []
    min_pt = getattr(bbox, "Min", None) or getattr(bbox, "Minimum", None)
    max_pt = getattr(bbox, "Max", None) or getattr(bbox, "Maximum", None)
    if min_pt is None or max_pt is None:
        return []
    try:
        outline = Outline(min_pt, max_pt)
        bbox_filter = BoundingBoxIntersectsFilter(outline)
        all_candidates = list(
            FilteredElementCollector(document)
            .WherePasses(bbox_filter)
            .ToElements()
        )
        floor_id = floor_element.Id.IntegerValue
        return [
            e for e in all_candidates
            if e is not None and e.IsValidObject
            and e.Id.IntegerValue != floor_id
            and e.Category is not None
            and e.Category.Id.IntegerValue == int(shaft_cat)
        ]
    except Exception:
        return []


def _shafts_by_solid_filter(document, floor_element, shaft_cat):
    """
    Shafts that intersect the floor's Solid geometry (no bbox pre-filter).
    Uses ElementIntersectsSolidFilter for each floor solid and merges results.
    """
    floor_solids = _get_floor_solids(floor_element)
    if not floor_solids:
        return []
    seen_ids = set()
    result = []
    for solid in floor_solids:
        try:
            solid_filter = ElementIntersectsSolidFilter(solid)
            for elem in (
                FilteredElementCollector(document)
                .OfCategory(shaft_cat)
                .WherePasses(solid_filter)
                .ToElements()
            ):
                if elem is not None and elem.IsValidObject and elem.Id.IntegerValue not in seen_ids:
                    seen_ids.add(elem.Id.IntegerValue)
                    result.append(elem)
        except Exception:
            pass
    return result


def _shafts_by_inverted_filter(document, floor_element, shaft_cat):
    """
    For each shaft: get elements that intersect this shaft; if the floor is among
    them, the shaft cuts the floor. Slower but different API semantics.
    """
    all_shafts = list(
        FilteredElementCollector(document)
        .OfCategory(shaft_cat)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    floor_id = floor_element.Id
    result = []
    for shaft in all_shafts:
        if shaft is None or not shaft.IsValidObject or shaft.Id == floor_id:
            continue
        try:
            inter_filter = ElementIntersectsElementFilter(shaft)
            intersecting = list(
                FilteredElementCollector(document)
                .WherePasses(inter_filter)
                .ToElements()
            )
            for e in intersecting:
                if e is not None and e.Id == floor_id:
                    result.append(shaft)
                    break
        except Exception:
            pass
    return result


def _proyectar_curva_a_plano_z(curve, z_plano):
    """Proyecta una curva al plano Z=z_plano. Soporta Line y Arc. Retorna None si falla."""
    if curve is None or not curve.IsBound:
        return None
    try:
        if isinstance(curve, Line):
            p1, p2 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            return Line.CreateBound(XYZ(p1.X, p1.Y, z_plano), XYZ(p2.X, p2.Y, z_plano))
        if isinstance(curve, Arc):
            c = curve.Center
            p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            centro = XYZ(c.X, c.Y, z_plano)
            pt0 = XYZ(p0.X, p0.Y, z_plano)
            pt1 = XYZ(p1.X, p1.Y, z_plano)
            return Arc.Create(pt0, pt1, centro)
    except Exception:
        pass
    return None


def _curvas_desde_shaft(opening, z_plano=None):
    """
    Obtiene las curvas del contorno de un shaft opening (BoundaryCurves o BoundaryRect).
    Si z_plano se indica, proyecta al plano Z para que las curvas sean coplanares (model lines).
    """
    curvas = []
    try:
        is_rect = getattr(opening, "IsRectBoundary", False)
        if is_rect:
            rect = getattr(opening, "BoundaryRect", None)
            if rect is not None:
                mn = getattr(rect, "Min", None) or getattr(rect, "Minimum", None)
                mx = getattr(rect, "Max", None) or getattr(rect, "Maximum", None)
                if mn is not None and mx is not None:
                    z = z_plano if z_plano is not None else mn.Z
                    pts = [
                        XYZ(mn.X, mn.Y, z),
                        XYZ(mx.X, mn.Y, z),
                        XYZ(mx.X, mx.Y, z),
                        XYZ(mn.X, mx.Y, z),
                    ]
                    for i in range(4):
                        c = Line.CreateBound(pts[i], pts[(i + 1) % 4])
                        if c is not None:
                            curvas.append(c)
            return curvas
    except Exception:
        pass
    try:
        boundary = getattr(opening, "BoundaryCurves", None)
        if boundary is not None:
            for c in boundary:
                if c is not None and c.IsBound:
                    if z_plano is not None:
                        c = _proyectar_curva_a_plano_z(c, z_plano)
                    if c is not None:
                        curvas.append(c)
    except Exception:
        pass
    return curvas


def _obtener_plano_cara_superior_losa(floor_element):
    """
    Obtiene el plano de la cara superior de la losa (Floor/FootPrintRoof).
    Busca la PlanarFace con normal apuntando hacia arriba (+Z).
    Returns:
        tuple: (Plane, z_plano) para usar en SketchPlane y proyección, o (None, None) si falla.
    """
    if floor_element is None:
        return None, None
    opts = Options()
    opts.ComputeReferences = False
    try:
        geom_elem = floor_element.get_Geometry(opts)
    except Exception:
        return None, None
    if geom_elem is None:
        return None, None
    cara_superior = None
    for obj in geom_elem:
        if obj is None:
            continue
        solid = obj if isinstance(obj, Solid) else None
        if solid is None or solid.Faces.Size == 0:
            if isinstance(obj, GeometryInstance):
                try:
                    inst = obj.GetInstanceGeometry()
                    if inst:
                        for g in inst:
                            if isinstance(g, Solid) and g.Faces.Size > 0:
                                solid = g
                                break
                except Exception:
                    pass
        if solid is None:
            continue
        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue
            normal = face.FaceNormal
            if normal and normal.Z >= 0.9:
                cara_superior = face
                break
        if cara_superior is not None:
            break
    if cara_superior is None:
        return None, None
    try:
        loops = cara_superior.GetEdgesAsCurveLoops()
        if loops is None:
            return None, None
        try:
            first_loop = loops.get_Item(0) if getattr(loops, "Size", 0) > 0 else None
        except Exception:
            first_loop = next(iter(loops), None) if loops else None
        if first_loop is None:
            return None, None
        n = getattr(first_loop, "NumberOfCurves", lambda: 0)()
        if n is None:
            n = 0
        if n > 0:
            try:
                first_curve = first_loop.get_Item(0)
            except Exception:
                first_curve = next(iter(first_loop), None)
            if first_curve is not None and first_curve.IsBound:
                pt = first_curve.GetEndPoint(0)
                origin = XYZ(pt.X, pt.Y, pt.Z)
                z_plano = origin.Z
                normal = cara_superior.FaceNormal
                plano = Plane.CreateByNormalAndOrigin(normal, origin)
                return plano, z_plano
    except Exception:
        pass
    return None, None


def _vertices_perimetro_exterior_losa(floor_element, document):
    """
    Obtiene los vértices 2D (x, y) del perímetro exterior de la losa (primer bucle del sketch).
    Sirve para point-in-polygon: solo se generan líneas de shafts que estén dentro del perímetro.
    Returns:
        list: [(x, y), ...] en orden cerrado, o [] si falla.
    """
    if floor_element is None or document is None or not isinstance(floor_element, Floor):
        return []
    try:
        sketch_id = floor_element.SketchId
        if sketch_id is None or sketch_id == ElementId.InvalidElementId:
            return []
        sketch = document.GetElement(sketch_id)
        if sketch is None or not isinstance(sketch, Sketch):
            return []
        profile = sketch.Profile
        if profile is None:
            return []
        n_loops = getattr(profile, "Size", 0) or getattr(profile, "Count", 0)
        if n_loops < 1:
            return []
        try:
            curve_array = profile.get_Item(0)
        except Exception:
            return []
        if curve_array is None:
            return []
        vertices = []
        n_curves = getattr(curve_array, "Size", 0) or getattr(curve_array, "Count", 0)
        for j in range(n_curves):
            try:
                c = curve_array.get_Item(j)
            except Exception:
                continue
            if c is not None and c.IsBound:
                pt = c.GetEndPoint(0)
                vertices.append((pt.X, pt.Y))
        return vertices if len(vertices) >= 3 else []
    except Exception:
        return []


def _punto_dentro_poligono_2d(px, py, vertices):
    """
    Ray casting: punto (px, py) dentro del polígono cerrado definido por vertices [(x,y), ...].
    Retorna True si el punto está dentro (o en el borde).
    """
    if not vertices or len(vertices) < 3:
        return False
    n = len(vertices)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = vertices[i][0], vertices[i][1]
        xj, yj = vertices[j][0], vertices[j][1]
        if (yj != yi) and ((yi > py) != (yj > py)):
            x_cruce = (xj - xi) * (py - yi) / (yj - yi) + xi
            if px < x_cruce:
                inside = not inside
        j = i
    return inside


def _punto_medio_curva(curve):
    """Punto en el medio del parámetro de la curva (0.5). Retorna XYZ o None."""
    if curve is None or not curve.IsBound:
        return None
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return XYZ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5, (p0.Z + p1.Z) * 0.5)
        except Exception:
            return None


def _curvas_perimetro_exterior_solo(floor_element, document, z_plano):
    """
    Solo el primer bucle del sketch (perímetro exterior). Para Area Reinforcement
    se listan primero las curvas perimetrales.
    """
    if floor_element is None or document is None or not isinstance(floor_element, Floor):
        return []
    try:
        sketch_id = floor_element.SketchId
        if sketch_id is None or sketch_id == ElementId.InvalidElementId:
            return []
        sketch = document.GetElement(sketch_id)
        if sketch is None or not isinstance(sketch, Sketch):
            return []
        profile = sketch.Profile
        if profile is None:
            return []
        n_loops = getattr(profile, "Size", 0) or getattr(profile, "Count", 0)
        if n_loops < 1:
            return []
        curve_array = profile.get_Item(0)
        if curve_array is None:
            return []
        curvas = []
        n_curves = getattr(curve_array, "Size", 0) or getattr(curve_array, "Count", 0)
        for j in range(n_curves):
            try:
                c = curve_array.get_Item(j)
            except Exception:
                continue
            if c is not None and c.IsBound:
                c_proy = _proyectar_curva_a_plano_z(c, z_plano)
                if c_proy is not None:
                    curvas.append(c_proy)
        return curvas
    except Exception:
        return []


def _curvas_desde_sketch_losa(floor_element, document, z_plano):
    """
    Obtiene las curvas del contorno de la losa desde su Sketch (perímetro y huecos).
    Solo aplica a Floor con SketchId válido. Las curvas se proyectan al plano z_plano.
    Returns:
        list: Curvas (Line/Arc) proyectadas a z_plano, o [] si no hay sketch o falla.
    """
    if floor_element is None or document is None or not isinstance(floor_element, Floor):
        return []
    try:
        sketch_id = floor_element.SketchId
        if sketch_id is None or sketch_id == ElementId.InvalidElementId:
            return []
        sketch = document.GetElement(sketch_id)
        if sketch is None or not isinstance(sketch, Sketch):
            return []
        profile = sketch.Profile
        if profile is None:
            return []
        n_loops = getattr(profile, "Size", 0) or getattr(profile, "Count", 0)
        if n_loops < 1:
            return []
        curvas = []
        for i in range(n_loops):
            try:
                curve_array = profile.get_Item(i) if hasattr(profile, "get_Item") else list(profile)[i]
            except Exception:
                continue
            if curve_array is None:
                continue
            n_curves = getattr(curve_array, "Size", 0) or getattr(curve_array, "Count", 0)
            for j in range(n_curves):
                try:
                    c = curve_array.get_Item(j) if hasattr(curve_array, "get_Item") else list(curve_array)[j]
                except Exception:
                    continue
                if c is not None and c.IsBound:
                    c_proy = _proyectar_curva_a_plano_z(c, z_plano)
                    if c_proy is not None:
                        curvas.append(c_proy)
        return curvas
    except Exception:
        return []


def _obtener_curvas_ordenadas_para_area_reinforcement(document, floor_element, shaft_elements, z_plano, vertices_perimetro):
    """
    Curvas en el orden requerido por el método Create(7 params): primero las perimetrales
    (primer bucle del sketch), luego las de cada shaft filtrado (dentro del perímetro).
    Returns:
        list[Curve]: curvas ordenadas (perímetro + shafts).
    """
    ordenadas = []
    curvas_perimetro = _curvas_perimetro_exterior_solo(floor_element, document, z_plano)
    for c in curvas_perimetro:
        if c is not None and c.IsBound:
            ordenadas.append(c)
    for opening in shaft_elements:
        curvas = _curvas_desde_shaft(opening, z_plano)
        if not curvas:
            curvas = _curvas_desde_shaft(opening, None)
        for c in curvas:
            if c is None or not c.IsBound:
                continue
            c = _proyectar_curva_a_plano_z(c, z_plano)
            if c is None:
                continue
            if vertices_perimetro:
                pt_medio = _punto_medio_curva(c)
                if pt_medio is None or not _punto_dentro_poligono_2d(pt_medio.X, pt_medio.Y, vertices_perimetro):
                    continue
            ordenadas.append(c)
    return ordenadas


def _direccion_desde_curvas(curves):
    """Calcula XYZ de dirección principal desde la primera curva (igual que crear_area_reinforcement_rps.obtener_direccion)."""
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
    except Exception:
        pass
    return XYZ(1, 0, 0)


def _get_area_reinforcement_type_id(document):
    """Obtiene el tipo por defecto o el primer AreaReinforcementType (igual que crear_area_reinforcement_rps)."""
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


def _get_first_rebar_bar_type_id(document):
    """Obtiene el ID del primer RebarBarType (igual que crear_area_reinforcement_rps)."""
    try:
        for elem in FilteredElementCollector(document).OfClass(RebarBarType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return None


def _get_first_rebar_hook_type_id(document):
    """Obtiene el ID del primer RebarHookType (igual que crear_area_reinforcement_rps)."""
    try:
        for elem in FilteredElementCollector(document).OfClass(RebarHookType):
            if elem:
                return elem.Id
    except Exception:
        pass
    return ElementId.InvalidElementId


def _crear_gancho_por_defecto(document):
    """Crea RebarHookType 90° por defecto si no hay ninguno (igual que crear_area_reinforcement_rps). Debe ejecutarse dentro de Transaction."""
    try:
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
    except Exception:
        pass
    return ElementId.InvalidElementId


def _asignar_hook_a_area_reinforcement(area_rein, hook_type_id):
    """Asigna el RebarHookType a las capas (igual que crear_area_reinforcement_rps: BIP + LookupParameter + fallback por iteración)."""
    if not area_rein or not hook_type_id or hook_type_id == ElementId.InvalidElementId:
        return
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
        except Exception:
            continue
    try:
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
            except Exception:
                continue
    except Exception:
        pass
    try:
        for p in area_rein.Parameters:
            if p is None or p.IsReadOnly or p.StorageType != StorageType.ElementId:
                continue
            try:
                nombre = p.Definition.Name if p.Definition else ""
                if "ook" in nombre.lower() or "gancho" in nombre.lower():
                    p.Set(hook_type_id)
            except Exception:
                continue
    except Exception:
        pass


def _crear_area_reinforcement_7params(document, floor_element, curvas_ordenadas):
    """
    AreaReinforcement.Create(Document, Element, IList<Curve>, XYZ, ElementId, ElementId, ElementId)
    La API exige que IList<Curve> forme UN SOLO bucle cerrado y contiguo. No acepta varios bucles
    (perímetro + shafts) en una sola lista. Por tanto se usa solo la PRIMERA lista: curvas del
    perímetro de la losa. Las curvas de los shafts serían bucles adicionales y esta sobrecarga
    no los admite; el Area Reinforcement se crea solo con el perímetro.
    """
    if document is None or floor_element is None or not isinstance(floor_element, Floor):
        return None
    # Solo la primera lista de curvas = perímetro (un único bucle contiguo)
    z_plano = curvas_ordenadas[0].GetEndPoint(0).Z if curvas_ordenadas else 0.0
    curvas_perimetro = _curvas_perimetro_exterior_solo(floor_element, document, z_plano)
    if not curvas_perimetro or len(curvas_perimetro) < 3:
        print("Area Reinforcement: FALLO - No se pudieron obtener curvas del perimetro de la losa.")
        return None
    area_type_id = _get_area_reinforcement_type_id(document)
    bar_type_id = _get_first_rebar_bar_type_id(document)
    hook_type_id = _get_first_rebar_hook_type_id(document)
    if not area_type_id or not bar_type_id:
        print("Area Reinforcement: FALLO - No hay AreaReinforcementType o RebarBarType en el proyecto.")
        return None
    if not hook_type_id or hook_type_id == ElementId.InvalidElementId:
        hook_type_id = _crear_gancho_por_defecto(document)
        if hook_type_id == ElementId.InvalidElementId:
            print("Area Reinforcement: AVISO - No hay RebarHookType; se intentó crear uno por defecto y falló.")
    print("Area Reinforcement: Create con PRIMERA lista (perimetro, {} curvas). Shafts no admitidos en esta API.".format(len(curvas_perimetro)))
    print("  areaTypeId={}, barTypeId={}, hookTypeId={}.".format(
        area_type_id.IntegerValue if area_type_id else 0,
        bar_type_id.IntegerValue if bar_type_id else 0,
        hook_type_id.IntegerValue if hook_type_id else 0
    ))
    layout_dir = _direccion_desde_curvas(curvas_perimetro)
    curve_list = List[Curve](curvas_perimetro)
    t = Transaction(document, "Area Reinforcement (7 params)")
    t.Start()
    try:
        ar = AreaReinforcement.Create(
            document, floor_element, curve_list, layout_dir,
            area_type_id, bar_type_id, hook_type_id
        )
        if ar and hook_type_id and hook_type_id != ElementId.InvalidElementId:
            _asignar_hook_a_area_reinforcement(ar, hook_type_id)
        t.Commit()
        return ar
    except Exception as ex:
        if t.HasStarted():
            try:
                t.RollBack()
            except Exception:
                pass
        print("Area Reinforcement: FALLO al llamar Create(...)")
        print("  Excepcion: {}".format(str(ex)))
        return None


def _crear_model_lines_shafts(document, uidoocument, shaft_elements, floor_element=None):
    """
    Crea model lines (ModelCurve): contorno de la losa (desde su sketch) y contorno de cada shaft.
    Si se pasa floor_element (losa seleccionada), las líneas se dibujan sobre la cara superior
    de la losa; si no, se usa el plano de la vista activa.
    Retorna el número de curvas creadas.
    """
    if document is None or uidoocument is None or not shaft_elements:
        return 0
    plano = None
    z_plano = None
    if floor_element is not None:
        plano, z_plano = _obtener_plano_cara_superior_losa(floor_element)
    if plano is None or z_plano is None:
        vista = uidoocument.ActiveView
        if vista is not None and vista.Origin is not None:
            z_plano = getattr(vista.Origin, "Z", 0.0)
        else:
            z_plano = 0.0
        plano = Plane.CreateByNormalAndOrigin(XYZ(0, 0, 1), XYZ(0, 0, z_plano))
    todas_las_curvas = []
    # Contorno de la losa desde su sketch (perímetro + huecos)
    if floor_element is not None:
        curvas_losa = _curvas_desde_sketch_losa(floor_element, document, z_plano)
        for c in curvas_losa:
            if c is not None and c.IsBound:
                todas_las_curvas.append(c)
    # Contorno de cada shaft: solo curvas cuyo centro queda dentro del perímetro de la losa
    vertices_perimetro = []
    if floor_element is not None:
        vertices_perimetro = _vertices_perimetro_exterior_losa(floor_element, document)
    for opening in shaft_elements:
        curvas = _curvas_desde_shaft(opening, z_plano)
        if not curvas:
            curvas = _curvas_desde_shaft(opening, None)
        for c in curvas:
            if c is None or not c.IsBound:
                continue
            c = _proyectar_curva_a_plano_z(c, z_plano)
            if c is None:
                continue
            if vertices_perimetro:
                pt_medio = _punto_medio_curva(c)
                if pt_medio is None or not _punto_dentro_poligono_2d(pt_medio.X, pt_medio.Y, vertices_perimetro):
                    continue
            todas_las_curvas.append(c)
    if not todas_las_curvas:
        return 0
    ids_creados = []
    t = Transaction(document, "Losa + shafts (model lines)")
    t.Start()
    try:
        sketch_plane = SketchPlane.Create(document, plano)
        for curve in todas_las_curvas:
            try:
                mc = document.Create.NewModelCurve(curve, sketch_plane)
                if mc is not None:
                    ids_creados.append(mc.Id)
            except Exception:
                pass
        t.Commit()
    except Exception:
        t.RollBack()
        return 0
    if ids_creados:
        try:
            uidoocument.Selection.SetElementIds(List[ElementId](ids_creados))
        except Exception:
            pass
    return len(ids_creados)


def run(document=None, uidoocument=None):
    """
    Main entry: get selection, validate one Floor/FootPrintRoof,
    find shaft openings that cut that floor via DB filters, and print results.

    If document/uidoocument are None, uses __revit__.ActiveUIDocument.
    """
    if document is None:
        document = doc
    if uidoocument is None:
        uidoocument = uidoc

    if document is None or uidoocument is None:
        print("Error: No active document. Run from Revit (pyRevit/RPS).")
        return

    # ── Selection: exactly one element, type Floor or FootPrintRoof ───────────
    try:
        selection_ids = list(uidoocument.Selection.GetElementIds())
    except Exception:
        print("Error: Could not get current selection.")
        return

    if not selection_ids:
        print("No selection. Please select exactly one Floor or FootPrintRoof (losa or cubierta) and run again.")
        return

    if len(selection_ids) > 1:
        print("Multiple elements selected. Please select exactly one Floor or FootPrintRoof.")
        return

    selected_element = document.GetElement(selection_ids[0])
    if selected_element is None or not selected_element.IsValidObject:
        print("Error: Selected element is invalid or was deleted.")
        return

    if not isinstance(selected_element, (Floor, FootPrintRoof)):
        print(
            "Invalid type. Selected element must be a Floor or FootPrintRoof (losa/cubierta). "
            "Got: {}.".format(type(selected_element).__name__)
        )
        return

    # ── Target category: Shaft Openings ─────────────────────────────────────
    shaft_cat = _get_shaft_opening_category()
    if shaft_cat is None:
        print("Error: Shaft opening category (OST_ShaftOpening / OST_Opening) not found in this Revit version.")
        return

    # Diagnostic: total shaft openings in document
    total_shafts = len(list(FilteredElementCollector(document).OfCategory(shaft_cat).ToElementIds()))
    print("Diagnostic: {} shaft opening(s) in document. Selected element: Id {} ({}).".format(
        total_shafts, selected_element.Id.IntegerValue, type(selected_element).__name__
    ))

    # ── Multiple strategies (ElementIntersectsElementFilter can miss in some setups) ─
    seen_ids = set()
    intersecting_final = []

    # Strategy 1: ElementIntersectsElementFilter(floor) — standard DB filter
    try:
        for e in (
            FilteredElementCollector(document)
            .OfCategory(shaft_cat)
            .WherePasses(ElementIntersectsElementFilter(selected_element))
            .ToElements()
        ):
            if e is not None and e.IsValidObject and e.Id.IntegerValue not in seen_ids:
                seen_ids.add(e.Id.IntegerValue)
                intersecting_final.append(e)
    except Exception:
        pass

    # Strategy 2: BoundingBoxIntersectsFilter(Outline(floor_bbox)) — same as pick_face_losa_superior
    for e in _shafts_by_bbox_intersects_filter(document, selected_element, shaft_cat):
        if e.Id.IntegerValue not in seen_ids:
            seen_ids.add(e.Id.IntegerValue)
            intersecting_final.append(e)

    # Strategy 3: ElementIntersectsSolidFilter(floor solid) — solid-level, no bbox
    for e in _shafts_by_solid_filter(document, selected_element, shaft_cat):
        if e.Id.IntegerValue not in seen_ids:
            seen_ids.add(e.Id.IntegerValue)
            intersecting_final.append(e)

    # Strategy 4: Inverted — for each shaft, elements that intersect shaft; if floor in set, shaft cuts floor
    for e in _shafts_by_inverted_filter(document, selected_element, shaft_cat):
        if e.Id.IntegerValue not in seen_ids:
            seen_ids.add(e.Id.IntegerValue)
            intersecting_final.append(e)

    # ── Output ──────────────────────────────────────────────────────────────
    if not intersecting_final:
        print("No shaft openings cut the selected floor/roof (tried 4 strategies).")
        return

    print("Shaft openings that cut the selected floor/roof ({}):".format(len(intersecting_final)))
    for elem in intersecting_final:
        name = elem.Name if getattr(elem, "Name", None) else "(no name)"
        print("  Id: {}, Name: {}".format(elem.Id.IntegerValue, name))

    # Curvas ordenadas: perimetrales primero, luego shafts filtrados (para model lines y Area Reinforcement)
    z_plano = None
    if selected_element is not None:
        plano, z_plano = _obtener_plano_cara_superior_losa(selected_element)
    if z_plano is None:
        try:
            z_plano = getattr(uidoocument.ActiveView.Origin, "Z", 0.0) if uidoocument and uidoocument.ActiveView else 0.0
        except Exception:
            z_plano = 0.0
    vertices_perimetro = _vertices_perimetro_exterior_losa(selected_element, document) if selected_element else []
    curvas_ordenadas = []
    if isinstance(selected_element, Floor):
        curvas_ordenadas = _obtener_curvas_ordenadas_para_area_reinforcement(
            document, selected_element, intersecting_final, z_plano, vertices_perimetro
        )

    # Dibujar model lines (contorno losa + shafts sobre la cara superior)
    num_lineas = _crear_model_lines_shafts(document, uidoocument, intersecting_final, selected_element)
    if num_lineas > 0:
        print("Model lines creadas: {} (contorno de la losa + contorno de los shafts sobre la cara superior).".format(num_lineas))
    else:
        print("No se pudieron crear model lines (revisar BoundaryCurves/BoundaryRect de los shafts).")

    # Area Reinforcement con método de 7 parámetros (curvas ordenadas: perimetrales + shafts)
    if isinstance(selected_element, Floor) and curvas_ordenadas:
        ar = _crear_area_reinforcement_7params(document, selected_element, curvas_ordenadas)
        if ar is not None:
            print("Area Reinforcement creado (ID: {}).".format(ar.Id.IntegerValue))
        else:
            print("No se pudo crear el Area Reinforcement. Ver mensajes anteriores para el motivo (excepcion, tipos o curvas no contiguas).")


# Run when executed as script (e.g. from RPS or pyRevit)
if __name__ == "__main__" or (doc is not None and uidoc is not None):
    run()
