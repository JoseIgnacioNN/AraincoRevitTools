# -*- coding: utf-8 -*-
"""
Script RPS: Extraer CurveLoop(s) de una losa (Floor).
Ejecutable en Revit Python Shell (RPS) — Revit 2024+ | IronPython 3.4

Obtiene los CurveLoops de la cara superior de la losa mediante geometría sólida
(GetEdgesAsCurveLoops). Primer bucle = perímetro exterior; bucles adicionales = huecos.

Uso en RPS:
  1. Selecciona una losa (Floor) en Revit.
  2. Ejecuta el script.
  3. curve_loops = extraer_curveloops_losa(doc, elem)  # o usar la selección automática
"""

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    CurveLoop,
    Options,
    PlanarFace,
    Solid,
    UnitUtils,
    UnitTypeId,
)

# ── Boilerplate RPS ─────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


def _obtener_cara_superior(geom_elem):
    """
    Busca la cara planar con normal orientada hacia arriba (+Z).
    Retorna la primera PlanarFace con FaceNormal.Z >= umbral.
    """
    if geom_elem is None:
        return None
    umbral_z = 0.9
    for geom_obj in geom_elem:
        solid = geom_obj if isinstance(geom_obj, Solid) else None
        if solid is None or solid.Faces.Size == 0:
            continue
        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue
            normal = face.FaceNormal
            if normal is not None and normal.Z >= umbral_z:
                return face
    return None


def _longitud_curveloop(cl):
    """Longitud total de un CurveLoop (unidades internas Revit)."""
    if cl is None:
        return 0.0
    total = 0.0
    for c in cl:
        if c is not None and c.IsBound:
            total += c.Length
    return total


def extraer_curveloops_losa(document, floor_element):
    """
    Extrae los CurveLoops de la cara superior de una losa (Floor).

    Args:
        document: Document (Revit DB Document).
        floor_element: Elemento Floor.

    Returns:
        tuple: (list[CurveLoop], z_elevation) — lista de CurveLoops (perímetro + huecos)
               y elevación Z de la cara. (None, None) si falla.
    """
    if document is None or floor_element is None:
        return None, None
    cat = floor_element.Category
    if cat is None or cat.Id.IntegerValue != int(BuiltInCategory.OST_Floors):
        return None, None

    opts = Options()
    opts.ComputeReferences = False
    geom_elem = floor_element.get_Geometry(opts)
    if geom_elem is None:
        return None, None

    cara_superior = _obtener_cara_superior(geom_elem)
    if cara_superior is None:
        return None, None

    z_elevation = cara_superior.Origin.Z if cara_superior.Origin else None
    curve_loops_raw = cara_superior.GetEdgesAsCurveLoops()
    curve_loops = list(curve_loops_raw) if curve_loops_raw else []
    if not curve_loops:
        return None, z_elevation

    # Ordenar por longitud descendente: el más largo = perímetro exterior
    loops_con_longitud = []
    for cl in curve_loops:
        lng = _longitud_curveloop(cl)
        loops_con_longitud.append((cl, lng))
    loops_con_longitud.sort(key=lambda x: x[1], reverse=True)
    ordered_loops = [cl for cl, _ in loops_con_longitud]

    return ordered_loops, z_elevation


def run(document, uidocument):
    """
    Ejecuta usando la selección actual: toma el primer elemento seleccionado
    si es un Floor y extrae sus CurveLoops. Retorna (curve_loops, z) para uso en RPS.
    """
    if document is None or uidocument is None:
        print("Error: doc/uidoc no disponibles. Ejecuta dentro de RPS.")
        return None, None

    elem_ids = list(uidocument.Selection.GetElementIds())
    if not elem_ids:
        print("No hay selección. Selecciona una losa (Floor) y vuelve a ejecutar.")
        return None, None

    elem = document.GetElement(elem_ids[0])
    if elem is None:
        print("No se pudo obtener el elemento.")
        return None, None

    curve_loops, z_elev = extraer_curveloops_losa(document, elem)
    if curve_loops is None:
        if elem.Category and elem.Category.Id.IntegerValue != int(BuiltInCategory.OST_Floors):
            print("El elemento no es un Floor (OST_Floors).")
        else:
            print("No se pudieron extraer CurveLoops de la losa.")
        return None, None

    perimetro = _longitud_curveloop(curve_loops[0])
    n_huecos = len(curve_loops) - 1
    perimetro_m = UnitUtils.ConvertFromInternalUnits(perimetro, UnitTypeId.Meters)

    print("-" * 50)
    print("CurveLoops extraídos de losa: {} (ID: {})".format(
        elem.Name or "(sin nombre)", elem.Id.IntegerValue))
    print("  Perímetro: 1 loop, {:.4f} m".format(perimetro_m))
    print("  Huecos: {} loop(s)".format(n_huecos))
    if z_elev is not None:
        z_m = UnitUtils.ConvertFromInternalUnits(z_elev, UnitTypeId.Meters)
        print("  Z cara superior: {:.4f} m".format(z_m))
    print("-" * 50)
    print("Uso: curve_loops tiene {} CurveLoop(s). Perímetro = curve_loops[0]".format(len(curve_loops)))

    return curve_loops, z_elev


# ── Ejecución al cargar en RPS ──────────────────────────────────────────────
if doc is not None and uidoc is not None:
    curve_loops, z_elevation = run(doc, uidoc)
    # Dejar en global para uso en consola: curve_loops, z_elevation
else:
    curve_loops = None
    z_elevation = None
    print("Ejecuta este script dentro de Revit Python Shell (RPS).")
