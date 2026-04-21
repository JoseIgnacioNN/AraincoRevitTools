# -*- coding: utf-8 -*-
"""
Script: Obtener perímetro y huecos de losa desde geometría sólida.
Revit 2024+ | pyRevit | IronPython 3.4 / CPython 3

Requisitos:
- Seleccionar un elemento Floor antes de ejecutar.
- Obtiene la geometría sólida, identifica la cara superior (+Z) y extrae
  los CurveLoops mediante GetEdgesAsCurveLoops (estándar API para perímetros/huecos).
- Primer bucle = perímetro exterior; bucles adicionales = huecos interiores.
- Crea DetailCurves en la vista actual para análisis visual.
"""

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    Arc,
    BuiltInCategory,
    CurveLoop,
    Line,
    Options,
    PlanarFace,
    Solid,
    Transaction,
    UnitUtils,
    UnitTypeId,
    XYZ,
)

# ── Boilerplate pyRevit ──────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


def _proyectar_curva_a_plano_z(curve, z_plano):
    """
    Proyecta una curva al plano horizontal Z=z_plano.
    Soporta Line y Arc. Retorna None si no se puede proyectar.
    """
    if curve is None or not curve.IsBound:
        return None
    try:
        if isinstance(curve, Line):
            p1 = curve.GetEndPoint(0)
            p2 = curve.GetEndPoint(1)
            pt1 = XYZ(p1.X, p1.Y, z_plano)
            pt2 = XYZ(p2.X, p2.Y, z_plano)
            return Line.CreateBound(pt1, pt2)
        if isinstance(curve, Arc):
            c = curve.Center
            centro = XYZ(c.X, c.Y, z_plano)
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            pt0 = XYZ(p0.X, p0.Y, z_plano)
            pt1 = XYZ(p1.X, p1.Y, z_plano)
            return Arc.Create(pt0, pt1, centro)
    except Exception:
        pass
    return None


def _longitud_curveloop(cl):
    """Calcula la longitud total de un CurveLoop (unidades internas Revit)."""
    if cl is None:
        return 0.0
    total = 0.0
    for c in cl:
        if c is not None and c.IsBound:
            total += c.Length
    return total


def _obtener_cara_superior(geom_elem):
    """
    Busca la cara planar con normal orientada hacia arriba (+Z).
    Retorna la primera PlanarFace encontrada con FaceNormal.Z > umbral.
    """
    if geom_elem is None:
        return None
    umbral_z = 0.9  # Normal apuntando hacia arriba
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


def main():
    if doc is None or uidoc is None:
        print("Error: Ejecuta este script dentro de pyRevit (doc/uidoc no disponibles).")
        return

    elem_ids = list(uidoc.Selection.GetElementIds())
    if not elem_ids:
        print("No hay elementos seleccionados. Selecciona una losa (Floor) y vuelve a ejecutar.")
        return

    elem_id = elem_ids[0]
    elem = doc.GetElement(elem_id)
    if elem is None:
        print("No se pudo obtener el elemento seleccionado.")
        return

    cat = elem.Category
    if cat is None or cat.Id.IntegerValue != int(BuiltInCategory.OST_Floors):
        print("El elemento seleccionado no es de categoría Floor (OST_Floors).")
        return

    opts = Options()
    opts.ComputeReferences = True
    geom_elem = elem.get_Geometry(opts)
    if geom_elem is None:
        print("No se pudo obtener la geometría del suelo.")
        return

    cara_superior = _obtener_cara_superior(geom_elem)
    if cara_superior is None:
        print("No se encontró la cara superior del suelo (normal +Z).")
        return

    # GetEdgesAsCurveLoops: estándar API para perímetros y huecos
    curve_loops_raw = cara_superior.GetEdgesAsCurveLoops()
    curve_loops = list(curve_loops_raw) if curve_loops_raw else []
    if not curve_loops:
        print("No se obtuvieron bucles de curvas en la cara superior.")
        return

    # Ordenar por longitud descendente: el más largo = perímetro exterior
    loops_con_longitud = []
    for cl in curve_loops:
        lng = _longitud_curveloop(cl)
        loops_con_longitud.append((cl, lng))
    loops_con_longitud.sort(key=lambda x: x[1], reverse=True)

    perimetro_interno = loops_con_longitud[0][1]
    huecos = loops_con_longitud[1:]

    # Conversión a unidades legibles (metros y pies)
    perimetro_m = UnitUtils.ConvertFromInternalUnits(perimetro_interno, UnitTypeId.Meters)
    perimetro_ft = UnitUtils.ConvertFromInternalUnits(perimetro_interno, UnitTypeId.Feet)

    print("-" * 60)
    print("Perímetro y huecos de losa (geometría sólida)")
    print("-" * 60)
    print("Elemento: {} (ID: {})".format(elem.Name or "(sin nombre)", elem.Id.IntegerValue))
    print("")
    print("Perímetro exterior:")
    print("  {:.4f} m  ({:.4f} ft)".format(perimetro_m, perimetro_ft))
    print("")
    print("Huecos interiores: {} encontrado(s)".format(len(huecos)))
    for i, (_, lng) in enumerate(huecos, 1):
        lng_m = UnitUtils.ConvertFromInternalUnits(lng, UnitTypeId.Meters)
        lng_ft = UnitUtils.ConvertFromInternalUnits(lng, UnitTypeId.Feet)
        print("  Hueco {}: {:.4f} m ({:.4f} ft)".format(i, lng_m, lng_ft))
    print("-" * 60)

    # Crear DetailCurves en la vista actual para análisis visual
    vista = uidoc.ActiveView
    if vista is None:
        print("No hay vista activa. Las DetailCurves no se crearon.")
        return

    z_vista = vista.Origin.Z if vista.Origin else 0.0
    creadas = 0
    t = Transaction(doc, "DetailCurves perímetro losa")
    t.Start()
    try:
        for cl, _ in loops_con_longitud:
            for c in cl:
                if c is None or not c.IsBound:
                    continue
                try:
                    doc.Create.NewDetailCurve(vista, c)
                    creadas += 1
                except Exception:
                    curva_proy = _proyectar_curva_a_plano_z(c, z_vista)
                    if curva_proy is not None:
                        try:
                            doc.Create.NewDetailCurve(vista, curva_proy)
                            creadas += 1
                        except Exception as ex:
                            print("  Aviso: no se pudo crear DetailCurve: {}".format(str(ex)))
                    else:
                        print("  Aviso: curva no proyectable (tipo no soportado)")
        t.Commit()
        print("DetailCurves creadas en la vista actual: {}".format(creadas))
    except Exception as ex:
        t.RollBack()
        print("Error al crear DetailCurves: {}".format(str(ex)))


if __name__ == "__main__" or (doc is not None and uidoc is not None):
    main()
