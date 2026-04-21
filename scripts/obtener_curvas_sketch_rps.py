# -*- coding: utf-8 -*-
# Ejecutar en RPS: File > Run script (no pegar línea a línea).
"""
Script RPS: Obtener curvas del perímetro exterior del sketch de una losa (Floor).
Compatible con Create(7 params). Revit 2024+ | IronPython 3.4
"""

import clr
clr.AddReference("RevitAPI")

from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    Curve,
    CurveLoop,
    ElementId,
    Plane,
    Sketch,
    SketchPlane,
    Transaction,
    XYZ,
)

# Variables predefinidas RPS
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


def _obtener_curvas_sketch(floor, document):
    """Obtiene las curvas del perímetro exterior del sketch de la losa (compatible con Create(7 params))."""
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
        return curves if curves else None
    except Exception:
        return None


def _curvas_a_curveloop(curves):
    """Convierte una lista de Curve en un CurveLoop. Retorna None si falla."""
    if not curves or len(curves) < 2:
        return None
    try:
        curve_list = List[Curve](curves)
        return CurveLoop.Create(curve_list)
    except Exception:
        return None


def _crear_model_curves_y_resaltar(document, uidocument, curves):
    """
    Crea ModelCurves para cada curva en la vista actual y las selecciona para resaltarlas.
    Retorna True si se crearon y se resaltaron, False en caso contrario.
    """
    if document is None or uidocument is None or not curves:
        return False
    loop = _curvas_a_curveloop(curves)
    if loop is None:
        return False
    if not loop.HasPlane():
        return False
    plano = loop.GetPlane()
    if plano is None:
        origen = curves[0].GetEndPoint(0) if curves else XYZ(0, 0, 0)
        plano = Plane.CreateByNormalAndOrigin(XYZ(0, 0, 1), origen)
    ids_creados = []
    with Transaction(document, "RPS Resaltar curvas sketch") as t:
        t.Start()
        try:
            sketch_plane = SketchPlane.Create(document, plano)
            for c in curves:
                if c is not None and c.IsBound:
                    mc = document.Create.NewModelCurve(c, sketch_plane)
                    if mc is not None:
                        ids_creados.append(mc.Id)
        except Exception as ex:
            print("Error creando lineas: {}".format(ex))
        t.Commit()
    if ids_creados:
        uidocument.Selection.SetElementIds(List[ElementId](ids_creados))
        return True
    return False


# ── Ejecución en RPS: selecciona un Floor y ejecuta el script ─────────────────
if doc is not None and uidoc is not None:
    sel = list(uidoc.Selection.GetElementIds())
    if sel:
        elem = doc.GetElement(sel[0])
        if elem is not None and elem.Category is not None:
            from Autodesk.Revit.DB import BuiltInCategory
            if elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Floors):
                curvas = _obtener_curvas_sketch(elem, doc)
                if curvas:
                    print("Curvas del sketch: {} segmentos".format(len(curvas)))
                    if _crear_model_curves_y_resaltar(doc, uidoc, curvas):
                        print("Curvas creadas como lineas de modelo y resaltadas en la vista.")
                    else:
                        print("No se pudieron crear las lineas de modelo para resaltar.")
                else:
                    print("No se pudieron obtener curvas del sketch.")
            else:
                print("Selecciona un Suelo (Floor) y vuelve a ejecutar.")
        else:
            print("No se pudo obtener el elemento.")
    else:
        print("Selecciona un Suelo (Floor) y ejecuta el script.")
else:
    print("Ejecuta este script dentro de Revit Python Shell (RPS).")
