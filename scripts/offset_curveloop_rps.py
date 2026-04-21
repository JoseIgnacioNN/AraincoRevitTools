# -*- coding: utf-8 -*-
# IMPORTANTE: Ejecutar con File > Run script (archivo .py). No pegar el codigo linea a linea en la consola.
"""
Script RPS: Offset de un CurveLoop.
Ejecutable en Revit Python Shell (RPS) - Revit 2024-2026 | IronPython 3.4

Obtiene el primer elemento de la seleccion (Suelo/Floor o Habitacion/Room),
extrae su CurveLoop de contorno, aplica un offset con CurveLoop.CreateViaOffset()
y crea lineas de modelo temporales para visualizar el bucle original y el offset.

Unidades: La API de Revit trabaja internamente en PIES (feet). 1.5 pies es el
offset por defecto; el valor se pasa directamente a CreateViaOffset().

Uso en RPS:
  1. Selecciona un Floor o una Room en Revit.
  2. Ejecuta el script (File > Run script, no pegar linea a linea).
  3. Se dibujan en la vista activa el CurveLoop original y el offset.
"""

import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    CurveLoop,
    Options,
    PlanarFace,
    Plane,
    SketchPlane,
    Solid,
    SpatialElementBoundaryOptions,
    Transaction,
    UnitUtils,
    UnitTypeId,
    XYZ,
)

# Variables predefinidas RPS: doc = documento activo, uidoc = UIDocument, selection = seleccion.
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    selection = uidoc.Selection.GetElementIds()
except NameError:
    doc = uidoc = None
    selection = []


# Distancia de offset en unidades internas de Revit (PIES).
# La API usa pies para longitudes; 1.5 = 1.5 pies.
OFFSET_PIES = 1.5

# Vector normal del plano del offset: apunta hacia arriba (+Z).
# Define el lado derecho del bucle al caminar por las curvas (regla de la mano derecha).
NORMAL_OFFSET = XYZ(0, 0, 1)


def _obtener_cara_superior(geom_elem):
    """Busca la cara planar con normal hacia +Z en la geometria del elemento."""
    if geom_elem is None:
        return None
    for geom_obj in geom_elem:
        solid = geom_obj if isinstance(geom_obj, Solid) else None
        if solid is None or solid.Faces.Size == 0:
            continue
        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue
            if face.FaceNormal and face.FaceNormal.Z >= 0.9:
                return face
    return None


def _curveloop_desde_floor(document, floor_element):
    """
    Extrae el CurveLoop exterior (perimetro) de la cara superior de un Suelo/Floor.
    Retorna (CurveLoop, origen_plano) o (None, None).
    """
    if document is None or floor_element is None:
        return None, None
    if not floor_element.Category or floor_element.Category.Id.IntegerValue != int(BuiltInCategory.OST_Floors):
        return None, None

    opts = Options()
    opts.ComputeReferences = False
    geom_elem = floor_element.get_Geometry(opts)
    if geom_elem is None:
        return None, None

    cara = _obtener_cara_superior(geom_elem)
    if cara is None:
        return None, None

    loops_raw = cara.GetEdgesAsCurveLoops()
    loops = list(loops_raw) if loops_raw else []
    if not loops:
        return None, None

    # El bucle mas largo suele ser el perimetro exterior. (sum sin default: compatible Python 3.4)
    def _longitud(cl):
        return sum(c.Length for c in cl if c and c.IsBound)
    exterior = max(loops, key=_longitud)
    origen = cara.Origin if cara.Origin else exterior.GetCurveLoopIterator().Current.GetEndPoint(0)
    return exterior, origen


def _curveloop_desde_room(room_element):
    """
    Extrae el primer bucle de contorno de una Habitacion/Room como CurveLoop.
    Retorna (CurveLoop, origen_plano) o (None, None).
    """
    if room_element is None:
        return None, None
    if not room_element.Category or room_element.Category.Id.IntegerValue != int(BuiltInCategory.OST_Rooms):
        return None, None
    if not hasattr(room_element, "GetBoundarySegments"):
        return None, None

    opts = SpatialElementBoundaryOptions()
    segmentos = room_element.GetBoundarySegments(opts)
    # GetBoundarySegments retorna IList de IList de BoundarySegment; vacio si la habitacion no tiene contorno.
    if segmentos is None or len(list(segmentos)) == 0:
        return None, None

    # Primer bucle de contorno (lista de BoundarySegment).
    primer_bucle = list(segmentos[0])
    curvas = []
    for seg in primer_bucle:
        c = seg.GetCurve()
        if c is not None:
            curvas.append(c)
    if not curvas:
        return None, None

    try:
        loop = CurveLoop.Create(curvas)
    except Exception:
        return None, None

    # Origen del plano: primer punto del bucle.
    origen = curvas[0].GetEndPoint(0) if curvas else XYZ(0, 0, 0)
    return loop, origen


def _obtener_curveloop_y_plano(document, element):
    """
    Obtiene un CurveLoop y un punto origen segun el tipo de elemento.
    Soporta Floor (Suelo) y Room (Habitacion).
    """
    loop_floor, origen_floor = _curveloop_desde_floor(document, element)
    if loop_floor is not None:
        return loop_floor, origen_floor

    loop_room, origen_room = _curveloop_desde_room(element)
    if loop_room is not None:
        return loop_room, origen_room

    return None, None


def _crear_lineas_desde_curveloop(document, curve_loop, sketch_plane, nombre_transaccion):
    """
    Crea lineas de modelo (ModelCurve) para cada curva del CurveLoop,
    en el plano dado. Las lineas se crean en el documento para visualizacion.
    """
    if document is None or curve_loop is None or sketch_plane is None:
        return
    with Transaction(document, nombre_transaccion) as t:
        t.Start()
        try:
            for curve in curve_loop:
                if curve is not None and curve.IsBound:
                    document.Create.NewModelCurve(curve, sketch_plane)
        except Exception as ex:
            print("Error creando lineas: {}".format(ex))
        t.Commit()


def run(document, uidocument):
    """
    Logica principal: toma el primer elemento de la seleccion, obtiene su CurveLoop,
    aplica offset con CreateViaOffset y crea lineas de modelo para verificacion visual.
    """
    if document is None or uidocument is None:
        print("Error: doc/uidoc no disponibles. Ejecuta el script dentro de Revit Python Shell (RPS).")
        return

    elem_ids = list(uidocument.Selection.GetElementIds())
    if not elem_ids:
        print("No hay seleccion. Selecciona un Suelo (Floor) o una Habitacion (Room) y vuelve a ejecutar.")
        return

    element = document.GetElement(elem_ids[0])
    if element is None:
        print("No se pudo obtener el elemento.")
        return

    # Obtener CurveLoop y origen del plano (segun Floor o Room).
    curve_loop_original, origen_plano = _obtener_curveloop_y_plano(document, element)
    if curve_loop_original is None:
        print("El elemento seleccionado no es un Floor ni una Room, o no se pudo extraer un contorno (CurveLoop).")
        return

    # Comprobar que el bucle sea planar; la API lo exige para CreateViaOffset.
    if not curve_loop_original.HasPlane():
        print("El CurveLoop no es planar. No se puede aplicar offset.")
        return

    # Crear el CurveLoop con offset usando el metodo estatico CreateViaOffset.
    # Parametros: bucle original, distancia en pies (unidades internas), normal del plano.
    try:
        curve_loop_offset = CurveLoop.CreateViaOffset(curve_loop_original, OFFSET_PIES, NORMAL_OFFSET)
    except Exception as ex:
        print("CreateViaOffset fallo (curvas complejas o offset invalido): {}".format(ex))
        return

    # Plano para las lineas de modelo: mismo plano que el bucle (vista en planta).
    plano = curve_loop_original.GetPlane()
    if plano is None:
        plano = Plane.CreateByNormalAndOrigin(NORMAL_OFFSET, origen_plano)

    sketch_plane = SketchPlane.Create(document, plano)

    # Transaccion: crear lineas de modelo para el bucle original y el offset.
    _crear_lineas_desde_curveloop(document, curve_loop_original, sketch_plane, "RPS Lineas CurveLoop original")
    _crear_lineas_desde_curveloop(document, curve_loop_offset, sketch_plane, "RPS Lineas CurveLoop offset")

    # Informacion en consola (unidades: metros para lectura).
    largo_orig = curve_loop_original.GetExactLength()
    largo_off = curve_loop_offset.GetExactLength()
    offset_m = UnitUtils.ConvertFromInternalUnits(OFFSET_PIES, UnitTypeId.Meters)
    largo_orig_m = UnitUtils.ConvertFromInternalUnits(largo_orig, UnitTypeId.Meters)
    largo_off_m = UnitUtils.ConvertFromInternalUnits(largo_off, UnitTypeId.Meters)

    print("-" * 50)
    print("Offset de CurveLoop completado.")
    print("  Elemento: {} (ID: {})".format(element.Name or "(sin nombre)", element.Id.IntegerValue))
    print("  Offset: {:.4f} m ({:.2f} pies, unidades internas)".format(offset_m, OFFSET_PIES))
    print("  Longitud original: {:.4f} m".format(largo_orig_m))
    print("  Longitud con offset: {:.4f} m".format(largo_off_m))
    print("  Se han creado lineas de modelo en la vista para verificacion.")
    print("-" * 50)


# Ejecucion al cargar en RPS
if doc is not None and uidoc is not None:
    run(doc, uidoc)
else:
    print("Ejecuta este script dentro de Revit Python Shell (RPS).")
