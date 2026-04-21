# -*- coding: utf-8 -*-
"""
RPS: Cotas de ancho y largo en planta para fundacion aislada seleccionada.

- Selecciona una fundacion aislada (FamilyInstance de categoria Structural Foundation).
- Genera dos cotas lineales en la vista activa en planta:
    · Horizontal (ancho): entre las dos caras verticales extremas en la direccion X.
    · Vertical  (largo) : entre las dos caras verticales extremas en la direccion Y.
- Las lineas de cota se desplazan fuera del contorno de la fundacion.
- Compatible con Revit 2024, ejecutable desde RevitPythonShell (RPS).

Estrategia de referencias (en orden de preferencia):
  1. FamilyInstance.GetReferences(Left/Right/Front/Back): funciona si la familia
     tiene planos de referencia con esos roles asignados explicitamente.
  2. Caras planas verticales del solido: fallback universal; agrupa las caras por
     direccion de normal y toma las dos mas extremas en cada eje.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    FamilyInstance,
    FamilyInstanceReferenceType,
    GeometryInstance,
    Line,
    Options,
    PlanarFace,
    ReferenceArray,
    Solid,
    Transaction,
    ViewDetailLevel,
    ViewPlan,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
_TITULO = u"BIMTools \u2014 Cotas Fundaci\u00f3n"
_OFFSET_MM = 500.0   # desplazamiento de la linea de cota respecto al borde
_MARGEN_MM = 100.0   # extension lateral de la linea de cota mas alla del borde
_MM_POR_PIE = 304.8


def _mm_a_pies(mm):
    return float(mm) / _MM_POR_PIE


# ---------------------------------------------------------------------------
# Helpers de vista
# ---------------------------------------------------------------------------

def _es_vista_planta(view):
    """
    True si la vista activa acepta cotas en planta.
    Usa isinstance(view, ViewPlan) para evitar el fallo de IronPython 3.4
    con ViewType enum (StructuralPlan no existe; el nombre correcto es EngineeringPlan).
    ViewPlan cubre FloorPlan, CeilingPlan y EngineeringPlan (planta estructural).
    """
    if view is None:
        return False
    try:
        return isinstance(view, ViewPlan)
    except Exception:
        return False


def _z_plano_vista(view):
    """
    Coordenada Z del plano de trabajo de la vista (para puntos de linea de cota).
    Intenta SketchPlane, luego nivel, luego Origin.Z como respaldo.
    """
    try:
        sp = view.SketchPlane
        if sp is not None:
            return float(sp.GetPlane().Origin.Z)
    except Exception:
        pass
    try:
        lvl = view.GenLevel
        if lvl is not None:
            return float(lvl.Elevation)
    except Exception:
        pass
    try:
        return float(view.Origin.Z)
    except Exception:
        return 0.0


# ---------------------------------------------------------------------------
# Estrategia 1: planos de referencia nombrados (Left/Right/Front/Back)
# ---------------------------------------------------------------------------

_TIPOS_REF_NAMED = [
    (u"left",  FamilyInstanceReferenceType.Left),
    (u"right", FamilyInstanceReferenceType.Right),
    (u"front", FamilyInstanceReferenceType.Front),
    (u"back",  FamilyInstanceReferenceType.Back),
]


def _refs_por_nombre(fi):
    """
    Intenta extraer referencias Left/Right/Front/Back de la instancia.
    Devuelve dict parcial; vacio si la familia no tiene esos roles.
    """
    resultado = {}
    for nombre, tipo in _TIPOS_REF_NAMED:
        try:
            lista = fi.GetReferences(tipo)
            if lista is not None and lista.Count > 0:
                resultado[nombre] = lista[0]
        except Exception:
            pass
    return resultado


# ---------------------------------------------------------------------------
# Estrategia 2: caras planas verticales del solido (fallback universal)
# ---------------------------------------------------------------------------

def _extraer_solidos(geom_elem):
    """Recorre un GeometryElement y devuelve todos los Solid con volumen > 0."""
    solidos = []
    if geom_elem is None:
        return solidos
    try:
        for obj in geom_elem:
            if isinstance(obj, Solid):
                try:
                    if float(obj.Volume) > 1e-9:
                        solidos.append(obj)
                except Exception:
                    pass
            elif isinstance(obj, GeometryInstance):
                try:
                    for sub in obj.GetInstanceGeometry():
                        if isinstance(sub, Solid):
                            try:
                                if float(sub.Volume) > 1e-9:
                                    solidos.append(sub)
                            except Exception:
                                pass
                except Exception:
                    pass
    except Exception:
        pass
    return solidos


def _normal_canon(face):
    """
    Devuelve la normal de la cara normalizada a la semiesfera positiva
    (componente X positiva; si X=0 entonces Y positiva).
    Retorna (nx, ny) redondeados a 1 decimal, o (None, None) si es invalido.
    """
    try:
        n = face.FaceNormal
        nx, ny = float(n.X), float(n.Y)
        length = (nx * nx + ny * ny) ** 0.5
        if length < 1e-6:
            return None, None
        nx /= length
        ny /= length
        if nx < -0.001 or (abs(nx) < 0.001 and ny < 0):
            nx, ny = -nx, -ny
        return round(nx, 1), round(ny, 1)
    except Exception:
        return None, None


def _refs_por_caras(fi, view):
    """
    Fallback: extrae referencias de las caras planas verticales mas extremas del solido.
    Agrupa caras por direccion de normal canonizada y, dentro de cada grupo, toma
    las dos caras con mayor separacion (proyeccion extrema sobre el eje de la normal).

    Devuelve dict con claves 'left'/'right' (primer eje) y 'front'/'back' (segundo eje).
    """
    solidos = []
    for usar_vista in [view, None]:
        try:
            opts = Options()
            opts.ComputeReferences = True
            opts.DetailLevel = ViewDetailLevel.Fine
            if usar_vista is not None:
                opts.View = usar_vista
            solidos = _extraer_solidos(fi.get_Geometry(opts))
            if solidos:
                break
        except Exception:
            pass

    if not solidos:
        return {}

    # Recopilar caras planas verticales con referencia valida
    caras_verticales = []
    for solid in solidos:
        try:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                ref = face.Reference
                if ref is None:
                    continue
                n = face.FaceNormal
                if abs(float(n.Z)) > 0.5:   # ignorar caras horizontales (tapa/fondo)
                    continue
                caras_verticales.append(face)
        except Exception:
            pass

    if len(caras_verticales) < 2:
        return {}

    # Agrupar por direccion canonizada de normal
    grupos = {}
    for face in caras_verticales:
        nx_r, ny_r = _normal_canon(face)
        if nx_r is None:
            continue
        key = (nx_r, ny_r)
        grupos.setdefault(key, []).append(face)

    # Para cada grupo: proyectar origenes sobre la normal y tomar los dos extremos
    pares = []   # list of (ref_min, ref_max, nx_r, ny_r)
    for (nx_r, ny_r), caras in grupos.items():
        if len(caras) < 2:
            continue
        length = (nx_r * nx_r + ny_r * ny_r) ** 0.5
        if length < 1e-6:
            continue
        nx_u = nx_r / length
        ny_u = ny_r / length
        proyecciones = []
        for face in caras:
            try:
                o = face.Origin
                proj = float(o.X) * nx_u + float(o.Y) * ny_u
                proyecciones.append((proj, face))
            except Exception:
                pass
        if len(proyecciones) < 2:
            continue
        proyecciones.sort(key=lambda x: x[0])
        ref_min = proyecciones[0][1].Reference
        ref_max = proyecciones[-1][1].Reference
        if ref_min is not None and ref_max is not None:
            pares.append((ref_min, ref_max, nx_r, ny_r))

    if not pares:
        return {}

    # Separar pares segun si la normal es mayoritariamente X o Y.
    # Normal mayoritaria en X → caras perpendiculares a X → miden el ANCHO → etiqueta left/right.
    # Normal mayoritaria en Y → caras perpendiculares a Y → miden el LARGO → etiqueta front/back.
    pares_x = [(r0, r1, nx, ny) for r0, r1, nx, ny in pares if abs(nx) >= abs(ny)]
    pares_y = [(r0, r1, nx, ny) for r0, r1, nx, ny in pares if abs(ny) > abs(nx)]

    resultado = {}
    if pares_x:
        ref_min, ref_max = pares_x[0][0], pares_x[0][1]
        resultado[u"left"]  = ref_min
        resultado[u"right"] = ref_max
    if pares_y:
        ref_min, ref_max = pares_y[0][0], pares_y[0][1]
        resultado[u"front"] = ref_min
        resultado[u"back"]  = ref_max

    # Si solo se encontro un par (fundacion no rectangular) se reutiliza para ambos ejes.
    if pares and not pares_x and not pares_y:
        ref_min, ref_max = pares[0][0], pares[0][1]
        resultado[u"left"]  = ref_min
        resultado[u"right"] = ref_max

    return resultado


# ---------------------------------------------------------------------------
# Obtener referencias (combina ambas estrategias)
# ---------------------------------------------------------------------------

def _obtener_referencias(fi, view):
    """
    Devuelve dict de referencias y el metodo usado ('nombradas' o 'caras').
    El dict puede tener claves: left, right, front, back.
    """
    refs = _refs_por_nombre(fi)
    tiene_lr = u"left" in refs and u"right" in refs
    tiene_fb = u"front" in refs and u"back" in refs
    if tiene_lr or tiene_fb:
        return refs, u"nombradas"

    refs = _refs_por_caras(fi, view)
    return refs, u"caras"


# ---------------------------------------------------------------------------
# Bounding box
# ---------------------------------------------------------------------------

def _bbox_elemento(elem, view):
    """Bounding box del elemento; intenta con la vista y sin ella."""
    bb = None
    try:
        bb = elem.get_BoundingBox(view)
    except Exception:
        pass
    if bb is None:
        try:
            bb = elem.get_BoundingBox(None)
        except Exception:
            pass
    return bb


# ---------------------------------------------------------------------------
# Creacion de cotas
# ---------------------------------------------------------------------------

def _crear_cota(doc, view, ref_a, ref_b, pt1, pt2):
    """
    Crea una Dimension lineal en la vista entre ref_a y ref_b.
    Devuelve la Dimension o None si falla.
    """
    try:
        ra = ReferenceArray()
        ra.Append(ref_a)
        ra.Append(ref_b)
        linea = Line.CreateBound(pt1, pt2)
        return doc.Create.NewDimension(view, linea, ra)
    except Exception:
        return None


def _cotar_fundacion(doc, view, fi):
    """
    Coloca cotas de ancho y largo en planta para la fundacion fi.
    Devuelve (creadas, errores).
    """
    refs, metodo = _obtener_referencias(fi, view)
    tiene_lr = u"left" in refs and u"right" in refs
    tiene_fb = u"front" in refs and u"back" in refs

    if not tiene_lr and not tiene_fb:
        return 0, [
            u"No se encontraron referencias v\u00e1lidas para cotar la fundaci\u00f3n.\n"
            u"Se intentaron planos de referencia nombrados (Left/Right/Front/Back)\n"
            u"y caras del s\u00f3lido sin \u00e9xito."
        ]

    bb = _bbox_elemento(fi, view)
    if bb is None:
        return 0, [u"No se pudo obtener el bounding box del elemento."]

    z = _z_plano_vista(view)
    offset = _mm_a_pies(_OFFSET_MM)
    margen = _mm_a_pies(_MARGEN_MM)

    creadas = 0
    errores = []

    # Cota horizontal (ancho): linea paralela a X, desplazada hacia Y- (abajo)
    if tiene_lr:
        y_linea = bb.Min.Y - offset
        pt1 = XYZ(bb.Min.X - margen, y_linea, z)
        pt2 = XYZ(bb.Max.X + margen, y_linea, z)
        dim = _crear_cota(doc, view, refs[u"left"], refs[u"right"], pt1, pt2)
        if dim is not None:
            creadas += 1
        else:
            errores.append(u"No se pudo crear la cota horizontal (ancho) [metodo: {}].".format(metodo))

    # Cota vertical (largo): linea paralela a Y, desplazada hacia X- (izquierda)
    if tiene_fb:
        x_linea = bb.Min.X - offset
        pt1 = XYZ(x_linea, bb.Min.Y - margen, z)
        pt2 = XYZ(x_linea, bb.Max.Y + margen, z)
        dim = _crear_cota(doc, view, refs[u"front"], refs[u"back"], pt1, pt2)
        if dim is not None:
            creadas += 1
        else:
            errores.append(u"No se pudo crear la cota vertical (largo) [metodo: {}].".format(metodo))

    return creadas, errores


# ---------------------------------------------------------------------------
# Validacion de seleccion
# ---------------------------------------------------------------------------

def _validar_seleccion(uidoc, doc):
    """
    Valida que haya exactamente una FamilyInstance de categoria Structural Foundation.
    Devuelve (FamilyInstance, None) o (None, mensaje_error).
    """
    try:
        ids = list(uidoc.Selection.GetElementIds())
    except Exception as ex:
        return None, u"Error al leer la selecci\u00f3n: {}".format(ex)

    if not ids:
        return None, (
            u"No hay ninguna fundaci\u00f3n seleccionada.\n"
            u"Seleccione una fundaci\u00f3n aislada en la vista activa y vuelva a ejecutar."
        )

    if len(ids) > 1:
        return None, (
            u"Hay {} elementos seleccionados.\n"
            u"Seleccione \u00fanicamente una fundaci\u00f3n aislada.".format(len(ids))
        )

    elem = doc.GetElement(ids[0])
    if elem is None:
        return None, u"El elemento seleccionado no existe en el documento."

    if not isinstance(elem, FamilyInstance):
        return None, (
            u"El elemento seleccionado no es una instancia de familia "
            u"(tipo: {}).\nSeleccione una fundaci\u00f3n aislada.".format(
                type(elem).__name__
            )
        )

    cat = elem.Category
    if cat is None or cat.Id.IntegerValue != int(BuiltInCategory.OST_StructuralFoundation):
        nombre_cat = cat.Name if cat else u"(sin categoria)"
        return None, (
            u"El elemento pertenece a la categor\u00eda '{}'.\n"
            u"Seleccione una fundaci\u00f3n aislada (Structural Foundation).".format(nombre_cat)
        )

    return elem, None


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def ejecutar(uidoc, doc):
    """Rutina principal. Llamada desde _main() con __revit__ activo."""
    view = doc.ActiveView

    if not _es_vista_planta(view):
        TaskDialog.Show(
            _TITULO,
            u"La vista activa no es una planta.\n"
            u"Abra o active una planta estructural o de piso y vuelva a ejecutar.",
        )
        return

    fi, error = _validar_seleccion(uidoc, doc)
    if fi is None:
        TaskDialog.Show(_TITULO, error)
        return

    creadas = 0
    errores = []

    with Transaction(doc, u"BIMTools: Cotas Fundaci\u00f3n en Planta") as t:
        t.Start()
        try:
            creadas, errores = _cotar_fundacion(doc, view, fi)
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            TaskDialog.Show(_TITULO, u"Error inesperado al crear cotas:\n{}".format(ex))
            return

    if creadas == 0:
        msg = u"No se pudo crear ninguna cota."
        if errores:
            msg += u"\n\nDetalles:\n" + u"\n".join(errores)
        TaskDialog.Show(_TITULO, msg)
        return

    msg = u"Se crearon {} cota(s) para la fundaci\u00f3n.".format(creadas)
    if errores:
        msg += u"\n\nAdvertencias:\n" + u"\n".join(errores)
    TaskDialog.Show(_TITULO, msg)
    print(msg)


def _main():
    try:
        _doc = __revit__.ActiveUIDocument.Document  # noqa: F821
        _uidoc = __revit__.ActiveUIDocument          # noqa: F821
    except NameError:
        TaskDialog.Show(
            _TITULO,
            u"Este script debe ejecutarse desde RevitPythonShell (RPS) o pyRevit "
            u"con la variable __revit__ disponible.",
        )
        return
    ejecutar(_uidoc, _doc)


if __name__ == "__main__":
    _main()
