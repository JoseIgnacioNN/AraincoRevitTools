# -*- coding: utf-8 -*-
"""
Script RPS: Losa + Shaft Openings por geometría (superficie y BBox como sólido).
Revit 2024+ | Revit Python Shell (RPS) | IronPython 3.4

Flujo:
  1. Seleccionas la losa (Floor).
  2. De la losa se obtiene la superficie superior y todas sus curvas (CurveLoops).
  3. Se obtienen todos los shaft openings del proyecto.
  4. De cada shaft se obtiene su BoundingBox, se convierte en un sólido (caja)
     en coordenadas de documento, y se evalúa intersección booleana con la
     geometría de la losa. Solo se conservan los que intersectan.

Tras ejecutar, en RPS quedan definidas:
  ref_losa                  — Reference al elemento losa
  elemento_losa              — Element (Floor) seleccionado
  cara_superior_losa         — PlanarFace de la cara superior
  curveloops_superficie_losa — list[CurveLoop]: todos los bucles de la cara (perímetro + huecos)
  curvas_superficie_losa     — list[Curve]: todas las curvas de esa superficie
  shaft_openings_proyecto   — list[Element]: todos los shaft openings del proyecto
  shaft_openings_que_intersectan — list[Element]: los que intersectan la losa (bbox→sólido vs losa)
"""

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Arc,
    BooleanOperationsType,
    BooleanOperationsUtils,
    BoundingBoxIntersectsFilter,
    BuiltInCategory,
    CurveLoop,
    FilteredElementCollector,
    GeometryCreationUtilities,
    GeometryInstance,
    Line,
    Options,
    Outline,
    PlanarFace,
    Solid,
    Transaction,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

# ── Boilerplate pyRevit / RPS ──────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
except NameError:
    doc = uidoc = None


class FiltroLosa(ISelectionFilter):
    """Filtro de selección: solo permite elementos de categoría Floor (losa)."""

    def AllowElement(self, element):
        if element is None or element.Category is None:
            return False
        return element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Floors)

    def AllowReference(self, reference, position):
        return True


def _categoria_opening_id():
    """
    Retorna el IntegerValue de la categoría Opening/Shaft para este Revit.
    En algunas versiones es OST_Opening, en otras OST_ShaftOpening. Si no existe, None.
    """
    for attr in ("OST_ShaftOpening", "OST_Opening"):
        try:
            cat = getattr(BuiltInCategory, attr, None)
            if cat is not None:
                return int(cat)
        except Exception:
            continue
    return None


class FiltroOpening(ISelectionFilter):
    """Filtro de selección: solo permite elementos de categoría Opening/Shaft."""

    def AllowElement(self, element):
        if element is None or element.Category is None:
            return False
        cid = _categoria_opening_id()
        if cid is None:
            return True
        return element.Category.Id.IntegerValue == cid

    def AllowReference(self, reference, position):
        return True


def pick_losa(document, uidoocument, mensaje=None):
    """
    Pide al usuario que seleccione un elemento (cualquier tipo). Sin filtro de categoría.

    Returns:
        tuple: (Reference, Element). (None, None) si cancela.
    """
    if document is None or uidoocument is None:
        return None, None
    if mensaje is None:
        mensaje = "Selecciona el primer elemento (losa)."
    try:
        ref = uidoocument.Selection.PickObject(ObjectType.Element, mensaje)
    except Exception:
        return None, None
    if ref is None:
        return None, None
    elem = document.GetElement(ref.ElementId)
    return ref, elem


def pick_opening(document, uidoocument, mensaje=None):
    """
    Pide al usuario que seleccione un elemento (cualquier tipo). Sin filtro de categoría.

    Returns:
        tuple: (Reference, Element). (None, None) si cancela.
    """
    if document is None or uidoocument is None:
        return None, None
    if mensaje is None:
        mensaje = "Selecciona el segundo elemento (shaft opening)."
    try:
        ref = uidoocument.Selection.PickObject(ObjectType.Element, mensaje)
    except Exception:
        return None, None
    if ref is None:
        return None, None
    elem = document.GetElement(ref.ElementId)
    return ref, elem


def _obtener_cara_superior_y_curvas(elemento_losa):
    """
    De la losa (Floor) obtiene la cara superior (PlanarFace con normal +Z)
    y todas las curvas de esa superficie (GetEdgesAsCurveLoops).

    Returns:
        tuple: (PlanarFace, list[CurveLoop], list[Curve]) o (None, [], []) si falla.
    """
    if elemento_losa is None:
        return None, [], []
    opts = Options()
    opts.ComputeReferences = False
    try:
        geom_elem = elemento_losa.get_Geometry(opts)
    except Exception:
        return None, [], []
    if geom_elem is None:
        return None, [], []
    cara = None
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
                cara = face
                break
        if cara is not None:
            break
    if cara is None:
        return None, [], []
    try:
        raw_loops = cara.GetEdgesAsCurveLoops()
        curveloops = list(raw_loops) if raw_loops else []
    except Exception:
        curveloops = []
    curvas = []
    for cl in curveloops:
        for c in cl:
            if c is not None and c.IsBound:
                curvas.append(c)
    return cara, curveloops, curvas


def _todos_shaft_openings(document):
    """Devuelve todos los elementos Opening/Shaft del proyecto."""
    if document is None:
        return []
    for attr in ("OST_ShaftOpening", "OST_Opening"):
        try:
            cat = getattr(BuiltInCategory, attr, None)
            if cat is not None:
                elems = FilteredElementCollector(document).OfCategory(cat).ToElements()
                return list(elems) if elems else []
        except Exception:
            continue
    cid = _categoria_opening_id()
    if cid is not None:
        col = FilteredElementCollector(document)
        return [e for e in col if e is not None and e.Category is not None and e.Category.Id.IntegerValue == cid]
    return []


def _solid_desde_bbox(bbox):
    """
    Convierte un BoundingBox (get_BoundingBox) en un Solid (caja) en coordenadas
    de documento. Aplica Transform si existe para Min/Max.
    """
    if bbox is None:
        return None
    mn = getattr(bbox, "Min", None) or getattr(bbox, "Minimum", None)
    mx = getattr(bbox, "Max", None) or getattr(bbox, "Maximum", None)
    if mn is None or mx is None:
        return None
    tr = getattr(bbox, "Transform", None)
    if tr is not None:
        try:
            mn = tr.OfPoint(mn)
            mx = tr.OfPoint(mx)
        except Exception:
            pass
    z_min, z_max = mn.Z, mx.Z
    if z_max <= z_min:
        z_max = z_min + 0.01
    pt1 = XYZ(mn.X, mn.Y, z_min)
    pt2 = XYZ(mx.X, mn.Y, z_min)
    pt3 = XYZ(mx.X, mx.Y, z_min)
    pt4 = XYZ(mn.X, mx.Y, z_min)
    line1 = Line.CreateBound(pt1, pt2)
    line2 = Line.CreateBound(pt2, pt3)
    line3 = Line.CreateBound(pt3, pt4)
    line4 = Line.CreateBound(pt4, pt1)
    try:
        loop = CurveLoop.Create([line1, line2, line3, line4])
        solid = GeometryCreationUtilities.CreateExtrusionGeometry(
            [loop],
            XYZ(0, 0, 1),
            z_max - z_min,
        )
    except Exception:
        return None
    if solid is None or solid.Volume <= 0:
        return None
    return solid


def _obtener_solidos_elemento(elemento, options=None):
    """
    Extrae todos los Solid de la geometría de un elemento.
    Maneja GeometryElement con Solid directo o con GeometryInstance.
    """
    if elemento is None:
        return []
    if options is None:
        options = Options()
        options.ComputeReferences = False
    try:
        geom_elem = elemento.get_Geometry(options)
    except Exception:
        return []
    if geom_elem is None:
        return []
    solidos = []
    for obj in geom_elem:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Volume > 0:
            solidos.append(obj)
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                if inst_geom is not None:
                    for g in inst_geom:
                        if isinstance(g, Solid) and g.Volume > 0:
                            solidos.append(g)
            except Exception:
                pass
    return solidos


def _solidos_intersectan(solid_a, solid_b, tol_volumen=1e-6):
    """
    True si la intersección booleana de los dos sólidos tiene volumen > tol_volumen.
    """
    if solid_a is None or solid_b is None or solid_a.Volume <= 0 or solid_b.Volume <= 0:
        return False
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid_a, solid_b, BooleanOperationsType.Intersect
        )
        return inter is not None and inter.Volume > tol_volumen
    except Exception:
        return False


# Margen en pies (unidades internas Revit) para exagerar la altura del sólido del
# opening y asegurar que la extrusión cruce la losa al evaluar intersección.
_EXAGERACION_ALTURA_PIES = 50.0


def _curveloop_transladar_z(curve_loop, delta_z):
    """
    Crea un nuevo CurveLoop con cada curva trasladada en Z por delta_z.
    Soporta Line y Arc. Retorna None si falla.
    """
    if curve_loop is None or delta_z == 0:
        return curve_loop
    nuevas = []
    for c in curve_loop:
        if c is None or not c.IsBound:
            return None
        try:
            if isinstance(c, Line):
                p1 = c.GetEndPoint(0)
                p2 = c.GetEndPoint(1)
                n1 = XYZ(p1.X, p1.Y, p1.Z + delta_z)
                n2 = XYZ(p2.X, p2.Y, p2.Z + delta_z)
                nuevas.append(Line.CreateBound(n1, n2))
            elif isinstance(c, Arc):
                cen = c.Center
                e0 = c.GetEndPoint(0)
                e1 = c.GetEndPoint(1)
                nc = XYZ(cen.X, cen.Y, cen.Z + delta_z)
                ne0 = XYZ(e0.X, e0.Y, e0.Z + delta_z)
                ne1 = XYZ(e1.X, e1.Y, e1.Z + delta_z)
                nuevas.append(Arc.Create(ne0, ne1, nc))
            else:
                return None
        except Exception:
            return None
    try:
        return CurveLoop.Create(nuevas)
    except Exception:
        return None


def _solid_desde_opening_boundary(opening):
    """
    Construye un sólido a partir del contorno del Opening (BoundaryRect o BoundaryCurves).
    La altura se exagera hacia arriba y hacia abajo (_EXAGERACION_ALTURA_PIES) para
    garantizar que la extrusión cruce la losa al hacer la intersección booleana.
    Retorna un Solid o None si falla.
    """
    if opening is None:
        return None
    loop = None
    z_min = 0.0
    z_max = 1.0
    margin = _EXAGERACION_ALTURA_PIES
    try:
        is_rect = getattr(opening, "IsRectBoundary", False)
        if is_rect:
            rect = getattr(opening, "BoundaryRect", None)
            if rect is not None:
                mn = getattr(rect, "Min", None) or getattr(rect, "Minimum", None)
                mx = getattr(rect, "Max", None) or getattr(rect, "Maximum", None)
                if mn is not None and mx is not None:
                    tr = getattr(rect, "Transform", None)
                    if tr is not None:
                        try:
                            mn = tr.OfPoint(mn)
                            mx = tr.OfPoint(mx)
                        except Exception:
                            pass
                    z_min, z_max = mn.Z, mx.Z
                    if z_max <= z_min:
                        z_max = z_min + 0.1
                    z_base = z_min - margin
                    pt1 = XYZ(mn.X, mn.Y, z_base)
                    pt2 = XYZ(mx.X, mn.Y, z_base)
                    pt3 = XYZ(mx.X, mx.Y, z_base)
                    pt4 = XYZ(mn.X, mx.Y, z_base)
                    line1 = Line.CreateBound(pt1, pt2)
                    line2 = Line.CreateBound(pt2, pt3)
                    line3 = Line.CreateBound(pt3, pt4)
                    line4 = Line.CreateBound(pt4, pt1)
                    loop = CurveLoop.Create([line1, line2, line3, line4])
        if loop is None:
            boundary = getattr(opening, "BoundaryCurves", None)
            if boundary is not None:
                curvas = [c for c in boundary if c is not None and c.IsBound]
                if curvas:
                    try:
                        loop = CurveLoop.Create(curvas)
                    except Exception:
                        pass
                    try:
                        bbox_op = opening.get_BoundingBox(None)
                        if bbox_op is not None:
                            mn = getattr(bbox_op, "Min", None)
                            mx = getattr(bbox_op, "Max", None)
                            if mn is not None and mx is not None:
                                z_min, z_max = mn.Z, mx.Z
                    except Exception:
                        pass
        if loop is None:
            return None
        z_base = z_min - margin
        altura = (z_max - z_min) + 2.0 * margin
        if is_rect is False and loop is not None:
            loop_t = _curveloop_transladar_z(loop, z_base - z_min)
            if loop_t is not None:
                loop = loop_t
        altura = max(altura, 0.01)
        try:
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                [loop],
                XYZ(0, 0, 1),
                altura,
            )
        except Exception:
            return None
        if solid is None or solid.Volume <= 0:
            return None
        return solid
    except Exception:
        return None


def _elemento_intersecta_geometria(elemento_losa, elemento_candidato):
    """
    True si la losa y el candidato se intersectan por geometría (sin usar BoundingBox).
    - Si ambos tienen sólidos en get_Geometry: intersección booleana.
    - Si el candidato no tiene sólidos (shaft openings): se construye un sólido
      desde el contorno del opening (BoundaryRect/BoundaryCurves) y se hace
      intersección booleana con la losa.
    """
    solidos_losa = _obtener_solidos_elemento(elemento_losa)
    if not solidos_losa:
        return False
    solidos_cand = _obtener_solidos_elemento(elemento_candidato)
    if solidos_cand:
        for s_losa in solidos_losa:
            for s_cand in solidos_cand:
                if _solidos_intersectan(s_losa, s_cand):
                    return True
        return False
    solid_opening = _solid_desde_opening_boundary(elemento_candidato)
    if solid_opening is not None:
        for s_losa in solidos_losa:
            if _solidos_intersectan(s_losa, solid_opening):
                return True
    return False


def _bbox_intersect(bbox_a, bbox_b):
    """
    Comprueba si dos BoundingBox (get_BoundingBox) se solapan en X, Y y Z.
    """
    if bbox_a is None or bbox_b is None:
        return False
    min_a = getattr(bbox_a, "Min", None)
    max_a = getattr(bbox_a, "Max", None)
    min_b = getattr(bbox_b, "Min", None)
    max_b = getattr(bbox_b, "Max", None)
    if min_a is None or max_a is None or min_b is None or max_b is None:
        return False
    if max_a.X < min_b.X or max_b.X < min_a.X:
        return False
    if max_a.Y < min_b.Y or max_b.Y < min_a.Y:
        return False
    if max_a.Z < min_b.Z or max_b.Z < min_a.Z:
        return False
    return True


def shaft_openings_que_intersectan_por_bbox_solid(document, elemento_losa, vista=None):
    """
    Obtiene todos los shaft openings del proyecto; para cada uno construye un sólido
    desde su BoundingBox y evalúa intersección booleana con la geometría de la losa.
    Devuelve solo los openings cuyo bbox (como sólido) intersecta la losa.

    Args:
        document: Document (Revit DB Document).
        elemento_losa: Element (Floor).
        vista: View opcional para get_BoundingBox (None = modelo).

    Returns:
        list: Elementos Opening cuya caja (bbox como sólido) intersecta la losa.
    """
    if document is None or elemento_losa is None:
        return []
    solidos_losa = _obtener_solidos_elemento(elemento_losa)
    if not solidos_losa:
        return []
    todos = _todos_shaft_openings(document)
    resultado = []
    for opening in todos:
        if opening is None or opening.Id == elemento_losa.Id:
            continue
        try:
            bbox = opening.get_BoundingBox(vista)
        except Exception:
            continue
        solid_bbox = _solid_desde_bbox(bbox)
        if solid_bbox is None:
            continue
        for s_losa in solidos_losa:
            if _solidos_intersectan(s_losa, solid_bbox):
                resultado.append(opening)
                break
    return resultado


def openings_que_intersectan_losa(document, elemento_losa, vista=None):
    """
    Devuelve solo los elementos Opening que intersectan la losa en geometría real.
    Primero obtiene candidatos por BoundingBoxIntersectsFilter + categoría Opening;
    luego filtra por intersección booleana de sólidos (solo los que realmente intersectan).

    Args:
        document: Document (Revit DB Document).
        elemento_losa: Element (Floor).
        vista: View opcional para get_BoundingBox (None = modelo).

    Returns:
        list: Elementos Opening cuya geometría (sólidos) intersecta la de la losa.
    """
    if document is None or elemento_losa is None:
        return []
    try:
        bbox_losa = elemento_losa.get_BoundingBox(vista)
    except Exception:
        bbox_losa = None
    if bbox_losa is None:
        return []
    min_pt = getattr(bbox_losa, "Min", None)
    max_pt = getattr(bbox_losa, "Max", None)
    if min_pt is None or max_pt is None:
        return []
    try:
        outline = Outline(min_pt, max_pt)
        filtro_bbox = BoundingBoxIntersectsFilter(outline)
        elementos = FilteredElementCollector(document).WherePasses(
            filtro_bbox
        ).ToElements()
    except Exception:
        return []
    id_losa = elemento_losa.Id
    cat_opening_id = _categoria_opening_id()
    if cat_opening_id is not None:
        candidatos = [
            e for e in elementos
            if e is not None and e.Id != id_losa
            and e.Category is not None
            and e.Category.Id.IntegerValue == cat_opening_id
        ]
    else:
        candidatos = [e for e in elementos if e is not None and e.Id != id_losa]
    resultado = [
        e for e in candidatos
        if _elemento_intersecta_geometria(elemento_losa, e)
    ]
    return resultado


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


def _curvas_desde_opening(opening, z_plano=None):
    """
    Obtiene las curvas del contorno de un Opening (BoundaryCurves o BoundaryRect).
    Retorna una lista de Curve. Si z_plano se indica, proyecta al plano Z (para vista en planta).
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


def mostrar_boundaries_openings_en_vista(document, uidoocument, opening_elements):
    """
    Crea DetailCurves en la vista activa con el contorno de cada Opening
    (BoundaryCurves o BoundaryRect). Retorna el número de curvas creadas.
    """
    if document is None or uidoocument is None or not opening_elements:
        return 0
    vista = uidoocument.ActiveView
    if vista is None:
        return 0
    z_vista = vista.Origin.Z if vista.Origin else 0.0
    creadas = 0
    t = Transaction(document, "RPS Shaft Openings boundaries")
    t.Start()
    try:
        for opening in opening_elements:
            curvas = _curvas_desde_opening(opening, z_vista)
            if not curvas:
                curvas = _curvas_desde_opening(opening, None)
            for curve in curvas:
                if curve is None or not curve.IsBound:
                    continue
                try:
                    document.Create.NewDetailCurve(vista, curve)
                    creadas += 1
                except Exception:
                    pass
        t.Commit()
    except Exception:
        t.RollBack()
        return 0
    return creadas


# ── Ejecución vía RPS: selección de losa → cara superior, curvas, shafts, intersección ─
ref_losa = None
elemento_losa = None
cara_superior_losa = None
curveloops_superficie_losa = []
curvas_superficie_losa = []
shaft_openings_proyecto = []
shaft_openings_que_intersectan = []

if doc is not None and uidoc is not None:
    rps_window = None
    try:
        rps_window = __window__
    except NameError:
        pass
    if rps_window is not None:
        rps_window.Hide()
    try:
        ref_losa, elemento_losa = pick_losa(doc, uidoc, "Selecciona la losa.")
    finally:
        if rps_window is not None:
            rps_window.Show()
            rps_window.Topmost = True
    if ref_losa and elemento_losa:
        cara_superior_losa, curveloops_superficie_losa, curvas_superficie_losa = (
            _obtener_cara_superior_y_curvas(elemento_losa)
        )
        shaft_openings_proyecto = _todos_shaft_openings(doc)
        shaft_openings_que_intersectan = shaft_openings_que_intersectan_por_bbox_solid(
            doc, elemento_losa, None
        )
        print("-" * 50)
        print("Losa: ElementId = {}".format(elemento_losa.Id.IntegerValue))
        print("  Cara superior: {}".format("OK" if cara_superior_losa else "No encontrada"))
        print("  CurveLoops en superficie: {}".format(len(curveloops_superficie_losa)))
        print("  Curvas en superficie: {}".format(len(curvas_superficie_losa)))
        print("  Shaft openings en proyecto: {}".format(len(shaft_openings_proyecto)))
        print("  Shaft openings que intersectan la losa (bbox como solido): {}".format(
            len(shaft_openings_que_intersectan)
        ))
        print("-" * 50)
        print("Variables: ref_losa, elemento_losa, cara_superior_losa,")
        print("           curveloops_superficie_losa, curvas_superficie_losa,")
        print("           shaft_openings_proyecto, shaft_openings_que_intersectan")
    else:
        print("Se cancelo la seleccion.")
else:
    print("Ejecuta este script desde Revit Python Shell (RPS).")
