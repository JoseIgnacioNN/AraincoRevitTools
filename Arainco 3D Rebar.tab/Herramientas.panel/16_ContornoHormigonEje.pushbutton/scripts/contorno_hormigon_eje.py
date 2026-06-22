# -*- coding: utf-8 -*-
"""
Contorno de hormigón en vista — unión booleana, corte por eje (Grid) y detail lines.

Revit 2024+ | pyRevit | IronPython 3.4

Paquete portable autocontenido (``16_ContornoHormigonEje.pushbutton/scripts/``).
Respaldo de desarrollo en ``BIMTools.extension/scripts/`` — sincronice tras editar.

Flujo:
  1. Elementos con Material for Model Behavior = Concrete visibles en la vista activa.
  2. Unión booleana de todos sus sólidos en uno solo.
  3. Plano de corte a partir del eje (Grid) elegido.
  4. ``CutWithHalfSpace`` → cara de corte → ``GetEdgesAsCurveLoops``.
  5. DetailCurves en la vista activa, agrupadas como ``CONTORNO`` + nombre del eje.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Arc,
    BooleanOperationsType,
    BooleanOperationsUtils,
    BuiltInCategory,
    DatumExtentType,
    ElementId,
    FilteredElementCollector,
    Grid,
    Line,
    Plane,
    PlanarFace,
    Transaction,
    UV,
    ViewSchedule,
    ViewSheet,
    ViewType,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List

from contorno_material_concrete import (
    _CATS_ESCANEO_MATERIAL_ESTRUCTURAL,
    material_estructural_es_concrete,
    obtener_solidos_elemento,
)

_DIALOG_TITLE = u"Arainco: Contorno hormigón por eje"
_TOL_VOLUMEN = 1e-12
_TOL_DIST_PLANO_FT = 0.02
_TOL_DOT_PARALELO = 0.08
_MIN_LINE_LEN_FT = 1.0 / 304.8


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _vector_unitario(v):
    if v is None:
        return None
    try:
        ln = v.GetLength()
        if ln < 1e-12:
            return None
        return XYZ(v.X / ln, v.Y / ln, v.Z / ln)
    except Exception:
        return None


def _punto_medio_curva(curve):
    if curve is None:
        return None
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return XYZ(
                (p0.X + p1.X) * 0.5,
                (p0.Y + p1.Y) * 0.5,
                (p0.Z + p1.Z) * 0.5,
            )
        except Exception:
            return None


def _distancia_punto_a_plano(p, plane):
    if p is None or plane is None:
        return 1e9
    try:
        fn = getattr(plane, "SignedDistanceTo", None)
        if fn is not None:
            return abs(float(fn(p)))
    except Exception:
        pass
    try:
        n = plane.Normal
        o = plane.Origin
        return abs(float(p.Subtract(o).DotProduct(n)))
    except Exception:
        return 1e9


def _proyectar_punto_al_plano(p, plane):
    if p is None or plane is None:
        return None
    try:
        n = _vector_unitario(plane.Normal)
        if n is None:
            return None
        v = p.Subtract(plane.Origin)
        dist = v.DotProduct(n)
        return p.Subtract(n.Multiply(dist))
    except Exception:
        return None


def vista_permitida(view):
    if view is None:
        return False, u"No hay vista activa."
    try:
        vt = view.ViewType
    except Exception:
        return False, u"No se pudo leer el tipo de vista."
    if isinstance(view, (ViewSheet, ViewSchedule)):
        return False, u"Usa una vista de modelo (planta, sección, alzado o detalle)."
    if vt == ViewType.ThreeD:
        return False, u"Usa una vista 2D (planta, sección, alzado o detalle)."
    if vt in (ViewType.DrawingSheet, ViewType.Legend, ViewType.Schedule):
        return False, u"Este tipo de vista no es compatible."
    return True, None


def listar_ejes_modelo(document):
    """Todos los ``Grid`` del documento, ordenados por nombre."""
    ejes = []
    try:
        for g in FilteredElementCollector(document).OfClass(Grid):
            if g is None or not isinstance(g, Grid):
                continue
            try:
                nombre = _as_unicode(g.Name).strip()
            except Exception:
                nombre = u""
            if not nombre:
                try:
                    nombre = u"Id {0}".format(g.Id.IntegerValue)
                except Exception:
                    nombre = u"(sin nombre)"
            ejes.append((nombre, g))
    except Exception:
        pass
    try:
        ejes.sort(key=lambda t: t[0].lower())
    except Exception:
        pass
    return ejes


def recoger_hormigon_en_vista(document, view):
    """Instancias de hormigón (Concrete) visibles en la vista activa."""
    out = []
    for cat in _CATS_ESCANEO_MATERIAL_ESTRUCTURAL:
        try:
            for el in (
                FilteredElementCollector(document, view.Id)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
            ):
                if material_estructural_es_concrete(el):
                    out.append(el)
        except Exception:
            pass
    return out


def seleccionar_elementos(uidoc, elementos):
    if uidoc is None:
        return
    ids = List[ElementId]()

    for el in elementos or []:
        if el is None:
            continue
        try:
            ids.Add(el.Id)
        except Exception:
            pass
    try:
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass


def unir_solidos_hormigon(elementos):
    """Une todos los sólidos de los elementos en uno solo."""
    solidos = []
    for el in elementos or []:
        for s in obtener_solidos_elemento(el):
            try:
                if s is not None and s.Volume > _TOL_VOLUMEN:
                    solidos.append(s)
            except Exception:
                continue
    if not solidos:
        return None
    if len(solidos) == 1:
        return solidos[0]
    try:
        solidos.sort(key=lambda x: -float(x.Volume))
    except Exception:
        pass
    acc = solidos[0]
    for i in range(1, len(solidos)):
        merged = None
        try:
            merged = BooleanOperationsUtils.ExecuteBooleanOperation(
                acc, solidos[i], BooleanOperationsType.Union
            )
        except Exception:
            try:
                merged = BooleanOperationsUtils.ExecuteBooleanOperation(
                    solidos[i], acc, BooleanOperationsType.Union
                )
            except Exception:
                return None
        if merged is None:
            return None
        acc = merged
    return acc


def _curva_mas_larga_grid(grid, view):
    candidatas = []
    try:
        for ext in (DatumExtentType.Model, DatumExtentType.ViewSpecific):
            try:
                crvs = grid.GetCurvesInView(ext, view)
            except Exception:
                crvs = None
            if crvs is None:
                continue
            try:
                n = int(crvs.Count)
            except Exception:
                n = 0
            for i in range(n):
                try:
                    c = crvs[i]
                    if c is not None and c.IsBound:
                        candidatas.append(c)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        c0 = grid.Curve
        if c0 is not None and c0.IsBound:
            candidatas.append(c0)
    except Exception:
        pass
    if not candidatas:
        return None
    try:
        return max(candidatas, key=lambda c: float(c.Length))
    except Exception:
        return candidatas[0]


def plano_corte_desde_eje(grid, view):
    """
    Plano vertical que contiene el eje (Grid) y es ⟂ a su trazo en planta.

    Returns:
        (Plane, nombre_eje) o (None, mensaje_error).
    """
    if grid is None:
        return None, u"No se indicó un eje."
    curve = _curva_mas_larga_grid(grid, view)
    if curve is None:
        return None, u"No se pudo obtener la curva del eje seleccionado."
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return None, u"La curva del eje no tiene extremos válidos."
    axis_dir = _vector_unitario(p1.Subtract(p0))
    origin = _punto_medio_curva(curve)
    if axis_dir is None or origin is None:
        return None, u"No se pudo definir dirección u origen del eje."

    horiz = XYZ(axis_dir.X, axis_dir.Y, 0.0)
    normal = None
    if horiz.GetLength() > 1e-6:
        horiz = _vector_unitario(horiz)
        if horiz is not None:
            normal = _vector_unitario(
                XYZ(-horiz.Y, horiz.X, 0.0)
            )
    if normal is None:
        try:
            rd = view.RightDirection
            normal = _vector_unitario(XYZ(rd.X, rd.Y, 0.0))
        except Exception:
            normal = None
    if normal is None:
        normal = XYZ(1.0, 0.0, 0.0)

    try:
        plane = Plane.CreateByNormalAndOrigin(normal, origin)
    except Exception:
        return None, u"No se pudo crear el plano de corte."
    try:
        nombre = _as_unicode(grid.Name).strip()
    except Exception:
        nombre = u""
    if not nombre:
        try:
            nombre = u"Id {0}".format(grid.Id.IntegerValue)
        except Exception:
            nombre = u"Eje"
    return plane, nombre


def _planar_face_mas_grande_sobre_plano(solid_cut, plane_ref, tol_dist):
    if solid_cut is None or plane_ref is None:
        return None
    pn = _vector_unitario(plane_ref.Normal)
    if pn is None:
        return None
    best = None
    best_a = -1.0
    try:
        for face in solid_cut.Faces:
            try:
                if not isinstance(face, PlanarFace) and type(face).__name__ != "PlanarFace":
                    continue
            except Exception:
                continue
            fn = _vector_unitario(face.FaceNormal)
            if fn is None:
                continue
            if abs(abs(float(fn.DotProduct(pn))) - 1.0) > _TOL_DOT_PARALELO:
                continue
            pt = None
            try:
                pt = face.Origin
            except Exception:
                pass
            if pt is None:
                try:
                    bbuv = face.GetBoundingBox()
                    if bbuv is not None:
                        u = (bbuv.Min.U + bbuv.Max.U) * 0.5
                        v = (bbuv.Min.V + bbuv.Max.V) * 0.5
                        pt = face.Evaluate(UV(u, v))
                except Exception:
                    pt = None
            if pt is None or _distancia_punto_a_plano(pt, plane_ref) > tol_dist:
                continue
            try:
                a = float(face.Area)
            except Exception:
                a = 0.0
            if a > best_a:
                best_a = a
                best = face
    except Exception:
        return None
    return best


def _buscar_cara_corte(solid_merged, plane_ref, origin):
    if solid_merged is None or plane_ref is None:
        return None
    pn = _vector_unitario(plane_ref.Normal)
    if pn is None:
        return None
    tols = [_TOL_DIST_PLANO_FT, 0.05, 0.12, 0.25]
    for td in tols:
        for flip in (False, True):
            nn = XYZ(-pn.X, -pn.Y, -pn.Z) if flip else pn
            try:
                cut_plane = Plane.CreateByNormalAndOrigin(nn, origin)
            except Exception:
                continue
            try:
                s_cut = BooleanOperationsUtils.CutWithHalfSpace(solid_merged, cut_plane)
            except Exception:
                s_cut = None
            if s_cut is None:
                continue
            try:
                if float(s_cut.Volume) <= _TOL_VOLUMEN:
                    continue
            except Exception:
                pass
            pf = _planar_face_mas_grande_sobre_plano(s_cut, cut_plane, td)
            if pf is not None:
                return pf
    return None


def _area_curve_loop(cl):
    try:
        return abs(float(cl.GetArea()))
    except Exception:
        return 0.0


def curveloops_perimetro(cara_planar):
    if cara_planar is None:
        return []
    try:
        loops_raw = cara_planar.GetEdgesAsCurveLoops()
    except Exception:
        return []
    if loops_raw is None:
        return []
    loops = []
    try:
        for i in range(loops_raw.Count):
            loops.append(loops_raw[i])
    except Exception:
        try:
            loops = list(loops_raw)
        except Exception:
            loops = []
    if not loops:
        return []
    loops.sort(key=_area_curve_loop, reverse=True)
    return loops


def _plano_vista(view):
    try:
        return Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
    except Exception:
        return None


def _proyectar_curva_a_plano(curve, plane):
    if curve is None or plane is None or not curve.IsBound:
        return None
    try:
        if isinstance(curve, Line):
            q0 = _proyectar_punto_al_plano(curve.GetEndPoint(0), plane)
            q1 = _proyectar_punto_al_plano(curve.GetEndPoint(1), plane)
            if q0 is None or q1 is None or q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
                return None
            return Line.CreateBound(q0, q1)
        if isinstance(curve, Arc):
            c = curve.Center
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            qc = _proyectar_punto_al_plano(c, plane)
            qp0 = _proyectar_punto_al_plano(p0, plane)
            qp1 = _proyectar_punto_al_plano(p1, plane)
            if qc is None or qp0 is None or qp1 is None:
                return None
            if qp0.DistanceTo(qp1) < _MIN_LINE_LEN_FT:
                return None
            return Arc.Create(qp0, qp1, qc)
    except Exception:
        pass
    try:
        q0 = _proyectar_punto_al_plano(curve.GetEndPoint(0), plane)
        q1 = _proyectar_punto_al_plano(curve.GetEndPoint(1), plane)
        if q0 is None or q1 is None or q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(q0, q1)
    except Exception:
        return None


def _nombre_grupo_unico(document, base):
    nombre = _as_unicode(base).strip()
    if not nombre:
        nombre = u"CONTORNO"
    existentes = set()
    try:
        from Autodesk.Revit.DB import GroupType

        for gt in FilteredElementCollector(document).OfClass(GroupType):
            try:
                existentes.add(_as_unicode(gt.Name))
            except Exception:
                pass
    except Exception:
        pass
    if nombre not in existentes:
        return nombre
    for i in range(2, 1000):
        candidato = u"{0} ({1})".format(nombre, i)
        if candidato not in existentes:
            return candidato
    return nombre


def crear_detail_lines_y_grupo(document, view, curve_loops, nombre_eje):
    """
    Crea DetailCurves en ``view`` y las agrupa.

    Returns:
        dict con claves ``detail_count``, ``group_name``, ``loops``.
    """
    plane_vista = _plano_vista(view)
    if plane_vista is None:
        raise ValueError(u"No se pudo obtener el plano de la vista activa.")

    detail_ids = []
    creadas = 0
    for cl in curve_loops or []:
        if cl is None:
            continue
        for c in cl:
            if c is None or not c.IsBound:
                continue
            curva = _proyectar_curva_a_plano(c, plane_vista)
            if curva is None:
                curva = c
            try:
                dc = document.Create.NewDetailCurve(view, curva)
            except Exception:
                curva2 = _proyectar_curva_a_plano(c, plane_vista)
                if curva2 is None:
                    continue
                dc = document.Create.NewDetailCurve(view, curva2)
            if dc is not None:
                detail_ids.append(dc.Id)
                creadas += 1

    group_name = _nombre_grupo_unico(
        document, u"CONTORNO" + _as_unicode(nombre_eje)
    )
    if detail_ids:
        ids = List[ElementId]()

        for eid in detail_ids:
            ids.Add(eid)
        grp = document.Create.NewGroup(ids)
        gt = document.GetElement(grp.GroupType.Id)
        gt.Name = group_name

    return {
        u"detail_count": creadas,
        u"group_name": group_name if detail_ids else None,
        u"loops": len(curve_loops or []),
    }


def ejecutar_contorno(uidoc, grid):
    """
    Flujo completo desde ``uidoc`` y el ``Grid`` elegido.

    Returns:
        (True, mensaje) o (False, mensaje_error).
    """
    if uidoc is None:
        return False, u"No hay documento activo."
    doc = uidoc.Document
    view = uidoc.ActiveView
    ok, msg = vista_permitida(view)
    if not ok:
        return False, msg

    elementos = recoger_hormigon_en_vista(doc, view)
    if not elementos:
        return False, (
            u"No hay elementos con Material for Model Behavior = Concrete "
            u"visibles en la vista activa."
        )

    seleccionar_elementos(uidoc, elementos)

    solido = unir_solidos_hormigon(elementos)
    if solido is None:
        return False, u"No se pudo unir la geometría sólida del hormigón."

    plane, nombre_eje = plano_corte_desde_eje(grid, view)
    if plane is None:
        return False, _as_unicode(nombre_eje)

    origin = plane.Origin
    cara = _buscar_cara_corte(solido, plane, origin)
    if cara is None:
        return False, (
            u"El plano del eje «{0}» no produce una sección válida sobre el sólido unificado.".format(
                nombre_eje
            )
        )

    loops = curveloops_perimetro(cara)
    if not loops:
        return False, u"No se obtuvieron curvas de perímetro en la sección."

    tx_name = u"Arainco: Contorno hormigón por eje"
    t = Transaction(doc, tx_name)
    t.Start()
    try:
        resultado = crear_detail_lines_y_grupo(doc, view, loops, nombre_eje)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        return False, _as_unicode(ex)

    return True, (
        u"Seleccionados {0} elemento(s) de hormigón.\n"
        u"Detail lines: {1} · bucles: {2}\n"
        u"Grupo: {3}".format(
            len(elementos),
            resultado.get(u"detail_count", 0),
            resultado.get(u"loops", 0),
            resultado.get(u"group_name") or u"(sin grupo)",
        )
    )


def run(revit):
    """Entrada pyRevit: abre la UI de selección de eje."""
    from contorno_hormigon_eje_ui import show_contorno_window

    show_contorno_window(revit)
