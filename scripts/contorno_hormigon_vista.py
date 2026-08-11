# -*- coding: utf-8 -*-
"""
Contorno de hormigón en vista — unión booleana, corte por plano de la vista activa.

Revit 2024+ | pyRevit | IronPython 3.4

Respaldo de desarrollo en ``BIMTools.extension/scripts/``.
Tras editar aquí, sincronice con ``20_ContornoHormigonVista.pushbutton/scripts/``
(ver ``ESTRUCTURA_PORTABLE.txt`` en el pushbutton).

Flujo:
  1. Elementos con Material for Model Behavior = Concrete visibles en la vista activa.
  2. Unión booleana de todos sus sólidos en uno solo.
  3. Plano de corte = plano de la vista activa (ViewDirection + Origin).
  4. ``CutWithHalfSpace`` → cara de corte → ``GetEdgesAsCurveLoops``.
  5. DetailCurves en la vista activa, agrupadas como ``CONTORNO`` + nombre de vista.
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Arc,
    BooleanOperationsType,
    BooleanOperationsUtils,
    ElementId,
    FilteredElementCollector,
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
    crear_options_geometria,
    material_estructural_es_concrete,
    obtener_solidos_elemento,
)

_DIALOG_TITLE = u"Arainco: Contorno hormigón por vista"
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


def _filtrar_solidos_utiles(solidos):
    out = []
    for s in solidos or []:
        if s is None:
            continue
        try:
            if float(s.Volume) > _TOL_VOLUMEN:
                out.append(s)
        except Exception:
            continue
    return out


def _solido_mayor_volumen(solidos):
    good = _filtrar_solidos_utiles(solidos)
    if not good:
        return None
    if len(good) == 1:
        return good[0]
    best = good[0]
    best_v = -1.0
    for s in good:
        try:
            v = float(s.Volume)
        except Exception:
            v = 0.0
        if v > best_v:
            best_v = v
            best = s
    return best


def _ordenar_solidos_por_volumen(solidos):
    try:
        return sorted(solidos, key=lambda x: -float(x.Volume))
    except Exception:
        return list(solidos)


def _unir_dos_solidos(a, b):
    if a is None:
        return b
    if b is None:
        return a
    try:
        merged = BooleanOperationsUtils.ExecuteBooleanOperation(
            a, b, BooleanOperationsType.Union
        )
        if merged is not None:
            return merged
    except Exception:
        pass
    try:
        return BooleanOperationsUtils.ExecuteBooleanOperation(
            b, a, BooleanOperationsType.Union
        )
    except Exception:
        return None


def _unir_solidos_lista(solidos):
    """Unión booleana estricta; falla si algún paso no puede unirse."""
    good = _ordenar_solidos_por_volumen(_filtrar_solidos_utiles(solidos))
    if not good:
        return None
    if len(good) == 1:
        return good[0]
    acc = good[0]
    for i in range(1, len(good)):
        acc = _unir_dos_solidos(acc, good[i])
        if acc is None:
            return None
        try:
            if float(acc.Volume) <= _TOL_VOLUMEN:
                return None
        except Exception:
            pass
    return acc


def _unir_solidos_greedy(solidos):
    """Unión parcial: omite piezas que no se puedan unir al acumulado."""
    good = _ordenar_solidos_por_volumen(_filtrar_solidos_utiles(solidos))
    if not good:
        return None
    if len(good) == 1:
        return good[0]
    acc = good[0]
    for i in range(1, len(good)):
        merged = _unir_dos_solidos(acc, good[i])
        if merged is not None:
            try:
                if float(merged.Volume) > _TOL_VOLUMEN:
                    acc = merged
            except Exception:
                acc = merged
    return acc


def _solidos_representantes_por_elemento(elementos, view=None):
    """Un sólido principal (mayor volumen) por elemento."""
    options = crear_options_geometria(view)
    reps = []
    for el in elementos or []:
        if el is None:
            continue
        sols = _filtrar_solidos_utiles(obtener_solidos_elemento(el, options))
        rep = _solido_mayor_volumen(sols)
        if rep is not None:
            reps.append(rep)
    return reps


def _solidos_todos_elementos(elementos, view=None):
    options = crear_options_geometria(view)
    solidos = []
    for el in elementos or []:
        if el is None:
            continue
        solidos.extend(_filtrar_solidos_utiles(obtener_solidos_elemento(el, options)))
    return solidos


def unir_solidos_hormigon(elementos, view=None):
    """
    Une la geometría del hormigón en un solo sólido.

    Estrategia (de más estable a más permisiva):
      1. Un cuerpo representativo por elemento.
      2. Todos los sólidos de instancia.
      3. Unión greedy omitiendo piezas conflictivas.
    """
    reps = _solidos_representantes_por_elemento(elementos, view)
    if not reps:
        return None

    merged = _unir_solidos_lista(reps)
    if merged is not None:
        return merged

    all_solids = _solidos_todos_elementos(elementos, view)
    if all_solids:
        merged = _unir_solidos_lista(all_solids)
        if merged is not None:
            return merged

    merged = _unir_solidos_greedy(reps)
    if merged is not None:
        return merged

    if all_solids:
        merged = _unir_solidos_greedy(all_solids)
        if merged is not None:
            return merged

    if len(reps) == 1:
        return reps[0]
    return None


def _plano_vista(view):
    try:
        return Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
    except Exception:
        return None


def plano_corte_desde_vista(view):
    """
    Plano de corte = plano de la vista activa.

    Returns:
        (Plane, nombre_vista) o (None, mensaje_error).
    """
    if view is None:
        return None, u"No hay vista activa."
    plane = _plano_vista(view)
    if plane is None:
        return None, u"No se pudo obtener el plano de la vista activa."
    try:
        nombre = _as_unicode(view.Name).strip()
    except Exception:
        nombre = u""
    if not nombre:
        try:
            nombre = u"Id {0}".format(view.Id.IntegerValue)
        except Exception:
            nombre = u"Vista"
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


def crear_detail_lines_y_grupo(document, view, curve_loops, nombre_vista):
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
        document, u"CONTORNO" + _as_unicode(nombre_vista)
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


def ejecutar_contorno(uidoc):
    """
    Flujo completo desde ``uidoc`` usando el plano de la vista activa.

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

    solido = unir_solidos_hormigon(elementos, view)
    if solido is None:
        return False, (
            u"No se pudo unir la geometría sólida del hormigón "
            u"({0} elemento(s) visibles). Comprueba que la geometría "
            u"sea sólida y que los elementos se solapen o estén unidos.".format(
                len(elementos)
            )
        )

    plane, nombre_vista = plano_corte_desde_vista(view)
    if plane is None:
        return False, _as_unicode(nombre_vista)

    origin = plane.Origin
    cara = _buscar_cara_corte(solido, plane, origin)
    if cara is None:
        return False, (
            u"El plano de la vista «{0}» no produce una sección válida "
            u"sobre el sólido unificado.".format(nombre_vista)
        )

    loops = curveloops_perimetro(cara)
    if not loops:
        return False, u"No se obtuvieron curvas de perímetro en la sección."

    tx_name = u"Arainco: Contorno hormigón por vista"
    t = Transaction(doc, tx_name)
    t.Start()
    try:
        resultado = crear_detail_lines_y_grupo(doc, view, loops, nombre_vista)
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
    """Entrada pyRevit: abre la UI de generación por vista."""
    from contorno_hormigon_vista_ui import show_contorno_window

    show_contorno_window(revit)
