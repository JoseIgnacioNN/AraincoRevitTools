# -*- coding: utf-8 -*-
"""
Redibujar contorno — unión booleana, corte por ``Armadura_Eje`` y detail lines.

Revit 2024+ | pyRevit | IronPython 3.4

Respaldo de desarrollo en ``BIMTools.extension/scripts/``.
Tras editar aquí, sincronice con ``02_RedibujarContorno.pushbutton`` (botón ligero).

Flujo:
  1. Elementos con Material for Model Behavior = Concrete visibles en la vista activa.
  2. Unión booleana de todos sus sólidos en uno solo.
  3. Plano de corte = plano vertical del eje (Grid) indicado en ``Armadura_Eje`` de la vista.
  4. ``CutWithHalfSpace`` → cara de corte → ``GetEdgesAsCurveLoops``.
  5. DetailCurves Medium Lines en la vista activa, agrupadas con el nombre de la vista
     (mismo criterio que Elevación Eje). Si ya existe ese grupo en la vista, se sustituye.
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
    Category,
    DatumExtentType,
    ElementId,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    GraphicsStyle,
    GraphicsStyleType,
    Grid,
    Group,
    GroupType,
    IFailuresPreprocessor,
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

_DIALOG_TITLE = u"Arainco: Redibujar contorno"
_PARAM_ARMADURA_EJE = u"Armadura_Eje"
_TOL_VOLUMEN = 1e-12
_TOL_DIST_PLANO_FT = 0.02
_TOL_DOT_PARALELO = 0.08
_MIN_LINE_LEN_FT = 1.0 / 304.8
_MEDIUM_LINES_NAMES = (
    u"Medium Lines",
    u"<Medium Lines>",
    u"Líneas medias",
    u"<Líneas medias>",
)
_FAILURE_SWALLOWER_SINGLETON = None


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


class _RedibujarContornoFailuresPreprocessor(IFailuresPreprocessor):
    """
    Silencia warnings al sustituir detail groups (p. ej. «group has been changed
    outside group edit mode»). La transacción solo crea detail lines y grupos.
    """

    def _iter_failure_msgs(self, failures_accessor):
        if failures_accessor is None:
            return
        try:
            fmsgs = failures_accessor.GetFailureMessages()
        except Exception:
            return
        if fmsgs is None:
            return
        try:
            n = int(fmsgs.Count)
        except Exception:
            n = 0
        for i in range(n):
            f = None
            try:
                f = fmsgs.get_Item(i)
            except Exception:
                try:
                    f = fmsgs[i]
                except Exception:
                    f = None
            if f is not None:
                yield f

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        msgs = list(self._iter_failure_msgs(failures_accessor))
        resolved_or_cleared = False
        for f in msgs:
            try:
                sev = f.GetSeverity()
            except Exception:
                continue
            if sev != FailureSeverity.Warning:
                continue
            try:
                failures_accessor.DeleteWarning(f)
                resolved_or_cleared = True
            except Exception:
                pass
        if resolved_or_cleared:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


def _attach_failure_swallower(txn):
    global _FAILURE_SWALLOWER_SINGLETON
    if txn is None or not isinstance(txn, Transaction):
        return False
    try:
        if _FAILURE_SWALLOWER_SINGLETON is None:
            _FAILURE_SWALLOWER_SINGLETON = _RedibujarContornoFailuresPreprocessor()
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_FAILURE_SWALLOWER_SINGLETON)
        try:
            opts.SetClearAfterRollback(True)
        except Exception:
            pass
        txn.SetFailureHandlingOptions(opts)
        return True
    except Exception:
        return False


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


def _clave_nombre_eje(nombre):
    return _as_unicode(nombre).strip().lower()


def leer_armadura_eje_desde_vista(view):
    """Lee ``Armadura_Eje`` de la vista activa (o la vista indicada)."""
    if view is None:
        return None
    p = None
    try:
        p = view.LookupParameter(_PARAM_ARMADURA_EJE)
    except Exception:
        p = None
    if p is None or not p.HasValue:
        return None
    try:
        val = p.AsString()
    except Exception:
        val = None
    if not val:
        try:
            val = p.AsValueString()
        except Exception:
            val = None
    if not val:
        return None
    try:
        text = unicode(val).strip()
    except Exception:
        try:
            text = str(val or u"").strip()
        except Exception:
            return None
    return text or None


def vista_tiene_parametro_armadura_eje(view):
    """True si la vista declara el parámetro de instancia ``Armadura_Eje``."""
    if view is None:
        return False
    try:
        return view.LookupParameter(_PARAM_ARMADURA_EJE) is not None
    except Exception:
        return False


def buscar_grid_por_nombre(document, nombre_eje):
    """
    Localiza un ``Grid`` del modelo cuyo nombre coincide con ``nombre_eje``.

    Returns:
        (Grid, None) o (None, mensaje_error).
    """
    clave = _clave_nombre_eje(nombre_eje)
    if not clave:
        return None, u"Armadura_Eje está vacío en la vista activa."
    candidatos = []
    try:
        for g in FilteredElementCollector(document).OfClass(Grid):
            if g is None or not isinstance(g, Grid):
                continue
            try:
                nombre = _as_unicode(g.Name).strip()
            except Exception:
                nombre = u""
            if _clave_nombre_eje(nombre) == clave:
                candidatos.append(g)
    except Exception:
        pass
    if not candidatos:
        return None, (
            u"No se encontró el eje «{0}» (Grid) en el modelo.".format(
                _as_unicode(nombre_eje).strip()
            )
        )
    if len(candidatos) == 1:
        return candidatos[0], None
    return candidatos[0], None


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


def limpiar_seleccion(uidoc):
    """Quita cualquier selección en el documento activo."""
    if uidoc is None:
        return
    empty = List[ElementId]()
    try:
        uidoc.Selection.SetElementIds(empty)
    except Exception:
        pass
    try:
        uidoc.RefreshActiveView()
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

    Extrae geometría una sola vez (reps + all solids del mismo pase).
    """
    options = crear_options_geometria(view)
    reps = []
    all_solids = []
    for el in elementos or []:
        if el is None:
            continue
        sols = _filtrar_solidos_utiles(obtener_solidos_elemento(el, options))
        if not sols:
            continue
        all_solids.extend(sols)
        rep = _solido_mayor_volumen(sols)
        if rep is not None:
            reps.append(rep)
    if not reps:
        return None

    merged = _unir_solidos_lista(reps)
    if merged is not None:
        return merged

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
        return candidatos[0]


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


def plano_corte_desde_armadura_eje(document, view):
    """
    Plano de corte a partir de ``Armadura_Eje`` en la vista activa.

    Returns:
        (Plane, nombre_eje) o (None, mensaje_error).
    """
    if view is None:
        return None, u"No hay vista activa."
    if not vista_tiene_parametro_armadura_eje(view):
        return None, u"La vista activa no tiene el parámetro Armadura_Eje."
    nombre_eje = leer_armadura_eje_desde_vista(view)
    if not nombre_eje:
        return None, u"Armadura_Eje está vacío en la vista activa."
    grid, err = buscar_grid_por_nombre(document, nombre_eje)
    if grid is None:
        return None, err or u"No se encontró el eje en el modelo."
    return plano_corte_desde_eje(grid, view)


def _plano_vista(view):
    try:
        return Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
    except Exception:
        return None


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


def _buscar_cara_corte(solid_merged, plane_ref, origin, preferred_normal=None):
    if solid_merged is None or plane_ref is None:
        return None
    pn = _vector_unitario(plane_ref.Normal)
    if pn is None:
        return None
    normals = []
    pref = _vector_unitario(preferred_normal) if preferred_normal is not None else None
    if pref is not None:
        normals.append(pref)
        normals.append(XYZ(-pref.X, -pref.Y, -pref.Z))
    normals.append(pn)
    normals.append(XYZ(-pn.X, -pn.Y, -pn.Z))
    ordered = []
    for nn in normals:
        if nn is None:
            continue
        dup = False
        for prev in ordered:
            try:
                if abs(abs(float(nn.DotProduct(prev))) - 1.0) < 1e-6:
                    dup = True
                    break
            except Exception:
                pass
        if not dup:
            ordered.append(nn)
    tols = [_TOL_DIST_PLANO_FT, 0.05, 0.12, 0.25]
    for td in tols:
        for nn in ordered:
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


def _element_id_igual(a, b):
    if a is None or b is None:
        return False
    try:
        if a == ElementId.InvalidElementId or b == ElementId.InvalidElementId:
            return False
    except Exception:
        pass
    try:
        return int(a.IntegerValue) == int(b.IntegerValue)
    except Exception:
        try:
            return a == b
        except Exception:
            return False


def _nombre_vista_grupo(view):
    if view is None:
        return u""
    try:
        return _as_unicode(view.Name).strip()
    except Exception:
        return u""


def _nombre_base_grupo_contorno(view, nombre_eje=None):
    """
    Nombre base del detail group — patrón Elevación Eje.

    Usa ``view.Name`` (p. ej. ``02_MA_ELEVACION EJE 1``). Si la vista no tiene
    nombre, respaldo ``CONTORNO {eje}``.
    """
    base = _nombre_vista_grupo(view)
    if base:
        return base
    eje = _as_unicode(nombre_eje).strip() if nombre_eje else u""
    if eje:
        return u"CONTORNO {0}".format(eje)
    return u"CONTORNO"


def _clave_nombre_grupo(text):
    return _as_unicode(text).strip().lower()


def _bases_nombre_contorno(view, nombre_eje=None):
    """
    Nombres base posibles del detail group de contorno en la vista.

    Incluye ``view.Name`` (Elevación Eje) y variantes legacy ``CONTORNO`` + eje.
    """
    bases = []
    vistos = set()

    def _agregar(nombre):
        nombre = _as_unicode(nombre).strip()
        if not nombre:
            return
        clave = _clave_nombre_grupo(nombre)
        if clave in vistos:
            return
        vistos.add(clave)
        bases.append(nombre)

    _agregar(_nombre_vista_grupo(view))
    eje = _as_unicode(nombre_eje).strip() if nombre_eje else u""
    if eje:
        _agregar(u"CONTORNO" + eje)
        _agregar(u"CONTORNO {0}".format(eje))
    return bases


def _nombre_grupo_variante_coincide(nombre_grupo, base):
    """True si ``nombre_grupo`` es ``base`` o ``base (2)``, ``base (3)``, …"""
    nombre = _clave_nombre_grupo(nombre_grupo)
    base = _clave_nombre_grupo(base)
    if not base:
        return False
    if nombre == base:
        return True
    pref = base + u" ("
    if nombre.startswith(pref) and nombre.endswith(u")"):
        try:
            return int(nombre[len(pref) : -1].strip()) >= 2
        except Exception:
            return False
    return False


def _nombre_grupo_contorno_coincide(nombre_grupo, view, nombre_eje=None):
    """True si el nombre de grupo sigue el patrón de contorno para la vista."""
    for base in _bases_nombre_contorno(view, nombre_eje):
        if _nombre_grupo_variante_coincide(nombre_grupo, base):
            return True
    return False


def _grupo_es_de_vista(document, view, grupo):
    if document is None or view is None or grupo is None:
        return False
    try:
        view_id = view.Id
    except Exception:
        return False
    try:
        owner_id = grupo.OwnerViewId
        if owner_id is not None and _element_id_igual(owner_id, view_id):
            return True
    except Exception:
        pass
    try:
        member_ids = grupo.GetMemberIds()
    except Exception:
        return False
    if member_ids is None:
        return False
    try:
        n = int(member_ids.Count)
    except Exception:
        try:
            n = len(member_ids)
        except Exception:
            n = 0
    if n == 0:
        return False
    for eid in member_ids:
        try:
            el = document.GetElement(eid)
            if el is None:
                continue
            owner_id = getattr(el, u"OwnerViewId", None)
            if owner_id is not None and _element_id_igual(owner_id, view_id):
                return True
        except Exception:
            continue
    return False


def _recoger_grupos_en_vista(document, view):
    """Grupos de detalle/modelo asociados a ``view`` (collector por vista + respaldo)."""
    if document is None or view is None:
        return []
    grupos = []
    try:
        view_id = view.Id
        grupos = list(
            FilteredElementCollector(document, view_id).OfClass(Group).ToElements()
        )
    except Exception:
        grupos = []
    if grupos:
        return grupos
    try:
        todos = list(FilteredElementCollector(document).OfClass(Group).ToElements())
    except Exception:
        return []
    for grp in todos:
        if _grupo_es_de_vista(document, view, grp):
            grupos.append(grp)
    return grupos


def _purgar_grouptypes_contorno_huerfanos(document, view, nombre_eje=None):
    """Elimina ``GroupType`` de contorno sin instancias (libera el nombre base)."""
    if document is None:
        return 0
    eliminados = 0
    tipos = []
    try:
        tipos = list(FilteredElementCollector(document).OfClass(GroupType).ToElements())
    except Exception:
        tipos = []
    instancias_por_tipo = {}
    try:
        for grp in FilteredElementCollector(document).OfClass(Group):
            if grp is None:
                continue
            try:
                tid = grp.GroupType.Id
            except Exception:
                continue
            try:
                key = int(tid.IntegerValue)
            except Exception:
                key = tid
            instancias_por_tipo[key] = instancias_por_tipo.get(key, 0) + 1
    except Exception:
        pass
    for gt in tipos:
        if gt is None:
            continue
        try:
            nombre = _as_unicode(gt.Name).strip()
        except Exception:
            continue
        if not _nombre_grupo_contorno_coincide(nombre, view, nombre_eje):
            continue
        try:
            key = int(gt.Id.IntegerValue)
        except Exception:
            key = gt.Id
        if instancias_por_tipo.get(key, 0) > 0:
            continue
        try:
            document.Delete(gt.Id)
            eliminados += 1
        except Exception:
            pass
    return eliminados


def eliminar_grupos_contorno_en_vista(document, view, nombre_eje=None):
    """
    Borra el detail group de contorno de la vista (nombre = ``view.Name``).

    Returns:
        Número de instancias de grupo eliminadas.
    """
    if document is None or view is None:
        return 0
    to_delete = []
    for grp in _recoger_grupos_en_vista(document, view):
        if grp is None:
            continue
        try:
            gt = grp.GroupType
            if gt is None:
                continue
            nombre = _as_unicode(gt.Name).strip()
        except Exception:
            continue
        if not _nombre_grupo_contorno_coincide(nombre, view, nombre_eje):
            continue
        to_delete.append(grp.Id)
    eliminados = 0
    for gid in to_delete:
        try:
            if document.GetElement(gid) is None:
                continue
            document.Delete(gid)
            eliminados += 1
        except Exception:
            pass
    if eliminados:
        _purgar_grouptypes_contorno_huerfanos(document, view, nombre_eje)
    return eliminados


def _norm_upper(text):
    return _as_unicode(text).strip().upper()


def _lines_category(document):
    if document is None:
        return None
    try:
        return Category.GetCategory(document, BuiltInCategory.OST_Lines)
    except Exception:
        pass
    try:
        return document.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        return None


def resolver_medium_lines_style_id(document):
    """``GraphicsStyle.Id`` del estilo de línea «Medium Lines» (subcategoría Líneas)."""
    if document is None:
        return None
    targets = set()
    for name in _MEDIUM_LINES_NAMES:
        targets.add(_norm_upper(name))
        bare = name.strip(u"<>").strip()
        if bare:
            targets.add(_norm_upper(bare))

    lines_cat = _lines_category(document)
    fallback_id = None
    if lines_cat is not None:
        try:
            for sub in lines_cat.SubCategories:
                try:
                    nm_u = _norm_upper(sub.Name)
                except Exception:
                    continue
                style_id = None
                try:
                    gs = sub.GetGraphicsStyle(GraphicsStyleType.Projection)
                    if gs is not None:
                        style_id = gs.Id
                except Exception:
                    style_id = None
                if style_id is None:
                    continue
                if nm_u in targets:
                    return style_id
                for t in targets:
                    bare = t.strip(u"<>").strip()
                    if bare and (nm_u == bare or bare in nm_u):
                        fallback_id = style_id
                        break
        except Exception:
            pass
    if fallback_id is not None:
        return fallback_id

    parent_iv = None
    if lines_cat is not None:
        try:
            parent_iv = lines_cat.Id.IntegerValue
        except Exception:
            parent_iv = None
    try:
        for gs in FilteredElementCollector(document).OfClass(GraphicsStyle):
            try:
                cg = getattr(gs, u"GraphicsStyleCategory", None)
                if cg is None:
                    cg = getattr(gs, u"Category", None)
                if cg is None:
                    continue
                if parent_iv is not None:
                    pc = getattr(cg, u"Parent", None)
                    if pc is None or pc.Id.IntegerValue != parent_iv:
                        continue
                nm_u = _norm_upper(cg.Name)
                if nm_u in targets:
                    return gs.Id
                for t in targets:
                    bare = t.strip(u"<>").strip()
                    if bare and (nm_u == bare or bare in nm_u):
                        return gs.Id
            except Exception:
                continue
    except Exception:
        pass
    return None


def _aplicar_line_style(curve_element, style_id, style_verificado=False):
    if curve_element is None or style_id is None:
        return False
    try:
        if style_id == ElementId.InvalidElementId:
            return False
    except Exception:
        pass
    if style_verificado:
        try:
            curve_element.LineStyleId = style_id
            return True
        except Exception:
            pass
    try:
        applicable = list(curve_element.GetLineStyleIds())
    except Exception:
        applicable = []
    chosen = style_id
    if applicable:
        try:
            piv = style_id.IntegerValue
            for aid in applicable:
                if aid is not None and aid.IntegerValue == piv:
                    chosen = aid
                    break
            else:
                for aid in applicable:
                    try:
                        gs = curve_element.Document.GetElement(aid)
                        nm = u""
                        if gs is not None:
                            cg = getattr(gs, u"GraphicsStyleCategory", None)
                            if cg is not None:
                                nm = _norm_upper(cg.Name)
                            if not nm:
                                nm = _norm_upper(getattr(gs, u"Name", u""))
                        for name in _MEDIUM_LINES_NAMES:
                            t = _norm_upper(name)
                            bare = t.strip(u"<>").strip()
                            if nm == t or (bare and (nm == bare or bare in nm)):
                                chosen = aid
                                break
                        else:
                            continue
                        break
                    except Exception:
                        continue
        except Exception:
            chosen = style_id
    try:
        curve_element.LineStyleId = chosen
        return True
    except Exception:
        pass
    try:
        gs = curve_element.Document.GetElement(chosen)
        if gs is not None:
            curve_element.LineStyle = gs
            return True
    except Exception:
        pass
    return False


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
    Crea DetailCurves en ``view`` (estilo Medium Lines) y las agrupa.

    Si ya hay un grupo de contorno con el nombre de la vista, lo sustituye.

    Returns:
        dict con claves ``detail_count``, ``group_name``, ``loops``,
        ``replaced_groups``.
    """
    plane_vista = _plano_vista(view)
    if plane_vista is None:
        raise ValueError(u"No se pudo obtener el plano de la vista activa.")

    replaced = eliminar_grupos_contorno_en_vista(document, view, nombre_eje)
    style_id = resolver_medium_lines_style_id(document)
    style_ok = False

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
                try:
                    dc = document.Create.NewDetailCurve(view, curva2)
                except Exception:
                    continue
            if dc is not None:
                if _aplicar_line_style(dc, style_id, style_verificado=style_ok):
                    style_ok = True
                detail_ids.append(dc.Id)
                creadas += 1

    group_name = _nombre_grupo_unico(
        document, _nombre_base_grupo_contorno(view, nombre_eje)
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
        u"replaced_groups": replaced,
    }


def ejecutar_contorno(uidoc):
    """
    Flujo completo desde ``uidoc`` usando ``Armadura_Eje`` de la vista activa.

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

    solido = unir_solidos_hormigon(elementos, view)
    if solido is None:
        return False, (
            u"No se pudo unir la geometría sólida del hormigón "
            u"({0} elemento(s) visibles). Comprueba que la geometría "
            u"sea sólida y que los elementos se solapen o estén unidos.".format(
                len(elementos)
            )
        )

    plane, nombre_eje = plano_corte_desde_armadura_eje(doc, view)
    if plane is None:
        return False, _as_unicode(nombre_eje)

    origin = plane.Origin
    preferred = None
    try:
        preferred = view.ViewDirection
    except Exception:
        preferred = None
    cara = _buscar_cara_corte(solido, plane, origin, preferred_normal=preferred)
    if cara is None:
        return False, (
            u"El plano del eje «{0}» (Armadura_Eje) no produce una sección válida "
            u"sobre el sólido unificado.".format(nombre_eje)
        )

    loops = curveloops_perimetro(cara)
    if not loops:
        return False, u"No se obtuvieron curvas de perímetro en la sección."

    tx_name = u"Arainco: Redibujar contorno"
    t = Transaction(doc, tx_name)
    _attach_failure_swallower(t)
    t.Start()
    try:
        resultado = crear_detail_lines_y_grupo(doc, view, loops, nombre_eje)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        return False, _as_unicode(ex)

    limpiar_seleccion(uidoc)

    msg = (
        u"Hormigón procesado: {0} elemento(s).\n"
        u"Detail lines: {1} · bucles: {2}\n"
        u"Grupo: {3}".format(
            len(elementos),
            resultado.get(u"detail_count", 0),
            resultado.get(u"loops", 0),
            resultado.get(u"group_name") or u"(sin grupo)",
        )
    )
    replaced = int(resultado.get(u"replaced_groups") or 0)
    if replaced:
        msg = msg + u"\nGrupo(s) de contorno anterior(es) sustituido(s): {0}.".format(
            replaced
        )
    return True, msg


def run(revit):
    """Entrada pyRevit: abre la UI de generación por vista."""
    from contorno_hormigon_vista_ui import show_contorno_window

    show_contorno_window(revit)
