# -*- coding: utf-8 -*-
"""
Geometría de sección transversal para **estribos** en Structural Framing.

Flujo (preview con ModelLine):
- ``LocationCurve`` (cuerda si el eje es curvo) → **recorte** por las **caras extremas** del
  sólido de la viga (planos con normal ~ paralela al eje / tapas); se conserva el **segmento
  central** entre esos planos → **ModelLine** del eje para revisión.
- Desde el **inicio** del eje recortado, punto a **50 mm** en dirección del eje → plano de corte ⟂ al eje;
  intersección con el sólido y perímetro con offset (model lines de estribo).
- Polígono aproximado por **envolvente convexa** de los puntos de corte (vigas cajón OK;
  perfiles en I / huecos pueden requerir otro enfoque).
- ``CurveLoop.CreateViaOffset`` ± recubrimiento (mm), se elige el bucle de **menor perímetro** (interior).
- Opcionalmente ``NewModelCurve``/eje en preview (`crear_model_lines=True`); por defecto desactivado;
  si se solicita, ``Rebar`` estribo
  con ``RebarShape`` de nombre **«10»** (``CreateFromCurvesAndShape``): reparto en tramos
  con ``MaximumSpacing``; longitud total de reparto = ``L_eje − 50 mm (inicio) − 50 mm (fin)``
  (simétrico al plano de sección a **50 mm** del inicio y al mismo respeto en el **final** del eje).
  En cada **extremo** del vano el tramo de estribos de carta *Extremos* tiene longitud
  **2 × canto (altura ) de la viga** (en paralelo al eje), recortado si hace falta para no
  solaparse; el **centro** usa la carta *Centrales* en el tramo restante. Si
  esa longitud ``< 2 × canto``, un solo conjunto **centrales** con primera y última barra del set activas.
  Opcional: ``MultiReferenceAnnotation`` por conjunto en la vista activa
  (``crear_multi_rebar_tags=True`` + ``view``), si el proyecto incluye un tipo MRA.
  Con ``view`` en colocación, presentación del conjunto **First last** en esa vista.
  Otras herramientas pueden usar :func:`crear_multi_rebar_annotations_por_nombre_tipo` con el
  nombre del tipo MRA del proyecto (p. ej. zapata de muro — «Recorrido Barras»).

No gestiona transacciones: el llamador abre ``Transaction``.
"""

from __future__ import division

import clr
import System

clr.AddReference("RevitAPI")

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    CurveLoop,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Line,
    LocationCurve,
    Plane,
    PlanarFace,
    SketchPlane,
    StorageType,
    Transform,
    UnitUtils,
    UnitTypeId,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarHookOrientation,
    RebarHookType,
    RebarPresentationMode,
    RebarShape,
    StructuralMaterialType,
    StructuralType,
)

from geometria_colision_vigas import obtener_solidos_elemento
from geometria_viga_cara_superior_detalle import (
    _crear_model_curve,
    _line_bound_desde_location_curve,
    _plano_desde_face,
    _punto_sobre_cara_planar,
)

try:
    from estribos_viga_rps import (
        _apply_maximum_spacing_layout,
        _pick_first_hook_type,
    )
except Exception:
    _apply_maximum_spacing_layout = None
    _pick_first_hook_type = None

try:
    from rebar_fundacion_cara_inferior import (
        buscar_rebar_shape_por_nombre,
        rebar_shape_display_name,
    )
except Exception:
    buscar_rebar_shape_por_nombre = None
    rebar_shape_display_name = None

_TOL_PLANE_FT = 1.5e-4  # ~0.05 mm
_TOL_MERGE_PTS_FT = 2.5e-3  # ~0.76 mm
_MIN_EDGE_FT = 1.0e-6
_AXIS_CASI_VERTICAL_TOL = 0.92
# Cara “tapa” del sólido: |normal·eje| ≥ este valor (paralelismo eje–normal).
_TAPA_VIGA_PARALELO_EJE_MIN = 0.88
# Área mínima (ft²) para considerar una cara como candidata a tapa (filtra aristas finas).
_TAPA_VIGA_AREA_MIN_FT2 = 1.0e-4
# Dedupe de parámetros s (pies) a lo largo del eje.
_DEDUPE_S_FT = 0.02
# Plano de la sección tipo estribo: a esta distancia (mm) desde el **inicio** del eje recortado.
_ESTRIBO_PLANO_DESDE_INICIO_MM = 50.0
# Recorte al **final** del eje recortado (mm): el reparto ``MaximumSpacing`` no llega al tapón final;
# debe coincidir en resultado con el respeto por el **inicio** (``_ESTRIBO_PLANO_DESDE_INICIO_MM``).
_ESTRIBO_ARRAY_LONGITUD_MENOS_MM = 50.0
# Fracción del vano de reparto (L_array) en cada zona de **extremos** si no se obtiene canto **h**.
_ESTRIBO_FRACC_EXTREMOS = 0.25
# Nombre del ``RebarShape`` del proyecto para estribos BIMTools (``CreateFromCurvesAndShape``).
_ESTRIBO_REBAR_SHAPE_NOMBRE = u"10"
# ``MultiReferenceAnnotationType`` del proyecto para etiquetas multibar de estribo (nombre en UI).
_ESTRIBO_MRA_TYPE_NAME = u"Structural Rebar_Estribo"
# Separación (mm) entre el trazo de armadura y la línea de cota MRA, hacia **−UpDirection** de la vista.
_MRA_ESTRIBO_OFFSET_DEBAJO_MM = 450.0
# Hueco (mm) adicional bajo el rectángulo de la **cota de traslape** (schema vigas) para no solapar texto MRA.
_MRA_ESTRIBO_CLEARANCE_TRASLAPE_MM = 200.0


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _altura_viga_estribos_mm(document, host):
    """
    Canto (altura de sección) en mm, misma prioridad que enfierrado vigas
    (``armadura_vigas_capas._read_width_depth_ft``, luego parámetros / caja).
    """
    if host is None:
        return None
    if document is not None:
        try:
            from armadura_vigas_capas import _read_width_depth_ft

            loc = getattr(host, "Location", None)
            if isinstance(loc, LocationCurve):
                crv = loc.Curve
                if crv is not None:
                    _w_ft, d_ft = _read_width_depth_ft(document, host, crv)
                    if d_ft is not None and float(d_ft) > 0.0:
                        d_mm = float(d_ft) * 304.8
                        if d_mm > 0.5:
                            return d_mm
        except Exception:
            pass
    try:
        p = host.get_Parameter(BuiltInParameter.STRUCTURAL_DEPTH)
        if p is not None and p.HasValue and p.StorageType == StorageType.Double:
            v = float(
                UnitUtils.ConvertFromInternalUnits(
                    p.AsDouble(), UnitTypeId.Millimeters
                )
            )
            if v > 0.5:
                return v
    except Exception:
        pass
    if document is not None:
        try:
            et = document.GetElement(host.GetTypeId())
        except Exception:
            et = None
        if et is not None:
            for nm in (
                u"Height",
                u"Depth",
                u"Altura",
                u"Profundidad",
                u"Canto",
                u"h",
                u"H",
                u"d",
            ):
                try:
                    p = et.LookupParameter(nm)
                    if p is None or not p.HasValue:
                        continue
                    if p.StorageType != StorageType.Double:
                        continue
                    v = float(
                        UnitUtils.ConvertFromInternalUnits(
                            p.AsDouble(), UnitTypeId.Millimeters
                        )
                    )
                    if v > 0.5:
                        return v
                except Exception:
                    continue
    for nm in (u"Depth", u"Structural Depth", u"Profundidad", u"Altura", u"Canto"):
        try:
            p = host.LookupParameter(nm)
            if p is None or not p.HasValue:
                continue
            if p.StorageType != StorageType.Double:
                continue
            v = float(
                UnitUtils.ConvertFromInternalUnits(
                    p.AsDouble(), UnitTypeId.Millimeters
                )
            )
            if v > 0.5:
                return v
        except Exception:
            continue
    try:
        bb = host.get_BoundingBox(None)
        if bb is None:
            return None
        dv = bb.Max - bb.Min
        dx = abs(float(dv.X))
        dy = abs(float(dv.Y))
        dz = abs(float(dv.Z))
        dims = sorted([dx, dy, dz], reverse=True)
        if len(dims) >= 3:
            sec_ft = max(dims[1], dims[2])
        else:
            sec_ft = dims[-1]
        v = float(
            UnitUtils.ConvertFromInternalUnits(sec_ft, UnitTypeId.Millimeters)
        )
        return v if v > 0.5 else None
    except Exception:
        return None


def _curva_location_framing(elemento):
    loc = getattr(elemento, "Location", None)
    if not isinstance(loc, LocationCurve):
        return None
    try:
        crv = loc.Curve
    except Exception:
        return None
    if crv is None or not crv.IsBound:
        return None
    return crv


def _iter_planar_faces_solid(solid):
    if solid is None:
        return
    try:
        faces = solid.Faces
    except Exception:
        return
    try:
        for f in faces:
            if isinstance(f, PlanarFace):
                yield f
    except Exception:
        return


def _param_ray_plane_interseccion(p0, axis_unit, plane):
    """
    Escalar ``s`` tal que ``p0 + axis_unit * s`` pertenece al plano (sin acotar a un tramo).
    """
    if p0 is None or axis_unit is None or plane is None:
        return None
    try:
        n = plane.Normal
        if n.GetLength() < _MIN_EDGE_FT:
            return None
        n = n.Normalize()
        o = plane.Origin
        denom = float(axis_unit.DotProduct(n))
    except Exception:
        return None
    try:
        if abs(denom) < 1e-12:
            return None
        s = float(o.Subtract(p0).DotProduct(n)) / denom
        return s
    except Exception:
        return None


def _dedupe_params_eje(ss, tol_ft):
    if not ss:
        return []
    xs = sorted(float(x) for x in ss)
    out = []
    for x in xs:
        if not out or abs(x - out[-1]) > tol_ft:
            out.append(x)
        else:
            out[-1] = 0.5 * (out[-1] + x)
    return out


def _linea_entre_tapas_extremas_viga(line_full, solid):
    """
    Recorta la cuerda del eje al segmento **central** comprendido entre las intersecciones
    con caras del sólido cuya normal es prácticamente paralela al eje (tapas de extremo).

    Returns:
        ``Line`` entre los cortes más exteriores válidos, o ``None``.
    """
    if line_full is None or solid is None:
        return None
    try:
        p0 = line_full.GetEndPoint(0)
        p1 = line_full.GetEndPoint(1)
        d = p1.Subtract(p0)
        L = float(d.GetLength())
        if L < _MIN_EDGE_FT:
            return None
        axis = d.Normalize()
    except Exception:
        return None
    ss = []
    for face in _iter_planar_faces_solid(solid):
        if face is None:
            continue
        try:
            if float(face.Area) < _TAPA_VIGA_AREA_MIN_FT2:
                continue
            fn = face.FaceNormal
            if fn is None or fn.GetLength() < _MIN_EDGE_FT:
                continue
            fn = fn.Normalize()
            ad = abs(float(fn.DotProduct(axis)))
            if ad < _TAPA_VIGA_PARALELO_EJE_MIN:
                continue
        except Exception:
            continue
        pl = _plano_desde_face(face)
        if pl is None:
            continue
        s = _param_ray_plane_interseccion(p0, axis, pl)
        if s is None:
            continue
        try:
            pt = p0.Add(axis.Multiply(float(s)))
        except Exception:
            continue
        if not _punto_sobre_cara_planar(pt, face):
            continue
        ss.append(float(s))
    merged = _dedupe_params_eje(ss, _DEDUPE_S_FT)
    if len(merged) < 2:
        return None
    t_lo = float(merged[0])
    t_hi = float(merged[-1])
    if t_hi <= t_lo + _MIN_EDGE_FT:
        return None
    # Dejar margen numérico respecto a la cuerda analítica
    t_lo = max(t_lo, -0.01 * L)
    t_hi = min(t_hi, 1.01 * L)
    if t_hi <= t_lo + _MIN_EDGE_FT:
        return None
    try:
        return Line.CreateBound(
            p0.Add(axis.Multiply(t_lo)),
            p0.Add(axis.Multiply(t_hi)),
        )
    except Exception:
        return None


def _base_uv_en_plano(n_plano):
    """``u``, ``v`` ortonormales en el plano de normal ``n_plano`` (``n`` unitario)."""
    try:
        n = n_plano.Normalize()
    except Exception:
        return None, None
    if n is None or n.GetLength() < _MIN_EDGE_FT:
        return None, None
    for w in (XYZ.BasisZ, XYZ.BasisX, XYZ.BasisY):
        try:
            u = w.CrossProduct(n)
            if u.GetLength() >= _MIN_EDGE_FT:
                u = u.Normalize()
                v = n.CrossProduct(u)
                if v.GetLength() >= _MIN_EDGE_FT:
                    return u, v.Normalize()
        except Exception:
            continue
    return None, None


def _punto_linea_plano_segmento(line, plane):
    """
    Intersección de un segmento con un plano.

    Returns:
        ``list`` de 0, 1 o 2 ``XYZ`` (dos si el segmento yace en el plano: extremos).
    """
    out = []
    if line is None or plane is None:
        return out
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        n = plane.Normal
        if n.GetLength() > _MIN_EDGE_FT:
            n = n.Normalize()
        o = plane.Origin
        d0 = float((p0.Subtract(o)).DotProduct(n))
        d1 = float((p1.Subtract(o)).DotProduct(n))
    except Exception:
        return out

    if abs(d0) <= _TOL_PLANE_FT and abs(d1) <= _TOL_PLANE_FT:
        if p0.DistanceTo(p1) > _MIN_EDGE_FT:
            return [p0, p1]
        return [p0]
    if abs(d0) <= _TOL_PLANE_FT:
        return [p0]
    if abs(d1) <= _TOL_PLANE_FT:
        return [p1]
    if d0 * d1 > _TOL_PLANE_FT * _TOL_PLANE_FT:
        return out
    try:
        s = d0 / (d0 - d1)
        if s < -1e-9 or s > 1.0 + 1e-9:
            return out
        s = max(0.0, min(1.0, s))
        pt = p0.Add(p1.Subtract(p0).Multiply(s))
        return [pt]
    except Exception:
        return out


def _puntos_interseccion_solido_plano(solid, plane):
    pts = []
    if solid is None or plane is None:
        return pts
    try:
        edges = solid.Edges
        n_e = int(edges.Size)
    except Exception:
        return pts
    for i in range(n_e):
        try:
            edge = edges.get_Item(i)
            crv = edge.AsCurve()
        except Exception:
            continue
        if crv is None:
            continue
        try:
            tess = crv.Tessellate()
        except Exception:
            tess = None
        tess_list = list(tess) if tess is not None else []
        if len(tess_list) < 2:
            try:
                seg = [
                    Line.CreateBound(crv.GetEndPoint(0), crv.GetEndPoint(1)),
                ]
            except Exception:
                continue
        else:
            seg = []
            tlist = tess_list
            for j in range(len(tlist) - 1):
                try:
                    seg.append(Line.CreateBound(tlist[j], tlist[j + 1]))
                except Exception:
                    continue
        for ln in seg:
            for p in _punto_linea_plano_segmento(ln, plane):
                pts.append(p)
    return pts


def _dedupe_puntos(puntos, tol_ft):
    out = []
    for p in puntos:
        try:
            if any(q.DistanceTo(p) <= tol_ft for q in out):
                continue
        except Exception:
            continue
        out.append(p)
    return out


def _convex_hull_monotone_chain(xy):
    """``xy``: lista de ``(x, y)``. Retorna índices al casco en orden CCW sin repetir primer punto."""
    n = len(xy)
    if n < 3:
        return list(range(n))
    pts = sorted(range(n), key=lambda i: (xy[i][0], xy[i][1]))

    def cross(i, j, k):
        return (xy[j][0] - xy[i][0]) * (xy[k][1] - xy[j][1]) - (
            xy[j][1] - xy[i][1]
        ) * (xy[k][0] - xy[j][0])

    lower = []
    for i in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], i) <= 0:
            lower.pop()
        lower.append(i)
    upper = []
    for i in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], i) <= 0:
            upper.pop()
        upper.append(i)
    hull = lower[:-1] + upper[:-1]
    return hull


def _curveloop_desde_puntos_ordenados(puntos_3d, u, v, origin):
    if not puntos_3d or len(puntos_3d) < 3:
        return None
    xy = []
    for p in puntos_3d:
        try:
            d = p.Subtract(origin)
            xy.append((float(d.DotProduct(u)), float(d.DotProduct(v))))
        except Exception:
            return None
    hull_idx = _convex_hull_monotone_chain(xy)
    if len(hull_idx) < 3:
        return None
    ordered = [puntos_3d[i] for i in hull_idx]
    try:
        curve_list = List[Curve]()
        n = len(ordered)
        for i in range(n):
            a = ordered[i]
            b = ordered[(i + 1) % n]
            if a.DistanceTo(b) <= _MIN_EDGE_FT:
                continue
            curve_list.Add(Line.CreateBound(a, b))
        if curve_list.Count < 3:
            return None
        return CurveLoop.Create(curve_list)
    except Exception:
        return None


def _curveloop_recubrimiento_interior(loop, cover_mm, plane_normal):
    if loop is None:
        return None
    try:
        if not loop.HasPlane():
            return None
    except Exception:
        return None
    off = _mm_to_internal(cover_mm)
    if off <= 0:
        return loop
    try:
        n = plane_normal.Normalize()
    except Exception:
        return None
    candidates = []
    for sign in (1.0, -1.0):
        try:
            cl = CurveLoop.CreateViaOffset(loop, sign * off, n)
            if cl is not None:
                candidates.append(cl)
        except Exception:
            continue
    if not candidates:
        return None

    def _perim(cl):
        try:
            return float(cl.GetExactLength())
        except Exception:
            s = 0.0
            try:
                it = cl.GetCurveLoopIterator()
                while it.MoveNext():
                    c = it.Current
                    if c is not None:
                        s += float(c.Length)
            except Exception:
                return 1e99
            return s

    return min(candidates, key=_perim)


def _crear_model_curves_desde_loop(document, loop, plane_cut):
    """Crea una ``ModelCurve`` por tramo del bucle (plano = normal corte, origen del plano)."""
    n = 0
    if document is None or loop is None or plane_cut is None:
        return n
    try:
        sp = SketchPlane.Create(document, plane_cut)
    except Exception:
        return n
    try:
        it = loop.GetCurveLoopIterator()
        while it.MoveNext():
            c = it.Current
            if c is None or not c.IsBound:
                continue
            try:
                document.Create.NewModelCurve(c, sp)
                n += 1
            except Exception:
                continue
    except Exception:
        pass
    return n


def _solido_principal(elemento):
    solids = obtener_solidos_elemento(elemento)
    if not solids:
        return None
    try:
        return max(solids, key=lambda s: float(s.Volume))
    except Exception:
        return solids[0]


def _host_viga_permite_rebar(host):
    if host is None or not isinstance(host, FamilyInstance):
        return False
    try:
        if host.Category is None:
            return False
        st = getattr(host, "StructuralType", None)
        if st != StructuralType.Beam:
            return False
    except Exception:
        return False
    try:
        sm = host.StructuralMaterialType
        if sm == StructuralMaterialType.Steel:
            return False
        if sm == StructuralMaterialType.Wood:
            return False
    except Exception:
        pass
    return True


def _rebar_cantidad_posiciones(rebar):
    if rebar is None:
        return 0
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 1


def _presentacion_estribo_first_last_en_vista(rebar, db_view):
    """
    En la vista dada, presentación del conjunto **Show first and last** (API ``FirstLast``).
    """
    if rebar is None or db_view is None:
        return
    try:
        mode = RebarPresentationMode.FirstLast
    except Exception:
        return
    try:
        if not rebar.CanApplyPresentationMode(db_view):
            return
        rebar.SetPresentationMode(db_view, mode)
    except Exception:
        pass


def _vista_permite_multi_rebar_annotation(view):
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import View3D

        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _centro_rebar_para_mra(rebar, view):
    if rebar is None:
        return None
    try:
        bb = rebar.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb0 = rebar.get_BoundingBox(None)
        if bb0 is not None:
            return (bb0.Min + bb0.Max) * 0.5
    except Exception:
        pass
    return None


def _norm_texto_tipo_mra(s):
    if s is None:
        return u""
    try:
        t = unicode(s)
    except Exception:
        try:
            t = System.Convert.ToString(s)
        except Exception:
            return u""
    try:
        return t.replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def _nombres_candidatos_multi_reference_annotation_type(t):
    """Cadenas a comparar con el nombre buscado (UI / parámetros tipo)."""
    out = []
    if t is None:
        return out
    try:
        n = getattr(t, "Name", None)
        if n:
            out.append(_norm_texto_tipo_mra(n))
    except Exception:
        pass
    try:
        p = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None and p.HasValue:
            if p.StorageType == StorageType.String:
                out.append(_norm_texto_tipo_mra(p.AsString()))
    except Exception:
        pass
    try:
        p = t.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p is not None and p.HasValue:
            if p.StorageType == StorageType.String:
                out.append(_norm_texto_tipo_mra(p.AsString()))
    except Exception:
        pass
    seen = set()
    uniq = []
    for x in out:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


def _multi_reference_annotation_type_by_name(document, type_name):
    """
    ``MultiReferenceAnnotationType`` cuyo nombre de tipo (UI) coincide con ``type_name``
    (exacto o solo distinto por mayúsculas).
    """
    if document is None:
        return None
    try:
        from Autodesk.Revit.DB import MultiReferenceAnnotationType
    except Exception:
        return None
    target = _norm_texto_tipo_mra(type_name)
    if not target:
        return None
    target_lower = target.lower()
    try:
        col = FilteredElementCollector(document).OfClass(MultiReferenceAnnotationType)
        exact = []
        ci = []
        for t in col:
            if t is None:
                continue
            for cand in _nombres_candidatos_multi_reference_annotation_type(t):
                if cand == target:
                    exact.append(t)
                    break
                try:
                    if cand.lower() == target_lower:
                        ci.append(t)
                        break
                except Exception:
                    pass
        if exact:
            return exact[0]
        if ci:
            return ci[0]
    except Exception:
        return None
    return None


def _multi_reference_annotation_type_estribo_bimtools(document):
    """
    ``MultiReferenceAnnotationType`` cuyo nombre coincide con :data:`_ESTRIBO_MRA_TYPE_NAME`
    (exacto, o solo distinto por mayúsculas).
    """
    return _multi_reference_annotation_type_by_name(document, _ESTRIBO_MRA_TYPE_NAME)


def crear_multi_rebar_annotations_por_nombre_tipo(
    document, view, rebars, avisos, type_name
):
    """
    Una ``MultiReferenceAnnotation`` por cada ``Rebar`` en ``rebars`` (cada conjunto).

    ``type_name``: nombre del tipo en la categoría **Multi-Rebar Annotations** del proyecto
    (p. ej. «Recorrido Barras»). Misma lógica de colocación que estribos (línea de cota debajo del
    conjunto según ``UpDirection`` de la vista), sin uso de cotas de traslape vigas.

    El llamador debe tener abierta una ``Transaction``.

    Returns:
        Cantidad de anotaciones creadas correctamente.
    """
    if document is None or view is None or not rebars or avisos is None:
        return 0
    try:
        tn = unicode(type_name).strip() if type_name is not None else u""
    except Exception:
        tn = (type_name or u"").strip() if type_name else u""
    if not tn:
        return 0
    if not _vista_permite_multi_rebar_annotation(view):
        avisos.append(
            u"Multi-Rebar Annotation («{0}»): use planta/alzado/sección (no plantilla ni 3D).".format(
                tn
            )
        )
        return 0
    mrat_type = _multi_reference_annotation_type_by_name(document, tn)
    if mrat_type is None:
        avisos.append(
            u"Multi-Rebar Annotation: no existe el tipo «{0}» en el proyecto.".format(tn)
        )
        return 0
    v_up_ref = None
    try:
        v_up_ref = view.UpDirection
        if v_up_ref is None or v_up_ref.GetLength() < 1e-12:
            v_up_ref = XYZ.BasisZ
        else:
            v_up_ref = v_up_ref.Normalize()
    except Exception:
        v_up_ref = XYZ.BasisZ
    candidatos = [rb for rb in (rebars or []) if rb is not None and isinstance(rb, Rebar)]
    d_list = []
    for rb in candidatos:
        pm = _centro_rebar_para_mra(rb, view)
        if pm is None:
            continue
        try:
            d_list.append(
                float(
                    _distancia_vertical_mra_estribo_ft(
                        document, view, pm, v_up_ref, s_lap_low_global=None
                    )
                )
            )
        except Exception:
            pass
    d_comun = max(d_list) if d_list else None
    n_ok = 0
    for rb in candidatos:
        if _try_crear_multi_rebar_annotation(
            document,
            view,
            rb,
            mrat_type,
            avisos,
            s_lap_low_global=None,
            vertical_offset_ft_override=d_comun,
        ):
            n_ok += 1
    return int(n_ok)


def _bb_centro_bb(bb):
    if bb is None:
        return None
    try:
        return (bb.Min + bb.Max) * 0.5
    except Exception:
        return None


def _bb_min_dot_up(bb, v_up):
    if bb is None or v_up is None:
        return None
    try:
        v_up = v_up.Normalize()
    except Exception:
        pass
    try:
        x0 = float(bb.Min.X)
        y0 = float(bb.Min.Y)
        z0 = float(bb.Min.Z)
        x1 = float(bb.Max.X)
        y1 = float(bb.Max.Y)
        z1 = float(bb.Max.Z)
    except Exception:
        return None
    m = None
    for ix in (x0, x1):
        for iy in (y0, y1):
            for iz in (z0, z1):
                try:
                    d = XYZ(ix, iy, iz).DotProduct(v_up)
                except Exception:
                    continue
                m = d if m is None else min(m, d)
    return m


def _cotas_traslape_vigas_en_vista(document, view):
    """
    ``Dimension`` de traslape creadas por enfierrado vigas (``lap_detail_link_vigas_schema.dim``).
    """
    out = []
    seen = set()
    if document is None or view is None:
        return out
    try:
        from lap_detail_link_vigas_schema import iter_vigas_lap_linked_detail_instances
    except Exception:
        return out
    for _inst, link in iter_vigas_lap_linked_detail_instances(document):
        did = link.get("dim")
        if did is None:
            continue
        try:
            if did == ElementId.InvalidElementId:
                continue
        except Exception:
            pass
        try:
            kid = int(did.IntegerValue)
        except Exception:
            continue
        if kid in seen:
            continue
        dim = document.GetElement(did)
        if dim is None:
            continue
        try:
            oid = getattr(dim, "OwnerViewId", None)
            if oid is not None and oid != view.Id:
                continue
        except Exception:
            try:
                if dim.ViewId != view.Id:
                    continue
            except Exception:
                pass
        seen.add(kid)
        out.append(dim)
    return out


def _s_min_traslapo_vigas_todas_en_vista(document, view, v_up):
    """
    Mínimo de ``punto·v_up`` en los bbox de **todas** las cotas de traslape vigas en la vista
    (borde más «bajo» de la cota respecto a ``v_up``). Una sola referencia propagada a todos los MRA.
    """
    if document is None or view is None or v_up is None:
        return None
    cotas = _cotas_traslape_vigas_en_vista(document, view)
    s_lap_low = None
    for dim in cotas:
        bb = None
        try:
            bb = dim.get_BoundingBox(view)
        except Exception:
            pass
        if bb is None:
            try:
                bb = dim.get_BoundingBox(None)
            except Exception:
                pass
        sm = _bb_min_dot_up(bb, v_up)
        if sm is not None:
            s_lap_low = sm if s_lap_low is None else min(s_lap_low, sm)
    return s_lap_low


def _distancia_vertical_mra_estribo_ft(document, view, p_mid, v_up, s_lap_low_global=None):
    """
    Distancia (pies) según ``−v_up`` desde ``p_mid`` hasta la línea MRA: como mínimo el
    offset base; si hay cotas de traslape (schema vigas), usa ``s_lap_low_global`` (todas las
    cotas de la vista) para bajar lo necesario y no solapar + margen.
    """
    off_base_ft = UnitUtils.ConvertToInternalUnits(
        float(_MRA_ESTRIBO_OFFSET_DEBAJO_MM), UnitTypeId.Millimeters
    )
    clearance_ft = UnitUtils.ConvertToInternalUnits(
        float(_MRA_ESTRIBO_CLEARANCE_TRASLAPE_MM), UnitTypeId.Millimeters
    )
    if document is None or view is None or p_mid is None or v_up is None:
        return off_base_ft
    s_lap_low = s_lap_low_global
    if s_lap_low is None:
        s_lap_low = _s_min_traslapo_vigas_todas_en_vista(document, view, v_up)
    try:
        pmd = float(p_mid.DotProduct(v_up))
    except Exception:
        return off_base_ft
    if s_lap_low is None:
        return off_base_ft
    try:
        need = pmd - float(s_lap_low) + float(clearance_ft)
    except Exception:
        return off_base_ft
    return max(float(off_base_ft), float(need))


def _try_crear_multi_rebar_annotation(
    document,
    view,
    rebar,
    mrat_type,
    avisos,
    s_lap_low_global=None,
    vertical_offset_ft_override=None,
):
    try:
        from Autodesk.Revit.DB import (
            DimensionStyleType,
            MultiReferenceAnnotation,
            MultiReferenceAnnotationOptions,
        )
    except Exception:
        return False
    if document is None or view is None or rebar is None:
        return False
    if mrat_type is None:
        return False
    p_mid = _centro_rebar_para_mra(rebar, view)
    if p_mid is None:
        return False
    try:
        vd = view.ViewDirection
        if vd is None or vd.GetLength() < 1e-12:
            return False
        vd = vd.Normalize()
        rd = view.RightDirection
        if rd is None or rd.GetLength() < 1e-12:
            return False
        rd = rd.Normalize()
        v_up = view.UpDirection
        if v_up is None or v_up.GetLength() < 1e-12:
            v_up = XYZ.BasisZ
        else:
            v_up = v_up.Normalize()
    except Exception:
        return False
    try:
        opts = MultiReferenceAnnotationOptions(mrat_type)
    except Exception:
        try:
            opts = MultiReferenceAnnotationOptions()
            try:
                opts.MultiReferenceAnnotationType = mrat_type.Id
            except Exception:
                return False
        except Exception:
            return False
    try:
        opts.DimensionStyleType = DimensionStyleType.Linear
    except Exception:
        pass
    try:
        opts.DimensionPlaneNormal = vd
        opts.DimensionLineDirection = rd
        if vertical_offset_ft_override is not None:
            off_abajo_ft = float(vertical_offset_ft_override)
        else:
            off_abajo_ft = _distancia_vertical_mra_estribo_ft(
                document, view, p_mid, v_up, s_lap_low_global=s_lap_low_global
            )
        # Debajo de la armadura; referencia de traslape **única**; offset común opcional (misma fila).
        try:
            p_line = p_mid - v_up.Multiply(float(off_abajo_ft))
        except Exception:
            p_line = p_mid
        opts.DimensionLineOrigin = p_line
        opts.TagHeadPosition = p_line
        # Sin líder de llamada (tipo «Structural Rebar_Estribo» + política BIMTools).
        opts.TagHasLeader = False
    except Exception:
        return False
    ids = List[ElementId]()
    ids.Add(rebar.Id)
    try:
        opts.SetElementsToDimension(ids)
    except Exception:
        return False
    try:
        if hasattr(opts, "ElementsMatchReferenceCategory"):
            if not opts.ElementsMatchReferenceCategory(document):
                try:
                    rid = rebar.Id.IntegerValue
                except Exception:
                    rid = u"?"
                avisos.append(
                    u"Multi-Rebar Tag estribo Id {0}: elementos no válidos para el tipo MRA.".format(
                        rid
                    )
                )
                return False
    except Exception:
        pass
    try:
        mra = MultiReferenceAnnotation.Create(document, view.Id, opts)
        if mra is None:
            return False
        try:
            from Autodesk.Revit.DB import (
                BuiltInCategory,
                ElementCategoryFilter,
                IndependentTag,
            )

            flt = ElementCategoryFilter(BuiltInCategory.OST_RebarTags)
            dep_ids = mra.GetDependentElements(flt)
            if not dep_ids:
                try:
                    dep_ids = mra.GetDependentElements(None)
                except Exception:
                    dep_ids = None
            if dep_ids:
                for did in dep_ids:
                    el = document.GetElement(did)
                    if isinstance(el, IndependentTag):
                        try:
                            if el.HasLeader:
                                el.HasLeader = False
                        except Exception:
                            pass
        except Exception:
            pass
        return True
    except Exception as ex:
        try:
            rid = rebar.Id.IntegerValue
        except Exception:
            rid = u"?"
        try:
            msg = unicode(ex)
        except Exception:
            msg = str(ex)
        avisos.append(
            u"Multi-Rebar Tag estribo Id {0}: {1}".format(rid, msg)
        )
        return False


def crear_multi_rebar_tags_estribos_en_vista(document, view, rebars, avisos):
    """
    Una ``MultiReferenceAnnotation`` por cada ``Rebar`` en ``rebars`` (cada conjunto de estribo).
    """
    if document is None or view is None or not rebars or avisos is None:
        return 0
    if not _vista_permite_multi_rebar_annotation(view):
        avisos.append(
            u"Etiquetas estribo: use vista ortogonal (planta/alzado/sección), no plantilla ni 3D."
        )
        return 0
    mrat_type = _multi_reference_annotation_type_estribo_bimtools(document)
    if mrat_type is None:
        avisos.append(
            u"Etiquetas estribo: no hay MultiReferenceAnnotationType «{0}» en el proyecto.".format(
                _ESTRIBO_MRA_TYPE_NAME
            )
        )
        return 0
    v_up_ref = None
    try:
        v_up_ref = view.UpDirection
        if v_up_ref is None or v_up_ref.GetLength() < 1e-12:
            v_up_ref = XYZ.BasisZ
        else:
            v_up_ref = v_up_ref.Normalize()
    except Exception:
        v_up_ref = XYZ.BasisZ
    s_lap_global = _s_min_traslapo_vigas_todas_en_vista(document, view, v_up_ref)
    candidatos = [rb for rb in (rebars or []) if rb is not None and isinstance(rb, Rebar)]
    d_list = []
    for rb in candidatos:
        pm = _centro_rebar_para_mra(rb, view)
        if pm is None:
            continue
        try:
            d_list.append(
                float(
                    _distancia_vertical_mra_estribo_ft(
                        document, view, pm, v_up_ref, s_lap_low_global=s_lap_global
                    )
                )
            )
        except Exception:
            pass
    d_comun = max(d_list) if d_list else None
    n_ok = 0
    for rb in candidatos:
        if _try_crear_multi_rebar_annotation(
            document,
            view,
            rb,
            mrat_type,
            avisos,
            s_lap_low_global=s_lap_global,
            vertical_offset_ft_override=d_comun,
        ):
            n_ok += 1
    return int(n_ok)


def _curve_loop_to_list(loop):
    lst = List[Curve]()
    if loop is None:
        return None
    try:
        it = loop.GetCurveLoopIterator()
        while it.MoveNext():
            c = it.Current
            if c is not None and c.IsBound:
                lst.Add(c)
    except Exception:
        return None
    if lst.Count < 2:
        return None
    return lst


def _curve_list_translated(base_list, delta):
    if base_list is None or delta is None:
        return None
    try:
        tr = Transform.CreateTranslation(delta)
        out = List[Curve]()
        n = int(base_list.Count)
        for i in range(n):
            out.Add(base_list[i].CreateTransformed(tr))
        return out
    except Exception:
        return None


def _shape_compare_string(value):
    if value is None:
        return u""
    try:
        t = unicode(value)
    except Exception:
        try:
            t = System.Convert.ToString(value)
        except Exception:
            return u""
    try:
        return t.replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""


def _rebar_shape_display_name_visible(sh):
    """Nombre visible del ``RebarShape`` (mismo criterio que fundación)."""
    if rebar_shape_display_name is not None:
        try:
            return _shape_compare_string(rebar_shape_display_name(sh))
        except Exception:
            pass
    try:
        return _shape_compare_string(getattr(sh, "Name", None))
    except Exception:
        return u""


def _rebar_shape_por_nombre_exacto(document, nombre):
    """
    ``RebarShape`` por nombre mostrado (Navegador): ``SYMBOL_NAME_PARAM`` / ``ALL_MODEL_TYPE_NAME``,
    no solo ``Element.Name`` (a menudo vacío). Coincidencia exacta, luego sin mayúsculas, luego dígitos.
    """
    if document is None or not nombre:
        return None
    key = _shape_compare_string(nombre)
    if not key:
        return None
    try:
        key_lower = key.lower()
    except Exception:
        key_lower = key
    key_digits = u"".join(
        ch for ch in key if ch in u"0123456789"
    )
    if buscar_rebar_shape_por_nombre is not None:
        try:
            sh = buscar_rebar_shape_por_nombre(document, key)
            if sh is not None:
                return sh
        except Exception:
            pass
    candidates = []
    try:
        for sh in FilteredElementCollector(document).OfClass(RebarShape):
            try:
                sn = _rebar_shape_display_name_visible(sh)
                if not sn:
                    continue
                try:
                    sn_low = sn.lower()
                except Exception:
                    sn_low = sn
                digits_only = u"".join(
                    ch for ch in sn if ch in u"0123456789"
                )
                candidates.append((sh, sn, sn_low, digits_only))
            except Exception:
                continue
    except Exception:
        return None
    for sh, sn, sn_low, _dig in candidates:
        if sn == key:
            return sh
    for sh, sn, sn_low, _dig in candidates:
        if sn_low == key_lower:
            return sh
    for sh, sn, sn_low, dig in candidates:
        if dig and dig == key:
            return sh
    for sh, sn, sn_low, dig in candidates:
        if key_digits and dig == key_digits:
            return sh
    return None


def _curve_list_to_object_ilist(curves_list):
    """``List[object]`` con las mismas curvas (orden original)."""
    olist = List[object]()
    if curves_list is None:
        return olist
    try:
        n = int(curves_list.Count)
        for i in range(n):
            olist.Add(curves_list[i])
    except Exception:
        pass
    return olist


def _polygon_points_desde_curvas_cerradas(curves_list):
    """Vértices en orden (``GetEndPoint(0)`` de cada tramo de un polígono cerrado)."""
    pts = []
    if curves_list is None:
        return pts
    try:
        n = int(curves_list.Count)
        for i in range(n):
            c = curves_list[i]
            if c is None:
                return []
            pts.append(c.GetEndPoint(0))
    except Exception:
        return []
    return pts


def _object_ilist_desde_puntos_cerrados(pts):
    """Polígono cerrado como ``List[object]`` de ``Line``."""
    olist = List[object]()
    if not pts or len(pts) < 3:
        return olist
    m = len(pts)
    try:
        for i in range(m):
            a = pts[i]
            b = pts[(i + 1) % m]
            if a.DistanceTo(b) <= _MIN_EDGE_FT:
                continue
            olist.Add(Line.CreateBound(a, b))
    except Exception:
        return List[object]()
    return olist


def _variants_object_ilist_stirrup(curves_list):
    """
    Variantes del bucle para ``CreateFromCurvesAndShape``:
    orden original, sentido inverso en vértices (mismo polígono, otra dirección).
    """
    out = []
    direct = _curve_list_to_object_ilist(curves_list)
    if direct is not None and direct.Count >= 2:
        out.append(direct)
    pts = _polygon_points_desde_curvas_cerradas(curves_list)
    if len(pts) >= 3:
        try:
            pts_inv = [pts[0]] + [pts[i] for i in range(len(pts) - 1, 0, -1)]
            inv = _object_ilist_desde_puntos_cerrados(pts_inv)
            if inv is not None and inv.Count >= 2:
                out.append(inv)
        except Exception:
            pass
    return out


def _iter_rebar_hook_types_candidatos(document, max_n=12):
    """Hasta ``max_n`` ``RebarHookType`` del proyecto (los estribos a veces fallan con el primero)."""
    if document is None:
        return
    n = 0
    try:
        for ht in FilteredElementCollector(document).OfClass(RebarHookType):
            if ht is None:
                continue
            yield ht
            n += 1
            if n >= int(max_n):
                break
    except Exception:
        return


def _try_create_estribo_rebar_shape_nombre(
    document, host, bar_type, axis_norm, curves_list, shape_nombre
):
    """
    Crea estribo con ``Rebar.CreateFromCurvesAndShape`` usando el ``RebarShape`` por nombre.
    Prueba sobrecarga extendida (Revit 2021+), varias orientaciones, ganchos y sentidos del bucle.
    Returns:
        ``(rebar, mensaje_error o None)``.
    """
    if document is None or host is None or bar_type is None or curves_list is None:
        return None, u"Parámetros incompletos para CreateFromCurvesAndShape."
    shape = _rebar_shape_por_nombre_exacto(document, shape_nombre)
    if shape is None:
        return None, u"No se resolvió RebarShape «{0}» en el proyecto.".format(shape_nombre)
    curve_variants = _variants_object_ilist_stirrup(curves_list)
    if not curve_variants:
        return None, u"Lista de curvas vacía o insuficiente."
    try:
        ax = axis_norm.Normalize()
    except Exception:
        ax = axis_norm
    norms = [ax]
    try:
        norms.append(ax.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
    )
    hooks = list(_iter_rebar_hook_types_candidatos(document, 12))
    if not hooks and _pick_first_hook_type is not None:
        h0 = _pick_first_hook_type(document)
        hooks = [h0] if h0 is not None else []
    if not hooks:
        return None, u"No hay RebarHookType en el proyecto."
    last_err = None
    invalid = ElementId.InvalidElementId
    for curves_clr in curve_variants:
        if curves_clr is None or curves_clr.Count < 2:
            continue
        for hook in hooks:
            if hook is None:
                continue
            for nvec in norms:
                for so, eo in orient_pairs:
                    try:
                        r = Rebar.CreateFromCurvesAndShape(
                            document,
                            shape,
                            bar_type,
                            hook,
                            hook,
                            host,
                            nvec,
                            curves_clr,
                            so,
                            eo,
                            0.0,
                            0.0,
                            invalid,
                            invalid,
                        )
                        if r is not None:
                            return r, None
                    except Exception as ex:
                        try:
                            last_err = unicode(ex)
                        except Exception:
                            last_err = str(ex)
                    try:
                        r = Rebar.CreateFromCurvesAndShape(
                            document,
                            shape,
                            bar_type,
                            hook,
                            hook,
                            host,
                            nvec,
                            curves_clr,
                            so,
                            eo,
                        )
                        if r is not None:
                            return r, None
                    except Exception as ex:
                        try:
                            last_err = unicode(ex)
                        except Exception:
                            last_err = str(ex)
                        continue
    detail = u" Shape Id={0}.".format(shape.Id.IntegerValue)
    if last_err:
        detail += u" Último error API: {0}".format(last_err)
    return None, (
        u"CreateFromCurvesAndShape falló (geometría incompatible con la forma, ganchos u orientación)."
        + detail
    )


def _crear_rebar_estribos_multizonas(
    document,
    host,
    line_work,
    loop_draw,
    axis,
    bar_type_ext,
    bar_type_cent,
    spacing_ext_mm,
    spacing_cent_mm,
    avisos,
    rebars_creados=None,
    view=None,
):
    """
    Crea uno o más ``Rebar`` (estribo) desde el bucle geométrico, con
    ``SetLayoutAsMaximumSpacing``. Longitud total de reparto =
    ``line_work.Length`` menos ``_ESTRIBO_PLANO_DESDE_INICIO_MM`` y
    ``_ESTRIBO_ARRAY_LONGITUD_MENOS_MM`` (50 mm + 50 mm por defecto).

    Si el canto **h** (mm) de la viga es conocido y ``L_arr >= 2·h`` (en longitud física):
    tramo inicial = ``min(2·h, L_arr/2)`` con carta *Extremos*, tramo final igual,
    tramo central = resto con carta *Centrales*.     Si ``L_arr < 2·h``, un solo tramo con carta *Centrales* y ``includeFirstBar`` /
    ``includeLastBar`` en True (vano corto: barra de inicio y término encendidas). Si **h**
    no se obtiene, se mantiene el reparto 25 % / 50 % / 25 %. En el tramo *Central* intermedio
    (vano largo), esos flags van en False; en *Extremos*, en True.
    """
    if (
        document is None
        or host is None
        or line_work is None
        or loop_draw is None
        or axis is None
    ):
        return 0
    if _apply_maximum_spacing_layout is None or _pick_first_hook_type is None:
        avisos.append(
            u"Estribos: no se cargó estribos_viga_rps (layout/hook no disponibles)."
        )
        return 0
    if not _host_viga_permite_rebar(host):
        try:
            eid = host.Id.IntegerValue
        except Exception:
            eid = u"?"
        avisos.append(
            u"Viga {0}: host no válido para Rebar (tipo/material).".format(eid)
        )
        return 0
    bt_e = bar_type_ext
    bt_c = bar_type_cent
    if bt_e is None and bt_c is None:
        avisos.append(u"Estribos: sin RebarBarType (Ø extremos/centrales).")
        return 0
    if bt_e is None:
        bt_e = bt_c
    if bt_c is None:
        bt_c = bt_e
    try:
        sp_e = float(max(50.0, float(spacing_ext_mm or 200.0)))
    except Exception:
        sp_e = 200.0
    try:
        sp_c = float(max(50.0, float(spacing_cent_mm or 200.0)))
    except Exception:
        sp_c = 200.0
    try:
        Lw = float(line_work.Length)
        inset_ini = _mm_to_internal(_ESTRIBO_PLANO_DESDE_INICIO_MM)
        inset_fin = _mm_to_internal(_ESTRIBO_ARRAY_LONGITUD_MENOS_MM)
        L_arr = max(0.0, Lw - inset_ini - inset_fin)
    except Exception:
        return 0
    if L_arr < _MIN_EDGE_FT * 4.0:
        try:
            eid = host.Id.IntegerValue
        except Exception:
            eid = u"?"
        avisos.append(
            u"Viga {0}: vano estribo (L−50 mm−50 mm) demasiado corto; no se creó Rebar.".format(
                eid
            )
        )
        return 0
    base_template = _curve_loop_to_list(loop_draw)
    if base_template is None:
        avisos.append(u"Estribos: bucle de curvas inválido para CreateFromCurves.")
        return 0
    min_len = max(_mm_to_internal(sp_e), _mm_to_internal(sp_c)) * 0.2
    h_mm = _altura_viga_estribos_mm(document, host)
    zonas = []
    if h_mm is not None and float(h_mm) > 0.0:
        try:
            two_h_ft = _mm_to_internal(2.0 * float(h_mm))
        except Exception:
            two_h_ft = None
        if two_h_ft is not None and L_arr < float(two_h_ft) - 1e-9:
            # Vano corto: un solo set con Ø/esp. centrales; primera y última barra activas.
            zonas = [(L_arr, bt_c, sp_c, True)]
        elif two_h_ft is not None:
            try:
                L_ext_tgt = _mm_to_internal(2.0 * float(h_mm))
            except Exception:
                L_ext_tgt = None
            if L_ext_tgt is None or L_ext_tgt < _MIN_EDGE_FT:
                pass
            else:
                L_half = 0.5 * L_arr
                L_ext_each = min(float(L_ext_tgt), float(L_half))
                L_cent = max(0.0, L_arr - 2.0 * L_ext_each)
                if L_cent < min_len + _MIN_EDGE_FT:
                    zonas = [
                        (L_ext_each, bt_e, sp_e, True),
                        (L_ext_each, bt_e, sp_e, True),
                    ]
                else:
                    zonas = [
                        (L_ext_each, bt_e, sp_e, True),
                        (L_cent, bt_c, sp_c, False),
                        (L_ext_each, bt_e, sp_e, True),
                    ]
    if not zonas:
        f = _ESTRIBO_FRACC_EXTREMOS
        L1 = L_arr * f
        L2 = L_arr * max(0.0, 1.0 - 2.0 * f)
        L3 = L_arr * f
        if L2 < min_len + _MIN_EDGE_FT:
            zonas = [(L_arr, bt_e, sp_e, True)]
        else:
            zonas = [
                (L1, bt_e, sp_e, True),
                (L2, bt_c, sp_c, False),
                (L3, bt_e, sp_e, True),
            ]
    cum = 0.0
    total_qty = 0
    try:
        ax = axis.Normalize()
    except Exception:
        ax = axis
    for Lz, bt, sp_mm, include_end_bars in zonas:
        if Lz < _MIN_EDGE_FT or bt is None:
            continue
        try:
            dv = ax.Multiply(float(cum))
        except Exception:
            break
        curves_z = _curve_list_translated(base_template, dv)
        if curves_z is None:
            avisos.append(u"Estribos: CreateTransformed en curvas falló.")
            break
        rebar, err_txt = _try_create_estribo_rebar_shape_nombre(
            document,
            host,
            bt,
            ax,
            curves_z,
            _ESTRIBO_REBAR_SHAPE_NOMBRE,
        )
        if rebar is None:
            try:
                eid = host.Id.IntegerValue
            except Exception:
                eid = u"?"
            msg = u"Viga {0}: no se creó Rebar estribo (shape «{1}»).".format(
                eid,
                _ESTRIBO_REBAR_SHAPE_NOMBRE,
            )
            if err_txt:
                msg += u" {0}".format(err_txt)
            avisos.append(msg)
            break
        sp_ft = _mm_to_internal(float(sp_mm))
        if not _apply_maximum_spacing_layout(
            rebar, sp_ft, Lz, include_end_bars=include_end_bars
        ):
            try:
                document.Delete(rebar.Id)
            except Exception:
                pass
            avisos.append(
                u"Viga: estribo creado pero falló MaximumSpacing (se eliminó la instancia)."
            )
            break
        total_qty += _rebar_cantidad_posiciones(rebar)
        if rebars_creados is not None:
            rebars_creados.append(rebar)
        if view is not None:
            _presentacion_estribo_first_last_en_vista(rebar, view)
        cum += float(Lz)
    return int(total_qty)


def crear_model_lines_preview_estribo_viga(
    document,
    elemento_framing,
    cover_mm=25.0,
    obstaculos=None,
    crear_rebar_estribos=False,
    rebar_bar_type_ext=None,
    rebar_bar_type_cent=None,
    spacing_ext_mm=200,
    spacing_cent_mm=200,
    crear_model_lines=False,
    out_rebars_creados=None,
    view=None,
):
    """
    Geometría de estribo (eje entre tapas, sección a **50 mm** del inicio, bucle con recubrimiento).

    Si ``crear_model_lines`` es ``True``, dibuja **ModelLine/ModelCurve** de eje y sección.
    Por defecto ``False`` (solo ``Rebar`` si ``crear_rebar_estribos``).

    ``obstaculos`` se ignora (reservado); el recorte usa solo la geometría del framing.

    Si ``out_rebars_creados`` es una ``list``, se añaden ahí los ``Rebar`` de estribo creados
    (para ``MultiReferenceAnnotation`` u otros usos).

    Si ``view`` no es ``None`` y se crean estribos, se aplica ``RebarPresentationMode.FirstLast``
    en esa vista (equivalente UI *Show first and last*).

    Returns:
        ``(n_model_curves_seccion, n_model_curves_eje, n_rebar_posiciones, avisos)``.
    """
    _ = obstaculos  # API estable; recorte solo por tapas del sólido de la viga.
    avisos = []
    n_sec = 0
    n_axis = 0
    n_rb = 0
    if document is None or elemento_framing is None:
        return 0, 0, 0, avisos
    crv = _curva_location_framing(elemento_framing)
    if crv is None:
        try:
            eid = elemento_framing.Id.IntegerValue
        except Exception:
            eid = u"?"
        avisos.append(u"Viga {0}: sin LocationCurve.".format(eid))
        return 0, 0, 0, avisos
    line_full = _line_bound_desde_location_curve(crv)
    if line_full is None:
        avisos.append(u"Viga: no se pudo obtener la cuerda del LocationCurve.")
        return 0, 0, 0, avisos
    try:
        p0f = line_full.GetEndPoint(0)
        p1f = line_full.GetEndPoint(1)
        t = (p1f.Subtract(p0f)).Normalize()
    except Exception:
        avisos.append(u"Viga: dirección del eje inválida.")
        return 0, 0, 0, avisos
    if t is None or t.GetLength() < _MIN_EDGE_FT:
        avisos.append(u"Viga: dirección del eje inválida.")
        return 0, 0, 0, avisos
    if abs(float(t.Z)) > _AXIS_CASI_VERTICAL_TOL:
        avisos.append(
            u"Viga: eje casi vertical (|Z|>0,92); se omite preview de estribos."
        )
        return 0, 0, 0, avisos

    solid = _solido_principal(elemento_framing)
    if solid is None:
        try:
            eid = elemento_framing.Id.IntegerValue
        except Exception:
            eid = u"?"
        avisos.append(u"Viga {0}: sin sólido de volumen.".format(eid))
        return 0, 0, 0, avisos

    line_work = _linea_entre_tapas_extremas_viga(line_full, solid)
    if line_work is None:
        line_work = line_full
        try:
            eid = elemento_framing.Id.IntegerValue
        except Exception:
            eid = u"?"
        avisos.append(
            u"Viga {0}: no se detectaron dos tapas extremas en el sólido; "
            u"se usa la cuerda del LocationCurve completa.".format(eid)
        )
    try:
        p0 = line_work.GetEndPoint(0)
        p1 = line_work.GetEndPoint(1)
        t = (p1.Subtract(p0)).Normalize()
        Lw = float(line_work.Length)
    except Exception:
        avisos.append(u"Viga: tramo recortado inválido.")
        return 0, 0, 0, avisos
    offset_ft = _mm_to_internal(_ESTRIBO_PLANO_DESDE_INICIO_MM)
    if Lw < offset_ft + _MIN_EDGE_FT:
        try:
            eid = elemento_framing.Id.IntegerValue
        except Exception:
            eid = u"?"
        offset_ft = max(_MIN_EDGE_FT * 4.0, 0.5 * Lw - _MIN_EDGE_FT)
        avisos.append(
            u"Viga {0}: tramo < {1} mm; sección en ~{2:.0f}% del eje recortado.".format(
                eid,
                int(round(_ESTRIBO_PLANO_DESDE_INICIO_MM)),
                100.0 * float(offset_ft) / Lw if Lw > 1e-12 else 0.0,
            )
        )
    try:
        pt_seccion = p0.Add(t.Multiply(float(offset_ft)))
    except Exception:
        avisos.append(u"Viga: no se pudo el punto de sección sobre el eje.")
        return 0, 0, 0, avisos
    if crear_model_lines:
        if _crear_model_curve(document, line_work, None):
            n_axis = 1
        else:
            try:
                eid = elemento_framing.Id.IntegerValue
            except Exception:
                eid = u"?"
            avisos.append(
                u"Viga {0}: no se creó ModelLine del eje recortado.".format(eid)
            )

    try:
        plane_cut = Plane.CreateByNormalAndOrigin(t, pt_seccion)
    except Exception as ex:
        try:
            avisos.append(u"Plano de corte: {0}".format(unicode(ex)))
        except Exception:
            avisos.append(u"Plano de corte: error.")
        return n_sec, n_axis, n_rb, avisos
    raw_pts = _puntos_interseccion_solido_plano(solid, plane_cut)
    pts = _dedupe_puntos(raw_pts, _TOL_MERGE_PTS_FT)
    if len(pts) < 3:
        try:
            eid = elemento_framing.Id.IntegerValue
        except Exception:
            eid = u"?"
        avisos.append(
            u"Viga {0}: la intersección sólido–plano dio menos de 3 puntos.".format(eid)
        )
        return n_sec, n_axis, n_rb, avisos
    u, v = _base_uv_en_plano(t)
    if u is None:
        avisos.append(u"No se pudo construir base UV del plano de sección.")
        return n_sec, n_axis, n_rb, avisos
    loop_poly = _curveloop_desde_puntos_ordenados(pts, u, v, pt_seccion)
    if loop_poly is None:
        avisos.append(u"No se pudo formar polígono de sección (convexa).")
        return n_sec, n_axis, n_rb, avisos
    loop_off = _curveloop_recubrimiento_interior(loop_poly, cover_mm, t)
    if loop_off is None:
        avisos.append(
            u"Offset de recubrimiento ({0} mm) falló; se dibuja el perímetro bruto.".format(
                cover_mm
            )
        )
        loop_draw = loop_poly
    else:
        loop_draw = loop_off
    if crear_model_lines:
        n_sec = _crear_model_curves_desde_loop(document, loop_draw, plane_cut)
        if n_sec <= 0:
            avisos.append(u"No se crearon ModelCurve de sección en el documento.")
    if crear_rebar_estribos and loop_draw is not None:
        _rebars_local = []
        n_rb = _crear_rebar_estribos_multizonas(
            document,
            elemento_framing,
            line_work,
            loop_draw,
            t,
            rebar_bar_type_ext,
            rebar_bar_type_cent,
            spacing_ext_mm,
            spacing_cent_mm,
            avisos,
            rebars_creados=_rebars_local,
            view=view,
        )
        if out_rebars_creados is not None:
            try:
                out_rebars_creados.extend(_rebars_local)
            except Exception:
                pass
    return n_sec, n_axis, n_rb, avisos


def crear_model_lines_preview_estribos_lista(
    document,
    elementos_framing,
    cover_mm=25.0,
    obstaculos=None,
    crear_rebar_estribos=False,
    rebar_bar_type_ext=None,
    rebar_bar_type_cent=None,
    spacing_ext_mm=200,
    spacing_cent_mm=200,
    crear_model_lines=False,
    view=None,
    crear_multi_rebar_tags=False,
):
    """
    Igual que :func:`crear_model_lines_preview_estribo_viga` para una lista de vigas.

    Con ``crear_multi_rebar_tags`` y ``view`` válidos, crea una ``MultiReferenceAnnotation``
    por cada conjunto de estribo tras procesar todas las vigas (una pasada).

    Returns:
        ``(total_seccion, total_eje, total_rebar_posiciones, avisos)``.
    """
    tot_sec = 0
    tot_ax = 0
    tot_rb = 0
    all_av = []
    tag_bucket = [] if crear_multi_rebar_tags else None
    for el in elementos_framing or []:
        if el is None:
            continue
        ns, na, nr, av = crear_model_lines_preview_estribo_viga(
            document,
            el,
            cover_mm,
            obstaculos,
            crear_rebar_estribos,
            rebar_bar_type_ext,
            rebar_bar_type_cent,
            spacing_ext_mm,
            spacing_cent_mm,
            crear_model_lines,
            out_rebars_creados=tag_bucket,
            view=view,
        )
        tot_sec += int(ns)
        tot_ax += int(na)
        tot_rb += int(nr)
        all_av.extend(av or [])
    if (
        crear_multi_rebar_tags
        and view is not None
        and tag_bucket
        and tot_rb > 0
    ):
        try:
            crear_multi_rebar_tags_estribos_en_vista(
                document, view, tag_bucket, all_av
            )
        except Exception as ex:
            try:
                all_av.append(
                    u"Etiquetas estribo (Rebar ya colocada): {0}".format(
                        unicode(ex)
                    )
                )
            except Exception:
                all_av.append(
                    u"Etiquetas estribo: error; la armadura sí quedó creada."
                )
    return tot_sec, tot_ax, tot_rb, all_av
