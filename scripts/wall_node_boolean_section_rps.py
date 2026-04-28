# -*- coding: utf-8 -*-
"""
Nodo / encuentro estructural: unión booleana (muro + muros en extremos + suelos)
e intersección con plano vertical típico: **normal de la curva de ubicación** (no la tangente).

Pasos:
  1) Seleccionar un único muro; ``LocationCurve`` → ``Line`` (extremos si no es línea; la normal usa la curva original).
  2) Punto medio; **normal** = ``Arc.Normal`` o, si es línea, ``Direction × Z`` alineada con ``Wall.Orientation``.
  3) Candidatos unidos / que intersectan: ``JoinGeometryUtils`` + ``ElementIntersectsElementFilter``.
 - Muros laterales: unidos cerca de los **extremos** del eje, o con encuentro en **T / cara lateral**
   a lo largo del tramo (proyección interior del eje, casi ortogonal, sin unión longitudinal).
 - Suelos: en la **cara superior** del muro, los elementos unidos a esa cara **no** se unen al sólido del boceto.
     En la **cara inferior**, para unir sólidos del boceto **solo** fundación estructural o losa de cimentación;
     forjados que no son fundación en la inferior **no** se unen. Suelos cuyo bbox **corta el plano de sección**
     y se solapan en planta con el muro (además de los ya detectados).
 - Muros: un **muro apilado** solo bajo el muro principal (cara superior del vecino ≈ base del host) **no** se une
   al sólido del boceto (no aporta el nudo en planta como un T/L a extremos).
  4) Extraer ``Solid`` de cada elemento (filtrar nulos / volumen ~0); ``BooleanOperationsType.Union``.
  5) ``BooleanOperationsUtils.CutWithHalfSpace`` (prueba normal ±) para obtener la cara de corte.
  6) ``PlanarFace.GetEdgesAsCurveLoops`` → ``SketchPlane`` + ``NewModelCurve``.
  7) Si el muro tiene fundación / losa de cimentación unida (*Join Geometry*), la base del boceto (Z mínima) se estira en el plano de corte según la altura de bbox de esa fundación.
  8) Offset hacia el interior: **50 mm** (fundación, cota baja) / **25 mm** (resto). La dirección “interior” se
     fija por el **sentido del polígono** (área firmada en el plano del boceto), no solo por el baricentro —
     así en formas cóncavas o en L el offset no invierte hacia afuera. Por loop de ``Line``: intersección de
     paralelas; respaldo ``CreateOffset`` con la misma lógica por arista cuando hay vértices cerrados.

Uso en Revit Python Shell::

    main(__revit__)

Compatibilidad: **Revit 2024 en adelante** (y pyRevit 4.8+ / 5.x). Los ``ElementId`` usan
``Value`` (API actual); se mantiene respaldo a ``IntegerValue`` (IronPython 2.7 / RPS heredado).
"""

from __future__ import print_function

import math
import os

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Arc,
    BooleanOperationsType,
    BooleanOperationsUtils,
    BoundingBoxIntersectsFilter,
    BuiltInCategory,
    ElementId,
    ElementIntersectsElementFilter,
    FilteredElementCollector,
    Floor,
    FloorType,
    GeometryInstance,
    JoinGeometryUtils,
    Line,
    LocationCurve,
    Options,
    Outline,
    Plane,
    PlanarFace,
    SketchPlane,
    Solid,
    Transaction,
    Transform,
    UnitUtils,
    UnitTypeId,
    UV,
    ViewDetailLevel,
    Wall,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from bimtools_joined_geometry import get_joined_element_ids


def _element_id_to_int(eid):
    """
    ID numérico de ``ElementId`` (Revit 2024+ ``Value``; API clásica ``IntegerValue``).
    Evita fallos con pyRevit 4.8/5.x bajo distintas versiones de la API.
    """
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _element_ids_equal(eid_a, eid_b):
    """Compara dos ``ElementId`` con tolerancia a API antigua o nueva."""
    if eid_a is None or eid_b is None:
        return False
    if eid_a == ElementId.InvalidElementId or eid_b == ElementId.InvalidElementId:
        return False
    try:
        if eid_a == eid_b:
            return True
    except Exception:
        pass
    ia = _element_id_to_int(eid_a)
    ib = _element_id_to_int(eid_b)
    if ia is not None and ib is not None:
        return ia == ib
    return False


def _xyz_normalize(v):
    if v is None:
        return None
    try:
        ln = v.GetLength()
        if ln < 1e-12:
            return None
        return XYZ(v.X / ln, v.Y / ln, v.Z / ln)
    except Exception:
        return None


def _tangent_at_mid(curve):
    if curve is None:
        return None
    try:
        tr = curve.ComputeDerivatives(0.5, True)
        return _xyz_normalize(tr.BasisX)
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return _xyz_normalize(XYZ(p1.X - p0.X, p1.Y - p0.Y, p1.Z - p0.Z))
    except Exception:
        return None


def _midpoint_curve(curve):
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


def _is_line_curve(curve):
    try:
        if isinstance(curve, Line):
            return True
    except Exception:
        pass
    try:
        return type(curve).__name__ == "Line"
    except Exception:
        return False


def _is_arc_curve(curve):
    try:
        if isinstance(curve, Arc):
            return True
    except Exception:
        pass
    try:
        return type(curve).__name__ == "Arc"
    except Exception:
        return False


def _normal_de_location_curve(curve, wall):
    """
    Normal unitaria del plano de corte asociada a la **curva de ubicación**:
    - ``Arc``: ``Arc.Normal`` (normal del plano del arco).
    - ``Line``: perpendicular al eje en planta (``Direction × Z``), signo alineado con ``Wall.Orientation``.
    - Otros tipos: tangente en 0.5 × ``Z`` (misma lógica que línea recta).
    """
    if curve is None:
        return None

    if _is_arc_curve(curve):
        try:
            return _xyz_normalize(curve.Normal)
        except Exception:
            return None

    if _is_line_curve(curve):
        try:
            d = curve.Direction
            z = XYZ.BasisZ
            if abs(d.DotProduct(z)) > 0.99999:
                n = d.CrossProduct(XYZ.BasisX)
            else:
                n = d.CrossProduct(z)
            n = _xyz_normalize(n)
            if n is None:
                return None
            try:
                wo = _xyz_normalize(wall.Orientation)
                if wo is not None and float(n.DotProduct(wo)) < 0.0:
                    n = XYZ(-n.X, -n.Y, -n.Z)
            except Exception:
                pass
            return n
        except Exception:
            return None

    t = _tangent_at_mid(curve)
    if t is None:
        return None
    z = XYZ.BasisZ
    try:
        if abs(t.DotProduct(z)) > 0.99999:
            n = t.CrossProduct(XYZ.BasisX)
        else:
            n = t.CrossProduct(z)
        n = _xyz_normalize(n)
        if n is None:
            return None
        try:
            wo = _xyz_normalize(wall.Orientation)
            if wo is not None and float(n.DotProduct(wo)) < 0.0:
                n = XYZ(-n.X, -n.Y, -n.Z)
        except Exception:
            pass
        return n
    except Exception:
        return None


def _location_as_line(wall):
    """
    Devuelve (Line, curve_original) a partir de ``LocationCurve``.
    Si la ubicación ya es ``Line``, se reutiliza; si no, ``Line.CreateBound(p0,p1)``.
    """
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return None, None
    c = loc.Curve
    if c is None:
        return None, None
    try:
        if isinstance(c, Line):
            return c, c
    except Exception:
        pass
    try:
        if type(c).__name__ == "Line":
            return c, c
    except Exception:
        pass
    try:
        p0 = c.GetEndPoint(0)
        p1 = c.GetEndPoint(1)
        ln = Line.CreateBound(p0, p1)
        return ln, c
    except Exception:
        return None, None


def _dist_point_to_curve(pt, crv):
    try:
        ir = crv.Project(pt)
        if ir is None:
            return 1e30
        return pt.DistanceTo(ir.XYZPoint)
    except Exception:
        return 1e30


def _cerca_de_extremo_muro(e0, e1, other_curve, tol):
    try:
        oe0 = other_curve.GetEndPoint(0)
        oe1 = other_curve.GetEndPoint(1)
    except Exception:
        return False
    for p in (oe0, oe1):
        if e0.DistanceTo(p) <= tol or e1.DistanceTo(p) <= tol:
            return True
    if _dist_point_to_curve(e0, other_curve) <= tol:
        return True
    if _dist_point_to_curve(e1, other_curve) <= tol:
        return True
    return False


def _min_dist_point_aabb(p, bb):
    """Distancia mínima de un punto a una caja alineada a ejes (pies internos)."""
    if p is None or bb is None:
        return 1e30
    try:
        cx = max(bb.Min.X, min(p.X, bb.Max.X))
        cy = max(bb.Min.Y, min(p.Y, bb.Max.Y))
        cz = max(bb.Min.Z, min(p.Z, bb.Max.Z))
        dx = p.X - cx
        dy = p.Y - cy
        dz = p.Z - cz
        return (dx * dx + dy * dy + dz * dz) ** 0.5
    except Exception:
        return 1e30


def _tol_extremo_curva_muros(host, other, tol_base):
    """
    Tolerancia en esquinas / uniones Revit: los ejes rara vez coinciden en <80 mm
    si hay limpieza de unión o muros gruesos; se escala con el ancho.
    """
    w0 = 0.0
    w1 = 0.0
    try:
        w0 = abs(float(host.Width))
    except Exception:
        pass
    try:
        w1 = abs(float(other.Width))
    except Exception:
        pass
    return max(float(tol_base), (w0 + w1) * 0.42 + float(tol_base) * 0.5)


def _extremo_cerca_bbox_muro(e0, e1, other_wall, tol):
    """Respaldo si la curva falla: el volumen del otro muro queda cerca de un extremo del eje."""
    bb = other_wall.get_BoundingBox(None)
    if bb is None:
        return False
    d0 = _min_dist_point_aabb(e0, bb)
    d1 = _min_dist_point_aabb(e1, bb)
    return min(d0, d1) <= tol


def _param_01_y_sep_eje_muro(e0, e1, p):
    """
    Proyección de ``p`` sobre la recta e0--e1: ``t`` = fracción a lo largo del tramo
    (0 = e0, 1 = e1) sin clamp; ``sep`` = distancia perpendicular a esa recta.
    """
    try:
        sx = e1.X - e0.X
        sy = e1.Y - e0.Y
        sz = e1.Z - e0.Z
        wx = p.X - e0.X
        wy = p.Y - e0.Y
        wz = p.Z - e0.Z
        den = sx * sx + sy * sy + sz * sz
        if den < 1e-20:
            return 0.5, 1e30
        t = (wx * sx + wy * sy + wz * sz) / den
        cx = e0.X + t * sx
        cy = e0.Y + t * sy
        cz = e0.Z + t * sz
        sep = ((p.X - cx) ** 2 + (p.Y - cy) ** 2 + (p.Z - cz) ** 2) ** 0.5
        return t, sep
    except Exception:
        return 0.5, 1e30


def _es_muro_encuentro_cara_lateral_o_t(host, base_line, other_wall, oc, tol_end):
    """
    Encuentro en **T** o en la **cara lateral** a lo largo del tramo: el eje del otro muro
    (o sus extremos) queda a distancia de orden de los anchos del eje anfitrión, con
    proyección **interior** (no en los 5 % iniciales/finales), y casi **ortogonal** al eje.
    Así se incluyen muros con unión solo en cara o sin solape de volumen API, que antes excluía
    el filtro de “solo extremos e0/e1”.

    Excluye la unión larga a lo largo de la misma hoja (reutiliza ``_es_union_longitudinal_paralelo``).
    """
    if other_wall is None or not isinstance(other_wall, Wall) or base_line is None or oc is None:
        return False
    if _es_union_longitudinal_paralelo(host, base_line, oc, tol_end):
        return False
    try:
        bd = _xyz_normalize(base_line.Direction)
    except Exception:
        bd = None
    if bd is None:
        return False
    try:
        od = _tangent_at_mid(oc)
    except Exception:
        od = None
    if od is None:
        try:
            od = _xyz_normalize(
                Line.CreateBound(oc.GetEndPoint(0), oc.GetEndPoint(1)).Direction
            )
        except Exception:
            return False
    try:
        if abs(float(od.DotProduct(bd))) > 0.92:
            return False
    except Exception:
        return False
    e0 = base_line.GetEndPoint(0)
    e1 = base_line.GetEndPoint(1)
    try:
        seg_len = e0.DistanceTo(e1)
    except Exception:
        seg_len = 0.0
    if seg_len < 1e-9:
        return False
    wh = 0.0
    wo = 0.0
    try:
        wh = abs(float(host.Width))
    except Exception:
        pass
    try:
        wo = abs(float(other_wall.Width))
    except Exception:
        pass
    try:
        m150 = UnitUtils.ConvertToInternalUnits(150.0, UnitTypeId.Millimeters)
    except Exception:
        m150 = 0.5
    # Contacto en cara: los ejes suelen desalinearse ~ la suma de medios espesores.
    lim_close = max(0.55 * (wh + wo) + 2.0 * float(tol_end), m150)
    try:
        par_lim = 0.04
        if seg_len < UnitUtils.ConvertToInternalUnits(0.45, UnitTypeId.Meters):
            par_lim = 0.12
    except Exception:
        par_lim = 0.05

    puntos = []
    try:
        puntos.append(_midpoint_curve(oc))
    except Exception:
        pass
    try:
        puntos.append(oc.GetEndPoint(0))
    except Exception:
        pass
    try:
        puntos.append(oc.GetEndPoint(1))
    except Exception:
        pass
    for om in puntos:
        if om is None:
            continue
        t, sep = _param_01_y_sep_eje_muro(e0, e1, om)
        if par_lim < t < 1.0 - par_lim and sep <= lim_close:
            return True
    return False


def _es_union_longitudinal_paralelo(host, base_line, oc, tol_end):
    """
    Muro casi paralelo al anfitrión cuyo eje cae cerca del tramo interior:
    unión a lo largo de la cara larga (no cap en inicio/fin).
    """
    try:
        bd = _xyz_normalize(base_line.Direction)
        om = _midpoint_curve(oc)
        if bd is None or om is None:
            return False
        r = base_line.Project(om)
        if r is None:
            return False
        par = float(r.Parameter)
        sep = om.DistanceTo(r.XYZPoint)
        try:
            w = abs(float(host.Width))
        except Exception:
            w = 0.0
        lim = max(w * 2.0, float(tol_end) * 3.0)
        od = _tangent_at_mid(oc)
        if od is not None and abs(float(od.DotProduct(bd))) > 0.92:
            if 0.05 < par < 0.95 and sep < lim:
                return True
    except Exception:
        pass
    return False


def _esta_unido_por_join_geometry(doc, host, other):
    if doc is None or host is None or other is None:
        return False
    try:
        return JoinGeometryUtils.AreElementsJoined(doc, host, other)
    except Exception:
        try:
            return JoinGeometryUtils.AreElementsJoined(doc, host.Id, other.Id)
        except Exception:
            return False


def _es_muro_lateral_en_extremos(doc, host, base_line, other_wall, tol_end):
    """
    Muro “en encuentro” respecto al anfitrión:

    - Si ``AreElementsJoined``: se acepta salvo unión longitudinal paralela (cara larga).
    - Si no está unido: proximidad de curva o de bbox a e0/e1 con tolerancia ampliada, **o**
      encuentro en **T / cara lateral** a mitad de tramo (eje del otro casi ortogonal, proyección
      interior del eje, separación de orden de los espesores).
    - ``ElementIntersectsElementFilter`` a veces excluye muros que solo comparten cara/planar
      sin solape de volumen; el barrido por bbox y la heurística T compensan.
    """
    if other_wall is None or not isinstance(other_wall, Wall):
        return False
    if _element_ids_equal(other_wall.Id, host.Id):
        return False
    ol = other_wall.Location
    if not isinstance(ol, LocationCurve):
        return False
    oc = ol.Curve
    if oc is None:
        return False
    e0 = base_line.GetEndPoint(0)
    e1 = base_line.GetEndPoint(1)
    tol_curve = _tol_extremo_curva_muros(host, other_wall, tol_end)
    tol_bbox = max(tol_curve, float(tol_end) * 2.2)

    if _esta_unido_por_join_geometry(doc, host, other_wall):
        if _es_union_longitudinal_paralelo(host, base_line, oc, tol_end):
            return False
        return True

    prox = _cerca_de_extremo_muro(e0, e1, oc, tol_curve) or _extremo_cerca_bbox_muro(
        e0, e1, other_wall, tol_bbox
    )
    if not prox:
        prox = _es_muro_encuentro_cara_lateral_o_t(host, base_line, other_wall, oc, tol_end)
    if not prox:
        return False
    if _es_union_longitudinal_paralelo(host, base_line, oc, tol_end):
        return False
    return True


def _outline_host_inflado(host, pad_internal):
    bb = host.get_BoundingBox(None)
    if bb is None:
        return None
    try:
        p = float(pad_internal)
        mn = XYZ(bb.Min.X - p, bb.Min.Y - p, bb.Min.Z - p)
        mx = XYZ(bb.Max.X + p, bb.Max.Y + p, bb.Max.Z + p)
        return Outline(mn, mx)
    except Exception:
        return None


def _es_fundacion_estructural(element):
    """Categoría Fundación estructural (zapatas, muros de cimentación, etc.)."""
    if element is None:
        return False
    try:
        cat = element.Category
        if cat is None:
            return False
        cid = _element_id_to_int(cat.Id)
        if cid is None:
            return False
        return cid == int(BuiltInCategory.OST_StructuralFoundation)
    except Exception:
        return False


def _es_losa_cimentacion_por_tipo(doc, element):
    """``Floor`` cuyo ``FloorType`` es losa de cimentación (independiente de la categoría mostrada)."""
    if doc is None or element is None or not isinstance(element, Floor):
        return False
    try:
        tid = element.GetTypeId()
        if tid is None or tid == ElementId.InvalidElementId:
            return False
        ft = doc.GetElement(tid)
        if ft is None:
            return False
        try:
            if isinstance(ft, FloorType):
                return bool(ft.IsFoundationSlab)
        except Exception:
            pass
        return bool(getattr(ft, "IsFoundationSlab", False))
    except Exception:
        return False


def _excluir_de_union_fundaciones(doc, element):
    """
    No unir al sólido del nudo: fundaciones estructurales ni losas de cimentación.
    Aplica a todo elemento unido o candidato (no solo ``Floor``).
    """
    if element is None:
        return False
    if _es_fundacion_estructural(element):
        return True
    if _es_losa_cimentacion_por_tipo(doc, element):
        return True
    return False


def _es_fundacion_para_estirar_boceto(doc, element):
    """Fundación unida al muro: categoría estructural o losa de cimentación (para estirar el boceto)."""
    if element is None:
        return False
    if _es_fundacion_estructural(element):
        return True
    if _es_losa_cimentacion_por_tipo(doc, element):
        return True
    return False


def _altura_bbox_elemento(elem):
    if elem is None:
        return 0.0
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return 0.0
        return float(bb.Max.Z - bb.Min.Z)
    except Exception:
        return 0.0


def _altura_max_fundacion_unida_host(doc, host):
    """
    Máxima altura de bbox entre fundaciones / losas de cimentación unidas por *Join Geometry*
    únicamente al muro seleccionado.
    """
    if doc is None or host is None:
        return 0.0
    h_max = 0.0
    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or not _es_fundacion_para_estirar_boceto(doc, el):
            continue
        h = _altura_bbox_elemento(el)
        if h > h_max:
            h_max = h
    return h_max


def _vector_abajo_en_plano_boceto(plane):
    """Proyección de −Z sobre el plano del boceto (``plane.Normal``)."""
    if plane is None:
        return None
    try:
        n = _xyz_normalize(plane.Normal)
    except Exception:
        n = None
    if n is None:
        return None
    g = XYZ(0.0, 0.0, -1.0)
    try:
        d = float(n.DotProduct(g))
        vx = g.X - n.X * d
        vy = g.Y - n.Y * d
        vz = g.Z - n.Z * d
        ln = (vx * vx + vy * vy + vz * vz) ** 0.5
        if ln < 1e-12:
            return XYZ(0.0, 0.0, -1.0)
        return XYZ(vx / ln, vy / ln, vz / ln)
    except Exception:
        return None


def _traslacion_extender_boceto_fundacion(plane_face, altura):
    if altura <= 1e-12 or plane_face is None:
        return None
    v = _vector_abajo_en_plano_boceto(plane_face)
    if v is None:
        return XYZ(0.0, 0.0, -altura)
    try:
        return XYZ(v.X * altura, v.Y * altura, v.Z * altura)
    except Exception:
        return XYZ(0.0, 0.0, -altura)


def _extender_curvas_borde_inferior_fundacion(plane_face, curves, h_fund):
    """
    Desplaza solo el borde inferior del contorno (Z mínima) en sentido “hacia fundación”
    en el plano del boceto, una distancia igual a la altura de bbox de la fundación unida.
    """
    if not curves or h_fund <= 1e-12 or plane_face is None:
        return curves
    tvec = _traslacion_extender_boceto_fundacion(plane_face, h_fund)
    if tvec is None:
        return curves
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        z_tol = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)
    except Exception:
        z_tol = 1e-4
    pts = []
    for cv in curves:
        try:
            pts.append(cv.GetEndPoint(0))
            pts.append(cv.GetEndPoint(1))
        except Exception:
            pass
    if not pts:
        return curves
    z_min = min(p.Z for p in pts)

    def _punto_es_borde_inferior(p):
        try:
            return float(p.Z) <= z_min + z_tol
        except Exception:
            return False

    out = []
    for cv in curves:
        try:
            p0 = cv.GetEndPoint(0)
            p1 = cv.GetEndPoint(1)
        except Exception:
            out.append(cv)
            continue
        if _is_line_curve(cv):
            q0 = p0.Add(tvec) if _punto_es_borde_inferior(p0) else p0
            q1 = p1.Add(tvec) if _punto_es_borde_inferior(p1) else p1
            try:
                if q0.DistanceTo(p0) > z_tol or q1.DistanceTo(p1) > z_tol:
                    out.append(Line.CreateBound(q0, q1))
                else:
                    out.append(cv)
            except Exception:
                out.append(Line.CreateBound(q0, q1))
            continue
        if _is_arc_curve(cv):
            try:
                if _punto_es_borde_inferior(p0) and _punto_es_borde_inferior(p1):
                    tr = Transform.CreateTranslation(tvec)
                    c2 = cv.CreateTransformed(tr)
                    if c2 is not None:
                        out.append(c2)
                        continue
            except Exception:
                pass
            out.append(cv)
            continue
        out.append(cv)
    return out


def _offsets_boceto_mm_fund_y_resto():
    """Distancias internas en pies: 50 mm fundación, 25 mm resto."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        o_f = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters)
        o_r = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters)
        return o_f, o_r
    except Exception:
        return 50.0 / 304.8, 25.0 / 304.8


def _punto_referencia_interior_boceto(plane_face, curves):
    """Baricentro de puntos medios de las curvas, proyectado al plano del boceto (lado 'interior')."""
    if plane_face is None or not curves:
        return None
    pts = []
    for cv in curves:
        if cv is None:
            continue
        try:
            pts.append(cv.Evaluate(0.5, True))
        except Exception:
            try:
                pts.append(cv.GetEndPoint(0))
            except Exception:
                pass
    if not pts:
        return None
    n = float(len(pts))
    c = XYZ(sum(p.X for p in pts) / n, sum(p.Y for p in pts) / n, sum(p.Z for p in pts) / n)
    return _punto_en_plano(c, plane_face)


def _normal_offset_hacia_interior(curve, plane_normal, inside_ref):
    """
    Vector unitario en el plano, perpendicular a la tangente, orientado hacia el interior (inside_ref).
    """
    if curve is None or inside_ref is None:
        return None
    t = _tangent_at_mid(curve)
    if t is None:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            t = _xyz_normalize(p1.Subtract(p0))
        except Exception:
            t = None
    if t is None:
        return None
    n = _xyz_normalize(plane_normal)
    if n is None:
        return None
    c1 = _xyz_normalize(n.CrossProduct(t))
    if c1 is None:
        return None
    c2 = XYZ(-c1.X, -c1.Y, -c1.Z)
    try:
        pm = curve.Evaluate(0.5, True)
    except Exception:
        try:
            pm = curve.GetEndPoint(0)
        except Exception:
            return c1
    try:
        v = inside_ref.Subtract(pm)
        if float(v.DotProduct(c1)) >= float(v.DotProduct(c2)):
            return c1
        return c2
    except Exception:
        return c1


def _curva_relacionada_fundacion(cv, todas_curvas, h_fund):
    """
    Tramo ligado a la zona de fundación: solo si hay fundación unida (``h_fund``)
    y el tramo está en la cota mínima del boceto (borde inferior).
    """
    if cv is None or h_fund <= 1e-9 or not todas_curvas:
        return False
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        z_tol = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters)
    except Exception:
        z_tol = 0.01
    pts = []
    for c in todas_curvas:
        if c is None:
            continue
        try:
            pts.append(c.GetEndPoint(0))
            pts.append(c.GetEndPoint(1))
        except Exception:
            pass
    if not pts:
        return False
    z_min = min(float(p.Z) for p in pts)
    try:
        pm = cv.Evaluate(0.5, True)
    except Exception:
        return False
    return float(pm.Z) <= z_min + z_tol


def _indice_arista_curva_en_lineas(curves, cv, tol_cierre):
    """Índice de la arista en ``curves`` que coincide con ``cv`` (referencia o extremos)."""
    if cv is None or not curves:
        return None
    for i, c in enumerate(curves):
        if c is cv:
            return i
    try:
        p0 = cv.GetEndPoint(0)
        p1 = cv.GetEndPoint(1)
    except Exception:
        return None
    for i, c in enumerate(curves):
        if c is None:
            continue
        try:
            q0 = c.GetEndPoint(0)
            q1 = c.GetEndPoint(1)
        except Exception:
            continue
        if (p0.DistanceTo(q0) < tol_cierre and p1.DistanceTo(q1) < tol_cierre) or (
            p0.DistanceTo(q1) < tol_cierre and p1.DistanceTo(q0) < tol_cierre
        ):
            return i
    return None


def _aplicar_offsets_interiores_createoffset(plane_face, curves, h_fund):
    """
    Respaldo: ``Curve.CreateOffset`` (a menudo falla o no mueve curvas procedentes de sólidos en RPS).
    Si las curvas forman un polígono cerrado de líneas, la normal de offset usa el **sentido del polígono**
    (igual que ``_offset_poligono_lineas_cerrado``); si no, respaldo al baricentro de puntos medios.
    """
    if not curves or plane_face is None:
        return curves
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        tol_cierre = UnitUtils.ConvertToInternalUnits(6.0, UnitTypeId.Millimeters)
    except Exception:
        tol_cierre = 0.01
    V = _vertices_poligono_cerrado_desde_lineas(curves, tol_cierre)
    signed_area = _signed_area_2d_polygon(plane_face, V) if V is not None else 0.0
    inside_ref = _punto_referencia_interior_boceto(plane_face, curves)
    if inside_ref is None and abs(signed_area) < 1e-12:
        return curves
    off_fund, off_resto = _offsets_boceto_mm_fund_y_resto()
    out = []
    for cv in curves:
        if cv is None:
            continue
        dist = off_fund if _curva_relacionada_fundacion(cv, curves, h_fund) else off_resto
        off_n = None
        if V is not None and abs(signed_area) >= 1e-12:
            idx = _indice_arista_curva_en_lineas(curves, cv, tol_cierre)
            if idx is not None:
                off_n = _inward_perp_unit_3d_edge_from_winding(
                    plane_face, V, idx, signed_area
                )
        if off_n is None:
            if inside_ref is None:
                out.append(cv)
                continue
            off_n = _normal_offset_hacia_interior(cv, plane_face.Normal, inside_ref)
        if off_n is None or dist <= 1e-12:
            out.append(cv)
            continue
        new_cv = None
        for d_try, n_try in (
            (dist, off_n),
            (-dist, off_n),
            (dist, XYZ(-off_n.X, -off_n.Y, -off_n.Z)),
            (-dist, XYZ(-off_n.X, -off_n.Y, -off_n.Z)),
        ):
            try:
                new_cv = cv.CreateOffset(d_try, n_try)
                if new_cv is not None:
                    break
            except Exception:
                new_cv = None
        out.append(new_cv if new_cv is not None else cv)
    return out


def _base_ortonormal_plano(plane_face):
    if plane_face is None:
        return None, None, None
    n = _xyz_normalize(plane_face.Normal)
    o = plane_face.Origin
    if n is None:
        return None, None, None
    try:
        if abs(float(n.DotProduct(XYZ.BasisZ))) > 0.99:
            ux = _xyz_normalize(n.CrossProduct(XYZ.BasisX))
        else:
            ux = _xyz_normalize(n.CrossProduct(XYZ.BasisZ))
    except Exception:
        ux = None
    if ux is None:
        return None, None, None
    uy = _xyz_normalize(n.CrossProduct(ux))
    if uy is None:
        return None, None, None
    return o, ux, uy


def _punto_2d_en_plano(plane_face, pt):
    """Proyección de ``pt`` en coordenadas (u,v) del plano del boceto."""
    o, ux, uy = _base_ortonormal_plano(plane_face)
    if o is None or pt is None:
        return None
    try:
        q = pt.Subtract(o)
        return (float(q.DotProduct(ux)), float(q.DotProduct(uy)))
    except Exception:
        return None


def _signed_area_2d_polygon(plane_face, V):
    """Área firmada en 2D (fórmula de la cuerda). Signo >0 = CCW en (ux, uy)."""
    if not V or len(V) < 3:
        return 0.0
    pts = []
    for p in V:
        t = _punto_2d_en_plano(plane_face, p)
        if t is None:
            return 0.0
        pts.append(t)
    n = len(pts)
    s = 0.0
    for i in range(n):
        j = (i + 1) % n
        s += pts[i][0] * pts[j][1] - pts[j][0] * pts[i][1]
    return 0.5 * s


def _inward_perp_unit_3d_edge_from_winding(plane_face, V, edge_index, signed_area):
    """
    Perpendicular unitaria al arista ``V[i] -> V[i+1]`` en el plano, hacia el **interior**
    del polígono según el sentido del contorno (área firmada).
    """
    if not V or len(V) < 3:
        return None
    n = len(V)
    o, ux, uy = _base_ortonormal_plano(plane_face)
    if o is None:
        return None
    p0 = V[edge_index]
    p1 = V[(edge_index + 1) % n]
    try:
        du = float(p1.Subtract(p0).DotProduct(ux))
        dv = float(p1.Subtract(p0).DotProduct(uy))
    except Exception:
        return None
    len_e = math.sqrt(du * du + dv * dv)
    if len_e < 1e-12:
        return None
    if abs(signed_area) < 1e-12:
        return None
    if signed_area > 0.0:
        nu = -dv / len_e
        nv = du / len_e
    else:
        nu = dv / len_e
        nv = -du / len_e
    try:
        return _xyz_normalize(ux.Multiply(nu).Add(uy.Multiply(nv)))
    except Exception:
        return None


def _intersect_lines_in_plane(o1, d1, o2, d2, plane_face):
    """Intersección de dos rectas coplanares al plano del boceto (d1,d2 no necesariamente unitarios)."""
    if o1 is None or d1 is None or o2 is None or d2 is None:
        return None
    o, ux, uy = _base_ortonormal_plano(plane_face)
    if o is None:
        return None

    def to2d(pt):
        q = pt.Subtract(o)
        return (float(q.DotProduct(ux)), float(q.DotProduct(uy)))

    def dir2d(dv):
        return (float(dv.DotProduct(ux)), float(dv.DotProduct(uy)))

    p1 = to2d(o1)
    t1 = dir2d(d1)
    p2 = to2d(o2)
    t2 = dir2d(d2)
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    det = t1[0] * t2[1] - t1[1] * t2[0]
    if abs(det) < 1e-14:
        return None
    u = (dx * t2[1] - dy * t2[0]) / det
    x = p1[0] + u * t1[0]
    y = p1[1] + u * t1[1]
    try:
        return o.Add(ux.Multiply(x)).Add(uy.Multiply(y))
    except Exception:
        return None


def _vertices_poligono_cerrado_desde_lineas(curves, tol_cierre):
    """Vértices V[0..n-1] si todas las curvas son ``Line`` y forman anillo cerrado."""
    if not curves or len(curves) < 3:
        return None
    for cv in curves:
        if cv is None or not _is_line_curve(cv):
            return None
    n = len(curves)
    V = []
    for i in range(n):
        V.append(curves[i].GetEndPoint(0))
    p_cierre = curves[n - 1].GetEndPoint(1)
    if p_cierre.DistanceTo(V[0]) > tol_cierre:
        return None
    for i in range(n):
        p1 = curves[i].GetEndPoint(1)
        p0n = curves[(i + 1) % n].GetEndPoint(0)
        if p1.DistanceTo(p0n) > tol_cierre:
            return None
    return V


def _offset_poligono_lineas_cerrado(plane_face, curves, h_fund):
    """
    Offset interior por aristas (50 / 25 mm): rectas paralelas + intersección en vértices.
    Evita depender de ``CreateOffset`` sobre geometría de sólido.
    """
    if not curves or plane_face is None:
        return None
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        tol_cierre = UnitUtils.ConvertToInternalUnits(6.0, UnitTypeId.Millimeters)
        z_tol = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters)
    except Exception:
        tol_cierre = 0.01
        z_tol = 0.01
    V = _vertices_poligono_cerrado_desde_lineas(curves, tol_cierre)
    if V is None:
        return None
    n = len(V)
    z_min = min(float(p.Z) for p in V)
    off_fund, off_resto = _offsets_boceto_mm_fund_y_resto()
    signed_area = _signed_area_2d_polygon(plane_face, V)
    inside_ref = _punto_referencia_interior_boceto(plane_face, curves)
    if inside_ref is None and abs(signed_area) < 1e-12:
        return None
    Npl = _xyz_normalize(plane_face.Normal)
    if Npl is None:
        return None
    origins = []
    dirs = []
    for i in range(n):
        p0 = V[i]
        p1 = V[(i + 1) % n]
        try:
            T = _xyz_normalize(p1.Subtract(p0))
        except Exception:
            T = None
        if T is None:
            return None
        pm = XYZ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5, (p0.Z + p1.Z) * 0.5)
        Left = None
        if abs(signed_area) >= 1e-12:
            Left = _inward_perp_unit_3d_edge_from_winding(plane_face, V, i, signed_area)
        if Left is None:
            if inside_ref is None:
                return None
            Left = _xyz_normalize(Npl.CrossProduct(T))
            if Left is None:
                return None
            try:
                v_in = inside_ref.Subtract(pm)
                if float(v_in.DotProduct(Left)) < 0.0:
                    Left = XYZ(-Left.X, -Left.Y, -Left.Z)
            except Exception:
                pass
        zm = (float(p0.Z) + float(p1.Z)) * 0.5
        dist = (
            off_fund
            if (h_fund > 1e-9 and zm <= z_min + z_tol)
            else off_resto
        )
        try:
            Oi = p0.Add(Left.Multiply(dist))
        except Exception:
            return None
        origins.append(Oi)
        dirs.append(T)
    Q = []
    for i in range(n):
        pt = _intersect_lines_in_plane(
            origins[(i - 1 + n) % n],
            dirs[(i - 1 + n) % n],
            origins[i],
            dirs[i],
            plane_face,
        )
        if pt is None:
            return None
        Q.append(pt)
    out = []
    for i in range(n):
        a = Q[i]
        b = Q[(i + 1) % n]
        if a.DistanceTo(b) < 1e-9:
            continue
        try:
            out.append(Line.CreateBound(a, b))
        except Exception:
            return None
    return out if len(out) >= 3 else None


def _aplicar_offsets_interiores_boceto_loop(plane_face, curves, h_fund):
    """
    Por cada loop cerrado: intenta offset poligonal (solo ``Line``); si no, ``CreateOffset``.
    """
    if not curves or plane_face is None:
        return curves
    poly = _offset_poligono_lineas_cerrado(plane_face, curves, h_fund)
    if poly is not None:
        return poly
    return _aplicar_offsets_interiores_createoffset(plane_face, curves, h_fund)


def _es_suelo_banda_superior_inferior(doc, host, floor, tol_z):
    if floor is None or not isinstance(floor, Floor):
        return False
    if _excluir_de_union_fundaciones(doc, floor):
        return False
    try:
        hbb = host.get_BoundingBox(None)
        fbb = floor.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or fbb is None:
        return False
    band = max(float(tol_z), 1e-4)
    touches_top = (fbb.Min.Z <= hbb.Max.Z + band) and (fbb.Max.Z >= hbb.Max.Z - band)
    touches_bot = (fbb.Max.Z >= hbb.Min.Z - band) and (fbb.Min.Z <= hbb.Min.Z + band)
    return touches_top or touches_bot


def _bbox_solape_xy_mn_mx(mn_a, mx_a, mn_b, mx_b, tol):
    """Solape en planta (X/Y) entre dos cajas alineadas a ejes."""
    if mn_a is None or mx_a is None or mn_b is None or mx_b is None:
        return False
    try:
        t = float(tol)
        if mx_a.X + t < mn_b.X - t:
            return False
        if mx_b.X + t < mn_a.X - t:
            return False
        if mx_a.Y + t < mn_b.Y - t:
            return False
        if mx_b.Y + t < mn_a.Y - t:
            return False
        return True
    except Exception:
        return False


def _bbox_intersecta_plano(bb, plane):
    """True si el bbox cruza el plano infinito (vértices a ambos lados o sobre el plano)."""
    if bb is None or plane is None:
        return False
    try:
        o = plane.Origin
        n = plane.Normal
        vals = []
        for ix in (0, 1):
            for iy in (0, 1):
                for iz in (0, 1):
                    x = bb.Min.X if ix == 0 else bb.Max.X
                    y = bb.Min.Y if iy == 0 else bb.Max.Y
                    z = bb.Min.Z if iz == 0 else bb.Max.Z
                    p = XYZ(x, y, z)
                    vals.append(float(p.Subtract(o).DotProduct(n)))
        return min(vals) <= 1e-5 and max(vals) >= -1e-5
    except Exception:
        return False


def _contacto_cara_superior_inferior_estricto(host, floor, tol_z):
    """
    Losa en **cara superior o inferior** del muro: solape XY y caras horizontales casi coincidentes.
    """
    if floor is None or not isinstance(floor, Floor):
        return False
    try:
        hbb = host.get_BoundingBox(None)
        fbb = floor.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or fbb is None:
        return False
    try:
        tol_xy = max(float(tol_z), UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
    except Exception:
        tol_xy = 0.15
    if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, fbb.Min, fbb.Max, tol_xy):
        return False
    try:
        d_face = UnitUtils.ConvertToInternalUnits(4.0, UnitTypeId.Millimeters)
    except Exception:
        d_face = 0.01
    band = max(float(tol_z), 1e-4)
    if abs(fbb.Min.Z - hbb.Max.Z) <= d_face + band:
        return True
    if abs(fbb.Max.Z - hbb.Min.Z) <= d_face + band:
        return True
    return False


def _suelo_sobre_cara_superior_muro(host, floor, tol_z):
    """
    ``Floor`` cuya cara inferior se apoya en la **cara superior** del muro
    (forjado encima, sin incluir suelo en la inferior del muro).
    Misma lógica XY que ``_contacto_cara_superior_inferior_estricto``, solo rama Z superior.
    """
    if floor is None or not isinstance(floor, Floor) or not isinstance(host, Wall):
        return False
    try:
        hbb = host.get_BoundingBox(None)
        fbb = floor.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or fbb is None:
        return False
    try:
        tol_z = float(tol_z)
        tol_xy = max(tol_z, UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
    except Exception:
        tol_z = 0.12
        tol_xy = 0.15
    if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, fbb.Min, fbb.Max, tol_xy):
        return False
    try:
        d_face = UnitUtils.ConvertToInternalUnits(4.0, UnitTypeId.Millimeters)
    except Exception:
        d_face = 0.01
    band = max(float(tol_z), 1e-4)
    return abs(fbb.Min.Z - hbb.Max.Z) <= d_face + band


def _muro_apilado_bajo_muro_principal(host, other, tol_z):
    u"""
    True si ``other`` es un muro cuyo techo coincide con la base de ``host`` (apilación vertical),
    con solape en planta: no debe formar parte de la unión booleana del boceto por la *cara inferior*.
    """
    if not isinstance(host, Wall) or not isinstance(other, Wall):
        return False
    try:
        if other.Id == host.Id:
            return False
    except Exception:
        pass
    try:
        hbb = host.get_BoundingBox(None)
        obb = other.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or obb is None:
        return False
    try:
        d_face = UnitUtils.ConvertToInternalUnits(4.0, UnitTypeId.Millimeters)
        band = max(float(tol_z), 1e-4)
        tol_xy = max(float(tol_z), UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
    except Exception:
        d_face = 0.01
        band = 0.12
        tol_xy = 0.15
    if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, obb.Min, obb.Max, tol_xy):
        return False
    if abs(obb.Max.Z - hbb.Min.Z) > d_face + band:
        return False
    if obb.Min.Z >= hbb.Min.Z - 1e-5:
        return False
    return True


def _muro_apilado_sobre_muro_principal(host, other, tol_z):
    u"""
    True si ``other`` es un muro cuyo piso coincide con la cabeza de ``host`` (apilación vertical),
    con solape en planta. Se usa para excluirlo de la unión booleana del boceto por la *cara superior*.
    """
    if not isinstance(host, Wall) or not isinstance(other, Wall):
        return False
    try:
        if other.Id == host.Id:
            return False
    except Exception:
        pass
    try:
        hbb = host.get_BoundingBox(None)
        obb = other.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or obb is None:
        return False
    try:
        d_face = UnitUtils.ConvertToInternalUnits(4.0, UnitTypeId.Millimeters)
        band = max(float(tol_z), 1e-4)
        tol_xy = max(float(tol_z), UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
    except Exception:
        d_face = 0.01
        band = 0.12
        tol_xy = 0.15
    if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, obb.Min, obb.Max, tol_xy):
        return False
    if abs(obb.Min.Z - hbb.Max.Z) > d_face + band:
        return False
    if obb.Max.Z <= hbb.Max.Z + 1e-5:
        return False
    return True


def _suelo_bajo_cara_inferior_muro(host, floor, tol_z):
    """
    ``Floor`` cuya cara superior se apoya en la **cara inferior** del muro
    (losa o fundación debajo, solape en planta y contacto Z en la base).
    """
    if floor is None or not isinstance(floor, Floor) or not isinstance(host, Wall):
        return False
    try:
        hbb = host.get_BoundingBox(None)
        fbb = floor.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or fbb is None:
        return False
    try:
        tol_z = float(tol_z)
        tol_xy = max(tol_z, UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
    except Exception:
        tol_z = 0.12
        tol_xy = 0.15
    if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, fbb.Min, fbb.Max, tol_xy):
        return False
    try:
        d_face = UnitUtils.ConvertToInternalUnits(4.0, UnitTypeId.Millimeters)
    except Exception:
        d_face = 0.01
    band = max(float(tol_z), 1e-4)
    return abs(fbb.Max.Z - hbb.Min.Z) <= d_face + band


def _incluir_suelo_unido_cara_superior_inferior(doc, host, floor, tol_z):
    """
    ``Floor`` unido al muro por *Join Geometry* y situado en cara superior o inferior.
    **Solo inferior:** para la unión de sólidos del boceto solo cuenta fundación estructural o
    losa de cimentación; forjado que no es fundación bajo el muro no se une.
    """
    if not isinstance(floor, Floor):
        return False
    try:
        es_sup = _suelo_sobre_cara_superior_muro(host, floor, tol_z)
        es_inf = _suelo_bajo_cara_inferior_muro(host, floor, tol_z)
    except Exception:
        es_sup = es_inf = False
    # Regla BIMTools: si hay elemento unido en cara superior del muro, NO entra en la unión
    # booleana para obtener el boceto (no recorta/define el nudo en alzado).
    if es_sup:
        return False
    if es_inf and not es_sup:
        return _es_fundacion_para_estirar_boceto(doc, floor)
    if _excluir_de_union_fundaciones(doc, floor):
        return False
    if _contacto_cara_superior_inferior_estricto(host, floor, tol_z):
        return True
    try:
        hbb = host.get_BoundingBox(None)
        fbb = floor.get_BoundingBox(None)
    except Exception:
        return False
    if hbb is None or fbb is None:
        return False
    try:
        tol_xy = max(float(tol_z), UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
    except Exception:
        tol_xy = 0.15
    if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, fbb.Min, fbb.Max, tol_xy):
        return False
    if not _es_suelo_banda_superior_inferior(doc, host, floor, tol_z):
        return False
    try:
        band = max(float(tol_z), 1e-4)
    except Exception:
        band = 0.12
    touches_top = (fbb.Min.Z <= hbb.Max.Z + band) and (fbb.Max.Z >= hbb.Max.Z - band)
    touches_bot = (fbb.Max.Z >= hbb.Min.Z - band) and (fbb.Min.Z <= hbb.Min.Z + band)
    if touches_top:
        return False
    if touches_bot and not touches_top:
        return _es_fundacion_para_estirar_boceto(doc, floor)
    return True


def _suelo_para_union_intersect_filter(doc, host, floor, tol_z, section_plane):
    """
    Criterio para suelos hallados con ``ElementIntersectsElementFilter(host)``:
    banda superior/inferior, o corte del plano de sección (mismo criterio que suelos extra por plano).
    **Solo cara inferior** del muro: misma regla que ``_incluir_suelo_unido_cara_superior_inferior`` (solo fundación).
    """
    if not isinstance(floor, Floor):
        return False
    try:
        es_sup = _suelo_sobre_cara_superior_muro(host, floor, tol_z)
        es_inf = _suelo_bajo_cara_inferior_muro(host, floor, tol_z)
    except Exception:
        es_sup = es_inf = False
    if es_sup:
        return False
    if es_inf and not es_sup:
        if not _es_fundacion_para_estirar_boceto(doc, floor):
            return False
    elif _excluir_de_union_fundaciones(doc, floor):
        return False
    if _es_suelo_banda_superior_inferior(doc, host, floor, tol_z):
        return True
    if section_plane is None:
        return False
    try:
        fbb = floor.get_BoundingBox(None)
    except Exception:
        fbb = None
    if fbb is None:
        return False
    return _bbox_intersecta_plano(fbb, section_plane)


def _agregar_floors_intersectan_plano_corte(doc, host, elementos, seen, section_plane, tol_z, pad_bb):
    """
    Suelos adicionales cuyo bbox **corta el plano de sección** y se solapan en planta con el muro,
    para unirlos al resto de geometrías del corte (además de los ya detectados por unión/intersección).
    """
    if doc is None or host is None or section_plane is None:
        return
    ol = _outline_host_inflado(host, pad_bb)
    if ol is None:
        return
    try:
        bf = BoundingBoxIntersectsFilter(ol)
        coll = FilteredElementCollector(doc).OfClass(Floor).WherePasses(bf).ToElementIds()
    except Exception:
        return
    try:
        tol_xy = max(float(tol_z), UnitUtils.ConvertToInternalUnits(80.0, UnitTypeId.Millimeters))
    except Exception:
        tol_xy = 0.25
    for eid in coll:
        if eid in seen:
            continue
        fl = doc.GetElement(eid)
        if fl is None or not isinstance(fl, Floor):
            continue
        try:
            es_sup = _suelo_sobre_cara_superior_muro(host, fl, tol_z)
            es_inf = _suelo_bajo_cara_inferior_muro(host, fl, tol_z)
        except Exception:
            es_sup = es_inf = False
        if es_sup:
            continue
        if es_inf and not es_sup:
            if not _es_fundacion_para_estirar_boceto(doc, fl):
                continue
        elif _excluir_de_union_fundaciones(doc, fl):
            continue
        try:
            fbb = fl.get_BoundingBox(None)
            hbb = host.get_BoundingBox(None)
        except Exception:
            continue
        if fbb is None or hbb is None:
            continue
        if not _bbox_intersecta_plano(fbb, section_plane):
            continue
        if not _bbox_solape_xy_mn_mx(hbb.Min, hbb.Max, fbb.Min, fbb.Max, tol_xy):
            continue
        elementos.append(fl)
        seen.add(eid)


def _solidos_desde_elemento(element):
    u"""
    ``ViewDetailLevel.Fine`` añade carga (múrmuro con *openings* → muchas caras/fragmentos
    y uniones booleanas excesivas). En **muro** se usa *Medium* (suficiente para el nudo);
    resto de elementos, *Fine*.
    """
    opts = Options()
    opts.ComputeReferences = False
    try:
        if isinstance(element, Wall):
            opts.DetailLevel = ViewDetailLevel.Medium
        else:
            opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    try:
        geom = element.get_Geometry(opts)
    except Exception:
        return []
    if geom is None:
        return []
    out = []
    for obj in geom:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if obj.Volume > 1e-12:
                    out.append(obj)
            except Exception:
                pass
        elif isinstance(obj, GeometryInstance):
            inst = None
            try:
                inst = obj.GetInstanceGeometry(obj.Transform)
            except Exception:
                inst = None
            if inst is None:
                try:
                    inst = obj.GetInstanceGeometry()
                except Exception:
                    inst = None
            if inst is None:
                continue
            for g in inst:
                if isinstance(g, Solid):
                    try:
                        if g.Volume > 1e-12:
                            out.append(g)
                    except Exception:
                        pass
    return out


def _filtrar_solidos_utiles(solids):
    good = []
    for s in solids:
        if s is None:
            continue
        try:
            if float(s.Volume) <= 1e-12:
                continue
        except Exception:
            continue
        good.append(s)
    return good


def _solido_mayor_volumen(solids):
    """Si la unión booleana falla, usa el cuerpo principal (mayor volumen)."""
    good = _filtrar_solidos_utiles(solids)
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


def _unir_solidos_representantes_por_elemento(elementos):
    """
    Un cuerpo representativo (mayor volumen) por elemento, luego unión booleana.
    Si falla juntar *todos* los sólidos de instancia (muchos cuerpos), a menudo la unión
    de 2-3 cuerpos principales (muro+vecino+suelo) **sí** funciona e incluye el nudo en L.
    """
    reps = []
    for el in elementos:
        if el is None:
            continue
        sols = _filtrar_solidos_utiles(_solidos_desde_elemento(el))
        if not sols:
            continue
        r = _solido_mayor_volumen(sols)
        if r is not None:
            reps.append(r)
    if not reps:
        return None
    return _unir_solidos(reps)


def _unir_solidos(solids):
    solids = _filtrar_solidos_utiles(solids)
    if not solids:
        return None
    if len(solids) == 1:
        return solids[0]
    merged = solids[0]
    for i in range(1, len(solids)):
        try:
            merged = BooleanOperationsUtils.ExecuteBooleanOperation(
                merged, solids[i], BooleanOperationsType.Union
            )
        except Exception:
            return None
        if merged is None:
            return None
        try:
            if float(merged.Volume) <= 1e-12:
                return None
        except Exception:
            pass
    return merged


def _punto_en_plano(p, plane):
    if p is None or plane is None:
        return None
    try:
        n = plane.Normal
        o = plane.Origin
        d = float(p.Subtract(o).DotProduct(n))
        return p.Subtract(n.Multiply(d))
    except Exception:
        return None


def _plane_from_planar_face(pf):
    if pf is None:
        return None
    try:
        if not isinstance(pf, PlanarFace) and type(pf).__name__ != "PlanarFace":
            return None
    except Exception:
        return None
    try:
        surf = pf.GetSurface()
        if surf is not None:
            try:
                if isinstance(surf, Plane):
                    return surf
            except Exception:
                pass
            try:
                if surf.GetType().Name == "Plane":
                    return surf
            except Exception:
                pass
    except Exception:
        pass
    try:
        n = _xyz_normalize(pf.FaceNormal)
    except Exception:
        n = None
    if n is None:
        return None
    pt = None
    try:
        pt = pf.Origin
    except Exception:
        pt = None
    if pt is None:
        try:
            bbuv = pf.GetBoundingBox()
            if bbuv is not None:
                u = (bbuv.Min.U + bbuv.Max.U) * 0.5
                v = (bbuv.Min.V + bbuv.Max.V) * 0.5
                pt = pf.Evaluate(UV(u, v))
        except Exception:
            pt = None
    if pt is None:
        return None
    try:
        return Plane.CreateByNormalAndOrigin(n, pt)
    except Exception:
        return None


def _curvas_para_model_document(curve, plane):
    out = []
    if curve is None or plane is None:
        return out
    try:
        if isinstance(curve, Line) or type(curve).__name__ == "Line":
            p0 = _punto_en_plano(curve.GetEndPoint(0), plane)
            p1 = _punto_en_plano(curve.GetEndPoint(1), plane)
            if p0 is None or p1 is None or p0.DistanceTo(p1) < 1e-9:
                return out
            out.append(Line.CreateBound(p0, p1))
            return out
        if isinstance(curve, Arc) or type(curve).__name__ == "Arc":
            p0 = _punto_en_plano(curve.GetEndPoint(0), plane)
            p1 = _punto_en_plano(curve.GetEndPoint(1), plane)
            pm = _punto_en_plano(curve.Evaluate(0.5, True), plane)
            if p0 is None or p1 is None or pm is None:
                return out
            try:
                out.append(Arc.Create(p0, p1, pm))
                return out
            except Exception:
                if p0.DistanceTo(p1) >= 1e-9:
                    out.append(Line.CreateBound(p0, p1))
                return out
    except Exception:
        pass
    try:
        pts = curve.Tessellate()
        n_pts = int(pts.Count)
        for i in range(n_pts - 1):
            p0 = _punto_en_plano(pts[i], plane)
            p1 = _punto_en_plano(pts[i + 1], plane)
            if p0 is None or p1 is None or p0.DistanceTo(p1) < 1e-9:
                continue
            out.append(Line.CreateBound(p0, p1))
        if out:
            return out
    except Exception:
        pass
    try:
        p0 = _punto_en_plano(curve.GetEndPoint(0), plane)
        p1 = _punto_en_plano(curve.GetEndPoint(1), plane)
        if p0 is None or p1 is None or p0.DistanceTo(p1) < 1e-9:
            return out
        out.append(Line.CreateBound(p0, p1))
    except Exception:
        pass
    return out


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


def _planar_face_mas_grande_sobre_plano(solid_cut, plane_ref, tol_dist, tol_dot_parallel=0.02):
    if solid_cut is None or plane_ref is None:
        return None
    try:
        pn = _xyz_normalize(plane_ref.Normal)
    except Exception:
        pn = None
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
            try:
                fn = _xyz_normalize(face.FaceNormal)
                if fn is None:
                    continue
                if abs(abs(float(fn.DotProduct(pn))) - 1.0) > tol_dot_parallel:
                    continue
            except Exception:
                continue
            try:
                pt = face.Origin
            except Exception:
                pt = None
            if pt is None:
                try:
                    bbuv = face.GetBoundingBox()
                    if bbuv is not None:
                        u = (bbuv.Min.U + bbuv.Max.U) * 0.5
                        v = (bbuv.Min.V + bbuv.Max.V) * 0.5
                        pt = face.Evaluate(UV(u, v))
                except Exception:
                    pt = None
            if pt is None:
                continue
            if _distancia_punto_a_plano(pt, plane_ref) > tol_dist:
                continue
            try:
                a = face.Area
            except Exception:
                a = 0.0
            if a > best_a:
                best_a = a
                best = face
    except Exception:
        return None
    return best


def _area_curve_loop(cl):
    try:
        return abs(cl.GetArea())
    except Exception:
        return 0.0


def _curve_loops_desde_face(planar_face):
    if planar_face is None:
        return []
    try:
        loops = planar_face.GetEdgesAsCurveLoops()
    except Exception:
        return []
    if loops is None or loops.Count == 0:
        return []
    out = []
    try:
        for i in range(loops.Count):
            out.append(loops[i])
    except Exception:
        pass
    if not out:
        return []
    out.sort(key=_area_curve_loop, reverse=True)
    return out


def _curvas_desde_curve_loop(cl):
    curvas = []
    if cl is None:
        return curvas
    try:
        n = int(cl.NumberOfCurves)
    except Exception:
        n = 0
    if n > 0:
        for i in range(n):
            try:
                c = cl.get_Item(i)
            except Exception:
                try:
                    c = cl[i]
                except Exception:
                    c = None
            if c is not None:
                curvas.append(c)
    if not curvas:
        try:
            for c in cl:
                if c is not None:
                    curvas.append(c)
        except Exception:
            pass
    return curvas


def _buscar_cara_corte(solid_merged, p_mid, n_unit, tol_dist):
    if solid_merged is None:
        return None, None
    tols = [float(tol_dist)]
    try:
        t50 = max(float(tol_dist) * 2.0, UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters))
        t120 = max(float(tol_dist) * 3.0, UnitUtils.ConvertToInternalUnits(120.0, UnitTypeId.Millimeters))
        t250 = max(float(tol_dist) * 4.0, UnitUtils.ConvertToInternalUnits(250.0, UnitTypeId.Millimeters))
        for t in (t50, t120, t250):
            if t > tols[0] * 1.01:
                tols.append(t)
    except Exception:
        for t in (0.12, 0.25, 0.4):
            if t > tols[0] * 1.01:
                tols.append(t)
    for td in tols:
        for flip in (False, True):
            nn = XYZ(-n_unit.X, -n_unit.Y, -n_unit.Z) if flip else n_unit
            try:
                cut_plane = Plane.CreateByNormalAndOrigin(nn, p_mid)
            except Exception:
                continue
            for tol_dot in (0.02, 0.08, 0.15, 0.25, 0.4):
                try:
                    s_cut = BooleanOperationsUtils.CutWithHalfSpace(solid_merged, cut_plane)
                except Exception:
                    s_cut = None
                if s_cut is None:
                    continue
                try:
                    if float(s_cut.Volume) <= 1e-12:
                        continue
                except Exception:
                    pass
                pf = _planar_face_mas_grande_sobre_plano(s_cut, cut_plane, td, tol_dot)
                if pf is not None:
                    return pf, cut_plane
    return None, None


def _punto_interior_muro_mitad_espesor(host, p_eje):
    """
    Un punto hacia el **interior** del muro (≈ 1/4 de espesor respecto al eje en −Orientation
    hacia el exterior, convención Revit). A veces el corte en el eje no atraviesa el sólido
    o deja ``CutWithHalfSpace`` con volumen nulo; el interior mejora el corte.
    """
    if host is None or p_eje is None:
        return None
    try:
        w = abs(float(host.Width))
    except Exception:
        w = 0.0
    if w <= 1e-9:
        return None
    try:
        o = _xyz_normalize(host.Orientation)
    except Exception:
        o = None
    if o is None:
        return None
    try:
        d = w * 0.24
        return p_eje.Subtract(o.Multiply(d))
    except Exception:
        return None


def _buscar_cara_corte_con_nudge(
    solid_merged, p_mid, n_unit, tol_dist, wall_line, host=None
):
    """
    Igual que ``_buscar_cara_corte``; si no hay cara, prueba orígenes desplazados: interior del
    espesor, a lo largo del eje y a lo largo de la normal del plano (el punto medio a veces queda
    mal con muro+fundación o uniones no coplanares).
    """
    if solid_merged is None or p_mid is None or n_unit is None:
        return None, None
    tdir = None
    if wall_line is not None:
        try:
            tdir = _tangent_at_mid(wall_line)
        except Exception:
            tdir = None
    if tdir is None and wall_line is not None:
        try:
            tdir = _xyz_normalize(wall_line.Direction)
        except Exception:
            tdir = None
    try:
        nrm = _xyz_normalize(n_unit)
    except Exception:
        nrm = None

    def _corte_desde_punto(pb):
        if pb is None:
            return None, None
        bf, cp = _buscar_cara_corte(solid_merged, pb, n_unit, tol_dist)
        if bf is not None:
            return bf, cp
        if tdir is not None:
            try:
                steps_mm = (10.0, 20.0, 40.0, 80.0, 150.0, 300.0, 500.0, 800.0)
                steps = [
                    UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters) for mm in steps_mm
                ]
            except Exception:
                steps = (0.03, 0.08, 0.16, 0.35, 0.6, 1.0, 1.2)
            for dist in steps:
                for sgn in (1.0, -1.0):
                    try:
                        p2 = XYZ(
                            pb.X + tdir.X * dist * sgn,
                            pb.Y + tdir.Y * dist * sgn,
                            pb.Z + tdir.Z * dist * sgn,
                        )
                    except Exception:
                        continue
                    bf, cp = _buscar_cara_corte(solid_merged, p2, n_unit, tol_dist)
                    if bf is not None:
                        return bf, cp
        if nrm is not None:
            try:
                steps_n = (5.0, 15.0, 30.0, 60.0, 100.0, 200.0)
                ds = [
                    UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters) for mm in steps_n
                ]
            except Exception:
                ds = (0.02, 0.05, 0.1, 0.2)
            for dist in ds:
                for sgn in (1.0, -1.0):
                    try:
                        p2 = XYZ(
                            pb.X + nrm.X * dist * sgn,
                            pb.Y + nrm.Y * dist * sgn,
                            pb.Z + nrm.Z * dist * sgn,
                        )
                    except Exception:
                        continue
                    bf, cp = _buscar_cara_corte(solid_merged, p2, n_unit, tol_dist)
                    if bf is not None:
                        return bf, cp
        return None, None

    bases = [p_mid]
    p_in = _punto_interior_muro_mitad_espesor(host, p_mid) if host is not None else None
    if p_in is not None:
        try:
            if p_in.DistanceTo(p_mid) > 1e-5:
                bases.append(p_in)
        except Exception:
            bases.append(p_in)

    for pb in bases:
        bf, cp = _corte_desde_punto(pb)
        if bf is not None:
            return bf, cp
    return None, None


def _coleccion_ids_unidas(doc, element):
    if doc is None or element is None:
        return []
    return get_joined_element_ids(doc, element)


class _WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element is not None and isinstance(element, Wall)

    def AllowReference(self, reference, position):
        return False


def _elementos_para_union(doc, host, wall_line, section_plane=None):
    try:
        tol_end = UnitUtils.ConvertToInternalUnits(80.0, UnitTypeId.Millimeters)
        tol_z = UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
        tol_plane = UnitUtils.ConvertToInternalUnits(30.0, UnitTypeId.Millimeters)
        pad_bb = UnitUtils.ConvertToInternalUnits(1.8, UnitTypeId.Meters)
    except Exception:
        tol_end = 0.25
        tol_z = 0.12
        tol_plane = 0.1
        pad_bb = 6.0

    elementos = [host]
    seen = set()
    try:
        seen.add(host.Id)
    except Exception:
        pass

    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or el.Id in seen:
            continue
        if _excluir_de_union_fundaciones(doc, el):
            continue
        if isinstance(el, Wall) and _es_muro_lateral_en_extremos(doc, host, wall_line, el, tol_end):
            if not _muro_apilado_bajo_muro_principal(
                host, el, tol_z
            ) and not _muro_apilado_sobre_muro_principal(host, el, tol_z):
                elementos.append(el)
                seen.add(el.Id)
        elif isinstance(el, Floor) and _incluir_suelo_unido_cara_superior_inferior(
            doc, host, el, tol_z
        ):
            elementos.append(el)
            seen.add(el.Id)

    try:
        xf = ElementIntersectsElementFilter(host)
        coll_w = (
            FilteredElementCollector(doc)
            .OfClass(Wall)
            .WherePasses(xf)
            .ToElementIds()
        )
        for eid in coll_w:
            if eid in seen:
                continue
            w2 = doc.GetElement(eid)
            if w2 is None or not isinstance(w2, Wall):
                continue
            if _es_muro_lateral_en_extremos(doc, host, wall_line, w2, tol_end):
                if not _muro_apilado_bajo_muro_principal(
                    host, w2, tol_z
                ) and not _muro_apilado_sobre_muro_principal(host, w2, tol_z):
                    elementos.append(w2)
                    seen.add(eid)
    except Exception:
        pass

    olh = _outline_host_inflado(host, pad_bb)
    if olh is not None:
        try:
            bf = BoundingBoxIntersectsFilter(olh)
            coll_bb = (
                FilteredElementCollector(doc)
                .OfClass(Wall)
                .WherePasses(bf)
                .ToElementIds()
            )
            for eid in coll_bb:
                if eid in seen:
                    continue
                w3 = doc.GetElement(eid)
                if w3 is None or not isinstance(w3, Wall):
                    continue
                if _es_muro_lateral_en_extremos(doc, host, wall_line, w3, tol_end):
                    if not _muro_apilado_bajo_muro_principal(
                        host, w3, tol_z
                    ) and not _muro_apilado_sobre_muro_principal(host, w3, tol_z):
                        elementos.append(w3)
                        seen.add(eid)
        except Exception:
            pass

    try:
        xf2 = ElementIntersectsElementFilter(host)
        coll_f = (
            FilteredElementCollector(doc)
            .OfClass(Floor)
            .WherePasses(xf2)
            .ToElementIds()
        )
        for eid in coll_f:
            if eid in seen:
                continue
            fl = doc.GetElement(eid)
            if fl is None or not isinstance(fl, Floor):
                continue
            if _suelo_para_union_intersect_filter(doc, host, fl, tol_z, section_plane):
                elementos.append(fl)
                seen.add(eid)
    except Exception:
        pass

    if section_plane is not None:
        _agregar_floors_intersectan_plano_corte(
            doc, host, elementos, seen, section_plane, tol_z, pad_bb
        )

    return elementos, tol_plane


def hay_suelo_unido_cara_superior(doc, host):
    """
    True si algún ``Floor`` está unido al muro vía *Join geometry* y queda
    en la **cara superior** del muro (forjado encima, solape en planta y contacto Z).

    Si es False, en post-malla (Armado Muros Nodo) las barras **verticales** pueden
    recibir L+135 en cabeza (sin empalme superior unido) en lugar de solo pata 90
    hacia fundación.
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return False
    try:
        tol_z = UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
    except Exception:
        tol_z = 0.12
    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or not isinstance(el, Floor):
            continue
        if _excluir_de_union_fundaciones(doc, el):
            continue
        if _suelo_sobre_cara_superior_muro(host, el, tol_z):
            return True
    return False


def hay_suelo_unido_cara_inferior(doc, host):
    """
    True si algún ``Floor`` (incl. fundación/losa ciment.) está unido al muro
    y contacta la **cara inferior** (base del muro) con solape en planta.

    Si es False, en post-malla (Armado Muros Nodo) las verticales pueden recibir
    L+135 en el pie (sin empalme inferior unido).
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return False
    try:
        tol_z = UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
    except Exception:
        tol_z = 0.12
    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or not isinstance(el, Floor):
            continue
        if _suelo_bajo_cara_inferior_muro(host, el, tol_z):
            return True
    return False


def hay_muro_unido_cara_inferior(doc, host):
    u"""
    True si hay un ``Wall`` unido bajo el muro (apilado/contiguo en Z) en la cara inferior.

    Se usa para suprimir L+135 (pata+gancho) en el **pie** de verticales: si existe empalme
    con otro muro en la base, no se debe añadir pata L ni gancho 135 en ese extremo.
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return False
    try:
        tol_z = UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
    except Exception:
        tol_z = 0.12
    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or not isinstance(el, Wall):
            continue
        if _muro_apilado_bajo_muro_principal(host, el, tol_z):
            return True
    return False


def hay_muro_unido_cara_superior(doc, host):
    u"""
    True si hay un ``Wall`` unido cuya base contacta con la **cara superior** del muro
    (apilación en Z), vía ``_muro_apilado_sobre_muro_principal``.

    Se usa en el post-malla para verticales de **cara interior**: en el **extremo inicial** del
    boceto hacia el encuentro (típ. cabeza) no añadir pata L ni gancho 135 al haber
    empalme con muro en cabeza.
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return False
    try:
        tol_z = UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
    except Exception:
        tol_z = 0.12
    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or not isinstance(el, Wall):
            continue
        if _muro_apilado_sobre_muro_principal(host, el, tol_z):
            return True
    return False


def hay_suelo_unido_cara_inferior_excluyendo_fundacion(doc, host):
    """
    Igual que ``hay_suelo_unido_cara_inferior`` pero **excluye** fundación estructural y
    losa de cimentación (misma idea que la cara superior con
    ``_excluir_de_union_fundaciones``). Sirve para decidir **L+135 en el pie** del post-malla:
    la cimentación unida no debe anular la pata/gancho como si hubiera forjado de empalme.

    Para *pata 90° a fundación* siga usando ``hay_suelo_unido_cara_inferior`` (incluye fund.).
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return False
    try:
        tol_z = UnitUtils.ConvertToInternalUnits(40.0, UnitTypeId.Millimeters)
    except Exception:
        tol_z = 0.12
    for jid in _coleccion_ids_unidas(doc, host):
        el = doc.GetElement(jid)
        if el is None or not isinstance(el, Floor):
            continue
        if _excluir_de_union_fundaciones(doc, el):
            continue
        if _suelo_bajo_cara_inferior_muro(host, el, tol_z):
            return True
    return False


def get_otros_element_ids_boceto_nudo(doc, host):
    """
    ``ElementId`` de los demás elementos que entran en ``_elementos_para_union`` (mismo criterio
    que el boceto por unión booleana). Incluye muros en encuentro por intersección/extremos aunque
    **no** estén en ``GetJoinedElements`` (caso típico: L sin *Unir geometría* entre muros).
    No incluye el propio ``host``.
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return []
    wall_line, curve_orig = _location_as_line(host)
    if wall_line is None:
        return []
    mid = _midpoint_curve(curve_orig)
    plane_normal = _normal_de_location_curve(curve_orig, host)
    if mid is None or plane_normal is None:
        return []
    try:
        section_plane = Plane.CreateByNormalAndOrigin(plane_normal, mid)
    except Exception:
        return []
    elementos, _tol = _elementos_para_union(doc, host, wall_line, section_plane)
    out = []
    host_id_int = _element_id_to_int(host.Id)
    for el in elementos:
        if el is None:
            continue
        try:
            if host_id_int is not None and _element_id_to_int(el.Id) == host_id_int:
                continue
        except Exception:
            try:
                if el.Id == host.Id:
                    continue
            except Exception:
                pass
        try:
            out.append(el.Id)
        except Exception:
            pass
    return out


def _mismo_id_elemento(a, b):
    if a is None or b is None:
        return False
    return _element_ids_equal(a.Id, b.Id)


def _texto_elemento_en_union(doc, el, host):
    """Una línea legible: rol, Id, categoría, tipo de familia."""
    if el is None:
        return "  - (nulo)"
    eid = _element_id_to_int(el.Id)
    if eid is None:
        eid = "?"

    cat = ""
    try:
        if el.Category is not None:
            cat = el.Category.Name
    except Exception:
        pass

    tipo = ""
    try:
        tid = el.GetTypeId()
        if tid is not None and tid != ElementId.InvalidElementId:
            et = doc.GetElement(tid)
            if et is not None:
                tipo = et.Name
    except Exception:
        pass

    if isinstance(el, Wall):
        rol = "Muro base" if _mismo_id_elemento(el, host) else "Muro (extremo)"
    elif isinstance(el, Floor):
        rol = "Suelo"
    else:
        rol = cat or "Elemento"

    nom = ""
    try:
        n = el.Name
        if n:
            nom = n
    except Exception:
        pass

    partes = [rol, "Id={}".format(eid)]
    if cat:
        partes.append(cat)
    if tipo:
        partes.append("Tipo: {}".format(tipo))
    if nom:
        partes.append('Nombre: "{}"'.format(nom))
    return "  - " + " | ".join(partes)


def _task_dialog_elementos_unidos(doc, elementos, host, n_lat, n_fl):
    """
    Muestra un ``TaskDialog`` con el detalle de elementos que entrarán en la unión booleana.
    Si la API falla, escribe el mismo texto en consola.
    """
    n = len(elementos)
    instr = (
        "Muro base + {0} muro(s) en extremo(s) + {1} suelo(s). Total: {2} elemento(s) en la unión."
    ).format(n_lat, n_fl, n)
    lines = [_texto_elemento_en_union(doc, el, host) for el in elementos]
    body = "\r\n".join(lines)
    try:
        td = TaskDialog("BIMTools.WallNodeBooleanUnion")
        try:
            td.Title = "Elementos detectados para la unión"
        except Exception:
            pass
        td.MainInstruction = instr
        td.MainContent = body
        td.Show()
    except Exception:
        try:
            TaskDialog.Show("Elementos detectados para la unión", instr + "\r\n\r\n" + body)
        except Exception:
            print(instr)
            for ln in lines:
                print(ln)


def _direccion_major_muro(host):
    """Dirección unitaria del eje del muro (luz mayor típica del refuerzo de área)."""
    wall_line, curve_orig = _location_as_line(host)
    if wall_line is None or curve_orig is None:
        return None
    ta = _tangent_at_mid(curve_orig)
    if ta is not None:
        return ta
    try:
        return _xyz_normalize(wall_line.Direction)
    except Exception:
        return None


def construir_boceto_nodo_union(doc, host, show_task_dialog=True, solo_bucle_exterior=False):
    """
    Construye el contorno del nudo (unión booleana + corte + fundación + offsets) para **Area Reinforcement**
    o para dibujar model lines.

    Retorna ``(plane_face, lista de Curve, major_dir XYZ, h_fund)`` o ``(None, None, None, 0.0)`` si falla.
    ``major_dir`` es la dirección del eje del muro (``LocationCurve``) para ``AreaReinforcement.Create``.

    ``solo_bucle_exterior``: si es True, solo se procesa el **perímetro exterior** de la cara de corte
    (mayor área). ``AreaReinforcement.Create`` exige una lista de curvas **contiguas** en un solo
    bucle cerrado; concatenar bucles interiores (huecos) con el exterior provoca el error
    *"These curves are not contiguous"*. Para ``False`` (p. ej. boceto con todas las aristas en
    ``main()``), se incluyen todos los ``CurveLoop`` de la cara.

    Los vecinos del nudo se resuelven con ``_elementos_para_union``: elementos unidos por
    *Join geometry*, más muros/suelos detectados por intersección de sólidos y por bbox
    (heurística T / extremos). Antes se omitía todo ese barrido si *Join geometry* no
    devolvía ningún muro o suelo válido; entonces el boceto quedaba solo con el host y
    un muro perpendicular en encuentro podía no integrarse aunque compartiera geometría.
    """
    if doc is None or host is None or not isinstance(host, Wall):
        return None, None, None, 0.0

    wall_line, curve_orig = _location_as_line(host)
    if wall_line is None:
        return None, None, None, 0.0

    mid = _midpoint_curve(curve_orig)
    plane_normal = _normal_de_location_curve(curve_orig, host)
    if mid is None or plane_normal is None:
        return None, None, None, 0.0

    try:
        section_plane = Plane.CreateByNormalAndOrigin(plane_normal, mid)
    except Exception:
        return None, None, None, 0.0

    # Siempre usar el mismo ensamblado de candidatos (join + intersección + bbox). Acotar solo
    # con ``[host]`` cuando *Join* devuelve vacío impedía incluir muros en T detectados por
    # ``ElementIntersectsElementFilter`` / ``_es_muro_lateral_en_extremos``.
    elementos, tol_plane = _elementos_para_union(doc, host, wall_line, section_plane)
    n_lat = sum(
        1
        for e in elementos
        if isinstance(e, Wall) and not _element_ids_equal(e.Id, host.Id)
    )
    n_fl = sum(1 for e in elementos if isinstance(e, Floor))
    if show_task_dialog:
        _task_dialog_elementos_unidos(doc, elementos, host, n_lat, n_fl)

    # Un **cuerpo representativo (mayor volumen) por elemento** primero. Con aberturas,
    # ``get_Geometry`` devuelve muchas piezas por muro; unir *todas* consecutivamente
    # dispara cientos de ``BooleanOperationsType.Union`` y congela la UI. La vía
    # representante ya se usaba como respaldo; aquí rinde bien el nudo y, si hace
    # falta, se repite la ruta “todas las piezas” (más lenta) solo al fallar.
    merged = _unir_solidos_representantes_por_elemento(elementos)
    all_solids = None
    if merged is None:
        all_solids = []
        for el in elementos:
            all_solids.extend(_solidos_desde_elemento(el))
        all_solids = _filtrar_solidos_utiles(all_solids)
        if not all_solids:
            return None, None, None, 0.0
        merged = _unir_solidos(all_solids)
    if merged is None and all_solids:
        merged = _solido_mayor_volumen(all_solids)
    if merged is None:
        return None, None, None, 0.0

    host_solo = _filtrar_solidos_utiles(_solidos_desde_elemento(host))
    host_solo = (
        [(_solido_mayor_volumen(host_solo) or host_solo[0])] if host_solo else []
    )

    best_face, cut_plane = _buscar_cara_corte_con_nudge(
        merged, mid, plane_normal, tol_plane, wall_line, host
    )
    if best_face is None and host_solo:
        m2 = host_solo[0]
        if m2 is not None:
            best_face, cut_plane = _buscar_cara_corte_con_nudge(
                m2, mid, plane_normal, tol_plane, wall_line, host
            )
    if best_face is None:
        return None, None, None, 0.0

    loops = _curve_loops_desde_face(best_face)
    if not loops:
        return None, None, None, 0.0

    # Un solo bucle contiguo para API que requiere perímetro único (p. ej. Area Reinforcement).
    loops_a_procesar = [loops[0]] if solo_bucle_exterior else loops

    plane_face = _plane_from_planar_face(best_face)
    if plane_face is None:
        plane_face = cut_plane
    if plane_face is None:
        plane_face = section_plane

    h_fund = _altura_max_fundacion_unida_host(doc, host)

    all_curves = []
    for cl in loops_a_procesar:
        sub = _curvas_desde_curve_loop(cl)
        if not sub:
            continue
        if h_fund > 1e-9:
            sub2 = _extender_curvas_borde_inferior_fundacion(plane_face, sub, h_fund)
            if sub2:
                sub = sub2
        sub_of = _aplicar_offsets_interiores_boceto_loop(plane_face, sub, h_fund)
        if not sub_of:
            sub_of = sub
        all_curves.extend(sub_of)

    if not all_curves and loops_a_procesar:
        for cl in loops_a_procesar:
            sub = _curvas_desde_curve_loop(cl)
            if not sub:
                continue
            all_curves.extend(sub)
        if not all_curves and loops_a_procesar:
            for cl in loops_a_procesar:
                sub0 = _curvas_desde_curve_loop(cl)
                if not sub0 or plane_face is None:
                    continue
                sub_l = _aplicar_offsets_interiores_boceto_loop(plane_face, sub0, 0.0)
                all_curves.extend(sub_l if sub_l else sub0)

    if not all_curves:
        return None, None, None, 0.0

    major_dir = _direccion_major_muro(host)
    if major_dir is None:
        return None, None, None, 0.0

    return plane_face, all_curves, major_dir, h_fund


def main(__revit__):
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        print("No hay documento activo.")
        return
    doc = uidoc.Document

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, _WallSelectionFilter(), "Selecciona un único muro."
        )
        host = doc.GetElement(ref.ElementId)
    except Exception as ex:
        err = str(ex).lower()
        if "cancel" in err or "operación" in err or "operation" in err or "cancelled" in err:
            return
        print("Selección: {}".format(ex))
        return

    if host is None or not isinstance(host, Wall):
        print("El elemento no es un muro.")
        return

    plane_face, all_curves, _, h_fund = construir_boceto_nodo_union(doc, host, show_task_dialog=True)
    if plane_face is None or not all_curves:
        print("No se pudo construir el boceto del nudo (geometría o unión booleana).")
        return

    if h_fund > 1e-9:
        try:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId

            hmm = UnitUtils.ConvertFromInternalUnits(h_fund, UnitTypeId.Millimeters)
            print(
                "Fundación unida al muro: altura bbox máx. ~{:.1f} mm; estirando base del boceto.".format(
                    hmm
                )
            )
        except Exception:
            print(
                "Fundación unida al muro: altura bbox máx. ~{:.4f} ft; estirando base del boceto.".format(
                    h_fund
                )
            )

    created = 0
    first_err = None
    t = Transaction(doc, "Nodo muro: sección transversal → model lines")
    t.Start()
    try:
        sp = SketchPlane.Create(doc, plane_face)
        for cv in all_curves:
            if cv is None:
                continue
            for c_doc in _curvas_para_model_document(cv, plane_face):
                if c_doc is None:
                    continue
                try:
                    doc.Create.NewModelCurve(c_doc, sp)
                    created += 1
                except Exception as ex:
                    if first_err is None:
                        first_err = str(ex)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        print("Transacción: {}".format(ex))
        return

    if created == 0:
        msg = "No se creó ninguna model line."
        if first_err:
            msg += " Primera excepción: {}".format(first_err)
        print(msg)
        return

    print(
        "Listo: {} model line(s). Offsets interiores 50 mm (fundación) / 25 mm (resto) cuando aplica.".format(
            created
        )
    )


if __name__ == "__main__":
    try:
        main(__revit__)
    except NameError:
        print("Ejecuta main(__revit__) en Revit (RPS / pyRevit).")
