# -*- coding: utf-8 -*-
"""
Geometría de zapata de muro (``WallFoundation``) por **corte con planos** definidos
exclusivamente por el **muro host** de la fundación.

Flujo (API Revit):

1. **Muro asociado** — ``WallFoundation.WallId`` → ``Wall``.
2. **Location curve** — ``wall.Location`` como ``LocationCurve`` → ``Curve`` del eje del muro.
3. **Planos de corte** — por el **centro** de la ``LocationCurve`` del muro
   (``Evaluate(0.5, True)``). Tangente 3D en ese punto: ``ComputeDerivatives(0.5, True).BasisX``.
   - *Transversal (U)*: ``Plane.CreateByNormalAndOrigin(tangente_3D, punto_medio)`` (igual
     que el script de referencia).
   - *Longitudinal*: normal ``tangente_3D × Z`` (plano que contiene eje del muro y vertical).
4. **Intersección y soleira** — aristas tipo ``Line``: intersección con el plano por
   **distancias firmadas**; otras curvas: API Revit. Se arma la nube de puntos; en la
   banda de **Z mínima** (``_SOLEIRA_Z_BAND_FT``) se elige el **par de puntos más
   separado**. Si falla, respaldo por tramos del perímetro.

Los recortes y offsets de armado se aplican con la misma cadena que
``aplicar_recubrimiento_inferior_completo_mm`` + ``offset_linea_eje_barra_*``.
"""

from __future__ import print_function

from Autodesk.Revit.DB import (
    ElementId,
    GeometryInstance,
    IntersectionResultArray,
    Line,
    LocationCurve,
    Options,
    Plane,
    SetComparisonResult,
    Solid,
    ViewDetailLevel,
    Wall,
    WallFoundation,
    XYZ,
)

from geometria_fundacion_cara_inferior import (
    aplicar_recubrimiento_inferior_completo_mm,
    evaluar_caras_paralelas_curva_mas_cercana,
    offset_linea_eje_barra_desde_cara_inferior_mm,
)

# Alineado con ``enfierrado_wall_foundation`` (evitar import circular).
_REC_OFF_PLANTA_INF_MM = 100.0
_REC_EXTREMOS_LONG_TANGENTE_MM = 50.0
_REC_EXTREMOS_INFERIOR_MM = 50.0
_RECO_HOR_MM = 50.0

_TOL_SEG_MERGE_FT = 0.02
_TOL_Z_BUCKET_FT = 0.12
# Coherente con script de referencia (soleira ~1,5 cm)
_SOLEIRA_Z_BAND_FT = 0.05
_PLANE_LINE_DIST_TOL_FT = 0.001


def _safe_volume(solid):
    try:
        return float(solid.Volume)
    except Exception:
        return None


def _line_soleira_planta_horizontal(crv):
    """
    Proyecta una ``Line`` al plano horizontal Z = min(Z) de sus extremos para
    que ``CreateOffset(..., BasisZ)`` del recubrimiento en planta no falle en
    aristas de corte ligeramente inclinadas.
    """
    if crv is None or not isinstance(crv, Line):
        return None
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        z = min(float(p0.Z), float(p1.Z))
        q0 = XYZ(float(p0.X), float(p0.Y), z)
        q1 = XYZ(float(p1.X), float(p1.Y), z)
        if float(q0.DistanceTo(q1)) < 1e-9:
            return None
        return Line.CreateBound(q0, q1)
    except Exception:
        return None


def _signed_plane_dist(plane, p):
    try:
        return float(plane.Normal.DotProduct(p - plane.Origin))
    except Exception:
        return 0.0


def _tangent_3d_normalized_curve_mid(lc):
    """Tangente unitaria 3D en el punto medio paramétrico (0…1) de la curva del muro."""
    if lc is None:
        return None
    try:
        dv = lc.ComputeDerivatives(0.5, True)
        if dv is None:
            return None
        tx = dv.BasisX
        if tx.GetLength() < 1e-12:
            return None
        return tx.Normalize()
    except Exception:
        return None


def _manual_intersect_segment_plane_points(p0, p1, plane):
    """
    Intersección de un segmento con un plano (lógica robusta tipo script BIMTools).
    Devuelve 0, 1 o 2 puntos **en el mismo sistema** que ``p0``/``p1``.
    """
    if plane is None or p0 is None or p1 is None:
        return []
    try:
        n = plane.Normal
        o = plane.Origin
        tol = max(float(_PLANE_LINE_DIST_TOL_FT), 1e-9)
        d0 = float(n.DotProduct(p0 - o))
        d1 = float(n.DotProduct(p1 - o))
        if abs(d0) <= tol and abs(d1) <= tol:
            return [p0, p1]
        if abs(d0) <= tol:
            return [XYZ(float(p0.X), float(p0.Y), float(p0.Z))]
        if abs(d1) <= tol:
            return [XYZ(float(p1.X), float(p1.Y), float(p1.Z))]
        if d0 * d1 < 0.0:
            denom = d0 - d1
            if abs(denom) < 1e-12:
                return []
            t = float(d0) / float(denom)
            if t < -tol or t > 1.0 + tol:
                return []
            t = max(0.0, min(1.0, t))
            v = p1 - p0
            pt = p0.Add(v.Multiply(t))
            return [XYZ(float(pt.X), float(pt.Y), float(pt.Z))]
        return []
    except Exception:
        return []


def _dedupe_points_xyz(pts, tol_ft=5e-4):
    if not pts:
        return []
    try:
        tol = max(float(tol_ft), 1e-7)
    except Exception:
        tol = 5e-4
    out = []
    for p in pts:
        if p is None:
            continue
        dup = False
        for q in out:
            try:
                if float(p.DistanceTo(q)) <= tol:
                    dup = True
                    break
            except Exception:
                continue
        if not dup:
            out.append(p)
    return out


def _linea_soleira_desde_puntos_interseccion(pts_doc):
    """
    De la nube de puntos del corte plano∩sólido, construye la línea de **soleira**
    (como en el script de referencia): puntos en la banda de Z mínima y par más
    separado entre ellos.
    """
    pts = _dedupe_points_xyz(pts_doc)
    if len(pts) < 2:
        return None
    try:
        z_min = min(float(p.Z) for p in pts)
    except Exception:
        return None
    band = float(_SOLEIRA_Z_BAND_FT)
    pts_base = [p for p in pts if abs(float(p.Z) - z_min) <= band + 1e-9]
    if len(pts_base) < 2:
        pts_base = pts
    best_a, best_b = None, None
    best_d = -1.0
    nbase = len(pts_base)
    for i in range(nbase):
        for j in range(i + 1, nbase):
            try:
                d = float(pts_base[i].DistanceTo(pts_base[j]))
            except Exception:
                continue
            if d > best_d:
                best_d = d
                best_a, best_b = pts_base[i], pts_base[j]
    if best_a is None or best_b is None or best_d < 1e-8:
        return None
    try:
        return Line.CreateBound(best_a, best_b)
    except Exception:
        return None


def _puntos_interseccion_arista_plano(edge, plane_local, trf_to_doc):
    """Puntos en **documento** de la intersección arista — plano (``plane_local``)."""
    if edge is None or plane_local is None:
        return []
    try:
        crv = edge.AsCurve()
    except Exception:
        crv = None
    if crv is None or not crv.IsBound:
        return []
    pts_local = []
    try:
        if isinstance(crv, Line):
            p0, p1 = crv.GetEndPoint(0), crv.GetEndPoint(1)
            pts_local = _manual_intersect_segment_plane_points(p0, p1, plane_local)
        else:
            pts_local = _curve_intersect_plane_points(crv, plane_local)
    except Exception:
        pts_local = []
    out = []
    for p in pts_local:
        q = _xyz_document_from_local(p, trf_to_doc)
        if q is not None:
            out.append(q)
    return out


def _puntos_interseccion_plano_solidos_doc(solid_pairs, plane_doc):
    """Unión de puntos de corte (documento) de todas las aristas ∩ ``plane_doc``."""
    if plane_doc is None:
        return []
    acc = []
    for sol, trf in solid_pairs:
        if sol is None:
            continue
        pl = _plane_document_to_local(plane_doc, trf)
        if pl is None:
            continue
        try:
            ea = sol.Edges
            if ea is None:
                continue
            ne = int(ea.Size)
        except Exception:
            continue
        for i in range(ne):
            try:
                ed = ea.get_Item(i)
            except Exception:
                ed = None
            acc.extend(_puntos_interseccion_arista_plano(ed, pl, trf))
    return acc


def _make_geometry_options():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
    return opts


def _iter_solids_wall_foundation(elem):
    """
    ``(Solid, Transform|None)``: prioriza ``GetInstanceGeometry()`` (coords.
    de modelo); si no hay sólidos, ``GetSymbolGeometry()`` + ``Transform``.
    """
    if elem is None:
        return
    try:
        geom = elem.get_Geometry(_make_geometry_options())
        if not geom:
            return
        for go in geom:
            try:
                if isinstance(go, GeometryInstance):
                    geoms_inst = []
                    try:
                        inst = go.GetInstanceGeometry()
                        if inst:
                            geoms_inst = [x for x in inst]
                    except Exception:
                        geoms_inst = []
                    solids_inst = []
                    for x in geoms_inst:
                        if not isinstance(x, Solid):
                            continue
                        v = _safe_volume(x)
                        if v is not None and abs(v) > 1e-12:
                            solids_inst.append(x)
                    if solids_inst:
                        for g in solids_inst:
                            yield g, None
                    else:
                        geoms = []
                        t = go.Transform
                        try:
                            sym = go.GetSymbolGeometry()
                            if sym:
                                geoms = [x for x in sym]
                        except Exception:
                            geoms = []
                        for g in geoms:
                            if not isinstance(g, Solid):
                                continue
                            vol = _safe_volume(g)
                            if vol is None or abs(vol) <= 1e-12:
                                continue
                            yield g, t
                elif isinstance(go, Solid):
                    v = _safe_volume(go)
                    if v is None or abs(v) <= 1e-12:
                        continue
                    yield go, None
            except Exception:
                continue
    except Exception:
        return


def _plane_document_to_local(plane_doc, trf_to_doc):
    """Lleva un plano en coordenadas de documento al espacio local del sólido."""
    if plane_doc is None:
        return None
    if trf_to_doc is None:
        return plane_doc
    try:
        inv = trf_to_doc.Inverse
        o = inv.OfPoint(plane_doc.Origin)
        n = inv.OfVector(plane_doc.Normal)
        if n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        return Plane.CreateByNormalAndOrigin(n, o)
    except Exception:
        return None


def _xyz_document_from_local(p, trf_to_doc):
    if p is None:
        return None
    if trf_to_doc is None:
        return XYZ(float(p.X), float(p.Y), float(p.Z))
    try:
        q = trf_to_doc.OfPoint(p)
        return XYZ(float(q.X), float(q.Y), float(q.Z))
    except Exception:
        return None


def _curve_intersect_plane_points(curve, plane):
    """Puntos de intersección de una arista acotada con un plano (coords. locales)."""
    if curve is None or not curve.IsBound or plane is None:
        return []
    pts = []
    try:
        arr = IntersectionResultArray()
        r = curve.Intersect(plane, arr)
        if arr is not None and arr.Size > 0:
            for i in range(int(arr.Size)):
                try:
                    it = arr.get_Item(i)
                    if it is not None and it.XYZPoint is not None:
                        pts.append(it.XYZPoint)
                except Exception:
                    continue
        if (
            r
            in (
                SetComparisonResult.Subset,
                SetComparisonResult.Superset,
                SetComparisonResult.Overlap,
            )
            and not pts
        ):
            try:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                d0 = _signed_plane_dist(plane, p0)
                d1 = _signed_plane_dist(plane, p1)
                tol_c = max(_TOL_SEG_MERGE_FT * 0.15, 1e-5)
                if abs(d0) <= tol_c and abs(d1) <= tol_c:
                    return [p0, p1]
            except Exception:
                pass
    except Exception:
        pass
    return pts


def _segment_from_points_local(pts, trf_to_doc):
    """Segmento 3D en documento: con 2+ puntos usa la **cuerda más larga** (robusto con arcos)."""
    if len(pts) < 2:
        return None
    try:
        out = []
        for p in pts:
            q = _xyz_document_from_local(p, trf_to_doc)
            if q is not None:
                out.append(q)
        if len(out) < 2:
            return None
        if len(out) == 2:
            return Line.CreateBound(out[0], out[1])
        best_i, best_j = 0, 1
        best_d = -1.0
        n = len(out)
        for i in range(n):
            for j in range(i + 1, n):
                d = float(out[i].DistanceTo(out[j]))
                if d > best_d:
                    best_d = d
                    best_i, best_j = i, j
        if best_d < 1e-8:
            return None
        return Line.CreateBound(out[best_i], out[best_j])
    except Exception:
        return None


def _collect_plane_intersection_segments(solid, plane_doc, trf_to_doc):
    segments = []
    if solid is None or plane_doc is None:
        return segments
    pl = _plane_document_to_local(plane_doc, trf_to_doc)
    if pl is None:
        return segments
    try:
        ea = solid.Edges
        if ea is None:
            return segments
        n = int(ea.Size)
    except Exception:
        return segments
    for i in range(n):
        try:
            edge = ea.get_Item(i)
            crv = edge.AsCurve() if edge is not None else None
        except Exception:
            crv = None
        pts = _curve_intersect_plane_points(crv, pl)
        if len(pts) >= 2:
            seg = _segment_from_points_local(pts, trf_to_doc)
            if seg is not None:
                segments.append(seg)
    return segments


def _tramos_perimetro_interseccion_plano_solidos(solid_pairs, plane_doc):
    """
    Perímetro de la sección **plano ∩ sólido**: tramos ``Line`` al intersecar cada
    arista del sólido con el plano (en coords. de documento).
    """
    all_segs = []
    for sol, trf in solid_pairs:
        all_segs.extend(_collect_plane_intersection_segments(sol, plane_doc, trf))
    out = []
    for s in all_segs:
        try:
            if s is None or float(s.Length) < 1e-8:
                continue
            out.append(s)
        except Exception:
            continue
    return out


def _pick_lowest_perimeter_segment_respaldo(segments):
    """Respaldo: menor Z medio y tramo más largo en esa banda."""
    if not segments:
        return None
    rows = []
    for crv in segments:
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            z = 0.5 * (float(p0.Z) + float(p1.Z))
            L = float(crv.Length)
            rows.append((z, L, crv))
        except Exception:
            continue
    if not rows:
        return None
    z_min = min(r[0] for r in rows)
    tier = [r for r in rows if r[0] <= z_min + _TOL_Z_BUCKET_FT]
    if not tier:
        tier = list(rows)
    tier.sort(key=lambda r: -r[1])
    return tier[0][2]


def _curva_inferior_desde_tramos_interseccion(segmentos):
    """
    De los tramos del perímetro de la sección (intersección plano — geometría fundación),
    obtiene la **curva inferior** representativa de la soleira: tramos cuyos **dos**
    extremos están en la banda de Z mínima global del corte; entre ellos el **más largo**.
    Si ninguno califica (mallado irregular), usa el criterio por Z medio (respaldo).
    """
    if not segmentos:
        return None
    eps = float(_TOL_Z_BUCKET_FT)
    z_pts = []
    for crv in segmentos:
        try:
            p0, p1 = crv.GetEndPoint(0), crv.GetEndPoint(1)
            z_pts.extend([float(p0.Z), float(p1.Z)])
        except Exception:
            continue
    if not z_pts:
        return None
    z_floor = min(z_pts)
    candidatos = []
    for crv in segmentos:
        try:
            p0, p1 = crv.GetEndPoint(0), crv.GetEndPoint(1)
            z0, z1 = float(p0.Z), float(p1.Z)
            if z0 <= z_floor + eps and z1 <= z_floor + eps:
                candidatos.append((float(crv.Length), crv))
        except Exception:
            continue
    if candidatos:
        candidatos.sort(key=lambda x: -x[0])
        return candidatos[0][1]
    return _pick_lowest_perimeter_segment_respaldo(segmentos)


def location_curve_muro_host(wall):
    """
    ``Curve`` del eje del muro host (``LocationCurve.Curve``).

    Returns:
        ``Curve`` | ``None``
    """
    if wall is None or not isinstance(wall, Wall):
        return None
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return None
    return loc.Curve


def punto_centro_location_curve_muro(wall):
    """
    Punto medio de la ``LocationCurve`` del muro en coordenadas de documento
    (``Curve.Evaluate(0.5, True)``, con respaldo a mitad de cuerda en línea).

    Returns:
        ``XYZ`` | ``None``
    """
    lc = location_curve_muro_host(wall)
    if lc is None:
        return None
    try:
        p = lc.Evaluate(0.5, True)
        if p is not None:
            return XYZ(float(p.X), float(p.Y), float(p.Z))
    except Exception:
        pass
    try:
        p0 = lc.GetEndPoint(0)
        p1 = lc.GetEndPoint(1)
        return XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
    except Exception:
        return None


def planos_corte_verticales_desde_location_muro(wall, punto_origen_doc=None):
    """
    A partir del **LocationCurve** del muro, construye los dos planos verticales de
    corte usados para intersectar la geometría de la fundación. Los planos pasan
    por ``punto_origen_doc`` (por defecto el **centro** de esa curva).

    Args:
        wall: ``Wall`` host.
        punto_origen_doc: origen 3D de ambos planos; si es ``None``, se usa
            :func:`punto_centro_location_curve_muro`. La tangente en planta se calcula
            proyectando ese punto sobre la ``LocationCurve``.

    Returns:
        dict | None: ``pl_long``, ``pl_trans``, ``tu``, ``perp``, ``location_curve``,
        ``origen_planos`` o None.
    """
    if wall is None:
        return None
    if punto_origen_doc is None:
        punto_origen_doc = punto_centro_location_curve_muro(wall)
    if punto_origen_doc is None:
        return None
    lc = location_curve_muro_host(wall)
    if lc is None:
        return None
    origin = XYZ(
        float(punto_origen_doc.X),
        float(punto_origen_doc.Y),
        float(punto_origen_doc.Z),
    )
    z_up = XYZ.BasisZ
    tu, perp = _tangent_and_perp_xy_from_wall(wall, punto_origen_doc)
    t3d = _tangent_3d_normalized_curve_mid(lc)
    if t3d is not None:
        n_trans_vec = t3d
        n_long_vec = t3d.CrossProduct(z_up)
        if float(n_long_vec.GetLength()) < 1e-12:
            if tu is not None:
                n_long_vec = tu.CrossProduct(z_up)
            else:
                tu_h = XYZ(float(t3d.X), float(t3d.Y), 0.0)
                if float(tu_h.GetLength()) < 1e-12:
                    return None
                tu_h = tu_h.Normalize()
                n_long_vec = tu_h.CrossProduct(z_up)
        if float(n_long_vec.GetLength()) < 1e-12:
            return None
        n_long = n_long_vec.Normalize()
        try:
            pl_long = Plane.CreateByNormalAndOrigin(n_long, origin)
            pl_trans = Plane.CreateByNormalAndOrigin(n_trans_vec, origin)
        except Exception:
            return None
        if tu is None or perp is None:
            tu = XYZ(float(t3d.X), float(t3d.Y), 0.0)
            if float(tu.GetLength()) < 1e-12:
                return None
            tu = tu.Normalize()
            perp = XYZ(-float(tu.Y), float(tu.X), 0.0).Normalize()
            try:
                ow = wall.Orientation
                ow_xy = XYZ(float(ow.X), float(ow.Y), 0.0)
                if ow_xy.GetLength() > 1e-12:
                    ow_xy = ow_xy.Normalize()
                    if float(perp.DotProduct(ow_xy)) < 0.0:
                        perp = perp.Negate()
            except Exception:
                pass
    else:
        if tu is None or perp is None:
            return None
        n_long = tu.CrossProduct(z_up)
        if n_long.GetLength() < 1e-12:
            return None
        n_long = n_long.Normalize()
        n_trans = XYZ(float(tu.X), float(tu.Y), 0.0)
        if n_trans.GetLength() < 1e-12:
            return None
        n_trans = n_trans.Normalize()
        try:
            pl_long = Plane.CreateByNormalAndOrigin(n_long, origin)
            pl_trans = Plane.CreateByNormalAndOrigin(n_trans, origin)
        except Exception:
            return None
    return {
        "pl_long": pl_long,
        "pl_trans": pl_trans,
        "tu": tu,
        "perp": perp,
        "location_curve": lc,
        "origen_planos": origin,
    }


def vector_transversal_planta_desde_muro_host(wf, q_doc):
    """
    Unitario en **planta**, perpendicular al ``LocationCurve`` del **muro host**
    de la zapata (sentido coherente con ``Wall.Orientation`` cuando existe).
    Es la dirección de referencia para **barras transversales / U**.

    Args:
        wf: ``WallFoundation``
        q_doc: punto de referencia (p. ej. centro del eje de zapata en XY); se
            usa para proyección sobre el muro si la línea es curva.

    Returns:
        ``XYZ`` | ``None``
    """
    wall = host_wall_from_wall_foundation(wf)
    if wall is None or q_doc is None:
        return None
    _tu, perp = _tangent_and_perp_xy_from_wall(wall, q_doc)
    return perp


def host_wall_from_wall_foundation(wf):
    """Muro host de la zapata o ``None`` (API pública para otros módulos)."""
    if wf is None or not isinstance(wf, WallFoundation):
        return None
    try:
        wid = wf.WallId
    except Exception:
        wid = None
    if wid is None or wid == ElementId.InvalidElementId:
        return None
    try:
        w = wf.Document.GetElement(wid)
    except Exception:
        w = None
    if w is None or not isinstance(w, Wall):
        return None
    return w


def _tangent_and_perp_xy_from_wall(wall, q_doc):
    """
    Tangente unitaria en planta al ``LocationCurve`` del muro en el punto más
    cercano a ``q_doc`` (XY) y perpendicular (``perp``) alineada con
    ``Wall.Orientation`` cuando aplica.
    """
    if wall is None or q_doc is None:
        return None, None
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        return None, None
    crv = loc.Curve
    if crv is None:
        return None, None
    try:
        p_mid = crv.Evaluate(0.5, True)
        q = XYZ(float(q_doc.X), float(q_doc.Y), float(p_mid.Z))
        pr = crv.Project(q)
        if pr is None:
            return None, None
        par = float(pr.Parameter)
        dv = crv.ComputeDerivatives(par, True)
        if dv is None:
            return None, None
        tx = dv.BasisX
        tu = XYZ(float(tx.X), float(tx.Y), 0.0)
        if tu.GetLength() < 1e-12:
            return None, None
        tu = tu.Normalize()
    except Exception:
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            d = p1 - p0
            tu = XYZ(float(d.X), float(d.Y), 0.0)
            if tu.GetLength() < 1e-12:
                return None, None
            tu = tu.Normalize()
        except Exception:
            return None, None
    perp = XYZ(-float(tu.Y), float(tu.X), 0.0)
    perp = perp.Normalize()
    try:
        ow = wall.Orientation
        ow_xy = XYZ(float(ow.X), float(ow.Y), 0.0)
        if ow_xy.GetLength() > 1e-12:
            ow_xy = ow_xy.Normalize()
            if float(perp.DotProduct(ow_xy)) < 0.0:
                perp = perp.Negate()
    except Exception:
        pass
    return tu, perp


def _wf_location_midpoint(wf):
    loc = wf.Location
    if isinstance(loc, LocationCurve) and loc.Curve is not None:
        try:
            return loc.Curve.Evaluate(0.5, True)
        except Exception:
            pass
    try:
        bb = wf.get_BoundingBox(None)
        if bb is not None:
            return XYZ(
                0.5 * (float(bb.Min.X) + float(bb.Max.X)),
                0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
                0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
            )
    except Exception:
        pass
    return None


def _z_range_ft(elem):
    bb = elem.get_BoundingBox(None) if elem is not None else None
    if bb is None:
        return None, None
    return float(bb.Min.Z), float(bb.Max.Z)


def geometria_inferior_wall_foundation_cortes_muro(
    wf, diam_long_mm, diam_trans_mm
):
    """
    Construye el mismo diccionario que ``_geometria_wf_cara_inferior_tol`` en
    ``enfierrado_wall_foundation``, a partir del flujo muro → LocationCurve →
    planos → intersección con sólidos → curvas inferiores de sección.

    Returns:
        dict | None
    """
    if wf is None or not isinstance(wf, WallFoundation):
        return None
    wall = host_wall_from_wall_foundation(wf)
    if wall is None:
        return None
    pm = punto_centro_location_curve_muro(wall)
    if pm is None:
        pm = _wf_location_midpoint(wf)
    if pm is None:
        return None
    z0, z1 = _z_range_ft(wf)
    if z0 is None or z1 is None:
        return None
    cortes = planos_corte_verticales_desde_location_muro(wall, pm)
    if cortes is None:
        return None
    tu = cortes["tu"]
    perp = cortes["perp"]
    pl_long = cortes["pl_long"]
    pl_trans = cortes["pl_trans"]
    try:
        op = cortes.get("origen_planos")
        if op is not None:
            origin = XYZ(float(op.X), float(op.Y), float(op.Z))
        else:
            origin = XYZ(float(pm.X), float(pm.Y), float(pm.Z))
    except Exception:
        origin = XYZ(float(pm.X), float(pm.Y), float(pm.Z))

    solid_pairs = list(_iter_solids_wall_foundation(wf))
    if not solid_pairs:
        return None

    pts_long = _puntos_interseccion_plano_solidos_doc(solid_pairs, pl_long)
    pts_trans = _puntos_interseccion_plano_solidos_doc(solid_pairs, pl_trans)
    c_long = _linea_soleira_desde_puntos_interseccion(pts_long)
    c_width = _linea_soleira_desde_puntos_interseccion(pts_trans)
    if c_long is None:
        seg_long = _tramos_perimetro_interseccion_plano_solidos(solid_pairs, pl_long)
        c_long = _curva_inferior_desde_tramos_interseccion(seg_long)
    if c_width is None:
        seg_trans = _tramos_perimetro_interseccion_plano_solidos(solid_pairs, pl_trans)
        c_width = _curva_inferior_desde_tramos_interseccion(seg_trans)
    if c_long is None or c_width is None:
        return None
    c_long = _line_soleira_planta_horizontal(c_long)
    c_width = _line_soleira_planta_horizontal(c_width)
    if c_long is None or c_width is None:
        return None

    try:
        d_l = float(diam_long_mm) if diam_long_mm else 0.0
        d_t = float(diam_trans_mm) if diam_trans_mm else 0.0
    except Exception:
        d_l = d_t = 0.0
    ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM)
    if d_l > 1e-6:
        ext_long_mm = float(_REC_EXTREMOS_LONG_TANGENTE_MM) + 0.5 * d_l

    n_cara = XYZ.BasisZ.Negate()
    marco_uvn = (
        origin,
        tu,
        perp,
        n_cara,
    )

    ct_width, _ = aplicar_recubrimiento_inferior_completo_mm(
        c_width, wf, _REC_OFF_PLANTA_INF_MM, _REC_EXTREMOS_INFERIOR_MM
    )
    ct_long, _ = aplicar_recubrimiento_inferior_completo_mm(
        c_long, wf, _REC_OFF_PLANTA_INF_MM, ext_long_mm
    )
    if ct_width is None or ct_long is None:
        return None

    long_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_long, n_cara, _RECO_HOR_MM, d_l
    )
    width_bar = offset_linea_eje_barra_desde_cara_inferior_mm(
        ct_width, n_cara, _RECO_HOR_MM, d_t
    )
    if long_bar is None or width_bar is None:
        return None

    ev = evaluar_caras_paralelas_curva_mas_cercana(wf, long_bar)
    cara_pp = None
    if isinstance(ev, dict):
        cara_pp = ev.get("mejor")
        if cara_pp is None:
            cara_pp = ev.get(u"mejor")

    usable_w = float(width_bar.Length)
    return {
        "long_line": long_bar,
        "width_line": width_bar,
        "marco_uvn": marco_uvn,
        "cara_pp": cara_pp,
        "n_cara": n_cara,
        "z0": z0,
        "z1": z1,
        "usable_w_ft": usable_w,
        # No sustituir estas líneas en ``_wf_geo_alinear_strip_*``: son las del corte
        # muro/planos; las mismas deben alimentar ``CreateFromCurves`` / U.
        "use_cortes_lines_for_rebar": True,
    }

