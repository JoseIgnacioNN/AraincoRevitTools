# -*- coding: utf-8 -*-
"""
Cotas en planta para zapatas de muro (WallFoundation).

- Revit 2024+ | Motor: IronPython (pyRevit).
- Vista: planta (ViewPlan).
- Por cada WallFoundation en la vista, **dos** cotas alineadas (linea paralela al eje del muro):

  1) **Solo** las dos **caras laterales** de la soleira (caras paralelas al muro que definen el ancho).
  2) Las **mismas caras laterales** y, si corresponde, **todos los Grids** que sean **paralelos al eje
     del muro** y cuya representacion en planta **intersecte** la zapata (solape en planta con la
     huella de la fundacion).
"""

from __future__ import print_function

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Category,
    DimensionStyleType,
    DimensionType,
    DatumExtentType,
    ElementId,
    ElementReferenceType,
    ElementTypeGroup,
    FilteredElementCollector,
    GeometryInstance,
    Grid,
    Line,
    LocationCurve,
    Options,
    PlanarFace,
    Reference,
    ReferenceArray,
    Solid,
    Transaction,
    ViewDetailLevel,
    ViewPlan,
    Wall,
    WallFoundation,
    XYZ,
)

_MM_POR_PIE = 304.8
_TOL_PARALELO = 0.02
_TOL_NORMAL = 0.18
_MARGEN_LINEA_MM = 150.0
_OFFSET_ANCHO_MM = 450.0
_OFFSET_SEPARACION_TANGENTE_MM = 400.0
_TOL_INTERSECCION_PIES = 0.1


def _mm_a_pies(mm):
    return float(mm) / _MM_POR_PIE


def _z_plano_vista(view):
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


def _obtener_tipo_cota_alineada(doc):
    candidatos = []
    try:
        for dt in FilteredElementCollector(doc).OfClass(DimensionType):
            try:
                if dt.StyleType == DimensionStyleType.Linear:
                    candidatos.append(dt)
            except Exception:
                continue
    except Exception:
        pass
    if candidatos:
        return candidatos[0]
    try:
        tid = doc.GetDefaultElementTypeId(ElementTypeGroup.LinearDimensionType)
        if tid is not None and tid != ElementId.InvalidElementId:
            el = doc.GetElement(tid)
            if isinstance(el, DimensionType):
                try:
                    if el.StyleType == DimensionStyleType.Linear:
                        return el
                except Exception:
                    pass
    except Exception:
        pass
    return None


def _extraer_solidos(geom_elem):
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


def _tangente_planta_desde_curva(crv):
    if crv is None:
        return None
    for u in (0.5, 0.25, 0.75):
        try:
            dv = crv.ComputeDerivatives(u, True)
            if dv is None:
                continue
            tx = dv.BasisX
            v = XYZ(float(tx.X), float(tx.Y), 0.0)
            if v.GetLength() < 1e-12:
                continue
            return v.Normalize()
        except Exception:
            continue
    return None


def _tangente_desde_muro_host(wf, doc):
    """
    La zapata de muro sigue el eje del muro; en muchos proyectos solo el muro
    expone bien la curva de ubicacion (Location de WallFoundation puede fallar
    en IronPython con isinstance o venir vacia).
    """
    try:
        wid = wf.WallId
        if wid is None or wid == ElementId.InvalidElementId:
            return None
        w = doc.GetElement(wid)
        if not isinstance(w, Wall):
            return None
        loc = w.Location
        crv = getattr(loc, "Curve", None)
        if crv is None and isinstance(loc, LocationCurve):
            try:
                crv = loc.Curve
            except Exception:
                pass
        if crv is not None:
            return _tangente_planta_desde_curva(crv)
    except Exception:
        pass
    return None


def _tangente_desde_bbox_plano(bb):
    """
    Respaldo: eje del mayor lado del bounding box en XY (zapata alargada = eje del muro).
    Falla en muros muy diagonales respecto a X/Y; priorizar muro host.
    """
    if bb is None:
        return None
    try:
        dx = abs(float(bb.Max.X - bb.Min.X))
        dy = abs(float(bb.Max.Y - bb.Min.Y))
    except Exception:
        return None
    if dx < 1e-9 and dy < 1e-9:
        return None
    if dx >= dy:
        return XYZ(1.0, 0.0, 0.0)
    return XYZ(0.0, 1.0, 0.0)


def _cluster_key_normal(nxy):
    """Clave estable para agrupar n y -n (hemisferio: componente X positiva preferida)."""
    try:
        x, y = float(nxy.X), float(nxy.Y)
        if x < -1e-9 or (abs(x) <= 1e-9 and y < 0):
            x, y = -x, -y
        return (round(x, 3), round(y, 3))
    except Exception:
        return None


def _tangente_desde_caras_verticales(wf, view):
    """
    Sin Location valida: agrupa normales de caras verticales; la direccion con mayor
    separacion entre caras paralelas (largo de la zapata) es el eje del muro en planta.
    """
    solidos = []
    for usar_vista in (view, None):
        try:
            opts = Options()
            opts.ComputeReferences = True
            opts.DetailLevel = ViewDetailLevel.Fine
            if usar_vista is not None:
                opts.View = usar_vista
            solidos = _extraer_solidos(wf.get_Geometry(opts))
            if solidos:
                break
        except Exception:
            pass
    if not solidos:
        return None

    grupos = {}
    for solid in solidos:
        try:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                if face.Reference is None:
                    continue
                try:
                    n = face.FaceNormal
                    nxy = XYZ(float(n.X), float(n.Y), 0.0)
                    if nxy.GetLength() < 1e-9:
                        continue
                    nxy = nxy.Normalize()
                    o = face.Origin
                    key = _cluster_key_normal(nxy)
                    if key is None:
                        continue
                    n_can = XYZ(float(key[0]), float(key[1]), 0.0)
                    if n_can.GetLength() < 1e-9:
                        continue
                    n_can = n_can.Normalize()
                    proy = float(o.DotProduct(n_can))
                    if key not in grupos:
                        grupos[key] = {"n": n_can, "projs": []}
                    grupos[key]["projs"].append(proy)
                except Exception:
                    continue
        except Exception:
            continue

    best_span = -1.0
    best_tangent = None
    for key, gd in grupos.items():
        projs = gd.get("projs") or []
        n_can = gd.get("n")
        if len(projs) < 2 or n_can is None:
            continue
        try:
            span = max(projs) - min(projs)
        except Exception:
            continue
        if span > best_span:
            best_span = span
            best_tangent = n_can

    if best_tangent is not None and best_span > 1e-6:
        return best_tangent
    return None


def _tangente_wall_foundation(wf, doc):
    t = _tangente_desde_muro_host(wf, doc)
    if t is not None:
        return t
    try:
        loc = wf.Location
        crv = getattr(loc, "Curve", None)
        if crv is None and isinstance(loc, LocationCurve):
            try:
                crv = loc.Curve
            except Exception:
                crv = None
        if crv is not None:
            t = _tangente_planta_desde_curva(crv)
            if t is not None:
                return t
    except Exception:
        pass
    try:
        vista = doc.ActiveView
    except Exception:
        vista = None
    t = _tangente_desde_caras_verticales(wf, vista)
    if t is not None:
        return t
    bb = wf.get_BoundingBox(None)
    if bb is not None:
        t = _tangente_desde_bbox_plano(bb)
        if t is not None:
            return t
    return None


def _perpendicular_planta(tangent):
    return XYZ(-float(tangent.Y), float(tangent.X), 0.0).Normalize()


def _paralelos_xy(u, v):
    try:
        a = XYZ(float(u.X), float(u.Y), 0.0).Normalize()
        b = XYZ(float(v.X), float(v.Y), 0.0).Normalize()
        if a is None or b is None:
            return False
        cr = a.CrossProduct(b)
        return float(cr.GetLength()) < _TOL_PARALELO
    except Exception:
        return False


def _orden_datum_extent(grid, vista):
    orden = []
    try:
        ext = grid.GetDatumExtentTypeInView(vista)
        if ext is not None:
            orden.append(ext)
    except Exception:
        pass
    for ext in (DatumExtentType.ViewSpecific, DatumExtentType.Model):
        if ext not in orden:
            orden.append(ext)
    return orden


def _es_ref_superficie(ref):
    try:
        return ref.ElementReferenceType == ElementReferenceType.REFERENCE_TYPE_SURFACE
    except Exception:
        return False


def _refs_desde_curva(crv):
    salida = []
    try:
        r = crv.Reference
        if r is not None:
            salida.append(r)
    except Exception:
        pass
    for ei in (0, 1):
        try:
            r = crv.GetEndPointReference(ei)
            if r is not None:
                salida.append(r)
        except Exception:
            pass
    return salida


def _referencia_grid(grid, vista):
    if grid is None or vista is None:
        return None
    try:
        r = Reference(grid)
        if r is not None and not _es_ref_superficie(r):
            return r
    except Exception:
        pass
    for ext in _orden_datum_extent(grid, vista):
        try:
            curves = grid.GetCurvesInView(ext, vista)
        except Exception:
            continue
        if curves is None or curves.Count == 0:
            continue
        for i in range(int(curves.Count)):
            try:
                c = curves[i]
            except Exception:
                continue
            if c is None:
                continue
            for r in _refs_desde_curva(c):
                if r is not None and not _es_ref_superficie(r):
                    return r
    try:
        crv = grid.Curve
        if crv is not None:
            for r in _refs_desde_curva(crv):
                if r is not None and not _es_ref_superficie(r):
                    return r
    except Exception:
        pass
    try:
        r = Reference(grid)
        if r is not None:
            return r
    except Exception:
        pass
    return None


def _referencia_grid_completa(grid, vista):
    r = _referencia_grid(grid, vista)
    if r is not None:
        return r
    for i in range(0, 32):
        try:
            r = grid.ComputeReferenceBySubElement(i)
            if r is not None:
                return r
        except Exception:
            break
    return None


def _refs_ancho_por_caras(wf, view, tangent):
    width_dir = _perpendicular_planta(tangent)
    solidos = []
    for usar_vista in (view, None):
        try:
            opts = Options()
            opts.ComputeReferences = True
            opts.DetailLevel = ViewDetailLevel.Fine
            if usar_vista is not None:
                opts.View = usar_vista
            solidos = _extraer_solidos(wf.get_Geometry(opts))
            if solidos:
                break
        except Exception:
            pass
    if not solidos:
        return None, None, None, None, None

    candidatas = []
    for solid in solidos:
        try:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                ref = face.Reference
                if ref is None:
                    continue
                try:
                    n = face.FaceNormal
                except Exception:
                    continue
                nxy = XYZ(float(n.X), float(n.Y), 0.0)
                if nxy.GetLength() < 1e-9:
                    continue
                nxy = nxy.Normalize()
                if abs(float(nxy.DotProduct(tangent))) > _TOL_NORMAL:
                    continue
                try:
                    o = face.Origin
                    proj = float(o.DotProduct(width_dir))
                except Exception:
                    continue
                candidatas.append((proj, ref))
        except Exception:
            continue

    if len(candidatas) < 2:
        return None, None, None, None, None

    candidatas.sort(key=lambda t: t[0])
    proj_lo = candidatas[0][0]
    proj_hi = candidatas[-1][0]
    return (
        candidatas[0][1],
        candidatas[-1][1],
        proj_hi - proj_lo,
        proj_lo,
        proj_hi,
    )


def _rango_proyeccion_bbox(bb, axis):
    """Min y max de proyeccion de las esquinas del bbox sobre un eje unitario XY."""
    if bb is None:
        return None, None
    try:
        mn = bb.Min
        mx = bb.Max
        corners = [
            XYZ(mn.X, mn.Y, mn.Z),
            XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z),
            XYZ(mx.X, mx.Y, mn.Z),
            XYZ(mn.X, mn.Y, mx.Z),
            XYZ(mx.X, mn.Y, mx.Z),
            XYZ(mn.X, mx.Y, mx.Z),
            XYZ(mx.X, mx.Y, mx.Z),
        ]
    except Exception:
        return None, None
    projs = []
    for c in corners:
        try:
            projs.append(float(c.DotProduct(axis)))
        except Exception:
            pass
    if not projs:
        return None, None
    return min(projs), max(projs)


def _rango_grid_tangent_en_vista(grid, vista, tangent):
    """Proyeccion del tramo del grid visible en la vista sobre el eje del muro."""
    projs = []
    for ext in _orden_datum_extent(grid, vista):
        try:
            curves = grid.GetCurvesInView(ext, vista)
        except Exception:
            continue
        if curves is None or curves.Count == 0:
            continue
        for i in range(int(curves.Count)):
            try:
                c = curves[i]
            except Exception:
                continue
            if c is None:
                continue
            for u in (0.0, 1.0):
                try:
                    p = c.Evaluate(u, True)
                    projs.append(float(p.DotProduct(tangent)))
                except Exception:
                    pass
    if not projs:
        try:
            crv = grid.Curve
            if crv is not None:
                for u in (0.0, 0.5, 1.0):
                    try:
                        p = crv.Evaluate(u, True)
                        projs.append(float(p.DotProduct(tangent)))
                    except Exception:
                        pass
        except Exception:
            pass
    if not projs:
        return None, None
    return min(projs), max(projs)


def _intervalos_se_solapan(lo_a, hi_a, lo_b, hi_b, tol):
    try:
        a0, a1 = min(lo_a, hi_a), max(lo_a, hi_a)
        b0, b1 = min(lo_b, hi_b), max(lo_b, hi_b)
        return max(a0, b0) <= min(a1, b1) + float(tol)
    except Exception:
        return False


def _grids_paralelos_que_intersectan_zapata(
    doc, vista, tangent, width_dir, p_lo, p_hi, bb_foot
):
    """
    Grids paralelos al eje del muro (|| tangent) que en planta cortan la huella:
    - proyeccion constante sobre width_dir dentro de [p_lo, p_hi] (entre caras laterales);
    - solape del tramo del grid (en vista) con el rango del bbox de la zapata sobre tangent.
    """
    w0 = min(float(p_lo), float(p_hi))
    w1 = max(float(p_lo), float(p_hi))
    s_foot_lo, s_foot_hi = _rango_proyeccion_bbox(bb_foot, tangent)
    if s_foot_lo is None:
        return []

    tol = float(_TOL_INTERSECCION_PIES)
    salida = []
    vistos = set()

    try:
        grids = FilteredElementCollector(doc, vista.Id).OfClass(Grid)
    except Exception:
        return []

    for g in grids:
        if not isinstance(g, Grid):
            continue
        try:
            crv = g.Curve
            if crv is None:
                continue
            dv = crv.ComputeDerivatives(0.5, True)
            gd = dv.BasisX
        except Exception:
            continue
        if not _paralelos_xy(gd, tangent):
            continue
        try:
            pm = crv.Evaluate(0.5, True)
            w_g = float(pm.DotProduct(width_dir))
        except Exception:
            continue
        if w_g < w0 - tol or w_g > w1 + tol:
            continue
        tg_lo, tg_hi = _rango_grid_tangent_en_vista(g, vista, tangent)
        if tg_lo is None:
            continue
        if not _intervalos_se_solapan(tg_lo, tg_hi, s_foot_lo, s_foot_hi, tol):
            continue
        gid = g.Id.IntegerValue
        if gid in vistos:
            continue
        vistos.add(gid)
        r = _referencia_grid_completa(g, vista)
        if r is None:
            continue
        salida.append((w_g, r, g))

    salida.sort(key=lambda t: t[0])
    return salida


def _reference_array_cadena_lateral_y_grids(ref_lo, ref_hi, p_lo, p_hi, grids_ordenados):
    """
    Orden monotono en direccion de ancho: caras laterales + grids intermedios.
    grids_ordenados: lista de (w_g, ref, grid) ordenada por w_g.
    """
    items = [(ref_lo, float(p_lo)), (ref_hi, float(p_hi))]
    for w_g, r, _g in grids_ordenados:
        items.append((r, float(w_g)))
    items.sort(key=lambda t: t[1])
    ra = ReferenceArray()
    for ref, _pw in items:
        ra.Append(ref)
    return ra


def _bbox_centro_y_extensión_tangente(bb, center, tangent):
    try:
        mn = bb.Min
        mx = bb.Max
        corners = [
            XYZ(mn.X, mn.Y, mn.Z),
            XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z),
            XYZ(mx.X, mx.Y, mn.Z),
        ]
    except Exception:
        return 2.0
    ext = 0.0
    for c in corners:
        try:
            v = float(c.Subtract(center).DotProduct(tangent))
            av = abs(v)
            if av > ext:
                ext = av
        except Exception:
            continue
    return max(ext, 0.5)


def _new_dimension(doc, vista, linea, ra, dim_type):
    try:
        if dim_type is not None:
            return doc.Create.NewDimension(vista, linea, ra, dim_type)
    except Exception:
        pass
    try:
        return doc.Create.NewDimension(vista, linea, ra)
    except Exception:
        return None


def ejecutar_cotas(uidoc):
    from pyrevit import forms

    doc = uidoc.Document
    vista = doc.ActiveView
    if not isinstance(vista, ViewPlan):
        forms.alert(
            u"La vista activa debe ser una planta (ViewPlan).",
            title=u"Cotas zapata de muro",
        )
        return

    avisos = []

    try:
        cat = Category.GetCategory(doc, BuiltInCategory.OST_Grids)
        if cat is not None and vista.GetCategoryHidden(cat.Id):
            avisos.append(
                u"Aviso: Ejes ocultos en V/G; la cota parcial respecto al grid puede fallar."
            )
    except Exception:
        pass

    try:
        wfs = [
            e
            for e in FilteredElementCollector(doc, vista.Id).OfClass(WallFoundation)
            if isinstance(e, WallFoundation)
        ]
    except Exception:
        wfs = []

    if not wfs:
        forms.alert(
            u"No hay zapatas de muro (Wall Foundation) en la vista activa.",
            title=u"Cotas zapata de muro",
        )
        return

    dim_type = _obtener_tipo_cota_alineada(doc)
    if dim_type is None:
        forms.alert(
            u"No hay un tipo de cota alineada (Linear) en el proyecto.",
            title=u"Cotas zapata de muro",
        )
        return

    z = _z_plano_vista(vista)
    margen = _mm_a_pies(_MARGEN_LINEA_MM)
    off_ancho = _mm_a_pies(_OFFSET_ANCHO_MM)
    off_tan = _mm_a_pies(_OFFSET_SEPARACION_TANGENTE_MM)

    creadas = 0

    with Transaction(doc, u"BIMTools: Cotas zapata de muro en planta") as t:
        t.Start()
        for wf in wfs:
            tangent = _tangente_wall_foundation(wf, doc)
            if tangent is None:
                avisos.append(
                    u"Id {0}: sin curva de ubicacion valida.".format(
                        wf.Id.IntegerValue
                    )
                )
                continue

            (
                ref_lo,
                ref_hi,
                ancho_aprox,
                p_lo,
                p_hi,
            ) = _refs_ancho_por_caras(wf, vista, tangent)
            if ref_lo is None or ref_hi is None:
                avisos.append(
                    u"Id {0}: no se hallaron caras para el ancho.".format(
                        wf.Id.IntegerValue
                    )
                )
                continue

            width_dir = _perpendicular_planta(tangent)
            bb = wf.get_BoundingBox(vista)
            if bb is None:
                bb = wf.get_BoundingBox(None)
            if bb is None:
                avisos.append(
                    u"Id {0}: sin bounding box.".format(wf.Id.IntegerValue)
                )
                continue

            center = bb.Min.Add(bb.Max.Subtract(bb.Min).Multiply(0.5))
            ext_t = _bbox_centro_y_extensión_tangente(bb, center, tangent)
            half_line = ext_t + margen

            base_neg = (
                center.Add(tangent.Multiply(-off_tan)).Add(width_dir.Multiply(-off_ancho))
            )
            base_pos = (
                center.Add(tangent.Multiply(off_tan)).Add(width_dir.Multiply(-off_ancho))
            )

            p1a = base_neg.Add(tangent.Multiply(-half_line))
            p2a = base_neg.Add(tangent.Multiply(half_line))
            p1a = XYZ(float(p1a.X), float(p1a.Y), z)
            p2a = XYZ(float(p2a.X), float(p2a.Y), z)

            ra2 = ReferenceArray()
            ra2.Append(ref_lo)
            ra2.Append(ref_hi)
            linea_a = Line.CreateBound(p1a, p2a)
            dim_total = _new_dimension(doc, vista, linea_a, ra2, dim_type)
            if dim_total is not None:
                creadas += 1
            else:
                avisos.append(
                    u"Id {0}: fallo cota de ancho total.".format(wf.Id.IntegerValue)
                )

            if dim_total is None:
                continue

            grids_datos = _grids_paralelos_que_intersectan_zapata(
                doc, vista, tangent, width_dir, p_lo, p_hi, bb
            )
            if not grids_datos:
                avisos.append(
                    u"Id {0}: sin grids paralelos al muro que intersecten la zapata; "
                    u"solo se creo la cota de ancho entre caras laterales.".format(
                        wf.Id.IntegerValue
                    )
                )
                continue

            ra3 = _reference_array_cadena_lateral_y_grids(
                ref_lo, ref_hi, p_lo, p_hi, grids_datos
            )

            p1b = base_pos.Add(tangent.Multiply(-half_line))
            p2b = base_pos.Add(tangent.Multiply(half_line))
            p1b = XYZ(float(p1b.X), float(p1b.Y), z)
            p2b = XYZ(float(p2b.X), float(p2b.Y), z)
            linea_b = Line.CreateBound(p1b, p2b)
            if _new_dimension(doc, vista, linea_b, ra3, dim_type) is not None:
                creadas += 1
            else:
                avisos.append(
                    u"Id {0}: fallo la 2. cota (caras laterales + grids).".format(
                        wf.Id.IntegerValue
                    )
                )

        t.Commit()

    msg = u"Se crearon {0} cota(s).".format(creadas)
    if avisos:
        msg += u"\n\n" + u"\n".join(avisos[:12])
        if len(avisos) > 12:
            msg += u"\n..."
    forms.alert(msg, title=u"Cotas zapata de muro")
    try:
        print(msg)
    except Exception:
        pass
