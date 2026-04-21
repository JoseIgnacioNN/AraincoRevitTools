# -*- coding: utf-8 -*-
"""
Vista en sección transversal al eje de la viga (corte por el punto medio del Location)
para revisar armadura longitudinal y estribos tras enfierrado.
"""

try:
    unicode
except NameError:
    unicode = str

import unicodedata

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    ElementTransformUtils,
    FilterElement,
    FilteredElementCollector,
    GeometryInstance,
    Level,
    Line,
    LocationCurve,
    Options,
    Solid,
    StorageType,
    Transaction,
    Transform,
    View,
    ViewDetailLevel,
    ViewFamily,
    ViewFamilyType,
    ViewPlan,
    ViewSection,
    ViewType,
    Wall,
    WallFoundation,
    XYZ,
    UnitUtils,
    UnitTypeId,
)

# Si es False, ``crear_vistas_seccion_revision_wall_foundation`` no crea vistas (solo zapata corrida).
CREAR_SECCION_REVISION_WALL_FOUNDATION = True

# Far Clip Offset por defecto en vistas de sección generadas (mm).
_FAR_CLIP_OFFSET_MM_DEFECTO = 100.0

# Mismo prefijo que ``_asignar_nombre_vista_unico`` / creación (para limpieza por nombre).
_NOMBRE_PREFIJO_SECCION_REVISION = u"BIMTools — Rev. enfierrado —"

# Muestreo de geometría para caja de sección alineada al elemento (no AABB mundo).
_GEOM_SAMPLE_POINTS_MAX = 2048


def _element_id_a_int(eid):
    if eid is None:
        return None
    try:
        return int(eid.Value)
    except Exception:
        try:
            return int(eid.IntegerValue)
        except Exception:
            return None


def _es_vista_seccion_revision(el):
    if el is None:
        return False
    try:
        if isinstance(el, ViewSection):
            return True
    except Exception:
        pass
    try:
        if getattr(el, "ViewType", None) == ViewType.Section:
            return True
    except Exception:
        pass
    return False


def _mm_a_interno(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _primer_view_family_type_seccion(document):
    col = FilteredElementCollector(document).OfClass(ViewFamilyType)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    for vft in col:
        try:
            if vft is not None and vft.ViewFamily == ViewFamily.Section:
                return vft.Id
        except Exception:
            continue
    return None


def _normalize_compare_name(value):
    if value is None:
        return ""
    s = str(value).strip()
    try:
        return unicodedata.normalize("NFC", s)
    except Exception:
        return s


def _view_family_type_display_name(vft):
    if vft is None:
        return ""
    try:
        n = vft.Name
        if n is not None:
            s = _normalize_compare_name(n)
            if s:
                return s
    except Exception:
        pass
    for bip in (
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
        BuiltInParameter.SYMBOL_NAME_PARAM,
    ):
        try:
            p = vft.get_Parameter(bip)
            if p and p.HasValue:
                s = _normalize_compare_name(p.AsString())
                if s:
                    return s
        except Exception:
            continue
    return ""


def _iter_view_family_types_section(document):
    col = FilteredElementCollector(document).OfClass(ViewFamilyType)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    for vft in col:
        try:
            if vft is not None and vft.ViewFamily == ViewFamily.Section:
                yield vft
        except Exception:
            continue


def _find_view_family_type_section_by_name(document, exact_name):
    """
    Busca ``ViewFamilyType`` con ``ViewFamily.Section``:

    1. Nombre exacto (normalizado).
    2. Mismo nombre sin distinguir mayúsculas.
    3. El **nombre del tipo contiene** el texto del filtro (p. ej. ``Building Section (00_PG_GENERAL)``
       para ``Section Filter`` = ``00_PG_GENERAL``). Si hay varias, se usa la de **nombre más corto**.
    """
    target = _normalize_compare_name(exact_name)
    if not target:
        return None, u"«Section Filter» no tiene texto válido para buscar el tipo de sección."
    for vft in _iter_view_family_types_section(document):
        try:
            if _view_family_type_display_name(vft) == target:
                return vft.Id, None
        except Exception:
            continue
    tl = target.lower()
    for vft in _iter_view_family_types_section(document):
        try:
            if _view_family_type_display_name(vft).lower() == tl:
                return vft.Id, None
        except Exception:
            continue
    contains_matches = []
    for vft in _iter_view_family_types_section(document):
        try:
            n = _view_family_type_display_name(vft)
            if not n:
                continue
            nl = n.lower()
            if tl in nl:
                contains_matches.append((len(n), vft))
        except Exception:
            continue
    if contains_matches:
        contains_matches.sort(key=lambda x: x[0])
        return contains_matches[0][1].Id, None
    sample = []
    for vft in _iter_view_family_types_section(document):
        try:
            n = _view_family_type_display_name(vft)
            if n:
                sample.append(n)
        except Exception:
            continue
    sample = sorted(set(sample))[:12]
    msg = (
        u"No se encontró un tipo de vista «Section» cuyo nombre sea «{0}» o contenga ese texto. "
        u"Defina un ViewFamilyType de sección acorde o ajuste «Section Filter»."
    ).format(target)
    if sample:
        msg += u" Ejemplos en el proyecto: {0}.".format(u", ".join(sample))
    return None, msg


def _resolver_view_family_type_seccion_desde_vista(document, view):
    """
    Usa el parámetro de instancia «Section Filter» de ``view`` para obtener un
    ``ElementId`` de ``ViewFamilyType`` (familia Section).

    - Si el parámetro guarda texto: se busca por nombre exacto o por **subcadena** (el nombre del
      tipo contiene el texto del filtro; p. ej. prefijo ``Building Section (… )``).
    - Si guarda ``ElementId`` a un ``ViewFamilyType`` Section: se usa directamente.
    - Si guarda ``ElementId`` a un ``FilterElement``: se usa el nombre del filtro
      para buscar el ``ViewFamilyType`` Section homónimo.

    Returns:
        ``(ElementId, None)`` o ``(None, mensaje_error)``.
    """
    if document is None or view is None:
        return None, u"Vista o documento no válidos."
    p = view.LookupParameter(u"Section Filter")
    if p is None:
        return None, (
            u"No se encontró el parámetro «Section Filter» en la vista activa. "
            u"Debe existir en la categoría Vistas (instancia), como en las plantas de revisión."
        )
    if not p.HasValue:
        return None, u"«Section Filter» está vacío en la vista activa."

    try:
        if p.StorageType == StorageType.ElementId:
            eid = p.AsElementId()
            if eid is None or eid == ElementId.InvalidElementId:
                return None, u"«Section Filter» no tiene referencia válida."
            el = document.GetElement(eid)
            if el is None:
                return None, u"No se encontró el elemento referenciado por «Section Filter»."
            if isinstance(el, ViewFamilyType):
                try:
                    if el.ViewFamily == ViewFamily.Section:
                        return el.Id, None
                except Exception:
                    pass
                return None, u"«Section Filter» no apunta a un tipo de vista «Section»."
            if isinstance(el, FilterElement):
                try:
                    fn = _normalize_compare_name(getattr(el, "Name", None))
                except Exception:
                    fn = u""
                if not fn:
                    return None, u"El filtro referenciado por «Section Filter» no tiene nombre."
                return _find_view_family_type_section_by_name(document, fn)
            return None, (
                u"«Section Filter» apunta a un elemento que no es tipo Section ni FilterElement; "
                u"use texto o referencia a tipo de sección / filtro homónimo."
            )

        if p.StorageType == StorageType.String:
            s = p.AsString()
            if s is None or not str(s).strip():
                return None, u"«Section Filter» (texto) está vacío."
            return _find_view_family_type_section_by_name(document, s)

        vs = None
        try:
            vs = p.AsValueString()
        except Exception:
            pass
        if vs and str(vs).strip():
            return _find_view_family_type_section_by_name(document, str(vs).strip())
    except Exception as ex:
        try:
            return None, u"Error al leer «Section Filter»: {0}".format(unicode(ex))
        except Exception:
            return None, u"Error al leer «Section Filter»."

    return None, u"No se pudo interpretar el valor de «Section Filter» (tipo de almacenamiento no soportado)."


def _linea_acotada_en_bbox(bb, origin, direction_unit):
    """
    Segmento en la recta ``origin + t * direction`` que cubre la proyección de las esquinas
    de ``bb`` (``BoundingBoxXYZ`` con ``Min``/``Max`` en coordenadas modelo).
    """
    if bb is None or bb.Min is None or bb.Max is None or origin is None or direction_unit is None:
        return None
    try:
        d = direction_unit
        if float(d.GetLength()) < 1e-12:
            return None
        d = d.Normalize()
        o = origin
        corners = (
            XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
            XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
            XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
            XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
            XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
            XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
            XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
            XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
        )
        ts = []
        for c in corners:
            try:
                ts.append(float((c - o).DotProduct(d)))
            except Exception:
                continue
        if len(ts) < 2:
            return None
        t0, t1 = min(ts), max(ts)
        if t1 - t0 < 1e-9:
            return None
        return Line.CreateBound(o + d.Multiply(t0), o + d.Multiply(t1))
    except Exception:
        return None


def _linea_acotada_en_bbox_para_direccion(elem, origin, direction_unit):
    """
    Igual que :func:`_linea_acotada_en_bbox` usando ``elem.get_BoundingBox(None)``.
    """
    if elem is None:
        return None
    try:
        bb = elem.get_BoundingBox(None)
        return _linea_acotada_en_bbox(bb, origin, direction_unit)
    except Exception:
        return None


def _expandir_extremos_bbox(mn, mx, bb):
    """Amplía ``(mn, mx)`` con las 8 esquinas de ``bb``; retorna nuevo par o None si bb inválido."""
    if bb is None or bb.Min is None or bb.Max is None:
        return mn, mx
    corners = (
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    )
    try:
        for c in corners:
            if mn is None:
                mn = XYZ(float(c.X), float(c.Y), float(c.Z))
                mx = XYZ(float(c.X), float(c.Y), float(c.Z))
            else:
                mn = XYZ(
                    min(float(mn.X), float(c.X)),
                    min(float(mn.Y), float(c.Y)),
                    min(float(mn.Z), float(c.Z)),
                )
                mx = XYZ(
                    max(float(mx.X), float(c.X)),
                    max(float(mx.Y), float(c.Y)),
                    max(float(mx.Z), float(c.Z)),
                )
    except Exception:
        return mn, mx
    return mn, mx


def _union_bounding_box_solids(elem):
    """
    Caja mínima eje-alineada que envuelve los ``Solid`` del elemento (modelo). Para
    ``WallFoundation``, ``get_BoundingBox(None)`` suele ser el del **muro host**; los sólidos
    de la zapata dan un recorte acotado al volumen real.
    """
    if elem is None:
        return None
    opt = Options()
    try:
        opt.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opt.ComputeReferences = False
    except Exception:
        pass
    try:
        geom = elem.get_Geometry(opt)
    except Exception:
        geom = None
    if geom is None:
        return None
    # Lista mutable (sin ``nonlocal``, incompatible con IronPython 2.7 / pyRevit clásico).
    _mn_mx = [None, None]

    def absorb_solid(solid):
        if solid is None:
            return
        try:
            if float(solid.Volume) <= 1e-12:
                return
        except Exception:
            pass
        try:
            bb = solid.GetBoundingBox()
        except Exception:
            bb = None
        _mn_mx[0], _mn_mx[1] = _expandir_extremos_bbox(_mn_mx[0], _mn_mx[1], bb)

    try:
        for obj in geom:
            if isinstance(obj, Solid):
                absorb_solid(obj)
            elif isinstance(obj, GeometryInstance):
                try:
                    sub = obj.GetInstanceGeometry()
                except Exception:
                    sub = None
                if sub is None:
                    continue
                try:
                    for o2 in sub:
                        if isinstance(o2, Solid):
                            absorb_solid(o2)
                except Exception:
                    pass
    except Exception:
        return None
    mn, mx = _mn_mx[0], _mn_mx[1]
    if mn is None or mx is None:
        return None
    try:
        out = BoundingBoxXYZ()
        out.Min = mn
        out.Max = mx
        return out
    except Exception:
        return None


def _bbox_seccion_desde_solida_o_elemento(elem):
    """
    Prioridad: unión de bboxes de sólidos del elemento; si no hay sólidos útiles,
    ``get_BoundingBox(None)``.
    """
    if elem is None:
        return None
    u = _union_bounding_box_solids(elem)
    if u is not None and u.Min is not None and u.Max is not None:
        try:
            dx = float(u.Max.X) - float(u.Min.X)
            dy = float(u.Max.Y) - float(u.Min.Y)
            dz = float(u.Max.Z) - float(u.Min.Z)
            if dx > 1e-9 or dy > 1e-9 or dz > 1e-9:
                return u
        except Exception:
            pass
    try:
        return elem.get_BoundingBox(None)
    except Exception:
        return None


def _collect_geometry_vertex_sample_points(elem, max_points=None):
    """
    Puntos 3D (modelo) sobre aristas de ``Solid`` del elemento: extremos de cada arista.
    Sirve para acotar la caja de sección en el **triedro local** del corte sin usar solo el
    AABB alineado a ejes globales (que envuelve mal geometrías rotadas / alargadas).
    """
    cap = int(max_points) if max_points is not None else _GEOM_SAMPLE_POINTS_MAX
    if cap < 8:
        cap = 8
    pts = []
    opt = Options()
    try:
        opt.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opt.ComputeReferences = False
    except Exception:
        pass
    try:
        geom = elem.get_Geometry(opt)
    except Exception:
        geom = None
    if geom is None:
        return pts

    def harvest_solid(solid):
        if solid is None or len(pts) >= cap:
            return
        try:
            if float(solid.Volume) <= 1e-12:
                return
        except Exception:
            pass
        try:
            for edge in solid.Edges:
                if len(pts) >= cap:
                    return
                crv = edge.AsCurve()
                if crv is None:
                    continue
                for u in (0.0, 1.0):
                    if len(pts) >= cap:
                        return
                    try:
                        pts.append(crv.Evaluate(u, True))
                    except Exception:
                        pass
        except Exception:
            pass

    try:
        for obj in geom:
            if isinstance(obj, Solid):
                harvest_solid(obj)
            elif isinstance(obj, GeometryInstance):
                try:
                    sub = obj.GetInstanceGeometry()
                except Exception:
                    sub = None
                if sub is None:
                    continue
                try:
                    for o2 in sub:
                        if isinstance(o2, Solid):
                            harvest_solid(o2)
                except Exception:
                    pass
    except Exception:
        pass
    return pts


def _centro_aprox_puntos_xyz(puntos):
    if not puntos:
        return None
    n = 0
    sx = sy = sz = 0.0
    for p in puntos:
        try:
            sx += float(p.X)
            sy += float(p.Y)
            sz += float(p.Z)
            n += 1
        except Exception:
            pass
    if n < 1:
        return None
    try:
        return XYZ(sx / float(n), sy / float(n), sz / float(n))
    except Exception:
        return None


def _linea_acotada_puntos_proyectados(puntos, origin, direction_unit):
    """
    Segmento en ``origin + t * direction`` que cubre la proyección escalar de ``puntos``.
    """
    if not puntos or origin is None or direction_unit is None:
        return None
    try:
        d = direction_unit
        if float(d.GetLength()) < 1e-12:
            return None
        d = d.Normalize()
        o = origin
        ts = []
        for c in puntos:
            try:
                ts.append(float((c - o).DotProduct(d)))
            except Exception:
                continue
        if len(ts) < 1:
            return None
        t0, t1 = min(ts), max(ts)
        if t1 - t0 < 1e-9:
            return None
        return Line.CreateBound(o + d.Multiply(t0), o + d.Multiply(t1))
    except Exception:
        return None


def _origen_y_linea_eje_seccion_en_medio_location(elem):
    """
    Origen de la sección en el **centro del volumen del host a lo largo del eje** de la
    ``LocationCurve``: se proyectan las esquinas del bbox sobre la tangente y se toma el
    punto medio de ese tramo. Así el corte no queda fuera de la zapata cuando la curva de
    ubicación acotada es más larga que el sólido (caso habitual en ``WallFoundation``).

    Para curvas **no acotadas** (línea infinita), se acota con el bbox como en
    :func:`_linea_desde_location_curve_host``.

    Returns:
        ``(origen, linea_eje)`` o ``(None, None)`` si no aplica.
    """
    if elem is None:
        return None, None
    loc = elem.Location
    if loc is None:
        return None, None
    crv = None
    if isinstance(loc, LocationCurve):
        crv = loc.Curve
    else:
        crv = getattr(loc, "Curve", None)
    if crv is None:
        return None, None
    try:
        is_bound = bool(crv.IsBound)
    except Exception:
        is_bound = True
    if not is_bound:
        try:
            if isinstance(crv, Line):
                dire = crv.Direction
                orig = getattr(crv, "Origin", None)
                if orig is None:
                    try:
                        orig = crv.Evaluate(0.0, True)
                    except Exception:
                        orig = None
                if orig is not None and dire is not None:
                    ln_bb = _linea_acotada_en_bbox_para_direccion(elem, orig, dire)
                    if ln_bb is not None:
                        p0 = ln_bb.GetEndPoint(0)
                        p1 = ln_bb.GetEndPoint(1)
                        seg = p1 - p0
                        lu = float(seg.GetLength())
                        if lu < 1e-9:
                            return None, None
                        origen = p0 + seg.Multiply(0.5)
                        t = seg.Normalize()
                        return origen, Line.CreateBound(origen, origen + t)
        except Exception:
            pass
        return None, None
    t = None
    try:
        der = crv.ComputeDerivatives(0.5, True)
        if der is not None:
            tx = der.BasisX
            if float(tx.GetLength()) > 1e-9:
                t = tx.Normalize()
    except Exception:
        pass
    if t is None:
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            dh = p1 - p0
            if float(dh.GetLength()) < 1e-9:
                return None, None
            t = dh.Normalize()
        except Exception:
            return None, None
    ref = None
    try:
        ref = crv.Evaluate(0.5, True)
    except Exception:
        pass
    if ref is None:
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            ref = p0 + (p1 - p0) * 0.5
        except Exception:
            ref = None
    if ref is not None and t is not None:
        ln_bb = _linea_acotada_en_bbox_para_direccion(elem, ref, t)
        if ln_bb is not None:
            try:
                p0 = ln_bb.GetEndPoint(0)
                p1 = ln_bb.GetEndPoint(1)
                seg = p1 - p0
                if float(seg.GetLength()) > 1e-9:
                    origen = p0 + seg.Multiply(0.5)
                    t2 = seg.Normalize()
                    return origen, Line.CreateBound(origen, origen + t2)
            except Exception:
                pass
    try:
        if ref is None:
            return None, None
        return ref, Line.CreateBound(ref, ref + t)
    except Exception:
        return None, None


def _origen_seccion_medio_eje_wall_foundation(elem):
    """
    Punto medio a lo largo del **eje longitudinal** de la zapata (mitad del largo real).

    **Prioridad:** geometría de sólidos — unión de bbox de sólidos + tangente de
    ``LocationCurve``; el tramo de esa dirección que atraviesa el bbox se recorta con
    :func:`_linea_acotada_en_bbox` y el origen es el **punto medio de ese tramo**. Así el corte
    queda en el centro del volumen aunque la curva de ubicación esté desfasada o acotada de
    forma distinta al sólido (caso habitual con muros largos).

    **Respaldo:** ``LocationCurve`` acotada (``Evaluate(0.5)``) o recta no acotada recortada al
    bbox de sólidos (lógica previa).
    """
    if elem is None or not isinstance(elem, WallFoundation):
        return None
    try:
        bb_sol = _bbox_seccion_desde_solida_o_elemento(elem)
        muro_dir = _tangente_longitudinal_desde_location(elem, bb_sol)
        if (
            bb_sol is not None
            and bb_sol.Min is not None
            and bb_sol.Max is not None
            and muro_dir is not None
        ):
            mid_bb = _centro_bounding_box_xyz(bb_sol)
            if mid_bb is not None:
                ln = _linea_acotada_en_bbox(bb_sol, mid_bb, muro_dir)
                if ln is not None:
                    p0 = ln.GetEndPoint(0)
                    p1 = ln.GetEndPoint(1)
                    seg = p1.Subtract(p0)
                    if float(seg.GetLength()) > 1e-9:
                        return p0.Add(seg.Multiply(0.5))
    except Exception:
        pass
    loc = elem.Location
    if loc is None:
        return None
    crv = None
    if isinstance(loc, LocationCurve):
        crv = loc.Curve
    else:
        crv = getattr(loc, "Curve", None)
    if crv is None:
        return None
    try:
        is_bound = bool(crv.IsBound)
    except Exception:
        is_bound = True
    if is_bound:
        try:
            return crv.Evaluate(0.5, True)
        except Exception:
            try:
                p0 = crv.GetEndPoint(0)
                p1 = crv.GetEndPoint(1)
                return p0.Add(p1.Subtract(p0).Multiply(0.5))
            except Exception:
                return None
    try:
        if isinstance(crv, Line):
            dire = crv.Direction
            if dire is None or float(dire.GetLength()) < 1e-12:
                return None
            dire = dire.Normalize()
            orig = getattr(crv, "Origin", None)
            if orig is None:
                try:
                    orig = crv.Evaluate(0.0, True)
                except Exception:
                    orig = None
            bb = _bbox_seccion_desde_solida_o_elemento(elem)
            if orig is not None and bb is not None:
                ln_bb = _linea_acotada_en_bbox(bb, orig, dire)
                if ln_bb is not None:
                    p0 = ln_bb.GetEndPoint(0)
                    p1 = ln_bb.GetEndPoint(1)
                    seg = p1.Subtract(p0)
                    if float(seg.GetLength()) > 1e-9:
                        return p0.Add(seg.Multiply(0.5))
    except Exception:
        pass
    return None


def _linea_desde_location_curve_host(elem):
    """
    Elemento con ``LocationCurve``: viga ``Structural Framing``, ``WallFoundation``, etc.
    Algunos hosts (p. ej. zapata) exponen ``Curve`` **no acotada**; se acota con el bbox.
    """
    loc = elem.Location
    if loc is None:
        return None
    crv = None
    if isinstance(loc, LocationCurve):
        crv = loc.Curve
    else:
        crv = getattr(loc, "Curve", None)
    if crv is None:
        return None
    try:
        is_bound = bool(crv.IsBound)
    except Exception:
        is_bound = True
    if not is_bound:
        try:
            if isinstance(crv, Line):
                dire = crv.Direction
                orig = getattr(crv, "Origin", None)
                if orig is None:
                    try:
                        orig = crv.Evaluate(0.0, True)
                    except Exception:
                        orig = None
                if orig is not None and dire is not None:
                    ln_bb = _linea_acotada_en_bbox_para_direccion(elem, orig, dire)
                    if ln_bb is not None:
                        return ln_bb
        except Exception:
            pass
        return None
    if isinstance(crv, Line):
        return crv
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


# Alias histórico (solo vigas en nombre).
_linea_desde_location_viga = _linea_desde_location_curve_host


def _tangente_longitudinal_desde_location(elem, bb=None):
    """
    Solo la **dirección** del eje según ``LocationCurve`` (no un punto del muro largo).

    Si la curva es **no acotada**, se usa ``bb`` (o :func:`_bbox_seccion_desde_solida_o_elemento`)
    para acotar la dirección al tramo que atraviesa el volumen de la zapata.

    Returns:
        ``XYZ`` unitario o ``None``.
    """
    if elem is None:
        return None
    loc = elem.Location
    if loc is None:
        return None
    crv = None
    if isinstance(loc, LocationCurve):
        crv = loc.Curve
    else:
        crv = getattr(loc, "Curve", None)
    if crv is None:
        return None
    try:
        is_bound = bool(crv.IsBound)
    except Exception:
        is_bound = True
    if not is_bound:
        try:
            if not isinstance(crv, Line):
                return None
            dire = crv.Direction
            if dire is None or float(dire.GetLength()) < 1e-12:
                return None
            dire = dire.Normalize()
            bb_loc = bb if bb is not None else _bbox_seccion_desde_solida_o_elemento(elem)
            if bb_loc is None or bb_loc.Min is None or bb_loc.Max is None:
                return None
            mid_bb = _centro_bounding_box_xyz(bb_loc)
            if mid_bb is None:
                return None
            ln = _linea_acotada_en_bbox(bb_loc, mid_bb, dire)
            if ln is None:
                return None
            p0, p1 = ln.GetEndPoint(0), ln.GetEndPoint(1)
            dh = p1.Subtract(p0)
            if float(dh.GetLength()) < 1e-9:
                return None
            return dh.Normalize()
        except Exception:
            return None
    vz = None
    try:
        der = crv.ComputeDerivatives(0.5, True)
        if der is not None:
            tx = der.BasisX
            if float(tx.GetLength()) > 1e-9:
                vz = tx.Normalize()
    except Exception:
        pass
    if vz is None:
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            dh = p1.Subtract(p0)
            if float(dh.GetLength()) < 1e-9:
                return None
            vz = dh.Normalize()
        except Exception:
            return None
    return vz


def _transform_seccion_desde_geometria_arista_horizontal_y_centro_bbox(elem, bb=None):
    """
    **Único método** de orientación para la sección de revisión de ``WallFoundation``.

    - **Bounding box auxiliar** — si ``bb`` es ``None``, :func:`_bbox_seccion_desde_solida_o_elemento`
      (sólidos de la zapata; respaldo para tangente y AABB si no hay muestras).
    - **Dirección (prioridad)** — tangente de la ``LocationCurve`` (curva no acotada → línea
      acotada con ``bb`` o puntos de geometría).
    - **Respaldo** — arista ``Line`` horizontal larga por geometría; luego ejes del bbox.
    - **Origen (prioridad)** — :func:`_origen_seccion_medio_eje_wall_foundation`: mitad del eje de
      la ``LocationCurve`` (``Evaluate(0.5)`` si está acotada; si no, recorte con bbox de
      sólidos). Luego :func:`_origen_y_linea_eje_seccion_en_medio_location`. Respaldo: geometría.
    - Triedro: ``vec_x = muro_dir × Z``, ``vec_y = Z``, ``vec_z = vec_x × vec_y`` (alineado con
      ``muro_dir``).
    """
    if elem is None:
        return None
    if bb is None:
        bb = _bbox_seccion_desde_solida_o_elemento(elem)
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    geom_pts = _collect_geometry_vertex_sample_points(elem)
    mid_pt = _centro_aprox_puntos_xyz(geom_pts)
    if mid_pt is None:
        mid_pt = _centro_bounding_box_xyz(bb)
    if mid_pt is None:
        return None

    muro_dir = _tangente_longitudinal_desde_location(elem, bb)

    if muro_dir is None:
        state = [None, 0.0]

        def consider_edge(crv):
            if crv is None or not isinstance(crv, Line):
                return
            try:
                ln = float(crv.Length)
            except Exception:
                return
            try:
                v = (crv.GetEndPoint(1) - crv.GetEndPoint(0)).Normalize()
            except Exception:
                return
            try:
                if abs(float(v.Z)) >= 0.1:
                    return
            except Exception:
                return
            if ln > state[1]:
                state[1] = ln
                state[0] = v

        opt = Options()
        try:
            opt.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        try:
            geom = elem.get_Geometry(opt)
        except Exception:
            geom = None
        if geom is not None:
            try:
                for obj in geom:
                    if isinstance(obj, Solid):
                        try:
                            for edge in obj.Edges:
                                consider_edge(edge.AsCurve())
                        except Exception:
                            pass
                    elif isinstance(obj, GeometryInstance):
                        try:
                            for sub in obj.GetInstanceGeometry():
                                if isinstance(sub, Solid):
                                    try:
                                        for edge in sub.Edges:
                                            consider_edge(edge.AsCurve())
                                    except Exception:
                                        pass
                        except Exception:
                            pass
            except Exception:
                pass

        muro_dir = state[0]
        if muro_dir is None:
            try:
                size = bb.Max - bb.Min
                sx = abs(float(size.X))
                sy = abs(float(size.Y))
                muro_dir = XYZ.BasisX if sx >= sy else XYZ.BasisY
            except Exception:
                return None

    up = XYZ.BasisZ
    try:
        vec_x = muro_dir.CrossProduct(up)
        if float(vec_x.GetLength()) < 1e-9:
            return None
        vec_x = vec_x.Normalize()
        vec_y = up
        vec_z = vec_x.CrossProduct(vec_y).Normalize()
    except Exception:
        return None
    try:
        if float(vec_z.DotProduct(muro_dir)) < 0.0:
            vec_z = vec_z.Negate()
            vec_x = vec_x.Negate()
    except Exception:
        pass
    tr = Transform.Identity
    tr.BasisX = vec_x
    tr.BasisY = vec_y
    tr.BasisZ = vec_z
    o_eje = None
    try:
        o_eje = _origen_seccion_medio_eje_wall_foundation(elem)
    except Exception:
        o_eje = None
    if o_eje is None:
        try:
            o_eje, _ln_eje = _origen_y_linea_eje_seccion_en_medio_location(elem)
        except Exception:
            o_eje = None
    if o_eje is not None:
        tr.Origin = o_eje
    else:
        tr.Origin = mid_pt
        try:
            if len(geom_pts) >= 3:
                ln_bb = _linea_acotada_puntos_proyectados(geom_pts, mid_pt, vec_z)
            else:
                ln_bb = _linea_acotada_en_bbox(bb, mid_pt, vec_z)
            if ln_bb is not None:
                p0 = ln_bb.GetEndPoint(0)
                p1 = ln_bb.GetEndPoint(1)
                seg = p1.Subtract(p0)
                if float(seg.GetLength()) > 1e-9:
                    tr.Origin = p0.Add(seg.Multiply(0.5))
        except Exception:
            pass
    return tr


def _transform_seccion_transversal_punto_medio(line):
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    origen = p0 + (p1 - p0) * 0.5
    dh = p1 - p0
    if dh.GetLength() < 1e-9:
        return None
    t = dh.Normalize()
    z = XYZ.BasisZ
    h = z.CrossProduct(t)
    if h.GetLength() < 1e-6:
        h = XYZ.BasisX.CrossProduct(t)
    if h.GetLength() < 1e-6:
        return None
    h = h.Normalize()
    v = t.CrossProduct(h).Normalize()
    tr = Transform.Identity
    tr.Origin = origen
    tr.BasisX = h
    tr.BasisY = v
    tr.BasisZ = t
    return tr


def _centro_bounding_box_xyz(bb):
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    try:
        return XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
        )
    except Exception:
        return None


def _punto_medio_segmento_linea(line):
    """Punto medio geométrico de una ``Line`` acotada (eje de zapata)."""
    if line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        dh = p1.Subtract(p0)
        if float(dh.GetLength()) < 1e-9:
            return None
        return p0.Add(dh.Multiply(0.5))
    except Exception:
        return None


def _linea_eje_wall_foundation_span_vertices_sobre_location(elem):
    """
    Segmento longitudinal = dirección del ``LocationCurve`` acotado del elemento y extensión
    de las proyecciones de **todos** los vértices de geometría sobre ese eje. Evita usar el par
    de puntos más separados en XY (diagonal de marco con hueco, etc.).
    """
    if elem is None or not isinstance(elem, WallFoundation):
        return None
    pts = _collect_geometry_vertex_sample_points(elem)
    if len(pts) < 2:
        return None
    loc = elem.Location
    if not isinstance(loc, LocationCurve) or loc.Curve is None:
        return None
    crv = loc.Curve
    try:
        if not bool(crv.IsBound):
            return None
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        dh = p1.Subtract(p0)
        lu = float(dh.GetLength())
        if lu < 1e-9:
            return None
        u = dh.Normalize()
    except Exception:
        return None
    ref = _centro_aprox_puntos_xyz(pts)
    if ref is None:
        try:
            ref = pts[0]
        except Exception:
            return None
    ts = []
    for p in pts:
        try:
            ts.append(float((p - ref).DotProduct(u)))
        except Exception:
            continue
    if len(ts) < 2:
        return None
    t0, t1 = min(ts), max(ts)
    if t1 - t0 < 1e-9:
        return None
    try:
        return Line.CreateBound(
            ref + u.Multiply(t0),
            ref + u.Multiply(t1),
        )
    except Exception:
        return None


def _origen_wall_foundation_centro_solido_en_eje_longitudinal(elem, line_eje):
    """
    Punto de paso del plano de sección en el **centro físico** de la zapata en planta: el eje
    ``long_line`` / ``LocationCurve`` puede ir **desplazado** respecto al centro del ancho
    (p. ej. eje de armado). Se toma la dirección longitudinal ``u`` desde ``line_eje`` y el
    **punto medio** del segmento en que la **unión de sólidos** proyecta su caja sobre ``u``
    (:func:`_linea_acotada_en_bbox`). Así el corte es **transversal** y **centrado** en el
    volumen, como en la referencia manual del usuario.

    Respaldo: proyección del centro del bbox de sólidos sobre el segmento ``line_eje``; luego
    punto medio del segmento.
    """
    if elem is None or line_eje is None:
        return None
    try:
        p0 = line_eje.GetEndPoint(0)
        p1 = line_eje.GetEndPoint(1)
        dh = p1.Subtract(p0)
        lu = float(dh.GetLength())
        if lu < 1e-9:
            return None
        u = dh.Normalize()
    except Exception:
        return None
    bb = _bbox_seccion_desde_solida_o_elemento(elem)
    if bb is None or bb.Min is None or bb.Max is None:
        return _punto_medio_segmento_linea(line_eje)
    ref = _centro_bounding_box_xyz(bb)
    if ref is None:
        return _punto_medio_segmento_linea(line_eje)
    ln_seg = _linea_acotada_en_bbox(bb, ref, u)
    if ln_seg is not None:
        try:
            a0 = ln_seg.GetEndPoint(0)
            a1 = ln_seg.GetEndPoint(1)
            seg = a1.Subtract(a0)
            if float(seg.GetLength()) > 1e-9:
                return a0.Add(seg.Multiply(0.5))
        except Exception:
            pass
    try:
        c = ref
        tpar = float((c.Subtract(p0)).DotProduct(u))
        tpar = max(0.0, min(lu, tpar))
        return p0.Add(u.Multiply(tpar))
    except Exception:
        return _punto_medio_segmento_linea(line_eje)


def _long_line_desde_geometria_wall_foundation(wf, diam_long_mm=16.0, diam_trans_mm=10.0):
    """
    ``long_line`` como en Fundación Corrida: primero corte por muro host
    (``geometria_inferior_wall_foundation_cortes_muro``), luego cadena completa
    ``_geometria_wall_foundation_inferior`` en ``enfierrado_wall_foundation``.

    Los diámetros solo influyen en tolerancias de extracción; valores típicos 16/10 mm.
    """
    if wf is None or not isinstance(wf, WallFoundation):
        return None
    try:
        from geometria_wall_foundation_cortes_muro import (
            geometria_inferior_wall_foundation_cortes_muro,
        )

        g_cut = geometria_inferior_wall_foundation_cortes_muro(
            wf, float(diam_long_mm), float(diam_trans_mm)
        )
        if g_cut is not None:
            ln = g_cut.get("long_line")
            if ln is not None:
                return Line.CreateBound(ln.GetEndPoint(0), ln.GetEndPoint(1))
    except Exception:
        pass
    try:
        import enfierrado_wall_foundation as _ewf

        geo, _ = _ewf._geometria_wall_foundation_inferior(
            wf, float(diam_long_mm), float(diam_trans_mm)
        )
        if not geo:
            return None
        ln = geo.get("long_line")
        if ln is None:
            return None
        return Line.CreateBound(ln.GetEndPoint(0), ln.GetEndPoint(1))
    except Exception:
        return None


def _linea_eje_wall_foundation_por_muestra_vertices(elem):
    """
    Respaldo sin ``long_line``: eje longitudinal por **máxima separación en planta (XY)**
    entre vértices de aristas de sólidos. El segmento une el par más alejado; el corte va a
    la mitad de ese tramo (coherente con una zapata alargada cuando ``LocationCurve`` falla).
    """
    if elem is None or not isinstance(elem, WallFoundation):
        return None
    pts = _collect_geometry_vertex_sample_points(elem)
    if len(pts) < 2:
        return None
    if len(pts) > 96:
        pts = pts[:96]
    best_d2 = -1.0
    pa = pb = None
    n = len(pts)
    for i in range(n):
        for j in range(i + 1, n):
            try:
                ax, ay = float(pts[i].X), float(pts[i].Y)
                bx, by = float(pts[j].X), float(pts[j].Y)
                d2 = (bx - ax) * (bx - ax) + (by - ay) * (by - ay)
                if d2 > best_d2:
                    best_d2 = d2
                    pa, pb = pts[i], pts[j]
            except Exception:
                continue
    if pa is None or pb is None or best_d2 < 1e-12:
        return None
    try:
        return Line.CreateBound(pa, pb)
    except Exception:
        return None


def _origen_seccion_proyectando_centro_host_sobre_eje(line, elem):
    """
    Punto sobre el eje del corte (segmento ``line``) más cercano al centro del ``BoundingBox``
    del host, acotado al segmento. Así el plano transversal queda en la “mitad” del volumen de
    la fundación aunque el eje de armado (``long_line``) esté offset respecto al sólido.
    """
    if line is None or elem is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        seg = p1.Subtract(p0)
        lu = float(seg.GetLength())
        if lu < 1e-9:
            return None
        u = seg.Normalize()
    except Exception:
        return None
    bb = elem.get_BoundingBox(None)
    c = _centro_bounding_box_xyz(bb)
    if c is None:
        try:
            return p0.Add(seg.Multiply(0.5))
        except Exception:
            return None
    try:
        tpar = float((c.Subtract(p0)).DotProduct(u))
    except Exception:
        return c
    tpar = max(0.0, min(lu, tpar))
    try:
        return p0.Add(u.Multiply(tpar))
    except Exception:
        return None


def _transform_seccion_transversal_desde_origen_y_linea(origen, line):
    """
    Corte **transversal al eje** del muro (``BasisZ`` = eje ``line``): la vista mira **a lo largo**
    del muro y muestra el **perfil** ancho × profundidad de la zapata (uso habitual armado transversal).

    No confundir con :func:`_transform_seccion_longitudinal_desde_origen_y_linea` (**transversal al
    ancho** en planta: trazo de sección paralelo al eje).
    """
    if origen is None or line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        dh = p1.Subtract(p0)
        if float(dh.GetLength()) < 1e-9:
            return None
        t = dh.Normalize()
    except Exception:
        return None
    z = XYZ.BasisZ
    h = z.CrossProduct(t)
    if h.GetLength() < 1e-6:
        h = XYZ.BasisX.CrossProduct(t)
    if h.GetLength() < 1e-6:
        return None
    h = h.Normalize()
    v = t.CrossProduct(h).Normalize()
    tr = Transform.Identity
    tr.Origin = origen
    tr.BasisX = h
    tr.BasisY = v
    tr.BasisZ = t
    return tr


def _transform_seccion_longitudinal_desde_origen_y_linea(origen, line):
    """
    Corte **transversal al ancho** de la fundación corrida (en planta: perpendicular a la dirección
    del **ancho** de la zapata, es decir el trazo de sección es **paralelo** al eje ``line``).

    ``ViewDirection`` (``BasisZ``) queda en el plano horizontal **perpendicular** al eje del muro
    (mira “de canto” a lo largo de la corrida). Triedro: ``BasisX`` = eje del muro, ``BasisY`` = Z
    global, ``BasisZ`` = ``BasisX`` × ``BasisY``.

    Contrasta con :func:`_transform_seccion_transversal_desde_origen_y_linea` (**transversal al
    eje**: ``BasisZ`` = eje, vista que corta el **ancho** de la zapata).
    """
    if origen is None or line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        dh = p1.Subtract(p0)
        if float(dh.GetLength()) < 1e-9:
            return None
        t = dh.Normalize()
    except Exception:
        return None
    z = XYZ.BasisZ
    bx = t
    by = z
    bz = bx.CrossProduct(by)
    if float(bz.GetLength()) < 1e-9:
        return None
    bz = bz.Normalize()
    tr = Transform.Identity
    tr.Origin = origen
    tr.BasisX = bx
    tr.BasisY = by
    tr.BasisZ = bz
    return tr


def _extremos_caja_seccion_desde_bbox(
    elem, transform, margen_interno, media_profundidad_interna, bb_extents=None
):
    bb = bb_extents if bb_extents is not None else elem.get_BoundingBox(None)
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    corners = (
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    )
    ox = transform.Origin
    bx = transform.BasisX
    by = transform.BasisY
    bz = transform.BasisZ
    xs = []
    ys = []
    zs = []
    for c in corners:
        d = c - ox
        xs.append(d.DotProduct(bx))
        ys.append(d.DotProduct(by))
        zs.append(d.DotProduct(bz))
    m = margen_interno
    zhalf = max(float(media_profundidad_interna), _mm_a_interno(150.0) * 0.5)
    zabs = max(abs(min(zs)), abs(max(zs))) + m
    zhalf = max(zhalf, zabs)
    return (
        min(xs) - m,
        max(xs) + m,
        min(ys) - m,
        max(ys) + m,
        -zhalf,
        zhalf,
    )


def _extremos_caja_seccion_proyectando_geometria(
    elem, transform, margen_interno, media_profundidad_interna, bb_fallback=None
):
    """
    Extremos locales (``Min``/``Max`` del ``BoundingBoxXYZ`` de la sección) proyectando **puntos
    de la geometría real** (aristas de sólidos) sobre el triedro ``transform``.

    Así la caja queda alineada al **sistema local del corte** (OBB en ese marco), no a partir de
    las ocho esquinas del AABB **mundial** del elemento, que infla el recorte en fundaciones
    rotadas o estrechas.
    """
    if elem is None or transform is None:
        return None
    pts = _collect_geometry_vertex_sample_points(elem)
    if len(pts) < 3:
        return _extremos_caja_seccion_desde_bbox(
            elem,
            transform,
            margen_interno,
            media_profundidad_interna,
            bb_extents=bb_fallback,
        )
    ox = transform.Origin
    bx = transform.BasisX
    by = transform.BasisY
    bz = transform.BasisZ
    xs = []
    ys = []
    zs = []
    for c in pts:
        try:
            d = c - ox
            xs.append(float(d.DotProduct(bx)))
            ys.append(float(d.DotProduct(by)))
            zs.append(float(d.DotProduct(bz)))
        except Exception:
            pass
    if len(xs) < 2:
        return _extremos_caja_seccion_desde_bbox(
            elem,
            transform,
            margen_interno,
            media_profundidad_interna,
            bb_extents=bb_fallback,
        )
    m = margen_interno
    zhalf = max(float(media_profundidad_interna), _mm_a_interno(150.0) * 0.5)
    zabs = max(abs(min(zs)), abs(max(zs))) + m
    zhalf = max(zhalf, zabs)
    return (
        min(xs) - m,
        max(xs) + m,
        min(ys) - m,
        max(ys) + m,
        -zhalf,
        zhalf,
    )


def _recentrar_transform_y_extents_xy(tr, xmn, xmx, ymn, ymx):
    """
    Desplaza ``tr.Origin`` en el plano local X/Y para que el centro geométrico del
    ``BoundingBoxXYZ`` coincida con ese origen (Revit usa la caja completa; si ``Min``/``Max``
    son asimétricos en X o Y, el corte parecía estar corrido en planta).
    """
    if tr is None:
        return xmn, xmx, ymn, ymx
    xmid = 0.5 * (float(xmn) + float(xmx))
    ymid = 0.5 * (float(ymn) + float(ymx))
    if abs(xmid) < 1e-9 and abs(ymid) < 1e-9:
        return xmn, xmx, ymn, ymx
    try:
        ox = tr.Origin
        bx = tr.BasisX
        by = tr.BasisY
        tr.Origin = ox.Add(bx.Multiply(xmid)).Add(by.Multiply(ymid))
    except Exception:
        return xmn, xmx, ymn, ymx
    return xmn - xmid, xmx - xmid, ymn - ymid, ymx - ymid


def _asignar_nombre_vista_unico(view, document, nombre_base):
    existentes = set()
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v is None or v.Id == view.Id:
                continue
            n = v.Name
            if n:
                existentes.add(unicode(n).strip().lower())
        except Exception:
            continue
    cand = nombre_base
    try:
        cand_u = unicode(cand).strip().lower()
    except Exception:
        cand_u = str(cand).strip().lower()
    k = 0
    while cand_u in existentes:
        k += 1
        cand = u"{0} ({1})".format(nombre_base, k)
        try:
            cand_u = unicode(cand).strip().lower()
        except Exception:
            cand_u = str(cand).strip().lower()
    view.Name = cand


def _aplicar_far_clip_offset_mm(view, mm):
    """Establece el parámetro Far Clip Offset de la vista (longitud interna)."""
    try:
        val = UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)
    except Exception:
        return False
    p = None
    try:
        from Autodesk.Revit.DB import ParameterTypeId

        p = view.get_Parameter(ParameterTypeId.ViewerBoundOffsetFar)
    except Exception:
        pass
    if p is None:
        try:
            p = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
        except Exception:
            p = None
    if p is None:
        return False
    try:
        if p.IsReadOnly:
            return False
        p.Set(val)
        return True
    except Exception:
        return False


def _linea_wall_host_desde_wall_foundation(wf, doc):
    """
    Línea de ubicación del muro que hospeda la zapata (``WallFoundation.WallId``).
    Prioriza el host frente a la curva de la propia zapata para el desplazamiento longitudinal.
    """
    if wf is None or doc is None:
        return None
    try:
        wid = wf.WallId
        if wid is None or wid == ElementId.InvalidElementId:
            return None
        w = doc.GetElement(wid)
        if not isinstance(w, Wall):
            return None
        return _linea_desde_location_curve_host(w)
    except Exception:
        return None


def _punto_proyectado_en_segmento_linea(pt, line):
    """
    Punto más cercano a ``pt`` sobre el segmento ``Line`` (Revit), en coordenadas modelo.
    """
    if pt is None or line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        seg = p1.Subtract(p0)
        ln = float(seg.GetLength())
        if ln < 1e-12:
            return p0
        u = seg.Normalize()
        t = float(pt.Subtract(p0).DotProduct(u))
        if t < 0.0:
            t = 0.0
        elif t > ln:
            t = ln
        return p0.Add(u.Multiply(t))
    except Exception:
        return None


def _wf_trasladar_cropbox_seccion_segun_muro_host(wf, view_section, doc):
    """
    Tras ``ViewSection.CreateSection``, corrige la posición del corte respecto a la zapata.

    El vector se obtiene proyectando el centro del bbox de la WF y el centro de la caja
    de sección sobre la **línea de ubicación del muro host** (``WallFoundation.WallId``).

    Se usa ``ElementTransformUtils.MoveElement`` (SDK DynamicModelUpdate / Building Coder);
    el setter de ``CropBox`` no mueve el plano de corte en secciones y puede ignorar el
    ``Transform`` — solo se usa como respaldo si el movimiento falla.
    """
    if wf is None or view_section is None or doc is None:
        return False
    host_line = _linea_wall_host_desde_wall_foundation(wf, doc)
    if host_line is None:
        return False
    bb = _bbox_seccion_desde_solida_o_elemento(wf)
    if bb is None or bb.Min is None or bb.Max is None:
        try:
            bb = wf.get_BoundingBox(None)
        except Exception:
            bb = None
    if bb is None or bb.Min is None or bb.Max is None:
        return False
    c_wf = XYZ(
        0.5 * (float(bb.Min.X) + float(bb.Max.X)),
        0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
        0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
    )
    p_wf = _punto_proyectado_en_segmento_linea(c_wf, host_line)
    if p_wf is None:
        return False
    try:
        cb = view_section.CropBox
        if cb is None:
            return False
        tr = cb.Transform
        xm = 0.5 * (float(cb.Min.X) + float(cb.Max.X))
        ym = 0.5 * (float(cb.Min.Y) + float(cb.Max.Y))
        zm = 0.5 * (float(cb.Min.Z) + float(cb.Max.Z))
        p_sec = (
            tr.Origin.Add(tr.BasisX.Multiply(xm))
            .Add(tr.BasisY.Multiply(ym))
            .Add(tr.BasisZ.Multiply(zm))
        )
    except Exception:
        return False
    p_sec_proj = _punto_proyectado_en_segmento_linea(p_sec, host_line)
    if p_sec_proj is None:
        return False
    try:
        delta = p_wf.Subtract(p_sec_proj)
        dlen = float(delta.GetLength())
    except Exception:
        return False
    tol = _mm_a_interno(0.5)
    if dlen < tol:
        return True
    try:
        ElementTransformUtils.MoveElement(doc, view_section.Id, delta)
        return True
    except Exception:
        pass
    try:
        new_tr = Transform.Identity
        new_tr.BasisX = tr.BasisX
        new_tr.BasisY = tr.BasisY
        new_tr.BasisZ = tr.BasisZ
        new_tr.Origin = tr.Origin.Add(delta)
        new_cb = BoundingBoxXYZ()
        new_cb.Min = cb.Min
        new_cb.Max = cb.Max
        new_cb.Transform = new_tr
        view_section.CropBox = new_cb
        return True
    except Exception:
        return False


def _wf_align_section_head_to_wf_mid_span(
    tr,
    xmn,
    xmx,
    ymn,
    ymx,
    zmn,
    zmx,
    wf_line_eje,
    elem,
    margen_interno,
    media_prof,
    bb_fallback,
):
    """
    Revit coloca la cabeza en (Max.X, Min.Y, Min.Z) del BoundingBoxXYZ local.

    1) Desplazar el origen para llevar esa esquina al punto medio del tramo; reproyectar
       extremos y recéntrar XY.
    2) Repetir el desplazamiento de cabeza al mid con los nuevos extremos (la cabeza se
       había desplazado al recéntrar).
    3) Tras (2), el origen vuelve a moverse: los ``Min``/``Max`` locales quedan calculados
       para el origen anterior. Se **vuelve a proyectar** la geometría **sin** recéntrar XY
       otra vez (evita otro ciclo que desplace la cabeza), para que la caja abarque todo el
       ancho de la fundación en el plano del corte.

    Un bucle por tolerancia en la distancia convergía mal (corte a mitad de tramo o símbolo
    en T3); este esquema fijo de tres pasos evita eso.
    """
    init = (xmn, xmx, ymn, ymx, zmn, zmx)
    if tr is None or wf_line_eje is None or elem is None:
        return tr, xmn, xmx, ymn, ymx, zmn, zmx
    try:
        p0 = wf_line_eje.GetEndPoint(0)
        p1 = wf_line_eje.GetEndPoint(1)
        mid = XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
    except Exception:
        return tr, xmn, xmx, ymn, ymx, zmn, zmx
    o_orig = tr.Origin

    def _paso_alinear_y_extremos_recuentrar():
        bx = tr.BasisX
        by = tr.BasisY
        bz = tr.BasisZ
        w_head = (
            tr.Origin.Add(bx.Multiply(xmx))
            .Add(by.Multiply(ymn))
            .Add(bz.Multiply(zmn))
        )
        tr.Origin = tr.Origin.Add(mid.Subtract(w_head))
        ex = _extremos_caja_seccion_proyectando_geometria(
            elem, tr, margen_interno, media_prof, bb_fallback=bb_fallback
        )
        if ex is None:
            return False
        xmn_, xmx_, ymn_, ymx_, zmn_, zmx_ = ex
        if xmn_ >= xmx_:
            xmn_, xmx_ = xmx_ - 0.5, xmx_ + 0.5
        if ymn_ >= ymx_:
            ymn_, ymx_ = ymx_ - 0.5, ymx_ + 0.5
        if zmn_ >= zmx_:
            zmn_, zmx_ = -media_prof, media_prof
        xmn_, xmx_, ymn_, ymx_ = _recentrar_transform_y_extents_xy(
            tr, xmn_, xmx_, ymn_, ymx_
        )
        return xmn_, xmx_, ymn_, ymx_, zmn_, zmx_

    try:
        r1 = _paso_alinear_y_extremos_recuentrar()
        if r1 is False:
            tr.Origin = o_orig
            return tr, init[0], init[1], init[2], init[3], init[4], init[5]
        xmn, xmx, ymn, ymx, zmn, zmx = r1

        r2 = _paso_alinear_y_extremos_recuentrar()
        if r2 is False:
            tr.Origin = o_orig
            return tr, init[0], init[1], init[2], init[3], init[4], init[5]
        xmn, xmx, ymn, ymx, zmn, zmx = r2

        bx = tr.BasisX
        by = tr.BasisY
        bz = tr.BasisZ
        w_head_mid = (
            tr.Origin.Add(bx.Multiply(xmx))
            .Add(by.Multiply(ymn))
            .Add(bz.Multiply(zmn))
        )
        tr.Origin = tr.Origin.Add(mid.Subtract(w_head_mid))

        ex3 = _extremos_caja_seccion_proyectando_geometria(
            elem, tr, margen_interno, media_prof, bb_fallback=bb_fallback
        )
        if ex3 is None:
            tr.Origin = o_orig
            return tr, init[0], init[1], init[2], init[3], init[4], init[5]
        xmn, xmx, ymn, ymx, zmn, zmx = ex3
        if xmn >= xmx:
            xmn, xmx = xmx - 0.5, xmx + 0.5
        if ymn >= ymx:
            ymn, ymx = ymx - 0.5, ymx + 0.5
        if zmn >= zmx:
            zmn, zmx = -media_prof, media_prof
    except Exception:
        try:
            tr.Origin = o_orig
        except Exception:
            pass
        return tr, init[0], init[1], init[2], init[3], init[4], init[5]
    return tr, xmn, xmx, ymn, ymx, zmn, zmx


def _crear_vistas_seccion_revision_para_hosts(
    document,
    elementos,
    etiqueta_elemento,
    nombre_template,
    tx_label,
    margen_mm,
    profundidad_eje_mm,
    far_clip_offset_mm,
    uidocument,
    max_vistas,
    gestionar_transaccion,
    limite_msg_template=None,
    lineas_eje=None,
    origen_en_medio_location_curve=False,
    vft_section_id=None,
    parent_view_id_plan=None,
):
    """
    Núcleo compartido: sección transversal al eje del corte.

    Si ``origen_en_medio_location_curve`` es True y el elemento es ``WallFoundation``, el
    ``Transform`` es **transversal al eje del muro** (``_transform_seccion_transversal_desde_origen_y_linea``)
    a partir del eje ``long_line`` (cortes muro / enfierrado), ``LocationCurve`` o vértices; el
    origen es el **punto medio** de ese segmento (respaldo: centro sólido en eje). Fallback de
    orientación: ``_transform_seccion_desde_geometria_arista_horizontal_y_centro_bbox``.

    Si es False (vigas), el origen proyecta el centro del ``BoundingBox`` del host sobre el eje
    (o punto medio de la línea si falla), para centrar cuando el eje de armado va desfasado.

    ``nombre_template``: cadena con ``{0}`` → ``int(elem.Id.IntegerValue)``.
    ``limite_msg_template``: opcional ``unicode`` con ``{0}`` = max_vistas (mensaje al truncar).
    ``lineas_eje``: por elemento, línea de eje opcional. Para ``WallFoundation`` (revisión zapata),
        si se informa (p. ej. ``long_line`` desde ``_geometria_wall_foundation_inferior``), se usa
        para orientar y situar el corte; si no, se intenta la misma vía import; respaldo: geometría
        por aristas.
    ``vft_section_id``: opcional ``ElementId`` de ``ViewFamilyType`` (``ViewFamily.Section``).
        Si es ``None``, se usa el primer tipo de sección del proyecto.
    ``parent_view_id_plan``: reservado (compatibilidad); la cabeza de la sección WF transversal
        se alinea al punto medio del tramo vía ``_wf_align_section_head_to_wf_mid_span``.
    """
    vistas = []
    avisos = []
    _ = parent_view_id_plan
    if not elementos:
        return vistas, avisos

    vft_id = vft_section_id if vft_section_id is not None else _primer_view_family_type_seccion(document)
    if vft_id is None:
        avisos.append(
            u"No hay tipo de vista «Section» en el proyecto; no se creó la sección de revisión."
        )
        return vistas, avisos

    margen_i = _mm_a_interno(margen_mm)
    media_prof = _mm_a_interno(profundidad_eje_mm) * 0.5

    elems = list(elementos)[: int(max_vistas)]
    if lineas_eje is not None:
        try:
            lineas_eje = list(lineas_eje)[: len(elems)]
        except Exception:
            lineas_eje = None
    if len(elementos) > int(max_vistas) and limite_msg_template:
        try:
            avisos.append(limite_msg_template.format(int(max_vistas)))
        except Exception:
            pass

    tx = None
    if gestionar_transaccion:
        tx = Transaction(document, tx_label)
        tx.Start()
    try:
        for i, elem in enumerate(elems):
            wf_branch = None
            wf_line_eje = None
            line_explicit = None
            if lineas_eje is not None and i < len(lineas_eje):
                line_explicit = lineas_eje[i]
            line = line_explicit
            if line is None:
                line = _linea_desde_location_curve_host(elem)

            bb_for_extents = None
            tr = None
            if origen_en_medio_location_curve:
                if isinstance(elem, WallFoundation):
                    bb_for_extents = _bbox_seccion_desde_solida_o_elemento(elem)
                    if bb_for_extents is None or bb_for_extents.Min is None:
                        bb_for_extents = elem.get_BoundingBox(None)
                    tr = None
                    line_eje_wf = line_explicit
                    if line_eje_wf is None:
                        line_eje_wf = _long_line_desde_geometria_wall_foundation(elem)
                    if line_eje_wf is None:
                        line_eje_wf = _linea_desde_location_curve_host(elem)
                    if line_eje_wf is None:
                        line_eje_wf = _linea_eje_wall_foundation_span_vertices_sobre_location(
                            elem
                        )
                    if line_eje_wf is None:
                        line_eje_wf = _linea_eje_wall_foundation_por_muestra_vertices(
                            elem
                        )
                    wf_line_eje = line_eje_wf
                    wf_line_path_ok = False
                    if line_eje_wf is not None:
                        o_sec = _origen_seccion_medio_eje_wall_foundation(elem)
                        if o_sec is None:
                            o_sec = _punto_medio_segmento_linea(line_eje_wf)
                        if o_sec is None:
                            o_sec = _origen_wall_foundation_centro_solido_en_eje_longitudinal(
                                elem, line_eje_wf
                            )
                        if o_sec is not None:
                            try:
                                tr = _transform_seccion_transversal_desde_origen_y_linea(
                                    o_sec, line_eje_wf
                                )
                                if tr is not None:
                                    wf_line_path_ok = True
                            except Exception:
                                tr = None
                    if tr is None:
                        tr = _transform_seccion_desde_geometria_arista_horizontal_y_centro_bbox(
                            elem, bb_for_extents
                        )
                    wf_branch = (
                        "wf_transversal_eje_line"
                        if wf_line_path_ok
                        else "wf_fallback_bbox_arista"
                    )
                    if tr is None:
                        avisos.append(
                            u"{0} Id {1}: sección de revisión no creada (sin geometría válida para alinear la caja).".format(
                                etiqueta_elemento,
                                int(elem.Id.IntegerValue),
                            )
                        )
                        continue
                else:
                    o_loc, line_ax = _origen_y_linea_eje_seccion_en_medio_location(elem)
                    if o_loc is not None and line_ax is not None:
                        tr = _transform_seccion_transversal_desde_origen_y_linea(
                            o_loc, line_ax
                        )

            if tr is None:
                if line is None:
                    avisos.append(
                        u"{0} Id {1}: sin curva de ubicación válida; sin sección.".format(
                            etiqueta_elemento,
                            int(elem.Id.IntegerValue),
                        )
                    )
                    continue
                origen_sec = _origen_seccion_proyectando_centro_host_sobre_eje(line, elem)
                if origen_sec is not None:
                    tr = _transform_seccion_transversal_desde_origen_y_linea(
                        origen_sec, line
                    )
                else:
                    tr = _transform_seccion_transversal_punto_medio(line)
            if tr is None:
                avisos.append(
                    u"{0} Id {1}: eje degenerado (¿vertical?); sin sección.".format(
                        etiqueta_elemento,
                        int(elem.Id.IntegerValue),
                    )
                )
                continue
            if isinstance(elem, WallFoundation) and origen_en_medio_location_curve:
                ex = _extremos_caja_seccion_proyectando_geometria(
                    elem, tr, margen_i, media_prof, bb_fallback=bb_for_extents
                )
            else:
                ex = _extremos_caja_seccion_desde_bbox(
                    elem, tr, margen_i, media_prof, bb_extents=bb_for_extents
                )
            if ex is None:
                dft = 8.0
                ex = (-dft, dft, -dft, dft, -media_prof, media_prof)
            xmn, xmx, ymn, ymx, zmn, zmx = ex
            if xmn >= xmx:
                xmn, xmx = xmx - 0.5, xmx + 0.5
            if ymn >= ymx:
                ymn, ymx = ymx - 0.5, ymx + 0.5
            if zmn >= zmx:
                zmn, zmx = -media_prof, media_prof

            xmn, xmx, ymn, ymx = _recentrar_transform_y_extents_xy(tr, xmn, xmx, ymn, ymx)

            if (
                isinstance(elem, WallFoundation)
                and origen_en_medio_location_curve
                and wf_branch == "wf_transversal_eje_line"
                and wf_line_eje is not None
            ):
                tr, xmn, xmx, ymn, ymx, zmn, zmx = _wf_align_section_head_to_wf_mid_span(
                    tr,
                    xmn,
                    xmx,
                    ymn,
                    ymx,
                    zmn,
                    zmx,
                    wf_line_eje,
                    elem,
                    margen_i,
                    media_prof,
                    bb_for_extents,
                )

            box = BoundingBoxXYZ()
            box.Transform = tr
            box.Min = XYZ(xmn, ymn, zmn)
            box.Max = XYZ(xmx, ymx, zmx)

            try:
                vs = ViewSection.CreateSection(document, vft_id, box)
            except Exception as e:
                try:
                    avisos.append(
                        u"{0} Id {1}: CreateSection — {2}".format(
                            etiqueta_elemento,
                            int(elem.Id.IntegerValue),
                            unicode(e),
                        )
                    )
                except Exception:
                    avisos.append(
                        u"{0} Id {1}: falló la creación de la sección.".format(
                            etiqueta_elemento,
                            int(elem.Id.IntegerValue),
                        )
                    )
                continue
            try:
                vs.CropBoxVisible = False
            except Exception:
                pass
            try:
                _aplicar_far_clip_offset_mm(vs, far_clip_offset_mm)
            except Exception:
                pass
            if (
                isinstance(elem, WallFoundation)
                and origen_en_medio_location_curve
                and wf_branch == "wf_transversal_eje_line"
            ):
                try:
                    _wf_trasladar_cropbox_seccion_segun_muro_host(elem, vs, document)
                except Exception:
                    pass
            try:
                nombre_base = nombre_template.format(int(elem.Id.IntegerValue))
            except Exception:
                nombre_base = u"BIMTools — Rev. enfierrado — {0}".format(
                    int(elem.Id.IntegerValue)
                )
            try:
                _asignar_nombre_vista_unico(vs, document, nombre_base)
            except Exception:
                pass
            vistas.append(vs)
        if gestionar_transaccion and tx is not None:
            tx.Commit()
    except Exception as e:
        if gestionar_transaccion and tx is not None:
            try:
                tx.RollBack()
            except Exception:
                pass
            try:
                avisos.append(u"Transacción sección revisión: {0}".format(unicode(e)))
            except Exception:
                avisos.append(u"Transacción sección revisión: error.")
            return [], avisos
        raise

    if uidocument is not None and vistas:
        try:
            uidocument.ActiveView = vistas[-1]
        except Exception:
            avisos.append(
                u"Se creó la sección pero no se pudo cambiar a esa vista automáticamente."
            )

    return vistas, avisos


def crear_vistas_seccion_revision_enfierrado(
    document,
    elementos_framing,
    margen_mm=800.0,
    profundidad_eje_mm=2000.0,
    far_clip_offset_mm=_FAR_CLIP_OFFSET_MM_DEFECTO,
    uidocument=None,
    max_vistas=12,
    gestionar_transaccion=True,
):
    """
    Una ViewSection por viga, corte transversal al eje en el punto medio.

    Args:
        document: Document
        elementos_framing: iterable de vigas (Structural Framing)
        margen_mm: ampliación del recorte alrededor del bbox de la viga
        profundidad_eje_mm: grosor del corte a lo largo del eje de la viga
        far_clip_offset_mm: Far Clip Offset de la vista (por defecto 100 mm)
        uidocument: si se pasa, activa la última vista creada
        max_vistas: tope de vistas por ejecución
        gestionar_transaccion: si ``False``, el llamador debe tener abierta una
            ``Transaction``; no se abre ni confirma aquí (misma transacción que el enfierrado).

    Returns:
        (lista de ViewSection, lista de textos de aviso)
    """
    return _crear_vistas_seccion_revision_para_hosts(
        document,
        elementos_framing,
        u"Viga",
        u"BIMTools — Rev. enfierrado — {0}",
        u"BIMTools — Sección revisión enfierrado vigas",
        margen_mm,
        profundidad_eje_mm,
        far_clip_offset_mm,
        uidocument,
        max_vistas,
        gestionar_transaccion,
        limite_msg_template=u"Solo se generaron {0} vista(s) de sección (límite por cantidad de vigas).",
        lineas_eje=None,
        origen_en_medio_location_curve=False,
    )


def crear_vistas_seccion_revision_wall_foundation(
    document,
    elementos_wall_foundation,
    linea_eje=None,
    margen_mm=1200.0,
    profundidad_eje_mm=2000.0,
    far_clip_offset_mm=_FAR_CLIP_OFFSET_MM_DEFECTO,
    uidocument=None,
    max_vistas=12,
    gestionar_transaccion=True,
    forzar_creacion=False,
    view_family_type_section_id=None,
):
    """
    Una ``ViewSection`` por zapata de muro. **Prioridad de orientación / origen:** la misma
    ``long_line`` que ``enfierrado_wall_foundation`` (``_geometria_wall_foundation_inferior``):
    ``linea_eje`` si el llamador la pasa; si no, se calcula con import perezoso; respaldo:
    :func:`_transform_seccion_desde_geometria_arista_horizontal_y_centro_bbox`.

    La **caja de recorte** usa :func:`_extremos_caja_seccion_proyectando_geometria`: proyección
    de puntos de aristas del sólido al marco local del corte.

    ``linea_eje``: eje longitudinal real de la zapata (p. ej. desde geometría inferior); si es
        ``None``, se intenta obtener igual que en Fundación Corrida.

    ``forzar_creacion``: si es True, ignora ``CREAR_SECCION_REVISION_WALL_FOUNDATION`` (p. ej.
        herramienta independiente de sección transversal).

    ``view_family_type_section_id``: opcional ``ElementId`` de tipo de vista Section; si es
        ``None``, se usa el primer tipo Section del proyecto.
    """
    if not forzar_creacion and not CREAR_SECCION_REVISION_WALL_FOUNDATION:
        return [], []
    raw_wf = list(elementos_wall_foundation or [])
    le_list = None
    if linea_eje is not None and raw_wf:
        le_list = [linea_eje if i == 0 else None for i in range(len(raw_wf))]
    return _crear_vistas_seccion_revision_para_hosts(
        document,
        elementos_wall_foundation,
        u"Zapata",
        u"BIMTools — Rev. enfierrado — WF {0}",
        u"BIMTools — Sección revisión Wall Foundation",
        margen_mm,
        profundidad_eje_mm,
        far_clip_offset_mm,
        uidocument,
        max_vistas,
        gestionar_transaccion,
        limite_msg_template=u"Solo se generaron {0} vista(s) de sección (límite por cantidad de zapatas).",
        lineas_eje=le_list,
        origen_en_medio_location_curve=True,
        vft_section_id=view_family_type_section_id,
    )


def crear_seccion_transversal_wall_foundation_desde_filtro_vista(
    document,
    elementos_wall_foundation,
    vista_origen,
    linea_eje=None,
    margen_mm=1200.0,
    profundidad_eje_mm=2000.0,
    far_clip_offset_mm=_FAR_CLIP_OFFSET_MM_DEFECTO,
    uidocument=None,
    max_vistas=12,
    gestionar_transaccion=True,
):
    """
    Crea ``ViewSection`` **transversales al eje del muro** que hospeda la zapata (corte ⟂ al eje
    en planta; ``ViewDirection`` a lo largo del eje del muro). Usa el **tipo Section** que
    corresponda al parámetro «Section Filter» de ``vista_origen``.

    Returns:
        ``(lista ViewSection, lista de avisos str)``.
    """
    vft_id, err = _resolver_view_family_type_seccion_desde_vista(document, vista_origen)
    if vft_id is None:
        return [], [err] if err else [u"No se pudo resolver el tipo de sección."]
    raw_wf = list(elementos_wall_foundation or [])
    le_list = None
    if linea_eje is not None and raw_wf:
        le_list = [linea_eje if i == 0 else None for i in range(len(raw_wf))]
    vid_plan = None
    try:
        if vista_origen is not None:
            vid_plan = vista_origen.Id
    except Exception:
        vid_plan = None
    return _crear_vistas_seccion_revision_para_hosts(
        document,
        elementos_wall_foundation,
        u"Zapata",
        u"BIMTools — Rev. enfierrado — WF {0}",
        u"BIMTools — Sección WF transversal al eje muro (Section Filter)",
        margen_mm,
        profundidad_eje_mm,
        far_clip_offset_mm,
        uidocument,
        max_vistas,
        gestionar_transaccion,
        limite_msg_template=u"Solo se generaron {0} vista(s) de sección (límite por cantidad de zapatas).",
        lineas_eje=le_list,
        origen_en_medio_location_curve=True,
        vft_section_id=vft_id,
        parent_view_id_plan=vid_plan,
    )


_VFT_STRUCTURAL_PLAN_PISO = u"Structural Plan (Piso)"
_NOMBRE_PREFIJO_PLANTA_FUNDACION = u"DET. FUND."


def _num_str_elemento(elem):
    """Obtiene la numeración de la fundación desde el parámetro «Numeracion Fundacion»."""
    for name in (u"Numeracion Fundacion", u"Numeracion fundacion",
                 u"Foundation Numbering", u"Numeracion"):
        try:
            from Autodesk.Revit.DB import StorageType
            p = elem.LookupParameter(name)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = p.AsString()
                if s and str(s).strip():
                    return str(s).strip()
            elif p.StorageType == StorageType.Integer:
                return str(p.AsInteger())
            elif p.StorageType == StorageType.Double:
                return str(int(round(p.AsDouble())))
        except Exception:
            continue
    # Fallback: ID del elemento
    try:
        return str(int(elem.Id.IntegerValue))
    except Exception:
        return u"?"


def _nombre_vista_planta_fundacion_unico(document, elem_id_int, sufijo=None, elemento=None):
    """Nombre único para la vista de planta de revisión de fundación aislada.

    Formato: ``DETALLE FUNDACION F{numeracion}_ARM. INFERIOR`` /
             ``DETALLE FUNDACION F{numeracion}_ARM. SUPERIOR``.
    """
    # Numeración desde el parámetro del elemento; fallback al ID
    if elemento is not None:
        num = _num_str_elemento(elemento)
    else:
        num = str(elem_id_int)

    # Sufijo legible según ubicación de armadura
    _SUFIJO_READABLE = {u"— Inf": u"ARM. INFERIOR", u"— Sup": u"ARM. SUPERIOR"}
    sufijo_nombre = _SUFIJO_READABLE.get(sufijo, sufijo) if sufijo else None

    if sufijo_nombre:
        base = u"{0} F{1}_{2}".format(_NOMBRE_PREFIJO_PLANTA_FUNDACION, num, sufijo_nombre)
    else:
        base = u"{0} F{1}".format(_NOMBRE_PREFIJO_PLANTA_FUNDACION, num)

    used = set()
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v and v.Name:
                used.add(str(v.Name).strip())
        except Exception:
            continue
    if base not in used:
        return base
    for i in range(2, 100):
        candidate = u"{0} ({1})".format(base, i)
        if candidate not in used:
            return candidate
    return base


def _ocultar_rebars_por_ubicacion(vista, document, ubicacion_mostrar, param_nombre=u"Armadura_Ubicacion"):
    """
    Oculta en ``vista`` todos los ``Rebar`` cuyo parámetro ``param_nombre`` NO coincide
    con ``ubicacion_mostrar``. Los rebars sin ese parámetro (o con valor vacío) también
    se ocultan si ``ubicacion_mostrar`` está definido.

    Debe llamarse dentro de una transacción abierta.
    """
    if vista is None or document is None or not ubicacion_mostrar:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
        from System.Collections.Generic import List

        ids_ocultar = List[ElementId]()
        for rb in FilteredElementCollector(document).OfClass(Rebar).ToElements():
            if rb is None:
                continue
            try:
                p = rb.LookupParameter(param_nombre)
                val = p.AsString() if (p is not None and not p.IsReadOnly) else u""
            except Exception:
                val = u""
            if (val or u"").strip() != ubicacion_mostrar:
                ids_ocultar.Add(rb.Id)
        if ids_ocultar.Count > 0:
            vista.HideElements(ids_ocultar)
    except Exception:
        pass


def _find_vft_structural_plan_piso(document):
    """
    Resuelve el ViewFamilyType «Structural Plan (Piso)».
    Primer intento: filtrar por ViewFamily.StructuralPlan + nombre exacto.
    Segundo intento: cualquier tipo con ese nombre.
    Tercer intento: cualquier tipo StructuralPlan.
    """
    target = _VFT_STRUCTURAL_PLAN_PISO.strip().lower()
    col = FilteredElementCollector(document)
    try:
        col = col.WhereElementIsElementType()
    except Exception:
        pass
    structural_matches = []
    any_sp = []
    for vft in col.OfClass(ViewFamilyType):
        try:
            if vft is None:
                continue
            is_sp = vft.ViewFamily == ViewFamily.StructuralPlan
            n = u""
            try:
                n = str(vft.Name).strip()
            except Exception:
                pass
            if not n:
                for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM):
                    try:
                        p = vft.get_Parameter(bip)
                        if p and p.HasValue:
                            s = p.AsString()
                            if s:
                                n = str(s).strip()
                                break
                    except Exception:
                        continue
            if is_sp:
                any_sp.append(vft)
            if n.lower() == target:
                structural_matches.append(vft)
        except Exception:
            continue
    if structural_matches:
        return structural_matches[0]
    if any_sp:
        return any_sp[0]
    return None


def _nivel_mas_proximo(document, z_ref_ft):
    """Level cuya Elevation es más cercana a ``z_ref_ft`` (unidades internas)."""
    best = None
    best_d = None
    for lv in FilteredElementCollector(document).OfClass(Level):
        try:
            if lv is None:
                continue
            d = abs(float(lv.Elevation) - float(z_ref_ft))
            if best_d is None or d < best_d:
                best_d = d
                best = lv
        except Exception:
            continue
    return best


_CROPBOX_ESCALA = 0.25       # expansión total del bbox (25 %)
_CROPBOX_MARGEN_MIN_MM = 200.0  # mínimo absoluto por eje
_MRA_TYPE_NAME = u"Recorrido Barras"          # MRA para vistas de planta
_MRA_TYPE_NAME_SECCION = u"Structural Rebar_Malla"  # MRA para secciones de fundación aislada


def _bbox_corners_from_element(elemento):
    """Devuelve los puntos de geometría del elemento (sólidos primero, bbox como respaldo)."""
    pts = _collect_geometry_vertex_sample_points(elemento)
    if len(pts) >= 3:
        return pts
    try:
        bb = elemento.get_BoundingBox(None)
        if bb is not None and bb.Min is not None and bb.Max is not None:
            return [
                XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
                XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
                XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
                XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
                XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
                XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
                XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
            ]
    except Exception:
        pass
    return []


def _activar_visibilidad_armaduras(vista, document):
    """
    Activa el estado «View Unobscured» de las armaduras en ``vista`` mediante
    ``RebarInView.SetUnobscured``, que es el mecanismo que controla el diálogo
    «Reinforcement Element View Visibility States» de Revit.
    """
    if vista is None:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
        for rb in FilteredElementCollector(document).OfClass(Rebar).ToElements():
            try:
                rb.SetUnobscuredInView(vista, True)
            except Exception:
                pass
    except Exception:
        pass


def _aplicar_representacion_show_middle(vista, document):
    """
    Aplica la representación «Show Middle» (``RebarPresentationMode.Middle``) a todas
    las armaduras (Rebar) del documento en la vista ``vista`` via
    ``Rebar.SetPresentationMode``.
    """
    if vista is None:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar, RebarPresentationMode
        for rb in FilteredElementCollector(document).OfClass(Rebar).ToElements():
            try:
                rb.SetPresentationMode(vista, RebarPresentationMode.Middle)
            except Exception:
                pass
    except Exception:
        pass


_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG"


def _host_id_safe(rb):
    """Retorna el ElementId del host del rebar, o None si no aplica."""
    try:
        return rb.GetHostId()
    except Exception:
        return None


def _project_onto_view_plane(v, vd):
    """
    Proyecta el vector ``v`` sobre el plano de la vista (elimina la componente
    en la dirección ``vd``).  Devuelve el vector normalizado, o ``None`` si el
    resultado es degenerado (longitud < 1e-9).
    """
    try:
        dot = float(v.DotProduct(vd))
        proj = XYZ(
            float(v.X) - dot * float(vd.X),
            float(v.Y) - dot * float(vd.Y),
            float(v.Z) - dot * float(vd.Z),
        )
        if proj.GetLength() < 1e-9:
            return None
        return proj.Normalize()
    except Exception:
        return None


def _rebar_spacing_direction(rb, fallback_rd, view_vd=None, view_up=None):
    """
    Determina la dirección de distribución (spacing) del set de barras.

    Método A  : vector posición-0 → posición-n-1 del array (más fiable).
    Método B2 : cuando ``bar_dir ≈ vd`` (barra longitudinal "entrando en pantalla"
                en cortes/elevaciones), la distribución visible es ``view_up``.
                Esto resuelve las barras laterales en secciones.
    Método B1 : ``bar_dir × vd`` (funciona cuando la barra no es paralela a vd).
    Método C  : ``fallback_rd`` (último recurso).
    """
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption
        n = 0
        try:
            n = int(rb.NumberOfBarPositions)
        except Exception:
            pass

        vd_ref = view_vd if view_vd is not None else XYZ(0.0, 0.0, -1.0)
        v_up_ref = view_up if view_up is not None else XYZ(0.0, 0.0, 1.0)

        # Método A: vector entre la primera y la última posición del array.
        if n > 1:
            for mpo_name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
                mpo = getattr(MultiplanarOption, mpo_name, None)
                if mpo is None:
                    continue
                try:
                    cs0 = list(rb.GetCenterlineCurves(False, False, False, mpo, 0))
                    csn = list(rb.GetCenterlineCurves(False, False, False, mpo, n - 1))
                    if cs0 and csn:
                        c0 = cs0[0].Evaluate(0.5, True)
                        cn = csn[0].Evaluate(0.5, True)
                        v = cn - c0
                        if float(v.GetLength()) > 1e-6:
                            return v.Normalize()
                except Exception:
                    pass

        # Obtener la curva más larga para determinar bar_dir.
        curves = []
        for mpo_name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
            mpo = getattr(MultiplanarOption, mpo_name, None)
            if mpo is None:
                continue
            try:
                cs = list(rb.GetCenterlineCurves(False, False, False, mpo, 0))
                if cs:
                    curves = cs
                    break
            except Exception:
                pass
        if not curves:
            try:
                curves = list(rb.GetCenterlineCurves(False, False, False))
            except Exception:
                pass

        if curves:
            longest = max(curves, key=lambda c: float(c.Length))
            p0 = longest.Evaluate(0.0, True)
            p1 = longest.Evaluate(1.0, True)
            bar_dir = p1 - p0
            if bar_dir.GetLength() > 1e-9:
                bar_dir = bar_dir.Normalize()

                # Método B2: barra ≈ paralela a vd (entra en pantalla).
                # En cortes/elevaciones la distribución visible es vertical (v_up).
                dot_vd = abs(float(bar_dir.DotProduct(vd_ref)))
                if dot_vd > 0.8:
                    return v_up_ref

                # Método B1: bar_dir × vd.
                spacing = bar_dir.CrossProduct(vd_ref)
                if spacing.GetLength() > 1e-9:
                    return spacing.Normalize()

        return fallback_rd
    except Exception:
        return fallback_rd


def _try_mra_fundacion(document, vista, rb, mrat_type, offset_ft_override, is_section=False, elemento=None):
    """
    Crea una ``MultiReferenceAnnotation`` para ``rb`` en ``vista``.

    La dirección de distribución se determina con ``_rebar_spacing_direction``
    (métodos A/B2/B1/C) y se proyecta al plano de vista con
    ``_project_onto_view_plane``.  El offset de la línea MRA se calcula
    automáticamente a partir del bbox del rebar.

    ``is_section``: cuando es ``True`` usa el centro del array como origen del
    MRA y adapta la dirección de cota según el tipo de barra:

    - Malla inferior/superior (``Armadura_Ubicacion`` = ``"F"``/``"F'"``): dirección
      horizontal (``RightDirection``) porque las barras aparecen como puntos
      distribuidos horizontalmente en la sección.
    - Barras laterales (``"L"``): dirección vertical (``UpDirection``) calculada
      por el Método B2 de ``_rebar_spacing_direction``.
    """
    try:
        from Autodesk.Revit.DB import (
            DimensionStyleType,
            MultiReferenceAnnotation,
            MultiReferenceAnnotationOptions,
        )
        from System.Collections.Generic import List

        if rb is None or mrat_type is None:
            return False

        vd = vista.ViewDirection.Normalize()
        rd = vista.RightDirection.Normalize()
        v_up = vista.UpDirection.Normalize()

        # Leer Armadura_Ubicacion para diferenciar el tipo de barra.
        ubicacion = u""
        try:
            _p_ub = rb.LookupParameter(u"Armadura_Ubicacion")
            ubicacion = (_p_ub.AsString() if _p_ub is not None else u"") or u""
        except Exception:
            pass
        _es_malla = ubicacion.strip() in (u"F", u"F'")

        # 1. Dirección de distribución 3D (A/B2/B1/C) → proyectada al plano de vista.
        spacing_dir_3d = _rebar_spacing_direction(rb, rd, vd, v_up)
        spacing_dir = _project_onto_view_plane(spacing_dir_3d, vd)
        if spacing_dir is None:
            spacing_dir = rd

        # Para barras de malla horizontal (F/F') en secciones:
        # - Aparecen como puntos distribuidos horizontalmente → spacing_dir = rd.
        # - Para barras paralelas al corte (líneas): la API fallará igualmente.
        if is_section and _es_malla:
            spacing_dir = rd

        # 2. Dirección de offset = perpendicular a spacing_dir en el plano de vista.
        #    Para F/F' en sección se usa ±v_up directamente (abajo/arriba).
        #    Para L en sección se usa spacing_dir × vd (lateral).
        if is_section and _es_malla:
            perp_dir = v_up  # se ajusta el signo al calcular p_line
        else:
            perp_dir = spacing_dir.CrossProduct(vd)
            if perp_dir.GetLength() < 1e-9:
                perp_dir = v_up
            else:
                perp_dir = perp_dir.Normalize()

        # ----------------------------------------------------------------
        # Cálculo del origen MRA (p_bar_03) y del offset (off).
        # En secciones se usa el CENTRO del array (midpoint barra 0 → barra n-1)
        # y un offset dinámico basado en el bbox del rebar + margen.
        # En planta se usa el punto al 0.26 de la curva de la barra media.
        # ----------------------------------------------------------------
        p_bar_03 = None
        off = float(offset_ft_override) if offset_ft_override is not None else UnitUtils.ConvertToInternalUnits(450.0, UnitTypeId.Millimeters)

        try:
            from Autodesk.Revit.DB.Structure import MultiplanarOption as _Mpo

            if is_section:
                # Para secciones se usa el centro del bbox del rebar en coordenadas
                # GLOBALES (get_BoundingBox(None)) como origen del tag.
                # get_BoundingBox(view) devuelve coordenadas locales de la vista, por
                # lo que no puede usarse directamente como punto en espacio modelo.
                # GetCenterlineCurves con barPositionIndex tampoco es fiable (devuelve
                # la misma posición para todos los índices). El bbox global cubre todo
                # el array, así que su centro es el punto más preciso para centrar el tag.
                _bb_rb = rb.get_BoundingBox(None)
                if _bb_rb is not None:
                    p_bar_03 = (_bb_rb.Min + _bb_rb.Max) * 0.5
                    # Offset: mitad de la extensión global en la dirección
                    # perpendicular a la cota + margen fijo.
                    _perp_ref = v_up if _es_malla else perp_dir
                    _dim = abs(float((_bb_rb.Max - _bb_rb.Min).DotProduct(_perp_ref)))
                    _margen = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters)
                    off = _dim * 0.5 + _margen
            else:
                # --- Planta: barra 0, punto al 0.26 ---
                _curves = []
                for _mpo_n in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
                    _mpo = getattr(_Mpo, _mpo_n, None)
                    if _mpo is None:
                        continue
                    try:
                        _cs = list(rb.GetCenterlineCurves(False, False, False, _mpo, 0))
                        if _cs:
                            _curves = _cs
                            break
                    except Exception:
                        pass
                if not _curves:
                    try:
                        _curves = list(rb.GetCenterlineCurves(False, False, False))
                    except Exception:
                        pass
                if _curves:
                    _longest = max(_curves, key=lambda c: float(c.Length))
                    p_bar_03 = _longest.Evaluate(0.26, True)
        except Exception:
            pass

        # Fallback: centro del bbox si no hay curvas
        if p_bar_03 is None:
            bb = rb.get_BoundingBox(None)
            if bb is None:
                return False
            p_bar_03 = XYZ(
                0.5 * (float(bb.Min.X) + float(bb.Max.X)),
                0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
                0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
            )

        # Posición de la línea MRA (origen de la cota):
        #   F  → debajo de la sección (-v_up)
        #   F' → arriba de la sección (+v_up)
        #   L  → lateral, siempre hacia afuera del hormigón
        if is_section and _es_malla:
            if ubicacion.strip() == u"F":
                p_line = p_bar_03 - v_up.Multiply(off)
            else:  # F'
                p_line = p_bar_03 + v_up.Multiply(off)
        elif is_section and not _es_malla:
            # Barras laterales: determinar dirección "hacia afuera" del elemento.
            # Se usa el centro del elemento host como referencia; si no está disponible
            # se usa la proyección de p_bar_03 sobre rd (positivo = hacia la derecha).
            _outward = None
            try:
                if elemento is not None:
                    _bb_el = elemento.get_BoundingBox(None)
                    if _bb_el is not None:
                        _elem_ctr = (_bb_el.Min + _bb_el.Max) * 0.5
                        _to_bar = p_bar_03 - _elem_ctr
                        _proj = _project_onto_view_plane(_to_bar, vd)
                        if _proj is not None:
                            # Proyectar sobre el eje horizontal de la vista (rd)
                            _horiz = float(_proj.DotProduct(rd))
                            _outward = rd if _horiz >= 0 else rd.Negate()
            except Exception:
                pass
            if _outward is None:
                # Fallback: usar perp_dir del CrossProduct (puede estar equivocado
                # para uno de los lados, pero es mejor que nada)
                _outward = perp_dir
            p_line = p_bar_03 + _outward.Multiply(off)
        else:
            p_line = p_bar_03 - perp_dir.Multiply(off)

        try:
            opts = MultiReferenceAnnotationOptions(mrat_type)
        except Exception:
            return False

        try:
            opts.DimensionStyleType = DimensionStyleType.Linear
        except Exception:
            pass

        opts.DimensionPlaneNormal = vd
        opts.DimensionLineDirection = spacing_dir
        opts.DimensionLineOrigin = p_line
        opts.TagHeadPosition = p_line
        try:
            opts.TagHasLeader = False
        except Exception:
            pass

        ids = List[ElementId]()
        ids.Add(rb.Id)
        opts.SetElementsToDimension(ids)

        try:
            if hasattr(opts, "ElementsMatchReferenceCategory"):
                if not opts.ElementsMatchReferenceCategory(document):
                    return False
        except Exception:
            pass

        try:
            mra = MultiReferenceAnnotation.Create(document, vista.Id, opts)
            return mra is not None
        except Exception:
            return False
    except Exception:
        return False


def _cotar_fundacion_en_vista(document, vista, elemento):
    """
    Coloca cotas de ancho y largo en planta para ``elemento`` en ``vista``,
    reutilizando la lógica de ``cota_fundacion_planta_rps._cotar_fundacion``.
    Requiere transacción abierta por el llamador.
    """
    if document is None or vista is None or elemento is None:
        return
    try:
        from cota_fundacion_planta_rps import _cotar_fundacion
        _cotar_fundacion(document, vista, elemento)
    except Exception:
        pass


def _multi_rebar_annotations_fundacion(vista, document, elemento, ubicacion_armadura=None):
    """
    Crea una ``MultiReferenceAnnotation`` por cada set de armadura alojado en ``elemento``
    detectando automáticamente la dirección de distribución de cada set, para que funcione
    correctamente con la malla (dos sets perpendiculares) de la fundación.
    La vista debe estar dentro de una transacción abierta.

    ``ubicacion_armadura``: si se indica, solo se anotan rebars cuyo parámetro
    ``Armadura_Ubicacion`` coincide con ese valor (p. ej. ``u"F"`` o ``u"F'"``).
    """
    if vista is None or elemento is None:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
        from geometria_estribos_viga import _multi_reference_annotation_type_by_name

        foundation_id = elemento.Id

        def _ubicacion_ok(rb):
            if not ubicacion_armadura:
                return True
            try:
                _p = rb.LookupParameter(u"Armadura_Ubicacion")
                _v = (_p.AsString() if _p is not None else u"") or u""
            except Exception:
                _v = u""
            return _v.strip() == ubicacion_armadura

        rebars_fundacion = [
            rb for rb in FilteredElementCollector(document).OfClass(Rebar).ToElements()
            if rb is not None and _host_id_safe(rb) == foundation_id and _ubicacion_ok(rb)
        ]
        if not rebars_fundacion:
            return

        mrat_type = _multi_reference_annotation_type_by_name(document, _MRA_TYPE_NAME)
        if mrat_type is None:
            return

        off_ft = UnitUtils.ConvertToInternalUnits(450.0, UnitTypeId.Millimeters)
        for rb in rebars_fundacion:
            _try_mra_fundacion(document, vista, rb, mrat_type, off_ft)
    except Exception:
        pass


def _etiquetar_sin_mra_seccion(document, vista, rebars_sin_mra):
    """
    Coloca etiquetas ``EST_A_STRUCTURAL REBAR TAG`` sobre barras que no pudieron
    ser anotadas con MRA en una vista de sección.

    El tipo de etiqueta se selecciona según el nombre del ``RebarShape``.
    La posición de la etiqueta se desplaza hacia AFUERA del hormigón:

    - ``F``  (inferior) → hacia abajo  (``-v_up``)
    - ``F'`` (superior) → hacia arriba (``+v_up``)
    - ``L``  (lateral)  → lateral, alejado del centro del elemento.

    Debe llamarse dentro de una transacción abierta (``use_transaction=False``).
    Tras la creación se activa ``HasLeader`` en cada etiqueta generada.
    """
    if document is None or vista is None or not rebars_sin_mra:
        return
    try:
        from Autodesk.Revit.DB import IndependentTag, TagMode, TagOrientation
        from enfierrado_shaft_hashtag import (
            _collect_rebar_tag_symbol_map,
            _primary_rebar_shape_tag_key,
            _rebar_shape_name_candidates,
            _rebar_reference_candidates_for_tag,
        )

        tag_map = _collect_rebar_tag_symbol_map(document, _TAG_FAMILY_NAME)
        if not tag_map:
            return

        rd     = vista.RightDirection.Normalize()
        vd     = vista.ViewDirection.Normalize()
        v_up   = vista.UpDirection.Normalize()

        off_ft = UnitUtils.ConvertToInternalUnits(450.0, UnitTypeId.Millimeters)

        created_tags = []

        for rb in rebars_sin_mra:
            if rb is None:
                continue
            try:
                # --- Tipo de tag según shape ---
                primary = _primary_rebar_shape_tag_key(document, rb)
                tag_type_id = tag_map.get(primary) if primary else None
                if tag_type_id is None:
                    for sk in _rebar_shape_name_candidates(document, rb):
                        tag_type_id = tag_map.get(sk)
                        if tag_type_id is not None:
                            break
                if tag_type_id is None:
                    continue

                try:
                    ts = document.GetElement(tag_type_id)
                    if ts is not None and not ts.IsActive:
                        ts.Activate()
                except Exception:
                    pass

                # --- Posición base: centro del bbox del rebar ---
                _bb = rb.get_BoundingBox(None)
                if _bb is None:
                    continue
                p_base = (_bb.Min + _bb.Max) * 0.5

                # --- Dirección outward según Armadura_Ubicacion ---
                ubicacion = u""
                try:
                    _p_ub = rb.LookupParameter(u"Armadura_Ubicacion")
                    ubicacion = (_p_ub.AsString() if _p_ub is not None else u"") or u""
                except Exception:
                    pass
                ub = ubicacion.strip()

                _margen = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters)
                _dim = abs(float((_bb.Max - _bb.Min).DotProduct(v_up)))
                off = _dim * 0.5 + _margen

                if ub == u"F":
                    p_tag = p_base - v_up.Multiply(off)
                elif ub == u"F'":
                    p_tag = p_base + v_up.Multiply(off)
                else:
                    # L o desconocido: lateral hacia afuera del centro
                    _outward = rd
                    try:
                        _bb_el = rb.Document.GetElement(rb.GetHostId()).get_BoundingBox(None) if rb.GetHostId() is not None else None
                        if _bb_el is not None:
                            _elem_ctr = (_bb_el.Min + _bb_el.Max) * 0.5
                            _to_bar = p_base - _elem_ctr
                            _proj = _to_bar - vd.Multiply(float(_to_bar.DotProduct(vd)))
                            if _proj.GetLength() > 1e-9:
                                _horiz = float(_proj.DotProduct(rd))
                                _outward = rd if _horiz >= 0 else rd.Negate()
                    except Exception:
                        pass
                    _dim_l = abs(float((_bb.Max - _bb.Min).DotProduct(rd)))
                    off_l = _dim_l * 0.5 + _margen
                    p_tag = p_base + _outward.Multiply(off_l)

                # --- Referencias y creación del tag ---
                refs = _rebar_reference_candidates_for_tag(document, vista, rb)
                if not refs:
                    continue

                tag_created = None
                for ref in refs:
                    for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                        for leader in (True, False):
                            try:
                                tag_created = IndependentTag.Create(
                                    document, tag_type_id, vista.Id,
                                    ref, leader, orient, p_tag)
                                if tag_created is not None:
                                    break
                            except Exception:
                                tag_created = None
                        if tag_created is not None:
                            break
                    if tag_created is not None:
                        break
                if tag_created is None:
                    for ref in refs:
                        for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                            for leader in (True, False):
                                try:
                                    tag_created = IndependentTag.Create(
                                        document, vista.Id, ref, leader,
                                        TagMode.TM_ADDBY_CATEGORY, orient, p_tag)
                                    if tag_created is not None:
                                        break
                                except Exception:
                                    tag_created = None
                            if tag_created is not None:
                                break
                        if tag_created is not None:
                            break
                if tag_created is not None:
                    try:
                        tag_created.HasLeader = True
                    except Exception:
                        pass
                    created_tags.append(tag_created)
            except Exception:
                continue
    except Exception:
        pass


def _mra_secciones_fundacion(document, vista, elemento):
    """
    Crea una ``MultiReferenceAnnotation`` (tipo «Structural Rebar_Malla») por cada set
    de armadura alojado en ``elemento`` visible en la sección ``vista``.
    Las barras que no puedan anotarse con MRA reciben un ``IndependentTag``
    (``EST_A_STRUCTURAL REBAR TAG``, tipo según RebarShape).
    Debe llamarse dentro de una transacción abierta.
    """
    if document is None or vista is None or elemento is None:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
        from geometria_estribos_viga import _multi_reference_annotation_type_by_name

        foundation_id = elemento.Id
        rebars = [
            rb for rb in FilteredElementCollector(document).OfClass(Rebar).ToElements()
            if rb is not None and _host_id_safe(rb) == foundation_id
        ]
        if not rebars:
            return

        mrat_type = _multi_reference_annotation_type_by_name(document, _MRA_TYPE_NAME_SECCION)

        off_ft = UnitUtils.ConvertToInternalUnits(450.0, UnitTypeId.Millimeters)
        rebars_sin_mra = []
        for rb in rebars:
            if mrat_type is not None:
                ok = _try_mra_fundacion(document, vista, rb, mrat_type, off_ft, is_section=True, elemento=elemento)
            else:
                ok = False
            if not ok:
                rebars_sin_mra.append(rb)

        if rebars_sin_mra:
            _etiquetar_sin_mra_seccion(document, vista, rebars_sin_mra)
    except Exception:
        pass


def _etiquetar_rebars_fundacion(vista, documento, elemento, ubicacion_armadura=None):
    """
    Coloca etiquetas ``EST_A_STRUCTURAL REBAR TAG`` sobre cada barra de armadura
    alojada en ``elemento`` dentro de ``vista``.

    El tipo de etiqueta se selecciona según el nombre del ``RebarShape`` usando
    ``etiquetar_rebars_creados_en_vista`` de ``enfierrado_shaft_hashtag`` (misma
    lógica probada en secciones).
    Debe llamarse dentro de una transacción abierta (``use_transaction=False``).
    """
    if vista is None or elemento is None:
        return
    try:
        from Autodesk.Revit.DB.Structure import Rebar
        from Autodesk.Revit.DB import IndependentTag, TagMode, TagOrientation, Reference
        from enfierrado_shaft_hashtag import (
            _collect_rebar_tag_symbol_map,
            _primary_rebar_shape_tag_key,
            _rebar_shape_name_candidates,
            _rebar_reference_candidates_for_tag,
        )

        tag_map = _collect_rebar_tag_symbol_map(documento, _TAG_FAMILY_NAME)
        if not tag_map:
            return

        foundation_id = elemento.Id

        for rb in FilteredElementCollector(documento).OfClass(Rebar).ToElements():
            try:
                host_id = None
                try:
                    host_id = rb.GetHostId()
                except Exception:
                    pass
                if host_id != foundation_id:
                    continue
                if ubicacion_armadura:
                    try:
                        _p_ub = rb.LookupParameter(u"Armadura_Ubicacion")
                        _val_ub = (_p_ub.AsString() if _p_ub is not None else u"") or u""
                    except Exception:
                        _val_ub = u""
                    if _val_ub.strip() != ubicacion_armadura:
                        continue

                # Tipo de tag según shape (misma lógica que enfierrado_shaft_hashtag)
                primary = _primary_rebar_shape_tag_key(documento, rb)
                tag_type_id = tag_map.get(primary) if primary else None
                if tag_type_id is None:
                    for sk in _rebar_shape_name_candidates(documento, rb):
                        tag_type_id = tag_map.get(sk)
                        if tag_type_id is not None:
                            break
                if tag_type_id is None:
                    continue

                # Activar tipo si es necesario
                try:
                    ts = documento.GetElement(tag_type_id)
                    if ts is not None and not ts.IsActive:
                        ts.Activate()
                except Exception:
                    pass

                # Posición: 0.27 sobre la curva de la barra media
                # Estrategia: obtener curva de barPositionIndex=0, luego trasladarla
                # al centro de distribución (bbox global) para construir la barra media.
                p = None
                try:
                    from Autodesk.Revit.DB.Structure import MultiplanarOption
                    # 1. Curva de la barra en posición 0 (la más larga = barra sin ganchos)
                    curvas_0 = []
                    for mpo_name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
                        mpo = getattr(MultiplanarOption, mpo_name, None)
                        if mpo is None:
                            continue
                        try:
                            curvas_0 = list(rb.GetCenterlineCurves(False, False, False, mpo, 0))
                            if curvas_0:
                                break
                        except Exception:
                            pass
                    if not curvas_0:
                        try:
                            curvas_0 = list(rb.GetCenterlineCurves(False, False, False))
                        except Exception:
                            pass
                    curva_dom = None
                    best_len = -1.0
                    for _c in curvas_0:
                        if _c is None:
                            continue
                        try:
                            _ln = float(_c.Length)
                        except Exception:
                            _ln = 0.0
                        if _ln > best_len:
                            curva_dom = _c
                            best_len = _ln

                    if curva_dom is not None:
                        # 2. Centro de distribución desde bbox global (todos los bars)
                        bb_all = rb.get_BoundingBox(None)
                        if bb_all is not None:
                            centro_dist = (bb_all.Min + bb_all.Max) * 0.5
                        else:
                            centro_dist = curva_dom.Evaluate(0.5, True)

                        # 3. Dirección de la barra y componente perpendicular al offset
                        pt0 = curva_dom.GetEndPoint(0)
                        pt1 = curva_dom.GetEndPoint(1)
                        bar_vec = pt1 - pt0
                        bar_len = float(bar_vec.GetLength())
                        if bar_len > 1e-6:
                            bar_dir = bar_vec.Normalize()
                            # Offset perpendicular desde barra-0 al centro de distribución
                            bar0_mid = curva_dom.Evaluate(0.5, True)
                            diff = centro_dist - bar0_mid
                            along = diff.DotProduct(bar_dir)
                            offset_perp = diff - bar_dir.Multiply(along)

                            # 4. Barra media trasladada → evaluar en 0.27
                            mid_start = pt0 + offset_perp
                            mid_end   = pt1 + offset_perp
                            p = mid_start + (mid_end - mid_start) * 0.27
                        else:
                            p = curva_dom.Evaluate(0.27, True)
                except Exception:
                    pass
                # Fallback: centro del bbox en la vista
                if p is None:
                    try:
                        bb = rb.get_BoundingBox(vista)
                        if bb is not None:
                            p = (bb.Min + bb.Max) * 0.5
                    except Exception:
                        pass
                if p is None:
                    try:
                        bb = rb.get_BoundingBox(None)
                        if bb is not None:
                            p = (bb.Min + bb.Max) * 0.5
                    except Exception:
                        pass
                if p is None:
                    continue

                # Referencias usando la función robusta de enfierrado_shaft_hashtag
                refs = _rebar_reference_candidates_for_tag(documento, vista, rb)
                if not refs:
                    continue

                # Crear tag
                created = None
                last_ex_msg = u""
                for ref in refs:
                    for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                        for leader in (False, True):
                            try:
                                created = IndependentTag.Create(
                                    documento, tag_type_id, vista.Id,
                                    ref, leader, orient, p)
                                if created is not None:
                                    break
                            except Exception as _cex:
                                last_ex_msg = str(_cex)
                                created = None
                        if created is not None:
                            break
                    if created is not None:
                        break
                if created is None:
                    for ref in refs:
                        for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
                            for leader in (False, True):
                                try:
                                    created = IndependentTag.Create(
                                        documento, vista.Id, ref, leader,
                                        TagMode.TM_ADDBY_CATEGORY, orient, p)
                                    if created is not None:
                                        try:
                                            created.SetTypeId(tag_type_id)
                                        except Exception:
                                            pass
                                        break
                                except Exception as _cex2:
                                    last_ex_msg = str(_cex2)
                                    created = None
                            if created is not None:
                                break
                        if created is not None:
                            break
            except Exception:
                pass
    except Exception:
        pass


def _aplicar_cropbox_fundacion(vista, elemento):
    """
    Activa y ajusta el ``CropBox`` de la ``ViewPlan`` ``vista`` para encuadrar ``elemento``.
    El marco de región de recorte queda oculto (``CropBoxVisible = False``).

    El área de recorte es un 25 % mayor que el bbox del elemento (12.5 % por cada lado),
    con un mínimo absoluto de ``_CROPBOX_MARGEN_MIN_MM`` por eje.

    Estrategia:
    1. Obtener el bbox del elemento en coordenadas mundo.
    2. Intentar mapear las 8 esquinas al sistema local de la vista via
       ``CropBox.Transform.Inverse.OfPoint()``.
    3. Si falla (Transform inválido / no inicializado), usar las coordenadas mundo
       directamente — válido para vistas sin rotación de planta (caso habitual).
    La vista debe estar completamente creada y el documento regenerado antes de llamar.
    """
    if vista is None or elemento is None:
        return

    m_min = _mm_a_interno(_CROPBOX_MARGEN_MIN_MM)

    # Obtener bbox del elemento en coordenadas mundo.
    try:
        bb = elemento.get_BoundingBox(None)
        if bb is None or bb.Min is None or bb.Max is None:
            return
    except Exception:
        return

    # Margen proporcional: 12.5 % del span en cada eje → +25 % total del bbox.
    mx = max(abs(float(bb.Max.X) - float(bb.Min.X)) * (_CROPBOX_ESCALA / 2.0), m_min)
    my = max(abs(float(bb.Max.Y) - float(bb.Min.Y)) * (_CROPBOX_ESCALA / 2.0), m_min)

    corners = [
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    ]

    # Intentar mapeo via Transform.Inverse del CropBox actual.
    applied = False
    try:
        cb_orig = vista.CropBox
        if cb_orig is not None and cb_orig.Transform is not None:
            tr = cb_orig.Transform
            tr_inv = tr.Inverse
            z_min_v = float(cb_orig.Min.Z)
            z_max_v = float(cb_orig.Max.Z)
            if abs(z_max_v - z_min_v) < 0.01:
                z_min_v = _mm_a_interno(-3000.0)
                z_max_v = _mm_a_interno(3000.0)
            xs, ys = [], []
            for pt in corners:
                lp = tr_inv.OfPoint(pt)
                xs.append(float(lp.X))
                ys.append(float(lp.Y))
            if xs and ys:
                span_x = max(xs) - min(xs)
                span_y = max(ys) - min(ys)
                px = max(span_x * (_CROPBOX_ESCALA / 2.0), m_min)
                py = max(span_y * (_CROPBOX_ESCALA / 2.0), m_min)
                cb_new = BoundingBoxXYZ()
                cb_new.Transform = tr
                cb_new.Min = XYZ(min(xs) - px, min(ys) - py, z_min_v)
                cb_new.Max = XYZ(max(xs) + px, max(ys) + py, z_max_v)
                vista.CropBox = cb_new
                applied = True
    except Exception:
        pass

    # Respaldo: coordenadas mundo directas (vistas de planta sin rotación).
    if not applied:
        try:
            z_min_fb = _mm_a_interno(-3000.0)
            z_max_fb = _mm_a_interno(3000.0)
            cb_fb = BoundingBoxXYZ()
            cb_fb.Min = XYZ(float(bb.Min.X) - mx, float(bb.Min.Y) - my, z_min_fb)
            cb_fb.Max = XYZ(float(bb.Max.X) + mx, float(bb.Max.Y) + my, z_max_fb)
            vista.CropBox = cb_fb
        except Exception:
            return

    try:
        vista.CropBoxActive = True
        # Recorte activo pero sin mostrar el marco de región de recorte en la vista.
        vista.CropBoxVisible = False
    except Exception:
        pass


def crear_vista_planta_fundacion_aislada(
    document,
    elemento,
    uidocument=None,
    gestionar_transaccion=True,
    ubicacion_armadura=None,
):
    """
    Crea una ``ViewPlan`` de tipo «Structural Plan (Piso)» para revisar la armadura
    de la fundación aislada ``elemento``.

    - El nivel asociado es el más cercano a la **cota de la cara inferior** del elemento
      (``get_BoundingBox(None).Min.Z`` como aproximación robusta).
    - La vista queda visible mirando hacia abajo.
    - Si ``uidocument`` se pasa, activa la vista recién creada.
    - ``ubicacion_armadura``: si se pasa (p. ej. ``u"F"`` o ``u"F'"``), la vista muestra
      solo los rebars con ese valor en el parámetro ``Armadura_Ubicacion`` (los demás se
      ocultan) y el nombre de vista incluye el sufijo correspondiente.

    Returns:
        (ViewPlan | None, aviso: str | None)
    """
    if document is None or elemento is None:
        return None, u"document o elemento nulo."

    vft = _find_vft_structural_plan_piso(document)
    if vft is None:
        return None, u"No se encontró ViewFamilyType «Structural Plan (Piso)» en el proyecto."

    z_ref_ft = None
    try:
        bb = elemento.get_BoundingBox(None)
        if bb is not None and bb.Min is not None:
            z_ref_ft = float(bb.Min.Z)
    except Exception:
        pass
    if z_ref_ft is None:
        try:
            from geometria_fundacion_cara_inferior import _inferior_planar_face_info
            r = _inferior_planar_face_info(elemento)
            if r is not None:
                z_ref_ft = float(r[2])
        except Exception:
            pass
    if z_ref_ft is None:
        return None, u"No se pudo determinar la cota de la cara inferior del elemento."

    level = _nivel_mas_proximo(document, z_ref_ft)
    if level is None:
        return None, u"No hay niveles en el documento."

    try:
        elem_id_int = int(elemento.Id.IntegerValue)
    except Exception:
        elem_id_int = 0

    # --- Paso 1: crear la vista y asignar nombre ---
    # Transaction o SubTransaction según si el caller ya tiene una abierta.
    _SUFIJO_MAP = {u"F": u"— Inf", u"F'": u"— Sup"}
    nombre_sufijo = _SUFIJO_MAP.get(ubicacion_armadura) if ubicacion_armadura else None

    vista = None
    aviso = None

    if gestionar_transaccion:
        from Autodesk.Revit.DB import Transaction as _Tx
        _tx1 = _Tx(document, u"BIMTools — Vista planta fundación aislada")
        try:
            _tx1.Start()
        except Exception as ex:
            return None, u"No se pudo iniciar la transacción: {0}".format(ex)
        try:
            vista = ViewPlan.Create(document, vft.Id, level.Id)
            try:
                vista.Name = _nombre_vista_planta_fundacion_unico(document, elem_id_int, nombre_sufijo, elemento)
            except Exception:
                pass
            _tx1.Commit()
        except Exception as ex:
            aviso = u"Error al crear vista de planta: {0}".format(ex)
            try:
                _tx1.RollBack()
            except Exception:
                pass
            return None, aviso
    else:
        from Autodesk.Revit.DB import SubTransaction as _STx
        _stx1 = _STx(document)
        try:
            _stx1.Start()
            vista = ViewPlan.Create(document, vft.Id, level.Id)
            try:
                vista.Name = _nombre_vista_planta_fundacion_unico(document, elem_id_int, nombre_sufijo, elemento)
            except Exception:
                pass
            _stx1.Commit()
        except Exception as ex:
            aviso = u"Error al crear vista de planta: {0}".format(ex)
            try:
                _stx1.RollBack()
            except Exception:
                pass
            return None, aviso

    # --- Paso 2: aplicar crop box + configuración (requiere regenerate) ---
    try:
        document.Regenerate()
    except Exception:
        pass

    def _aplicar_config_vista():
        _aplicar_cropbox_fundacion(vista, elemento)
        try:
            vista.Scale = 25
        except Exception:
            pass
        _activar_visibilidad_armaduras(vista, document)
        try:
            from Autodesk.Revit.DB import BuiltInCategory
            cat_floors = document.Settings.Categories.get_Item(BuiltInCategory.OST_Floors)
            if cat_floors is not None:
                vista.SetCategoryHidden(cat_floors.Id, True)
        except Exception:
            pass
        if ubicacion_armadura:
            _ocultar_rebars_por_ubicacion(vista, document, ubicacion_armadura)
        _aplicar_representacion_show_middle(vista, document)
        _etiquetar_rebars_fundacion(vista, document, elemento, ubicacion_armadura=ubicacion_armadura)
        _multi_rebar_annotations_fundacion(vista, document, elemento, ubicacion_armadura=ubicacion_armadura)
        _cotar_fundacion_en_vista(document, vista, elemento)

    if gestionar_transaccion:
        from Autodesk.Revit.DB import Transaction as _Tx
        _tx2 = _Tx(document, u"BIMTools — Crop box fundación aislada")
        try:
            _tx2.Start()
            _aplicar_config_vista()
            _tx2.Commit()
        except Exception:
            try:
                _tx2.RollBack()
            except Exception:
                pass
    else:
        from Autodesk.Revit.DB import SubTransaction as _STx
        _stx2 = _STx(document)
        try:
            _stx2.Start()
            _aplicar_config_vista()
            _stx2.Commit()
        except Exception:
            try:
                _stx2.RollBack()
            except Exception:
                pass

    if uidocument is not None and vista is not None:
        try:
            uidocument.ActiveView = vista
        except Exception:
            pass

    return vista, aviso


def crear_secciones_fundacion_aislada(
    document,
    elemento,
    uidocument=None,
    margen_mm=1200.0,
    profundidad_mm=800.0,
    far_clip_mm=200.0,
    nombre_prefijo=u"DET. FUND.",
    gestionar_transaccion=True,
):
    """
    Crea dos ``ViewSection`` perpendiculares que cortan la fundación aislada ``elemento``
    por su centro geométrico (lógica extraída de ``secciones_fundacion_aislada_rps.py``):

    - Corte A-A: plano local YZ (dir_corte = eje X local del elemento).
    - Corte B-B: plano local XZ (dir_corte = eje Y local del elemento).

    Los nombres incluyen el parámetro «Numeracion Fundacion» si existe.
    Si ``uidocument`` se pasa, activa la última vista creada.

    Returns:
        (lista de ViewSection creadas, lista de avisos str)
    """
    if document is None or elemento is None:
        return [], [u"document o elemento nulo."]

    # --- ViewFamilyType Section ---
    vft_id = None
    try:
        for vft in FilteredElementCollector(document).OfClass(ViewFamilyType):
            if vft is not None and vft.ViewFamily == ViewFamily.Section:
                vft_id = vft.Id
                break
    except Exception:
        pass
    if vft_id is None:
        return [], [u"No hay ningún tipo de vista «Section» en el proyecto."]

    # --- Helpers locales ---
    _m = _mm_a_interno

    def _nombre_unico_sec(vista, nombre_base):
        existentes = set()
        for v in FilteredElementCollector(document).OfClass(View):
            try:
                if v is None or v.Id == vista.Id:
                    continue
                n = v.Name
                if n:
                    existentes.add(str(n).strip().lower())
            except Exception:
                continue
        cand = nombre_base
        k = 0
        while cand.strip().lower() in existentes:
            k += 1
            cand = u"{0} ({1})".format(nombre_base, k)
        try:
            vista.Name = cand
        except Exception:
            pass

    def _puntos_geom(elem):
        pts = []
        opt = Options()
        try:
            opt.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        try:
            for obj in elem.get_Geometry(opt) or []:
                solids = []
                if isinstance(obj, Solid) and obj.Volume > 1e-9:
                    solids.append(obj)
                elif isinstance(obj, GeometryInstance):
                    try:
                        for sub in obj.GetInstanceGeometry():
                            if isinstance(sub, Solid) and sub.Volume > 1e-9:
                                solids.append(sub)
                    except Exception:
                        pass
                for s in solids:
                    try:
                        for face in s.Faces:
                            try:
                                for v in face.Triangulate().Vertices:
                                    pts.append(v)
                            except Exception:
                                pass
                    except Exception:
                        pass
        except Exception:
            pass
        return pts

    def _esquinas_bb(bb):
        mn, mx = bb.Min, bb.Max
        return [
            XYZ(mn.X, mn.Y, mn.Z), XYZ(mx.X, mn.Y, mn.Z),
            XYZ(mn.X, mx.Y, mn.Z), XYZ(mx.X, mx.Y, mn.Z),
            XYZ(mn.X, mn.Y, mx.Z), XYZ(mx.X, mn.Y, mx.Z),
            XYZ(mn.X, mx.Y, mx.Z), XYZ(mx.X, mx.Y, mx.Z),
        ]

    def _build_tr(origen, dir_corte):
        bz = dir_corte.Normalize()
        bx = XYZ.BasisZ.CrossProduct(bz)
        if bx.GetLength() < 1e-6:
            bx = XYZ.BasisX.CrossProduct(bz)
        if bx.GetLength() < 1e-6:
            return None
        bx = bx.Normalize()
        by = bz.CrossProduct(bx).Normalize()
        tr = Transform.Identity
        tr.Origin = origen
        tr.BasisX = bx
        tr.BasisY = by
        tr.BasisZ = bz
        return tr

    def _crear_vs(elem, origen, dir_corte, label):
        tr = _build_tr(origen, dir_corte)
        if tr is None:
            return None, u"No se pudo construir la orientación del corte."
        pts = _puntos_geom(elem)
        if len(pts) < 4:
            bb_el = elem.get_BoundingBox(None)
            if bb_el is None:
                return None, u"Sin geometría ni BoundingBox."
            pts = _esquinas_bb(bb_el)
        ox, bx, by, bz = tr.Origin, tr.BasisX, tr.BasisY, tr.BasisZ
        xs = [float((p - ox).DotProduct(bx)) for p in pts]
        ys = [float((p - ox).DotProduct(by)) for p in pts]
        zs = [float((p - ox).DotProduct(bz)) for p in pts]
        mm_m = _m(margen_mm)
        near_clip = _m(100.0)
        far_clip = max(max(zs) + mm_m, _m(profundidad_mm) * 0.5)
        xabs = max(abs(min(xs)), abs(max(xs))) + mm_m
        ymn_r = min(ys) - mm_m
        ymx_r = max(ys) + mm_m
        ymid = 0.5 * (ymn_r + ymx_r)
        if abs(ymid) > 1e-9:
            tr.Origin = ox.Add(by.Multiply(ymid))
        yabs = max(abs(ymn_r - ymid), abs(ymx_r - ymid))
        box = BoundingBoxXYZ()
        box.Transform = tr
        box.Min = XYZ(-xabs, -yabs, -near_clip)
        box.Max = XYZ(xabs, yabs, far_clip)
        try:
            vs = ViewSection.CreateSection(document, vft_id, box)
        except Exception as ex:
            return None, u"CreateSection falló: {0}".format(ex)
        try:
            vs.CropBoxVisible = False
        except Exception:
            pass
        try:
            p_far = vs.get_Parameter(BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
            if p_far is not None and not p_far.IsReadOnly:
                p_far.Set(_m(far_clip_mm))
        except Exception:
            pass
        try:
            vs.Scale = 25
        except Exception:
            pass
        try:
            _nombre_unico_sec(vs, label)
        except Exception:
            pass
        return vs, None

    def _ejes_locales(elem):
        try:
            opt = Options()
            try:
                opt.DetailLevel = ViewDetailLevel.Fine
            except Exception:
                pass
            candidatos = {}
            def _acumular(solido):
                for edge in solido.Edges:
                    try:
                        curva = edge.AsCurve()
                        if not isinstance(curva, Line):
                            continue
                        p0 = curva.GetEndPoint(0)
                        p1 = curva.GetEndPoint(1)
                        if abs(float(p1.Z - p0.Z)) > 1e-3:
                            continue
                        v = XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
                        lg = float(v.GetLength())
                        if lg < 1e-6:
                            continue
                        d = v.Normalize()
                        if float(d.X) < -1e-6 or (abs(float(d.X)) < 1e-6 and float(d.Y) < 0):
                            d = d.Negate()
                        key = (round(float(d.X), 4), round(float(d.Y), 4))
                        candidatos[key] = candidatos.get(key, 0.0) + lg
                    except Exception:
                        continue
            for obj in elem.get_Geometry(opt) or []:
                if isinstance(obj, Solid) and obj.Volume > 1e-9:
                    _acumular(obj)
                elif isinstance(obj, GeometryInstance):
                    try:
                        for sub in obj.GetInstanceGeometry():
                            if isinstance(sub, Solid) and sub.Volume > 1e-9:
                                _acumular(sub)
                    except Exception:
                        pass
            if candidatos:
                ordenados = sorted(candidatos.items(), key=lambda kv: -kv[1])
                kx, _ = ordenados[0]
                dir_x = XYZ(kx[0], kx[1], 0.0).Normalize()
                dir_y = XYZ.BasisZ.CrossProduct(dir_x).Normalize()
                if len(ordenados) >= 2:
                    k2, _ = ordenados[1]
                    d2 = XYZ(k2[0], k2[1], 0.0).Normalize()
                    if abs(float(dir_x.DotProduct(d2))) < 0.1:
                        dir_y = d2
                return dir_x, dir_y
        except Exception:
            pass
        try:
            from Autodesk.Revit.DB import FamilyInstance as _FI
            if isinstance(elem, _FI):
                fi_tr = elem.GetTotalTransform()
                if fi_tr is not None:
                    bx = fi_tr.BasisX
                    if bx is not None and bx.GetLength() > 1e-9:
                        bx_h = XYZ(float(bx.X), float(bx.Y), 0.0)
                        if bx_h.GetLength() > 1e-6:
                            dir_x = bx_h.Normalize()
                            return dir_x, XYZ.BasisZ.CrossProduct(dir_x).Normalize()
        except Exception:
            pass
        try:
            bb = elemento.get_BoundingBox(None)
            if bb is not None:
                sx = abs(float(bb.Max.X) - float(bb.Min.X))
                sy = abs(float(bb.Max.Y) - float(bb.Min.Y))
                if sy > sx:
                    return XYZ.BasisY, XYZ.BasisX
        except Exception:
            pass
        return XYZ.BasisX, XYZ.BasisY

    # --- Lectura de numeración ---
    def _num_str(elem):
        for name in (u"Numeracion Fundacion", u"Numeracion fundacion",
                     u"Foundation Numbering", u"Numeracion"):
            try:
                from Autodesk.Revit.DB import StorageType
                p = elem.LookupParameter(name)
                if p is None or not p.HasValue:
                    continue
                if p.StorageType == StorageType.String:
                    s = p.AsString()
                    if s and str(s).strip():
                        return str(s).strip()
                elif p.StorageType == StorageType.Integer:
                    return str(p.AsInteger())
                elif p.StorageType == StorageType.Double:
                    return str(int(round(p.AsDouble())))
                vs2 = p.AsValueString()
                if vs2 and str(vs2).strip():
                    return str(vs2).strip()
            except Exception:
                continue
        try:
            return str(int(elemento.Id.IntegerValue))
        except Exception:
            return u"?"

    # --- Centro geométrico ---
    pts_c = _puntos_geom(elemento)
    if len(pts_c) >= 4:
        xs_c = [float(p.X) for p in pts_c]
        ys_c = [float(p.Y) for p in pts_c]
        zs_c = [float(p.Z) for p in pts_c]
        centro = XYZ(
            0.5 * (min(xs_c) + max(xs_c)),
            0.5 * (min(ys_c) + max(ys_c)),
            0.5 * (min(zs_c) + max(zs_c)),
        )
    else:
        bb = elemento.get_BoundingBox(None)
        if bb is None:
            return [], [u"El elemento no tiene geometría ni BoundingBox."]
        centro = XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
        )

    dir_x, dir_y = _ejes_locales(elemento)
    num = _num_str(elemento)
    label_aa = u"{0} F{1}_1".format(nombre_prefijo, num)
    label_bb = u"{0} F{1}_2".format(nombre_prefijo, num)

    # Detectar si la fundación es cuadrada proyectando los puntos sobre cada eje local.
    # Tolerancia: 1 mm en unidades internas (~0.00328 ft).
    _TOL_CUADRADA_FT = 1.0 / 304.8
    es_cuadrada = False
    try:
        _pts = _puntos_geom(elemento) or []
        if len(_pts) >= 2:
            _proj_x = [float(p.X) * float(dir_x.X) + float(p.Y) * float(dir_x.Y) for p in _pts]
            _proj_y = [float(p.X) * float(dir_y.X) + float(p.Y) * float(dir_y.Y) for p in _pts]
            _dim_x = max(_proj_x) - min(_proj_x)
            _dim_y = max(_proj_y) - min(_proj_y)
            es_cuadrada = abs(_dim_x - _dim_y) <= _TOL_CUADRADA_FT
    except Exception:
        es_cuadrada = False

    avisos = []
    vistas = []

    def _abrir_tx_sec():
        if gestionar_transaccion:
            _tx = Transaction(document, u"BIMTools \u2014 Secciones fundaci\u00f3n aislada")
        else:
            from Autodesk.Revit.DB import SubTransaction
            _tx = SubTransaction(document)
        return _tx

    def _abrir_tx_mra():
        if gestionar_transaccion:
            _tx = Transaction(document, u"BIMTools \u2014 MRA secciones fundaci\u00f3n aislada")
        else:
            from Autodesk.Revit.DB import SubTransaction
            _tx = SubTransaction(document)
        return _tx

    tx = _abrir_tx_sec()
    try:
        tx.Start()
        vs_aa, err_aa = _crear_vs(elemento, centro, dir_x, label_aa)
        if vs_aa is not None:
            vistas.append(vs_aa)
        elif err_aa:
            avisos.append(u"A-A: " + err_aa)
        if not es_cuadrada:
            vs_bb, err_bb = _crear_vs(elemento, centro, dir_y, label_bb)
            if vs_bb is not None:
                vistas.append(vs_bb)
            elif err_bb:
                avisos.append(u"B-B: " + err_bb)
        tx.Commit()
    except Exception as ex:
        try:
            tx.RollBack()
        except Exception:
            pass
        return [], [u"Error al crear secciones: {0}".format(ex)]

    # Segunda transacción/subtransacción: MRA en cada sección creada.
    if vistas:
        tx2 = _abrir_tx_mra()
        try:
            tx2.Start()
            for _vs in vistas:
                _mra_secciones_fundacion(document, _vs, elemento)
            tx2.Commit()
        except Exception as _ex2:
            try:
                tx2.RollBack()
            except Exception:
                pass
            avisos.append(u"MRA secciones no creadas: {0}".format(_ex2))

    if uidocument is not None and vistas:
        try:
            uidocument.ActiveView = vistas[-1]
        except Exception:
            pass

    return vistas, avisos


def eliminar_vistas_seccion_revision_enfierrado(document, ids_vista, uidocument=None):
    """
    Elimina vistas de sección de revisión: Ids registrados + vistas cuyo nombre
    empieza por ``BIMTools — Rev. enfierrado —`` (respaldo si falló el registro).

    Si la vista activa está entre ellas, cambia a otra vista antes de borrar.

    Returns:
        Cantidad de Ids enviados a ``Delete`` (puede ser 0).
    """
    if document is None:
        return 0
    vistos = set()
    candidates = []
    for eid in ids_vista or []:
        ii = _element_id_a_int(eid)
        if ii is not None and ii not in vistos:
            vistos.add(ii)
            candidates.append(eid)
    for v in FilteredElementCollector(document).OfClass(View):
        try:
            if v is None or getattr(v, "IsTemplate", False):
                continue
            n = v.Name
            if not n:
                continue
            try:
                ns = unicode(n)
            except Exception:
                ns = str(n)
            if not ns.startswith(_NOMBRE_PREFIJO_SECCION_REVISION):
                continue
            ii = _element_id_a_int(v.Id)
            if ii is not None and ii not in vistos:
                vistos.add(ii)
                candidates.append(v.Id)
        except Exception:
            continue
    if not candidates:
        return 0
    to_del = []
    for eid in candidates:
        try:
            el = document.GetElement(eid)
            if _es_vista_seccion_revision(el):
                to_del.append(eid)
        except Exception:
            continue
    if not to_del:
        return 0
    del_ints = set()
    for eid in to_del:
        ii = _element_id_a_int(eid)
        if ii is not None:
            del_ints.add(ii)
    if uidocument is not None and del_ints:
        try:
            av = uidocument.ActiveView
            if av is not None:
                avi = _element_id_a_int(av.Id)
                if avi is not None and avi in del_ints:
                    for v in FilteredElementCollector(document).OfClass(View):
                        try:
                            if v is None or getattr(v, "IsTemplate", False):
                                continue
                            vi = _element_id_a_int(v.Id)
                            if vi is None or vi in del_ints:
                                continue
                            uidocument.ActiveView = v
                            break
                        except Exception:
                            continue
        except Exception:
            pass
    t = Transaction(document, u"BIMTools — Eliminar secciones revisión enfierrado")
    t.Start()
    try:
        try:
            from System.Collections.Generic import List as NetList

            lid = NetList[ElementId]()
            for eid in to_del:
                lid.Add(eid)
            document.Delete(lid)
        except Exception:
            for eid in to_del:
                try:
                    document.Delete(eid)
                except Exception:
                    pass
        t.Commit()
        return len(to_del)
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return 0
