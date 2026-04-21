# -*- coding: utf-8 -*-
"""
Geometría de fundación: cara inferior y contorno.

- Identificación de la cara inferior: misma lógica que el antiguo tool Fundación aislada,
  con criterio relajado si falla la clasificación por normal.
- Perímetro: Face.GetEdgesAsCurveLoops; si no hay loops, fallback por teselación.
- Una curva del lado más corto o del **más largo**; si no hay CurveLoop válido, segmento por teselación.
- Evaluación de caras paralelas a una curva y selección de la más cercana al eje.
- Caras laterales verticales: enumeración y aristas horizontales para armadura horizontal por cara.
"""

from Autodesk.Revit.DB import (
    Curve,
    CurveLoop,
    GeometryInstance,
    Line,
    Options,
    Plane,
    SketchPlane,
    UV,
    ViewPlan,
    ViewSection,
    XYZ,
)

# Distancia acotada en planta (mm): offset perpendicular hacia el interior (combo inferior).
RECUBRIMIENTO_DEFAULT_MM = 150.0
# Recorte a lo largo de la curva en cada extremo (mm); independiente de la cota en planta.
RECUBRIMIENTO_EXTREMOS_MM = 50.0
# Recubrimiento al hormigón de la cara inferior (mm). El eje de la barra se desplaza hacia el
# interior del elemento: cover + diámetro nominal / 2 (desde la cara al centro de la barra).
RECUBRIMIENTO_INFERIOR_CARA_HORMIGON_MM = 50.0
# Longitud de cada eje del marcador UVN en el detalle (mm).
MARCADOR_EJES_CARA_MM = 300.0
# Longitud visual del segmento que representa el vector ``norm`` de CreateFromCurves (mm).
MARCADOR_NORM_REBAR_MM = 400.0
# Longitud mínima (pies) de una ``Line`` proyectada para ``NewDetailCurve``.
_MIN_DETAIL_LINE_LEN_FT = 0.003
# Tolerancia: curva en el plano de la cara (p. ej. cara inferior) — no elegir como «más cercana otra».
_TOL_COPLANAR_CURVA_FT = 2.0 / 304.8

# Caché de geometría de caras por element.Id para una sola llamada a Execute().
# Se limpia explícitamente al inicio de cada operación con clear_face_cache().
_FACE_ENTRIES_CACHE = {}


def clear_face_cache():
    """Descarta el caché de entradas de caras. Llamar al inicio de cada ExternalEvent.Execute."""
    global _FACE_ENTRIES_CACHE
    _FACE_ENTRIES_CACHE = {}


def _obtener_plano_detalle_vista(view):
    """
    Plano sobre el que Revit dibuja ``DetailCurve`` / ``ModelCurve`` en la vista.

    En **sección, alzado y detalle** (``ViewSection``), el ``SketchPlane`` activo puede ser
    un nivel u otro plano distinto del plano de corte; proyectar ahí desplaza la geometría
    en cota (p. ej. cientos de mm o más). Por eso, para ``ViewSection`` se usa siempre el
    plano geométrico de la vista: normal ``ViewDirection``, punto ``Origin``.
    """
    if view is None:
        return None
    try:
        if isinstance(view, ViewSection):
            vd = view.ViewDirection
            if vd is not None and vd.GetLength() > 1e-12:
                o = view.Origin
                if o is not None:
                    pl = Plane.CreateByNormalAndOrigin(vd.Normalize(), o)
                    if pl is not None:
                        return pl
    except Exception:
        pass
    try:
        sp = view.SketchPlane
        if sp is not None:
            pl = sp.GetPlane()
            if pl is not None:
                return pl
    except Exception:
        pass
    try:
        if isinstance(view, ViewPlan):
            lvl = view.GenLevel
            if lvl is not None:
                z = float(lvl.Elevation)
                try:
                    from Autodesk.Revit.DB import PlanViewPlane

                    vr = view.GetViewRange()
                    z = z + float(vr.GetOffset(PlanViewPlane.CutPlane))
                except Exception:
                    pass
                return Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0.0, 0.0, z))
    except Exception:
        pass
    try:
        o = view.Origin
        if o is not None:
            return Plane.CreateByNormalAndOrigin(
                XYZ.BasisZ, XYZ(float(o.X), float(o.Y), float(o.Z))
            )
    except Exception:
        pass
    return None


def _proyectar_punto_al_plano(pt, plane):
    if pt is None or plane is None:
        return None
    try:
        n = plane.Normal
        if n is None or n.GetLength() < 1e-12:
            return pt
        n = n.Normalize()
        o = plane.Origin
        v = pt - o
        dist = v.DotProduct(n)
        return pt - n.Multiply(dist)
    except Exception:
        return None


def _proyectar_linea_al_plano_vista(line, plane):
    """Proyecta una ``Line`` al plano de la vista; None si degenera."""
    if line is None or plane is None or not isinstance(line, Line):
        return None
    try:
        q0 = _proyectar_punto_al_plano(line.GetEndPoint(0), plane)
        q1 = _proyectar_punto_al_plano(line.GetEndPoint(1), plane)
        if q0 is None or q1 is None:
            return None
        if q0.DistanceTo(q1) < _MIN_DETAIL_LINE_LEN_FT:
            return None
        return Line.CreateBound(q0, q1)
    except Exception:
        return None


def _centro_xy_bbox(elem):
    try:
        bbox = elem.get_BoundingBox(None)
        if bbox is None:
            return XYZ(0, 0, 0)
        mn, mx = bbox.Min, bbox.Max
        return XYZ((mn.X + mx.X) * 0.5, (mn.Y + mx.Y) * 0.5, (mn.Z + mx.Z) * 0.5)
    except Exception:
        return XYZ(0, 0, 0)


def _punto_medio_curva(curve):
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return XYZ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5, (p0.Z + p1.Z) * 0.5)
        except Exception:
            return None


def _distancia_xy(a, b):
    dx = a.X - b.X
    dy = a.Y - b.Y
    return (dx * dx + dy * dy) ** 0.5


def _offset_linea_plano_manual(curve, elemento, offset_ft):
    """Offset en el plano XY (normal Z) cuando CreateOffset no aplica (fallback)."""
    try:
        if not isinstance(curve, Line):
            return None
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < 1e-12:
            return None
        t = d.Normalize()
        perp = XYZ.BasisZ.CrossProduct(t)
        if perp.GetLength() < 1e-12:
            return None
        perp = perp.Normalize()
        cen = _centro_xy_bbox(elemento)
        mid = XYZ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5, (p0.Z + p1.Z) * 0.5)
        v = XYZ(cen.X - mid.X, cen.Y - mid.Y, 0.0)
        if v.GetLength() > 1e-12:
            v = v.Normalize()
            sign = 1.0 if perp.DotProduct(v) >= 0.0 else -1.0
        else:
            sign = 1.0
        off = perp.Multiply(float(offset_ft) * sign)
        return Line.CreateBound(p0.Add(off), p1.Add(off))
    except Exception:
        return None


def aplicar_offset_recubrimiento_mm(curve, elemento, offset_mm=None):
    """
    Offset en **planta** (plano horizontal, normal Z) hacia el **interior** de la fundación,
    equivalente a recubrimiento lateral respecto al borde inferior.

    Usa ``Curve.CreateOffset`` (Revit) y elige el sentido que acerca la curva al centro
    del bounding box en XY. Si falla, usa desplazamiento manual para ``Line``.

    Args:
        curve: ``Curve`` (típicamente ``Line`` o ``Arc``) en la cara inferior.
        elemento: elemento host (fundación) para orientar el offset hacia dentro.
        offset_mm: distancia en mm; por defecto :data:`RECUBRIMIENTO_DEFAULT_MM`.

    Returns:
        Curve | None: nueva curva offset o None.
    """
    if curve is None or elemento is None:
        return None
    if offset_mm is None:
        offset_mm = RECUBRIMIENTO_DEFAULT_MM
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        offset_ft = UnitUtils.ConvertToInternalUnits(float(offset_mm), UnitTypeId.Millimeters)
    except Exception:
        offset_ft = float(offset_mm) / 304.8
    normal_plano = XYZ.BasisZ
    centro = _centro_xy_bbox(elemento)
    candidatos = []
    for sign in (1.0, -1.0):
        try:
            off = float(sign) * offset_ft
            arr = curve.CreateOffset(off, normal_plano)
        except Exception:
            continue
        if arr is None:
            continue
        try:
            n = arr.Count
        except Exception:
            continue
        for i in range(n):
            try:
                c = arr[i]
            except Exception:
                continue
            if c is None:
                continue
            mid = _punto_medio_curva(c)
            if mid is None:
                candidatos.append((c, float("inf")))
            else:
                candidatos.append((c, _distancia_xy(mid, centro)))
    if candidatos:
        candidatos.sort(key=lambda x: x[1])
        return candidatos[0][0]
    return _offset_linea_plano_manual(curve, elemento, offset_ft)


def aplicar_recubrimiento_extremos_mm(curve, rec_mm):
    """
    Recorta la curva **desde cada extremo** una distancia ``rec_mm`` medida **a lo largo**
    de la curva (recubrimiento en extremos).

    Solo se implementa de forma **estable** para ``Line``: ``Arc.Create`` tras recortar
    arcos degeneró cierres duros de Revit al importar lógica de otro módulo.

    Para ``Arc`` u otras curvas, devuelve None (se usa la curva solo con offset en planta).

    Args:
        curve: ``Curve`` ya desplazada en planta si aplica.
        rec_mm: distancia en mm por extremo.

    Returns:
        Line | None: nueva ``Line``, o None.
    """
    if curve is None:
        return None
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        d_ft = UnitUtils.ConvertToInternalUnits(float(rec_mm), UnitTypeId.Millimeters)
    except Exception:
        d_ft = float(rec_mm) / 304.8
    try:
        L = float(curve.Length)
    except Exception:
        return None
    if L < 2.0 * d_ft + 1e-9:
        return None
    if not isinstance(curve, Line):
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        v = p1 - p0
        ln = v.GetLength()
        if ln < 2.0 * d_ft + 1e-12:
            return None
        u = v.Normalize()
        pn0 = p0.Add(u.Multiply(d_ft))
        pn1 = p1.Add(u.Multiply(-d_ft))
        for p in (pn0, pn1):
            for a in (p.X, p.Y, p.Z):
                if abs(float(a)) > 1e15:
                    return None
        return Line.CreateBound(pn0, pn1)
    except Exception:
        return None


def aplicar_recubrimiento_inferior_completo_mm(
    curve, elemento, offset_planta_mm=None, extremos_mm=None
):
    """
    Aplica recubrimiento en **planta** (hacia interior) y recorte en **extremos** (a lo largo de la curva).

    Orden: offset perpendicular → recorte en ambos extremos.

    Args:
        curve: curva del lado corto en la cara inferior.
        elemento: fundación (orienta el offset en planta).
        offset_planta_mm: distancia acotada en planta (mm); por defecto
            :data:`RECUBRIMIENTO_DEFAULT_MM`.
        extremos_mm: recorte por extremo a lo largo de la curva (mm), típicamente hasta el **eje**
            de la barra; por defecto :data:`RECUBRIMIENTO_EXTREMOS_MM`. Para recubrimiento nominal
            ``c`` a la fibra desde el canto, suele pasarse ``c + ø/2`` (eje respecto al hormigón).

    Returns:
        tuple: (curva_tratada, curva_solo_offset) donde ``curva_solo_offset`` es la curva
        tras el offset en planta **sin** recorte en extremos (útil para depuración);
        ``curva_tratada`` incluye extremos o coincide con ``curva_solo_offset`` si el
        recorte extremo no aplica (curva demasiado corta).
    """
    if offset_planta_mm is None:
        offset_planta_mm = RECUBRIMIENTO_DEFAULT_MM
    if extremos_mm is None:
        extremos_mm = RECUBRIMIENTO_EXTREMOS_MM
    c_off = aplicar_offset_recubrimiento_mm(curve, elemento, offset_planta_mm)
    if c_off is None:
        return (None, None)
    c_end = aplicar_recubrimiento_extremos_mm(c_off, extremos_mm)
    if c_end is None:
        return (c_off, c_off)
    return (c_end, c_off)


def offset_linea_eje_barra_desde_cara_inferior_mm(
    line, n_outward, cover_hormigon_mm=None, diametro_nominal_mm=None
):
    """
    Desplaza la ``Line`` (eje teórico de la barra) hacia el **interior** del hormigón respecto
    a la cara inferior: dirección ``-n`` (normal saliente de la cara).

    Distancia en mm: ``cover_hormigon_mm + (diámetro nominal / 2)`` — recubrimiento al
    hormigón más radio de barra para situar el **eje** de la armadura.

    Args:
        line: ``Line`` en el plano de la cara inferior (tras offset en planta / extremos).
        n_outward: normal unitaria de la cara hacia fuera del sólido (``N`` del marco UVN).
            Si es None, se asume cara horizontal con interior según ``+Z``.
        cover_hormigon_mm: por defecto :data:`RECUBRIMIENTO_INFERIOR_CARA_HORMIGON_MM`.
        diametro_nominal_mm: diámetro de la ``RebarBarType`` (mm); si es None, solo se usa el cover.

    Returns:
        ``Line`` nueva o la original si falla.
    """
    if line is None or not isinstance(line, Line):
        return line
    if cover_hormigon_mm is None:
        cover_hormigon_mm = RECUBRIMIENTO_INFERIOR_CARA_HORMIGON_MM
    try:
        d_half = max(0.0, float(diametro_nominal_mm) * 0.5)
    except Exception:
        d_half = 0.0
    try:
        total_mm = float(cover_hormigon_mm) + d_half
    except Exception:
        total_mm = float(RECUBRIMIENTO_INFERIOR_CARA_HORMIGON_MM) + d_half
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        dist_ft = UnitUtils.ConvertToInternalUnits(total_mm, UnitTypeId.Millimeters)
    except Exception:
        dist_ft = total_mm / 304.8
    try:
        if n_outward is not None and float(n_outward.GetLength()) > 1e-12:
            inward = n_outward.Normalize().Negate()
        else:
            inward = XYZ.BasisZ
        v = inward.Multiply(dist_ft)
        p0 = line.GetEndPoint(0).Add(v)
        p1 = line.GetEndPoint(1).Add(v)
        return Line.CreateBound(p0, p1)
    except Exception:
        return line


def offset_linea_adicional_hacia_interior_mm(line, n_outward, distancia_mm):
    """
    Traslada la ``Line`` una distancia adicional hacia el **interior** del hormigón
    (mismo sentido que :func:`offset_linea_eje_barra_desde_cara_inferior_mm`: ``-n`` saliente).

    Se usa para situar la segunda dirección de armadura (p. ej. lado largo) **sobre** la
    primera (lado corto), separando ejes una distancia típica igual al **diámetro nominal** mm.
    """
    if line is None or not isinstance(line, Line):
        return line
    try:
        d_mm = float(distancia_mm)
    except Exception:
        d_mm = 0.0
    if d_mm <= 1e-9:
        return line
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        dist_ft = UnitUtils.ConvertToInternalUnits(d_mm, UnitTypeId.Millimeters)
    except Exception:
        dist_ft = d_mm / 304.8
    try:
        if n_outward is not None and float(n_outward.GetLength()) > 1e-12:
            inward = n_outward.Normalize().Negate()
        else:
            inward = XYZ.BasisZ
        v = inward.Multiply(dist_ft)
        p0 = line.GetEndPoint(0).Add(v)
        p1 = line.GetEndPoint(1).Add(v)
        return Line.CreateBound(p0, p1)
    except Exception:
        return line


def _largo_ancho_planta_mm_desde_cara_inferior(elemento):
    """
    Largo y ancho en planta (mm) desde la **geometría real** de la cara inferior: longitudes
    de arista del lado más corto y del más largo del contorno exterior (véase
    :func:`extraer_curva_lado_menor_cara_inferior` / :func:`extraer_curva_lado_mayor_cara_inferior`).
    No usa ``BoundingBox`` alineado a ejes del proyecto.
    """
    if elemento is None:
        return None, None
    try:
        rm = extraer_curva_lado_menor_cara_inferior(elemento)
        rM = extraer_curva_lado_mayor_cara_inferior(elemento)
    except Exception:
        return None, None
    if rm is None or rM is None:
        return None, None
    try:
        cm, _ = rm
        cM, _ = rM
        len_m = float(cm.Length) * 304.8
        len_M = float(cM.Length) * 304.8
    except Exception:
        return None, None
    if len_m > len_M:
        len_m, len_M = len_M, len_m
    ancho_mm = len_m
    largo_mm = len_M
    return largo_mm, ancho_mm


def longitud_distribucion_perpendicular_barra_inferior_ft(
    elemento,
    linea_barra,
    rec_planta_mm=None,
    lado_malla=None,
):
    """
    Longitud en **pies** en la dirección perpendicular a la barra en planta (distribución del
    set de armadura) para ``RebarShapeDrivenAccessor.SetLayoutAsMaximumSpacing`` (parámetro
    *array length*).

    Para ``lado_malla`` en ``("menor","mayor")`` el método primario usa las **longitudes reales
    de aristas** de la cara inferior (:func:`_largo_ancho_planta_mm_desde_cara_inferior`).
    Este método es invariante a la rotación en planta: las longitudes de arista son propiedades
    geométricas intrínsecas, independientes de los ejes del proyecto.

    - ``"menor"``: barras van en la dirección corta → se distribuyen a lo largo del eje largo
      → ``array_length = largo_mm − 2×rec_planta``.
    - ``"mayor"``: barras van en la dirección larga → se distribuyen a lo largo del eje corto
      → ``array_length = ancho_mm − 2×rec_planta``.

    Si las aristas fallan (geometría no rectangular), se cae a la proyección del perímetro real
    (:func:`luz_proyeccion_perimetro_inferior_ft`).  La proyección del perímetro se usa como
    método primario para ``lado_malla`` distinto de ``"menor"``/``"mayor"``.

    Nunca usa ``BoundingBox`` global (que sobredimensiona elementos rotados).
    """
    if elemento is None or linea_barra is None:
        return 1.0
    try:
        rec_use = float(rec_planta_mm) if rec_planta_mm is not None else 0.0
    except Exception:
        rec_use = 0.0

    if lado_malla in (u"menor", u"mayor"):
        largo_mm, ancho_mm = _largo_ancho_planta_mm_desde_cara_inferior(elemento)
        if largo_mm is not None and ancho_mm is not None:
            try:
                from Autodesk.Revit.DB import UnitUtils, UnitTypeId

                base_mm = float(largo_mm) if lado_malla == u"menor" else float(ancho_mm)
                span_mm = max(base_mm - 2.0 * rec_use, 0.01)
                span_ft = UnitUtils.ConvertToInternalUnits(span_mm, UnitTypeId.Millimeters)
                return max(float(span_ft), 1e-6)
            except Exception:
                pass

    span_ft = luz_proyeccion_perimetro_inferior_ft(
        elemento,
        linea_barra,
        rec_use,
        perpendicular_a_tangente=True,
    )
    if span_ft is not None:
        return max(float(span_ft), 1e-6)

    span_ft = luz_proyeccion_perimetro_inferior_ft(
        elemento,
        linea_barra,
        rec_use,
        perpendicular_a_tangente=True,
        tolerancias=(
            (0.05, 0.18),
            (0.12, 0.30),
            (0.25, 0.45),
            (0.50, 0.70),
            (1.0, 1.0),
            (2.0, 2.0),
        ),
    )
    if span_ft is not None:
        return max(float(span_ft), 1e-6)

    return 1.0


def _normalizar_xyz_seguro(v):
    try:
        if v is None:
            return None
        if float(v.GetLength()) < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _corregir_marco_cara_horizontal_saliente(bx, by, bz, es_tapa_superior):
    """
    Fuerza ``bz`` = normal **saliente** del hormigón en caras casi horizontales:

    - Tapa **superior** de la fundación: saliente con componente Z **positiva** (hacia arriba).
    - Tapa **inferior**: saliente con Z **negativa** (hacia abajo).

    ``ComputeDerivatives`` / ``FaceNormal`` a veces devuelven la normal hacia el interior; entonces
    :func:`offset_linea_eje_barra_desde_cara_inferior_mm` desplaza al revés y el recubrimiento
    y la propagación del conjunto no coinciden entre mallas.

    Si se invierte ``bz``, se invierte también ``by`` para mantener marco derecho (U×V ≈ N).
    """
    if bz is None:
        return bx, by, bz
    try:
        nz = float(bz.Z)
        if abs(nz) < 0.55:
            return bx, by, bz
        if es_tapa_superior:
            if nz < 0.0:
                bz = XYZ(-bz.X, -bz.Y, -bz.Z)
                if by is not None:
                    by = XYZ(-by.X, -by.Y, -by.Z)
        else:
            if nz > 0.0:
                bz = XYZ(-bz.X, -bz.Y, -bz.Z)
                if by is not None:
                    by = XYZ(-by.X, -by.Y, -by.Z)
    except Exception:
        pass
    return bx, by, bz


def obtener_marco_coordenadas_cara_inferior(
    elem, tol=0.05, tol_normal_z=0.18, ultra_fallback=False
):
    """
    Marco local de la **cara inferior** en coordenadas de documento.

    Origen en el centro del dominio UV de la cara. Ejes: tangentes paramétricas **U**, **V**
    y **normal N** (saliente respecto del sólido), según ``Face.ComputeDerivatives``;
    si falla, respaldo con ``PlanarFace`` (XVector, YVector, FaceNormal).

    Args:
        ultra_fallback: ver :func:`_inferior_planar_face_info`.

    Returns:
        tuple | None: ``(origin, u_axis, v_axis, n_axis)`` con vectores unitarios.
    """
    r = _inferior_planar_face_info(elem, tol, tol_normal_z, ultra_fallback)
    if r is None:
        return None
    face, t, _z_inf = r
    o = None
    bx = by = bz = None
    try:
        bbox = face.GetBoundingBox()
        if bbox is not None:
            uu = 0.5 * (bbox.Min.U + bbox.Max.U)
            vv = 0.5 * (bbox.Min.V + bbox.Max.V)
            uv = UV(uu, vv)
        else:
            uv = UV(0.0, 0.0)
        deriv = face.ComputeDerivatives(uv)
        if deriv is not None:
            o = deriv.Origin
            bx = _normalizar_xyz_seguro(deriv.BasisX)
            by = _normalizar_xyz_seguro(deriv.BasisY)
            bz = _normalizar_xyz_seguro(deriv.BasisZ)
    except Exception:
        pass
    if bx is None or by is None or o is None:
        try:
            from Autodesk.Revit.DB import PlanarFace

            if isinstance(face, PlanarFace):
                o = face.Origin
                bx = _normalizar_xyz_seguro(face.XVector)
                by = _normalizar_xyz_seguro(face.YVector)
                bz = _normalizar_xyz_seguro(face.FaceNormal)
        except Exception:
            pass
    if o is None or bx is None or by is None:
        return None
    if bz is None:
        bz = _normalizar_xyz_seguro(bx.CrossProduct(by))
    if bz is None:
        return None
    if t is not None:
        try:
            o = t.OfPoint(o)
            bx = t.OfVector(bx)
            by = t.OfVector(by)
            bz = t.OfVector(bz)
        except Exception:
            return None
    bx, by, bz = _corregir_marco_cara_horizontal_saliente(bx, by, bz, False)
    return (o, bx, by, bz)


def obtener_marco_coordenadas_cara_superior(elem, tol=0.05, tol_normal_z=0.18):
    """
    Marco local de la **cara superior** en coordenadas de documento.

    Misma convención que :func:`obtener_marco_coordenadas_cara_inferior` (U, V, N saliente).

    Returns:
        tuple | None: ``(origin, u_axis, v_axis, n_axis)`` con vectores unitarios.
    """
    r = _superior_planar_face_info(elem, tol, tol_normal_z)
    if r is None:
        return None
    face, t, _z_sup = r
    o = None
    bx = by = bz = None
    try:
        bbox = face.GetBoundingBox()
        if bbox is not None:
            uu = 0.5 * (bbox.Min.U + bbox.Max.U)
            vv = 0.5 * (bbox.Min.V + bbox.Max.V)
            uv = UV(uu, vv)
        else:
            uv = UV(0.0, 0.0)
        deriv = face.ComputeDerivatives(uv)
        if deriv is not None:
            o = deriv.Origin
            bx = _normalizar_xyz_seguro(deriv.BasisX)
            by = _normalizar_xyz_seguro(deriv.BasisY)
            bz = _normalizar_xyz_seguro(deriv.BasisZ)
    except Exception:
        pass
    if bx is None or by is None or o is None:
        try:
            from Autodesk.Revit.DB import PlanarFace

            if isinstance(face, PlanarFace):
                o = face.Origin
                bx = _normalizar_xyz_seguro(face.XVector)
                by = _normalizar_xyz_seguro(face.YVector)
                bz = _normalizar_xyz_seguro(face.FaceNormal)
        except Exception:
            pass
    if o is None or bx is None or by is None:
        return None
    if bz is None:
        bz = _normalizar_xyz_seguro(bx.CrossProduct(by))
    if bz is None:
        return None
    if t is not None:
        try:
            o = t.OfPoint(o)
            bx = t.OfVector(bx)
            by = t.OfVector(by)
            bz = t.OfVector(bz)
        except Exception:
            return None
    bx, by, bz = _corregir_marco_cara_horizontal_saliente(bx, by, bz, True)
    return (o, bx, by, bz)


def crear_lineas_detalle_evaluacion_recubrimiento(
    document,
    view,
    pares_borde_offset,
    marcos_cara=None,
    longitud_marcador_mm=None,
    use_transaction=True,
):
    """
    Crea ``DetailCurve`` en la vista para revisar posición y recubrimiento en planta.

    Por cada fundación se dibuja:
      - la curva en el **borde** (lado corto inferior);
      - la curva tras **offset en planta + recubrimiento en extremos** (o solo offset si
        no hubo margen para recortar extremos), si no es None;
      - opcionalmente, un **marcador de ejes** U, V, N (tangentes y normal de la cara)
        desde el origen del marco local.

    Requiere vista que admita detalle (planta, alzado, sección, lámina). En vista 3D
    ``NewDetailCurve`` suele fallar.

    Args:
        document: ``Document``.
        view: ``View`` activa.
        pares_borde_offset: lista de ``(curve_borde, curve_offset | None)``.
        marcos_cara: lista de ``(origin, u, v, n)`` de
            :func:`obtener_marco_coordenadas_cara_inferior`, en el mismo orden que las
            fundaciones evaluadas (o ``None``).
        longitud_marcador_mm: longitud de cada eje del marcador; por defecto
            :data:`MARCADOR_EJES_CARA_MM`.
        use_transaction: si es False, no abre ``Transaction`` (el llamador debe tener
            una transacción activa).

    Returns:
        tuple: (n_creadas, err_texto) con ``err_texto`` None si no hubo fallo global;
        si algunas curvas fallan, se hace commit de las válidas y ``err_texto`` resume avisos.
    """
    from Autodesk.Revit.DB import Transaction

    if document is None or view is None:
        return 0, u"No hay documento o vista activa."
    if not pares_borde_offset:
        return 0, None
    if longitud_marcador_mm is None:
        longitud_marcador_mm = MARCADOR_EJES_CARA_MM
    creadas = [0]
    avisos = []

    def _dibujar():
        plano_v = _obtener_plano_detalle_vista(view)

        def _curva_detail(curva):
            """Curva coplanar con la vista (p. ej. planta: ajuste Z al plano de corte)."""
            if curva is None:
                return None
            if plano_v is not None and isinstance(curva, Line):
                cp = _proyectar_linea_al_plano_vista(curva, plano_v)
                if cp is not None:
                    return cp
                # No devolver la Line 3D original: NewDetailCurve exige curva en el plano de la vista.
                return None
            return curva

        for borde, off in pares_borde_offset:
            if borde is not None:
                cd_b = _curva_detail(borde)
                if cd_b is not None:
                    try:
                        document.Create.NewDetailCurve(view, cd_b)
                        creadas[0] += 1
                    except Exception as ex:
                        avisos.append(unicode(ex))
            if off is not None:
                cd_o = _curva_detail(off)
                if cd_o is not None:
                    try:
                        document.Create.NewDetailCurve(view, cd_o)
                        creadas[0] += 1
                    except Exception as ex:
                        avisos.append(unicode(ex))
        if marcos_cara:
            try:
                from Autodesk.Revit.DB import UnitUtils, UnitTypeId

                len_ft = UnitUtils.ConvertToInternalUnits(
                    float(longitud_marcador_mm), UnitTypeId.Millimeters
                )
            except Exception:
                len_ft = float(longitud_marcador_mm) / 304.8
            for marco in marcos_cara:
                if marco is None or len(marco) < 4:
                    continue
                origin, u_ax, v_ax, n_ax = marco[0], marco[1], marco[2], marco[3]
                for axis in (u_ax, v_ax, n_ax):
                    if axis is None:
                        continue
                    try:
                        p2 = origin.Add(axis.Multiply(len_ft))
                        ln = Line.CreateBound(origin, p2)
                        if ln is None:
                            continue
                        c_d = _curva_detail(ln)
                        if c_d is None:
                            continue
                        document.Create.NewDetailCurve(view, c_d)
                        creadas[0] += 1
                    except Exception as ex:
                        avisos.append(unicode(ex))

    if use_transaction:
        t = Transaction(document, u"BIMTools: detalle evaluación recubrimiento fundación")
        try:
            t.Start()
        except Exception as ex:
            return 0, unicode(ex)
        try:
            _dibujar()
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            return 0, unicode(ex)
    else:
        try:
            _dibujar()
        except Exception as ex:
            return 0, unicode(ex)
    if not avisos:
        return creadas[0], None
    txt = u"; ".join(avisos[:4])
    if len(avisos) > 4:
        txt += u"…"
    return creadas[0], txt


def crear_marcador_vector_norm_rebar(
    document,
    view,
    linea_barra,
    norm_vector,
    longitud_mm=None,
    use_transaction=True,
):
    """
    Dibuja una ``DetailLine`` desde el punto medio de ``linea_barra`` en la dirección del
    vector ``norm`` pasado a ``Rebar.CreateFromCurves*`` (mismo sentido que en la API).

    Returns:
        ``(n_creadas, err_texto)`` con ``n_creadas`` 0 o 1.
    """
    from Autodesk.Revit.DB import Transaction, UnitUtils, UnitTypeId

    if document is None or view is None or linea_barra is None or norm_vector is None:
        return 0, None
    if not isinstance(linea_barra, Line):
        return 0, None
    try:
        n = norm_vector.Normalize()
    except Exception:
        return 0, None
    if n is None or n.GetLength() < 1e-12:
        return 0, None
    if longitud_mm is None:
        longitud_mm = MARCADOR_NORM_REBAR_MM
    try:
        len_ft = UnitUtils.ConvertToInternalUnits(float(longitud_mm), UnitTypeId.Millimeters)
    except Exception:
        len_ft = float(longitud_mm) / 304.8

    try:
        mid = linea_barra.Evaluate(0.5, True)
    except Exception:
        try:
            p0 = linea_barra.GetEndPoint(0)
            p1 = linea_barra.GetEndPoint(1)
            mid = XYZ(
                (p0.X + p1.X) * 0.5,
                (p0.Y + p1.Y) * 0.5,
                (p0.Z + p1.Z) * 0.5,
            )
        except Exception:
            return 0, None
    try:
        p_end = mid.Add(n.Multiply(len_ft))
        ln = Line.CreateBound(mid, p_end)
    except Exception:
        return 0, None

    plano_v = _obtener_plano_detalle_vista(view)
    if plano_v is not None and isinstance(ln, Line):
        cp = _proyectar_linea_al_plano_vista(ln, plano_v)
        if cp is None:
            return 0, None
        ln_draw = cp
    else:
        ln_draw = ln

    def _commit():
        document.Create.NewDetailCurve(view, ln_draw)

    if use_transaction:
        t = Transaction(document, u"BIMTools: marcador vector norm Rebar (CreateFromCurves)")
        try:
            t.Start()
        except Exception as ex:
            return 0, unicode(ex)
        try:
            _commit()
            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            return 0, unicode(ex)
        return 1, None
    try:
        _commit()
    except Exception as ex:
        return 0, unicode(ex)
    return 1, None


def _make_geometry_options():
    opts = Options()
    try:
        opts.ComputeReferences = False
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ViewDetailLevel

        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    return opts


def _collect_face_entries(elem):
    """Lista de (face, transform, fz_min, fz_max, normal_es) en coords. documento (puntos)."""
    # Caché por elemento: la geometría no cambia dentro de una misma llamada Execute.
    try:
        eid_key = int(elem.Id.IntegerValue)
    except Exception:
        eid_key = None
    if eid_key is not None and eid_key in _FACE_ENTRIES_CACHE:
        return _FACE_ENTRIES_CACHE[eid_key]

    z_min_g, z_max_g = None, None
    entries = []
    try:
        opts = _make_geometry_options()
        geom = elem.get_Geometry(opts)
        if not geom:
            if eid_key is not None:
                _FACE_ENTRIES_CACHE[eid_key] = (None, None, None)
            return None, None, None
        for go in geom:
            try:
                if isinstance(go, GeometryInstance):
                    t = go.Transform
                    geoms_to_check = []
                    try:
                        sym_geom = go.GetSymbolGeometry()
                        if sym_geom:
                            try:
                                geoms_to_check = [x for x in sym_geom]
                            except Exception:
                                geoms_to_check = []
                    except Exception:
                        pass
                    if not geoms_to_check:
                        geoms_to_check = list(go.GetInstanceGeometry())
                        t = None
                else:
                    t = None
                    geoms_to_check = [go] if hasattr(go, "Faces") else []
                for g in geoms_to_check:
                    if not hasattr(g, "Faces"):
                        continue
                    for face in g.Faces:
                        face_pts = []
                        for edge_loop in list(face.EdgeLoops):
                            for edge in edge_loop:
                                for pt in edge.Tessellate():
                                    face_pts.append(t.OfPoint(pt) if t else pt)
                        if not face_pts:
                            continue
                        fz_min = min(p.Z for p in face_pts)
                        fz_max = max(p.Z for p in face_pts)
                        normal_es = None
                        try:
                            bbox = face.GetBoundingBox()
                            uv_mid = UV(
                                (bbox.Min.U + bbox.Max.U) / 2.0,
                                (bbox.Min.V + bbox.Max.V) / 2.0,
                            )
                            normal_es = face.ComputeNormal(uv_mid)
                            if t and hasattr(t, "OfVector"):
                                normal_es = t.OfVector(normal_es)
                        except Exception:
                            try:
                                uv_mid = UV(0.5, 0.5)
                                normal_es = face.ComputeNormal(uv_mid)
                                if t and hasattr(t, "OfVector"):
                                    normal_es = t.OfVector(normal_es)
                            except Exception:
                                pass
                        entries.append((face, t, fz_min, fz_max, normal_es))
                        if z_min_g is None or fz_min < z_min_g:
                            z_min_g = fz_min
                        if z_max_g is None or fz_max > z_max_g:
                            z_max_g = fz_max
            except Exception:
                continue
        if not entries or z_min_g is None or z_max_g is None:
            result = (None, None, None)
        else:
            result = (entries, z_min_g, z_max_g)
        if eid_key is not None:
            _FACE_ENTRIES_CACHE[eid_key] = result
        return result
    except Exception:
        return None, None, None


def _select_inferior_face(entries, z_min, tol, tol_normal_z):
    """Criterio estricto: normal ~Z y fz_min en contacto con z_min global."""
    face_sel, t_sel = None, None
    tol_z = max(float(tol), 1e-4)
    for face, t, fz_min, fz_max, normal_es in entries:
        if normal_es is not None:
            es_parallel_z = (
                abs(normal_es.X) < tol_normal_z and abs(normal_es.Y) < tol_normal_z
            )
        else:
            es_parallel_z = (fz_max - fz_min) < max(float(tol), 0.12)
        if not es_parallel_z:
            continue
        if abs(fz_min - z_min) < tol_z:
            face_sel, t_sel = face, t
    return (face_sel, t_sel)


def _select_inferior_face_relaxed(entries, z_min, tol):
    """
    Si falla la normal: cara casi horizontal por espesor pequeño en Z y fz_min ~ z_min.
    Cubre familias donde ComputeNormal o teselación no dan plano perfecto.
    """
    tol_z = max(float(tol), 0.1)
    thick_max = max(float(tol) * 12.0, 0.35)
    best = None
    best_fzmin = None
    for face, t, fz_min, fz_max, _n in entries:
        thick = fz_max - fz_min
        if thick > thick_max:
            continue
        if abs(fz_min - z_min) > tol_z:
            continue
        if best_fzmin is None or fz_min < best_fzmin - 1e-6:
            best_fzmin = fz_min
            best = (face, t)
    return best


def _select_superior_face(entries, z_max, tol, tol_normal_z):
    """Cara horizontal cuyo ``fz_max`` coincide con el ``z_max`` global (cara superior)."""
    face_sel, t_sel = None, None
    tol_z = max(float(tol), 1e-4)
    for face, t, fz_min, fz_max, normal_es in entries:
        if normal_es is not None:
            es_parallel_z = (
                abs(normal_es.X) < tol_normal_z and abs(normal_es.Y) < tol_normal_z
            )
        else:
            es_parallel_z = (fz_max - fz_min) < max(float(tol), 0.12)
        if not es_parallel_z:
            continue
        if abs(fz_max - z_max) < tol_z:
            face_sel, t_sel = face, t
    return (face_sel, t_sel)


def _select_superior_face_relaxed(entries, z_max, tol):
    """Respaldo: cara casi horizontal con ``fz_max`` ~ ``z_max`` global."""
    tol_z = max(float(tol), 0.1)
    thick_max = max(float(tol) * 12.0, 0.35)
    best = None
    best_fzmax = None
    for face, t, fz_min, fz_max, _n in entries:
        thick = fz_max - fz_min
        if thick > thick_max:
            continue
        if abs(fz_max - z_max) > tol_z:
            continue
        if best_fzmax is None or fz_max > best_fzmax + 1e-6:
            best_fzmax = fz_max
            best = (face, t)
    return best


def _select_inferior_face_ultra_loose(entries, z_min, z_max):
    """
    Último recurso antes de renunciar a la cara inferior: fondos teselados,
    ``WallFoundation`` con soleira inclinada en planta o símbolos donde
    ``fz_min`` no coincide con ``z_min`` global o el grosor Z local supera el
    umbral de :func:`_select_inferior_face_relaxed`.

    Prioriza la cara más baja con normal suficientemente vertical (|n·Z|).
    """
    if not entries or z_min is None or z_max is None:
        return None
    h_elem = float(z_max) - float(z_min)
    if h_elem < 1e-9:
        return None
    band = max(0.15, min(1.35, 0.08 + 0.28 * h_elem))
    thick_cap = max(4.0, min(14.0, 0.88 * h_elem + 2.25))
    best = None
    best_key = None
    z0 = float(z_min)
    for face, t, fz_min, fz_max, normal_es in entries:
        fz_mi = float(fz_min)
        fz_ma = float(fz_max)
        thick = fz_ma - fz_mi
        if thick > thick_cap:
            continue
        if fz_mi > z0 + band:
            continue
        if normal_es is not None:
            nz = abs(float(normal_es.Z))
            if nz < 0.42:
                continue
        else:
            nz = 0.66
        key = (fz_mi, -nz, thick)
        if best_key is None or key < best_key:
            best_key = key
            best = (face, t)
    return best


def _superior_planar_face_info(elem, tol=0.05, tol_normal_z=0.18):
    """Cara horizontal en ``z_max`` (tapa superior de la fundación)."""
    entries, z_min, z_max = _collect_face_entries(elem)
    if not entries or z_max is None:
        return None
    face_sel, t_sel = _select_superior_face(entries, z_max, tol, tol_normal_z)
    if face_sel is None:
        relaxed = _select_superior_face_relaxed(entries, z_max, tol)
        if relaxed is None:
            return None
        face_sel, t_sel = relaxed
    return (face_sel, t_sel, z_max)


def _inferior_planar_face_info(elem, tol=0.05, tol_normal_z=0.18, ultra_fallback=False):
    """
    tol_normal_z algo mayor (0.18) para normales no perfectamente verticales en símbolo.

    Args:
        ultra_fallback: si es True y fallan criterios estricto y relajado, intenta
            :func:`_select_inferior_face_ultra_loose` (p. ej. ``WallFoundation``
            con geometría irregular). Por defecto False para no cambiar fundación
            aislada u otros llamadores.
    """
    entries, z_min, z_max = _collect_face_entries(elem)
    if not entries or z_min is None:
        return None
    face_sel, t_sel = _select_inferior_face(entries, z_min, tol, tol_normal_z)
    if face_sel is None:
        relaxed = _select_inferior_face_relaxed(entries, z_min, tol)
        if relaxed is not None:
            face_sel, t_sel = relaxed
        elif ultra_fallback and z_max is not None:
            ultra = _select_inferior_face_ultra_loose(entries, z_min, z_max)
            if ultra is None:
                return None
            face_sel, t_sel = ultra
        else:
            return None
    return (face_sel, t_sel, z_min)


def _tangente_unitaria_curva(curve):
    """Tangente unitaria en el punto medio (``Line`` o ``Curve`` genérica)."""
    if curve is None:
        return None
    try:
        if isinstance(curve, Line):
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            d = p1 - p0
            if d.GetLength() < 1e-12:
                return None
            return d.Normalize()
    except Exception:
        pass
    try:
        d = curve.ComputeDerivatives(0.5, True)
        if d is None:
            return None
        tx = d.BasisX
        if tx.GetLength() < 1e-12:
            return None
        return tx.Normalize()
    except Exception:
        return None


def _punto_medio_curva(curve):
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


def _origen_documento_cara(face, transform):
    """Punto de referencia en la cara, coordenadas de documento."""
    try:
        bbox = face.GetBoundingBox()
        uu = 0.5 * (bbox.Min.U + bbox.Max.U)
        vv = 0.5 * (bbox.Min.V + bbox.Max.V)
        uv = UV(uu, vv)
        xyz = face.Evaluate(uv)
        return transform.OfPoint(xyz) if transform else xyz
    except Exception:
        try:
            from Autodesk.Revit.DB import PlanarFace

            if isinstance(face, PlanarFace):
                o = face.Origin
                return transform.OfPoint(o) if transform else o
        except Exception:
            pass
    return None


def _normal_documento_desde_cache(face, normal_es):
    """``normal_es`` viene de :func:`_collect_face_entries` (ya en coords. de documento)."""
    if normal_es is None:
        return None
    try:
        n = normal_es.Normalize()
        if n.GetLength() < 1e-12:
            return None
        return n
    except Exception:
        return None


def evaluar_caras_paralelas_curva_mas_cercana(
    elem,
    curve,
    tol_paralelismo_dot=0.08,
    omitir_coplanar_con_curva=True,
    tol_coplanar_ft=None,
):
    """
    Recorre las caras del sólido del elemento y selecciona, entre las que son **paralelas**
    a la tangente de ``curve`` (el plano de la cara es paralelo a la recta: ``|N·T| <= tol``),
    la cara cuyo plano está a **menor distancia perpendicular** del punto medio de la curva.

    Por defecto se excluyen las caras **coplanares** con la curva (distancia punto medio → plano
    casi cero), para no devolver la cara inferior sobre la que se apoya la curva extraída.

    Args:
        elem: ``Element`` con geometría sólida.
        curve: ``Curve`` (p. ej. la de :func:`extraer_curva_lado_menor_cara_inferior`).
        tol_paralelismo_dot: umbral en ``|N·T|`` (T y N unitarios); cercano a 0 = bien paralelo.
        omitir_coplanar_con_curva: si True, no considerar caras donde la curva yace en el plano.
        tol_coplanar_ft: umbral de distancia al plano para considerar «coplanar»;
            por defecto ~2 mm.

    Returns:
        dict | None: ``mejor`` — ``(face, transform)`` o None; ``distancia_ft``; ``tangente``;
        ``candidatos`` — lista de dicts con ``face``, ``transform``, ``distancia_ft``, ``dot_nt``,
        ``origen_plano``; ``descartados_coplanar`` — cantidad omitida por coplanaridad.
        None si no hay tangente o geometría.
    """
    if elem is None or curve is None:
        return None
    t_curve = _tangente_unitaria_curva(curve)
    mid = _punto_medio_curva(curve)
    if t_curve is None or mid is None:
        return None
    if tol_coplanar_ft is None:
        tol_coplanar_ft = _TOL_COPLANAR_CURVA_FT

    entries, _zmin, _zmax = _collect_face_entries(elem)
    if not entries:
        return None

    tol_dot = max(float(tol_paralelismo_dot), 1e-5)
    candidatos = []
    desc_coplanar = 0

    for face, t, fz_min, fz_max, normal_es in entries:
        n_doc = _normal_documento_desde_cache(face, normal_es)
        if n_doc is None:
            continue
        try:
            dot_nt = abs(float(n_doc.DotProduct(t_curve)))
        except Exception:
            continue
        if dot_nt > tol_dot:
            continue
        o_doc = _origen_documento_cara(face, t)
        if o_doc is None:
            continue
        try:
            dist = abs(float((mid - o_doc).DotProduct(n_doc)))
        except Exception:
            continue
        if omitir_coplanar_con_curva and dist <= tol_coplanar_ft:
            desc_coplanar += 1
            continue
        candidatos.append(
            {
                "face": face,
                "transform": t,
                "distancia_ft": dist,
                "dot_nt": dot_nt,
                "origen_plano": o_doc,
            }
        )

    if not candidatos:
        return {
            "mejor": None,
            "distancia_ft": None,
            "tangente": t_curve,
            "candidatos": [],
            "descartados_coplanar": desc_coplanar,
        }

    candidatos.sort(key=lambda x: x["distancia_ft"])
    best = candidatos[0]
    return {
        "mejor": (best["face"], best["transform"]),
        "distancia_ft": best["distancia_ft"],
        "tangente": t_curve,
        "candidatos": candidatos,
        "descartados_coplanar": desc_coplanar,
    }


def _normal_saliente_face_documento(face, transform):
    """Normal unitaria saliente de la cara en coordenadas de documento."""
    if face is None:
        return None
    n_out = None
    try:
        from Autodesk.Revit.DB import PlanarFace

        if isinstance(face, PlanarFace):
            n_out = face.FaceNormal
            if transform is not None and hasattr(transform, "OfVector"):
                n_out = transform.OfVector(n_out)
    except Exception:
        pass
    if n_out is None or n_out.GetLength() < 1e-12:
        try:
            bbox = face.GetBoundingBox()
            if bbox is not None:
                uv = UV(
                    0.5 * (bbox.Min.U + bbox.Max.U),
                    0.5 * (bbox.Min.V + bbox.Max.V),
                )
                d = face.ComputeDerivatives(uv)
                if d is not None:
                    n_out = d.BasisZ
                    if transform is not None and hasattr(transform, "OfVector"):
                        n_out = transform.OfVector(n_out)
        except Exception:
            pass
    if n_out is None or n_out.GetLength() < 1e-12:
        return None
    try:
        return n_out.Normalize()
    except Exception:
        return None


def vector_reverso_cara_paralela_mas_cercana_a_barra(
    elem,
    linea_barra,
    excluir_caras_tapas_horizontales=False,
    tol_abs_nz_tapa=0.82,
):
    """
    Vector unitario **reverso** de la normal saliente de la **cara paralela más cercana**
    a la barra (misma selección que :func:`evaluar_caras_paralelas_curva_mas_cercana`).

    Sirve como ``eje_referencia_z_ganchos`` en ``CreateFromCurves*``: el plano paralelo
    excluye la cara coplanar con la curva; la cara elegida es la de menor distancia
    perpendicular al eje de la barra entre las restantes.

    Si ``excluir_caras_tapas_horizontales`` es True, no se consideran caras cuya normal
    sea casi vertical (``|n.Z| >= tol_abs_nz_tapa``). Así, en **barras horizontales en
    paramentos verticales** no gana la tapa inferior/superior (muy cercana en distancia
    perpendicular pero con normal ±Z), que volcaba el eje de ganchos hacia arriba/abajo
    en lugar de hacia el interior en planta.
    """
    if elem is None or linea_barra is None:
        return None
    ev = evaluar_caras_paralelas_curva_mas_cercana(elem, linea_barra)
    if ev is None:
        return None
    candidatos = list(ev.get("candidatos") or [])
    if excluir_caras_tapas_horizontales and candidatos:
        verticales = []
        for c in candidatos:
            n = _normal_saliente_face_documento(c.get("face"), c.get("transform"))
            if n is None:
                continue
            try:
                if abs(float(n.Z)) < float(tol_abs_nz_tapa):
                    verticales.append(c)
            except Exception:
                continue
        candidatos = verticales
    if not candidatos:
        return None
    candidatos.sort(key=lambda x: x["distancia_ft"])
    best = candidatos[0]
    face = best.get("face")
    transform = best.get("transform")
    n_out = _normal_saliente_face_documento(face, transform)
    if n_out is None:
        return None
    try:
        rev = n_out.Negate()
        if rev.GetLength() < 1e-12:
            return None
        return rev.Normalize()
    except Exception:
        return None


def _proyecta_vector_perpendicular_a_tangente_barra(t_unit, vec):
    """Proyección de ``vec`` sobre el plano perpendicular a la tangente unitaria ``t_unit``."""
    if t_unit is None or vec is None:
        return None
    try:
        a = vec.Normalize()
        w = a - t_unit.Multiply(a.DotProduct(t_unit))
        if w.GetLength() < 1e-10:
            return None
        return w.Normalize()
    except Exception:
        return None


def _orientar_vector_hacia_centro_elemento(vec_unit, punto_ref, elemento):
    """Invierte ``vec_unit`` si apunta alejándose del centro del ``BoundingBox`` del elemento."""
    if vec_unit is None or elemento is None or punto_ref is None:
        return vec_unit
    try:
        bb = elemento.get_BoundingBox(None)
        if bb is None:
            return vec_unit
        c = XYZ(
            (bb.Min.X + bb.Max.X) * 0.5,
            (bb.Min.Y + bb.Max.Y) * 0.5,
            (bb.Min.Z + bb.Max.Z) * 0.5,
        )
        v = c - punto_ref
        if v.GetLength() < 1e-12:
            return vec_unit
        v = v.Normalize()
        if float(vec_unit.DotProduct(v)) < 0.0:
            return vec_unit.Negate()
        return vec_unit
    except Exception:
        return vec_unit


def obtener_norm_plano_barra_desde_basisz_cara_paralela(line, face, transform, elemento=None):
    """
    Usa la cara paralela más cercana a la curva: ``BasisZ`` de ``Face.ComputeDerivatives``
    (eje local Z de la parametrización de la cara, coherente con la normal del plano en
    muchas caras planas), transformado a coordenadas de documento y proyectado al plano
    perpendicular a la tangente de la barra — vector adecuado para ``norm`` en
    ``CreateFromCurves*`` y para el sentido de propagación del conjunto.

    Se orienta hacia el **interior** del ``elemento`` (hacia el centro del ``BoundingBox``)
    cuando se proporciona el host.
    """
    if line is None or face is None:
        return None
    t = _tangente_unitaria_curva(line)
    if t is None:
        return None
    basis_z = None
    try:
        bbox = face.GetBoundingBox()
        if bbox is None:
            return None
        uu = 0.5 * (bbox.Min.U + bbox.Max.U)
        vv = 0.5 * (bbox.Min.V + bbox.Max.V)
        uv = UV(uu, vv)
        d = face.ComputeDerivatives(uv)
        if d is None:
            return None
        basis_z = d.BasisZ
        if transform is not None and hasattr(transform, "OfVector"):
            basis_z = transform.OfVector(basis_z)
        if basis_z.GetLength() < 1e-12:
            return None
        basis_z = basis_z.Normalize()
    except Exception:
        return None
    w = _proyecta_vector_perpendicular_a_tangente_barra(t, basis_z)
    if w is None:
        return None
    mid = _punto_medio_curva(line)
    w = _orientar_vector_hacia_centro_elemento(w, mid, elemento)
    return w


def curva_lado_menor_desde_loops(curve_loops, tol_rel=1e-5):
    """
    Contorno exterior = loop de mayor perímetro; curva = arista de longitud mínima.
    """
    if not curve_loops:
        return None
    try:
        outer = max(curve_loops, key=lambda cl: cl.GetExactLength())
    except Exception:
        outer = curve_loops[0]
    curves = []
    try:
        it = outer.GetCurveIterator()
        while it.MoveNext():
            c = it.Current
            if c is not None:
                curves.append(c)
    except Exception:
        return None
    if not curves:
        return None
    lengths = []
    for c in curves:
        try:
            lengths.append(float(c.Length))
        except Exception:
            lengths.append(0.0)
    min_len = min(lengths)
    tol_abs = max(1e-9, min_len * float(tol_rel))
    for c, ln in zip(curves, lengths):
        if ln <= min_len + tol_abs:
            return c
    return curves[0]


def _curva_lado_menor_desde_teselacion_superior(face, t, z_sup, tol):
    """Segmento más corto del borde teselado en la cara superior → Line."""
    pts_inf, seg_inf = _tessellate_boundary_xy(
        face, t, z_sup, tol, prefer_upper_plane=True
    )
    if not seg_inf:
        return None
    best_pair = None
    best_len = None
    for p1, p2 in seg_inf:
        try:
            ln = float(p1.DistanceTo(p2))
        except Exception:
            continue
        if best_len is None or ln < best_len:
            best_len = ln
            best_pair = (p1, p2)
    if best_pair is None:
        return None
    try:
        return Line.CreateBound(best_pair[0], best_pair[1])
    except Exception:
        return None


def _curva_lado_menor_desde_teselacion(face, t, z_inf, tol):
    """Segmento más corto del borde teselado → Line."""
    pts_inf, seg_inf = _tessellate_boundary_xy(face, t, z_inf, tol)
    if not seg_inf:
        return None
    best_pair = None
    best_len = None
    for p1, p2 in seg_inf:
        try:
            ln = float(p1.DistanceTo(p2))
        except Exception:
            continue
        if best_len is None or ln < best_len:
            best_len = ln
            best_pair = (p1, p2)
    if best_pair is None:
        return None
    try:
        return Line.CreateBound(best_pair[0], best_pair[1])
    except Exception:
        return None


def extraer_curva_lado_menor_cara_inferior(
    elem, tol=0.05, tol_normal_z=0.18, ultra_fallback=False
):
    """
    Una curva del lado más corto. Orden: GetEdgesAsCurveLoops → teselación.

    Args:
        ultra_fallback: ver :func:`_inferior_planar_face_info`.

    Returns:
        tuple | None: (curve, z_inf)
    """
    r = _inferior_planar_face_info(elem, tol, tol_normal_z, ultra_fallback)
    if r is None:
        return None
    face, t, z_inf = r

    out_loops = []
    try:
        raw_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        raw_loops = None
    if raw_loops:
        for cl in raw_loops:
            try:
                if t is not None:
                    out_loops.append(CurveLoop.CreateViaTransform(cl, t))
                else:
                    out_loops.append(CurveLoop.CreateViaCopy(cl))
            except Exception:
                continue
    if out_loops:
        c = curva_lado_menor_desde_loops(out_loops)
        if c is not None:
            return (c, z_inf)

    c2 = _curva_lado_menor_desde_teselacion(face, t, z_inf, tol)
    if c2 is not None:
        return (c2, z_inf)
    return None


def curva_lado_mayor_desde_loops(curve_loops, tol_rel=1e-5, excluir_pts=None):
    """
    Contorno exterior = loop de mayor perímetro; curva = arista de longitud **máxima**.

    ``excluir_pts``: tupla ``(p0, p1)`` de XYZ en coordenadas documento; si la primera
    arista candidata tiene los mismos extremos (tolerancia 1e-3 ft), se salta y se
    devuelve la siguiente candidata. Útil para zapatas cuadradas donde todas las
    aristas tienen la misma longitud y se necesita la arista perpendicular.
    """
    if not curve_loops:
        return None
    try:
        outer = max(curve_loops, key=lambda cl: cl.GetExactLength())
    except Exception:
        outer = curve_loops[0]
    curves = []
    try:
        it = outer.GetCurveIterator()
        while it.MoveNext():
            c = it.Current
            if c is not None:
                curves.append(c)
    except Exception:
        return None
    if not curves:
        return None
    lengths = []
    for c in curves:
        try:
            lengths.append(float(c.Length))
        except Exception:
            lengths.append(0.0)
    max_len = max(lengths)
    tol_abs = max(1e-9, max_len * float(tol_rel))

    def _es_excluida(c):
        if excluir_pts is None:
            return False
        try:
            ep0, ep1 = excluir_pts
            tol_e = 1e-3
            ca0 = c.GetEndPoint(0)
            ca1 = c.GetEndPoint(1)
            return (
                ca0.DistanceTo(ep0) < tol_e and ca1.DistanceTo(ep1) < tol_e
            ) or (
                ca0.DistanceTo(ep1) < tol_e and ca1.DistanceTo(ep0) < tol_e
            )
        except Exception:
            return False

    candidatas = [c for c, ln in zip(curves, lengths) if ln >= max_len - tol_abs]
    for c in candidatas:
        if not _es_excluida(c):
            return c
    return candidatas[0] if candidatas else curves[0]


def _curva_lado_mayor_desde_teselacion_superior(face, t, z_sup, tol, excluir_pts=None):
    """Segmento más largo del borde teselado en la cara superior → Line.

    Args:
        excluir_pts: tupla ``(XYZ, XYZ)`` con los extremos a excluir (ver
            :func:`_curva_lado_mayor_desde_teselacion`).
    """
    pts_inf, seg_inf = _tessellate_boundary_xy(
        face, t, z_sup, tol, prefer_upper_plane=True
    )
    if not seg_inf:
        return None
    best_len = None
    for p1, p2 in seg_inf:
        try:
            ln = float(p1.DistanceTo(p2))
        except Exception:
            continue
        if best_len is None or ln > best_len:
            best_len = ln
    if best_len is None:
        return None
    tol_e = 1.0 / 304.8
    def _es_excl_tess(p1, p2):
        if excluir_pts is None:
            return False
        ea, eb = excluir_pts
        return (
            (p1.DistanceTo(ea) < tol_e and p2.DistanceTo(eb) < tol_e) or
            (p1.DistanceTo(eb) < tol_e and p2.DistanceTo(ea) < tol_e)
        )
    candidatos = []
    for p1, p2 in seg_inf:
        try:
            ln = float(p1.DistanceTo(p2))
        except Exception:
            continue
        if abs(ln - best_len) < tol_e:
            candidatos.append((p1, p2))
    for p1, p2 in candidatos:
        if not _es_excl_tess(p1, p2):
            try:
                return Line.CreateBound(p1, p2)
            except Exception:
                continue
    for p1, p2 in candidatos:
        try:
            return Line.CreateBound(p1, p2)
        except Exception:
            continue
    return None


def _curva_lado_mayor_desde_teselacion(face, t, z_inf, tol, excluir_pts=None):
    """Segmento más largo del borde teselado → Line.

    Args:
        excluir_pts: tupla ``(XYZ, XYZ)`` con los extremos de la curva a excluir.
            Si el candidato de mayor longitud coincide, se retorna el siguiente segmento
            de la misma longitud (útil para zapatas cuadradas donde todas las aristas
            tienen igual longitud y se necesita la arista perpendicular adyacente).
    """
    pts_inf, seg_inf = _tessellate_boundary_xy(face, t, z_inf, tol)
    if not seg_inf:
        return None
    best_len = None
    for p1, p2 in seg_inf:
        try:
            ln = float(p1.DistanceTo(p2))
        except Exception:
            continue
        if best_len is None or ln > best_len:
            best_len = ln
    if best_len is None:
        return None
    tol_e = 1.0 / 304.8
    def _es_excl_tess(p1, p2):
        if excluir_pts is None:
            return False
        ea, eb = excluir_pts
        return (
            (p1.DistanceTo(ea) < tol_e and p2.DistanceTo(eb) < tol_e) or
            (p1.DistanceTo(eb) < tol_e and p2.DistanceTo(ea) < tol_e)
        )
    candidatos = []
    for p1, p2 in seg_inf:
        try:
            ln = float(p1.DistanceTo(p2))
        except Exception:
            continue
        if abs(ln - best_len) < tol_e:
            candidatos.append((p1, p2))
    for p1, p2 in candidatos:
        if not _es_excl_tess(p1, p2):
            try:
                return Line.CreateBound(p1, p2)
            except Exception:
                continue
    for p1, p2 in candidatos:
        try:
            return Line.CreateBound(p1, p2)
        except Exception:
            continue
    return None


def extraer_curva_lado_mayor_cara_inferior(
    elem, tol=0.05, tol_normal_z=0.18, ultra_fallback=False, excluir_curva=None
):
    """
    Una curva del lado más largo del perímetro inferior. Orden: GetEdgesAsCurveLoops → teselación.

    Args:
        ultra_fallback: ver :func:`_inferior_planar_face_info`.
        excluir_curva: ``Curve`` a excluir (mismos extremos) de los candidatos. Útil para
            zapatas cuadradas, donde todas las aristas tienen la misma longitud; pasando la
            arista del ``lado_menor`` se obtiene la arista **perpendicular** adyacente.

    Returns:
        tuple | None: (curve, z_inf)
    """
    r = _inferior_planar_face_info(elem, tol, tol_normal_z, ultra_fallback)
    if r is None:
        return None
    face, t, z_inf = r

    excluir_pts = None
    if excluir_curva is not None:
        try:
            excluir_pts = (excluir_curva.GetEndPoint(0), excluir_curva.GetEndPoint(1))
        except Exception:
            excluir_pts = None

    out_loops = []
    try:
        raw_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        raw_loops = None
    if raw_loops:
        for cl in raw_loops:
            try:
                if t is not None:
                    out_loops.append(CurveLoop.CreateViaTransform(cl, t))
                else:
                    out_loops.append(CurveLoop.CreateViaCopy(cl))
            except Exception:
                continue
    if out_loops:
        c = curva_lado_mayor_desde_loops(out_loops, excluir_pts=excluir_pts)
        if c is not None:
            return (c, z_inf)

    c2 = _curva_lado_mayor_desde_teselacion(face, t, z_inf, tol, excluir_pts=excluir_pts)
    if c2 is not None:
        return (c2, z_inf)
    return None


def extraer_curva_lado_menor_cara_superior(elem, tol=0.05, tol_normal_z=0.18):
    """
    Una curva del lado más corto en la **cara superior**. Orden: GetEdgesAsCurveLoops → teselación.

    Returns:
        tuple | None: (curve, z_sup)
    """
    r = _superior_planar_face_info(elem, tol, tol_normal_z)
    if r is None:
        return None
    face, t, z_sup = r

    out_loops = []
    try:
        raw_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        raw_loops = None
    if raw_loops:
        for cl in raw_loops:
            try:
                if t is not None:
                    out_loops.append(CurveLoop.CreateViaTransform(cl, t))
                else:
                    out_loops.append(CurveLoop.CreateViaCopy(cl))
            except Exception:
                continue
    if out_loops:
        c = curva_lado_menor_desde_loops(out_loops)
        if c is not None:
            return (c, z_sup)

    c2 = _curva_lado_menor_desde_teselacion_superior(face, t, z_sup, tol)
    if c2 is not None:
        return (c2, z_sup)
    return None


def extraer_curva_lado_mayor_cara_superior(elem, tol=0.05, tol_normal_z=0.18, excluir_curva=None):
    """
    Una curva del lado más largo del perímetro **superior**. Orden: GetEdgesAsCurveLoops → teselación.

    Args:
        excluir_curva: ``Curve`` a excluir de los candidatos (ver
            :func:`extraer_curva_lado_mayor_cara_inferior`).

    Returns:
        tuple | None: (curve, z_sup)
    """
    r = _superior_planar_face_info(elem, tol, tol_normal_z)
    if r is None:
        return None
    face, t, z_sup = r

    excluir_pts = None
    if excluir_curva is not None:
        try:
            excluir_pts = (excluir_curva.GetEndPoint(0), excluir_curva.GetEndPoint(1))
        except Exception:
            excluir_pts = None

    out_loops = []
    try:
        raw_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        raw_loops = None
    if raw_loops:
        for cl in raw_loops:
            try:
                if t is not None:
                    out_loops.append(CurveLoop.CreateViaTransform(cl, t))
                else:
                    out_loops.append(CurveLoop.CreateViaCopy(cl))
            except Exception:
                continue
    if out_loops:
        c = curva_lado_mayor_desde_loops(out_loops, excluir_pts=excluir_pts)
        if c is not None:
            return (c, z_sup)

    c2 = _curva_lado_mayor_desde_teselacion_superior(face, t, z_sup, tol, excluir_pts=excluir_pts)
    if c2 is not None:
        return (c2, z_sup)
    return None


def _tessellate_boundary_xy(face, t, z_plane_ft, tol, prefer_upper_plane=False):
    """Puntos y segmentos en el plano horizontal (teselación de aristas).

    ``prefer_upper_plane=False``: puntos en ``min(Z)`` (cara inferior).
    ``prefer_upper_plane=True``: puntos en ``max(Z)`` (cara superior).
    """

    def _xform(p):
        return t.OfPoint(p) if t else p

    face_pts = []
    for edge_loop in list(face.EdgeLoops):
        for edge in edge_loop:
            for pt in edge.Tessellate():
                face_pts.append(_xform(pt))
    if not face_pts:
        return [], []
    if prefer_upper_plane:
        fz_ref = max(p.Z for p in face_pts)
        pts_inf = [p for p in face_pts if abs(p.Z - fz_ref) < tol]
    else:
        fz_min = min(p.Z for p in face_pts)
        pts_inf = [p for p in face_pts if abs(p.Z - fz_min) < tol]
    segs = []
    for edge_loop in list(face.EdgeLoops):
        loop_pts = []
        for edge in edge_loop:
            tess = list(edge.Tessellate())
            if tess:
                loop_pts.append(_xform(tess[0]))
        if len(loop_pts) >= 2:
            for i in range(len(loop_pts)):
                p1, p2 = loop_pts[i], loop_pts[(i + 1) % len(loop_pts)]
                segs.append((p1, p2))
    if not pts_inf:
        pts_inf = [p for p in face_pts if abs(p.Z - z_plane_ft) < tol * 4]
    if not segs and len(pts_inf) >= 2:
        dedup = []
        seen = set()
        for p in pts_inf:
            key = (round(p.X, 6), round(p.Y, 6))
            if key not in seen:
                seen.add(key)
                dedup.append(p)
        if len(dedup) >= 2:
            segs = [(dedup[i], dedup[(i + 1) % len(dedup)]) for i in range(len(dedup))]
    return pts_inf, segs


def muestras_xy_perimetro_inferior_doc(elem, tol=0.05, tol_normal_z=0.18):
    """
    Puntos del contorno de la cara inferior en coordenadas de **documento** (teselación de
    aristas), sin usar el ``BoundingBox`` alineado a ejes del proyecto.
    """
    r = _inferior_planar_face_info(elem, tol, tol_normal_z)
    if r is None:
        return None
    face, transform, z_inf = r
    try:
        pts_inf, _ = _tessellate_boundary_xy(face, transform, z_inf, tol)
    except Exception:
        pts_inf = []
    if not pts_inf:
        return None
    out = []
    seen = set()
    for p in pts_inf:
        try:
            key = (
                round(float(p.X), 5),
                round(float(p.Y), 5),
                round(float(p.Z), 5),
            )
        except Exception:
            continue
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out if len(out) >= 2 else None


def luz_proyeccion_perimetro_inferior_ft(
    elem,
    line_ref,
    rec_planta_mm,
    perpendicular_a_tangente=True,
    tolerancias=None,
):
    """
    Luz en **pies**: (max−min) de la proyección del perímetro inferior sobre un eje en planta,
    menos ``2×rec_planta``. El eje es ⟂ a la tangente de ``line_ref`` en XY si
    ``perpendicular_a_tangente``, si no coincide con dicha tangente (p. ej. largo del muro).

    Usa muestreo del borde real de la cara inferior — no la caja envolvente global del elemento.
    """
    if elem is None or line_ref is None:
        return None
    if tolerancias is None:
        tolerancias = (
            (0.05, 0.18),
            (0.12, 0.30),
            (0.25, 0.45),
            (0.50, 0.70),
        )
    try:
        p0 = line_ref.GetEndPoint(0)
        p1 = line_ref.GetEndPoint(1)
        t = XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
        if float(t.GetLength()) < 1e-12:
            return None
        t = t.Normalize()
        if perpendicular_a_tangente:
            u = XYZ(-float(t.Y), float(t.X), 0.0).Normalize()
        else:
            u = t
        pm = line_ref.Evaluate(0.5, True)
        ox, oy = float(pm.X), float(pm.Y)
        ux, uy = float(u.X), float(u.Y)
    except Exception:
        return None

    for tol, tnz in tolerancias:
        pts = muestras_xy_perimetro_inferior_doc(elem, tol, tnz)
        if pts is None or len(pts) < 2:
            continue
        dots = []
        for p in pts:
            try:
                dots.append(
                    (float(p.X) - ox) * ux + (float(p.Y) - oy) * uy
                )
            except Exception:
                continue
        if len(dots) < 2:
            continue
        span = float(max(dots) - min(dots))
        if span < 1e-9:
            continue
        try:
            rmm = float(rec_planta_mm) if rec_planta_mm is not None else 0.0
        except Exception:
            rmm = 0.0
        try:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId

            rec_ft = UnitUtils.ConvertToInternalUnits(rmm, UnitTypeId.Millimeters)
        except Exception:
            rec_ft = rmm / 304.8
        span = max(span - 2.0 * rec_ft, 1e-4)
        return max(span, 1e-6)
    return None


def span_bruto_proyeccion_perimetro_inferior_ft(
    elem, pm, u_xy, tolerancias=None
):
    """
    Longitud en **pies** de la proyección del perímetro inferior sobre el unitario ``u_xy``
    en planta: ``max((p-pm)·u) - min(...)`` sobre muestreos del borde — **sin** descontar
    recubrimiento. Sirve para acotar el eje de armado a la **huella real**.
    ``pm``: punto en planta (típ. punto medio del ``LocationCurve``).
    """
    if elem is None or pm is None or u_xy is None:
        return None
    if tolerancias is None:
        tolerancias = (
            (0.05, 0.18),
            (0.12, 0.30),
            (0.25, 0.45),
            (0.50, 0.70),
        )
    try:
        ox, oy = float(pm.X), float(pm.Y)
        ux, uy = float(u_xy.X), float(u_xy.Y)
        lu = (ux * ux + uy * uy) ** 0.5
        if lu < 1e-12:
            return None
        ux, uy = ux / lu, uy / lu
    except Exception:
        return None

    for tol, tnz in tolerancias:
        pts = muestras_xy_perimetro_inferior_doc(elem, tol, tnz)
        if pts is None or len(pts) < 2:
            continue
        dots = []
        for p in pts:
            try:
                dots.append(
                    (float(p.X) - ox) * ux + (float(p.Y) - oy) * uy
                )
            except Exception:
                continue
        if len(dots) < 2:
            continue
        span = float(max(dots) - min(dots))
        if span < 1e-9:
            continue
        return span
    return None


def centro_xy_perimetro_inferior_doc(elem, tolerancias=None):
    """
    Centroide en XY del perímetro inferior (muestreo en documento), sin ``BoundingBox``.
    """
    if elem is None:
        return None
    if tolerancias is None:
        tolerancias = (
            (0.05, 0.18),
            (0.12, 0.30),
            (0.25, 0.45),
            (0.50, 0.70),
        )
    for tol, tnz in tolerancias:
        pts = muestras_xy_perimetro_inferior_doc(elem, tol, tnz)
        if pts is None or len(pts) < 2:
            continue
        try:
            sx = sum(float(p.X) for p in pts)
            sy = sum(float(p.Y) for p in pts)
            n = float(len(pts))
            return (sx / n, sy / n)
        except Exception:
            continue
    return None


def extraer_curvas_perimetrales_cara_inferior(elem, tol=0.05, tol_normal_z=0.18):
    """
    Curvas del perímetro de la cara inferior (nativas Revit si es posible).

    Returns:
        tuple | None: (curve_loops, z_inf)
    """
    r = _inferior_planar_face_info(elem, tol, tol_normal_z)
    if r is None:
        return None
    face, t, z_inf = r
    try:
        raw_loops = face.GetEdgesAsCurveLoops()
    except Exception:
        raw_loops = None
    if not raw_loops:
        return None
    out_loops = []
    for cl in raw_loops:
        try:
            if t is not None:
                out_loops.append(CurveLoop.CreateViaTransform(cl, t))
            else:
                out_loops.append(CurveLoop.CreateViaCopy(cl))
        except Exception:
            continue
    if not out_loops:
        return None
    return (out_loops, z_inf)


def elegir_loop_mayor_perimetro(curve_loops):
    """
    De una lista de ``CurveLoop``, devuelve el de mayor longitud total de borde
    (típicamente el contorno exterior en zapatas con huecos).
    """
    if not curve_loops:
        return None
    best = None
    best_per = -1.0
    for cl in curve_loops:
        try:
            per = 0.0
            for c in _iter_curvas_en_curveloop(cl):
                if c is not None:
                    per += float(c.Length)
        except Exception:
            continue
        if per > best_per:
            best_per = per
            best = cl
    return best


def lineas_horizontales_perimetro_inferior_exterior(elem):
    """
    ``Line`` en planta (coords. de documento) por cada tramo horizontal del borde **exterior**
    de la cara inferior (misma lógica que el perímetro nativo de Revit).

    Returns:
        tuple | None: ``(lista_de_Line, z_inf)`` o None si no hay geometría usable.
    """
    r = extraer_curvas_perimetrales_cara_inferior(elem)
    if r is None:
        return None
    loops, z_inf = r
    best = elegir_loop_mayor_perimetro(loops)
    if best is None:
        return None
    tol_z = max(0.02 / 304.8, 1e-6)
    min_len = max(1.0 / 304.8, _MIN_DETAIL_LINE_LEN_FT)
    out = []
    for c in _iter_curvas_en_curveloop(best):
        line, _ln = _linea_corda_horizontal_desde_curva(c, None, tol_z)
        if line is None:
            continue
        try:
            if float(line.Length) < min_len:
                continue
        except Exception:
            continue
        out.append(line)
    if not out:
        return None
    return (out, z_inf)


def normal_saliente_horizontal_paramento_para_barra_horizontal(line, elemento):
    """
    Normal horizontal **saliente** del paramento vertical asociado a una barra horizontal
    en planta (para ``n_outward`` en :func:`offset_linea_eje_barra_desde_cara_inferior_mm`).

    Usa la proyección del vector hacia el centro del ``BoundingBox`` sobre la dirección
    perpendicular a la tangente de la barra (hacia el interior del hormigón en planta).
    """
    if line is None or elemento is None or not isinstance(line, Line):
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
        if d.GetLength() < 1e-12:
            return None
        t = d.Normalize()
        bb = elemento.get_BoundingBox(None)
        if bb is None:
            return None
        c = XYZ(
            (bb.Min.X + bb.Max.X) * 0.5,
            (bb.Min.Y + bb.Max.Y) * 0.5,
            (bb.Min.Z + bb.Max.Z) * 0.5,
        )
        mid = XYZ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5, (p0.Z + p1.Z) * 0.5)
        v_to_center = XYZ(c.X - mid.X, c.Y - mid.Y, 0.0)
        if v_to_center.GetLength() < 1e-12:
            nh = t.CrossProduct(XYZ.BasisZ)
            if nh.GetLength() < 1e-12:
                return None
            return nh.Normalize().Negate()
        v_to_center = v_to_center.Normalize()
        perp_in = v_to_center - t.Multiply(v_to_center.DotProduct(t))
        if perp_in.GetLength() < 1e-12:
            nh = t.CrossProduct(XYZ.BasisZ)
            if nh.GetLength() < 1e-12:
                return None
            nh = nh.Normalize()
        else:
            nh = perp_in.Normalize()
        return nh.Negate()
    except Exception:
        return None


def rango_z_caras_laterales_o_bbox(elem):
    """
    ``(z_min_ft, z_max_ft)`` para distribución en altura: caras verticales si existen;
    si no, ``BoundingBox`` del elemento.
    """
    if elem is None:
        return None, None
    try:
        lat = enumerar_caras_laterales_verticales(elem)
    except Exception:
        lat = []
    if lat:
        try:
            z_lo = min(float(x["fz_min"]) for x in lat if x.get("fz_min") is not None)
            z_hi = max(float(x["fz_max"]) for x in lat if x.get("fz_max") is not None)
            return z_lo, z_hi
        except Exception:
            pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None, None
        return float(bb.Min.Z), float(bb.Max.Z)
    except Exception:
        return None, None


def extraer_cara_inferior(elem, tol=0.05, tol_normal_z=0.18):
    """
    Puntos/segmentos por teselación y Z (compatibilidad).

    Returns:
        tuple | None: (pts_inf, seg_inf, z_inf)
    """
    r = _inferior_planar_face_info(elem, tol, tol_normal_z)
    if r is None:
        return None
    face, t, z_inf = r
    pts_inf, seg_inf = _tessellate_boundary_xy(face, t, z_inf, tol)
    if not pts_inf:
        return None
    return (pts_inf, seg_inf, z_inf)


def enumerar_caras_laterales_verticales(elem, tol_nz=0.82, min_altura_ft=0.01):
    """
    Caras planas **casi verticales** (normal no paralela a Z: ``|n·Z|`` bajo), típicamente
    perímetro lateral de la fundación. Excluye tapas horizontales.

    Returns:
        list[dict]: cada ítem con ``face``, ``transform``, ``fz_min``, ``fz_max``, ``normal``.
    """
    entries, _z_min, _z_max = _collect_face_entries(elem)
    if not entries:
        return []
    out = []
    for face, t, fz_min, fz_max, normal_es in entries:
        if normal_es is None:
            continue
        try:
            nz = abs(float(normal_es.Z))
        except Exception:
            continue
        if nz >= float(tol_nz):
            continue
        if fz_max - fz_min < float(min_altura_ft):
            continue
        out.append(
            {
                "face": face,
                "transform": t,
                "fz_min": fz_min,
                "fz_max": fz_max,
                "normal": normal_es,
            }
        )
    return out


def obtener_marco_coordenadas_cara_lateral(face, transform):
    """
    Marco UVN (``origin``, ``u``, ``v``, ``n`` saliente) para una cara lateral en coords. de documento.

    No aplica la corrección de tapas horizontales; en caras verticales ``|n·Z|`` es bajo y el
    marco se usa igual que en inferior/superior para ``CreateFromCurves``.
    """
    if face is None:
        return None
    o = None
    bx = by = bz = None
    try:
        bbox = face.GetBoundingBox()
        if bbox is not None:
            uu = 0.5 * (bbox.Min.U + bbox.Max.U)
            vv = 0.5 * (bbox.Min.V + bbox.Max.V)
            uv = UV(uu, vv)
        else:
            uv = UV(0.0, 0.0)
        deriv = face.ComputeDerivatives(uv)
        if deriv is not None:
            o = deriv.Origin
            bx = _normalizar_xyz_seguro(deriv.BasisX)
            by = _normalizar_xyz_seguro(deriv.BasisY)
            bz = _normalizar_xyz_seguro(deriv.BasisZ)
    except Exception:
        pass
    if bx is None or by is None or o is None:
        try:
            from Autodesk.Revit.DB import PlanarFace

            if isinstance(face, PlanarFace):
                o = face.Origin
                bx = _normalizar_xyz_seguro(face.XVector)
                by = _normalizar_xyz_seguro(face.YVector)
                bz = _normalizar_xyz_seguro(face.FaceNormal)
        except Exception:
            pass
    if o is None or bx is None or by is None:
        return None
    if bz is None:
        bz = _normalizar_xyz_seguro(bx.CrossProduct(by))
    if bz is None:
        return None
    if transform is not None:
        try:
            o = transform.OfPoint(o)
            bx = transform.OfVector(bx)
            by = transform.OfVector(by)
            bz = transform.OfVector(bz)
        except Exception:
            return None
    return (o, bx, by, bz)


def _iter_curvas_en_curveloop(cl):
    """Itera curvas de un ``CurveLoop`` (API varía según versión)."""
    if cl is None:
        return
    try:
        it = cl.GetCurveIterator()
        while it.MoveNext():
            c = it.Current
            if c is not None:
                yield c
        return
    except Exception:
        pass
    try:
        for c in cl:
            if c is not None:
                yield c
        return
    except Exception:
        pass
    try:
        n = cl.NumberOfCurves()
        for i in range(int(n)):
            yield cl.GetCurveAt(i)
    except Exception:
        pass


def _linea_corda_horizontal_desde_curva(c, transform, tol_z_ft):
    """
    Si ``c`` es una curva con extremos a la misma cota Z (horizontal), devuelve
    una ``Line`` cuerda en coords. de documento y una longitud para ordenar (cuerda en XY).
    """
    if c is None:
        return None, -1.0
    try:
        p0 = c.GetEndPoint(0)
        p1 = c.GetEndPoint(1)
    except Exception:
        return None, -1.0
    p0d = transform.OfPoint(p0) if transform else p0
    p1d = transform.OfPoint(p1) if transform else p1
    if abs(p0d.Z - p1d.Z) > tol_z_ft:
        return None, -1.0
    try:
        dx = p1d.X - p0d.X
        dy = p1d.Y - p0d.Y
        chord = (dx * dx + dy * dy) ** 0.5
    except Exception:
        return None, -1.0
    if chord < 1e-9:
        return None, -1.0
    try:
        ln = float(c.Length)
    except Exception:
        ln = chord
    try:
        line = Line.CreateBound(p0d, p1d)
    except Exception:
        return None, -1.0
    return line, max(chord, ln)


def _curvas_borde_desde_face(face):
    """Lista de ``Curve`` del borde: ``GetEdgesAsCurveLoops`` y, si falla o vacío, ``EdgeLoops``."""
    out = []
    try:
        raw_loops = face.GetEdgesAsCurveLoops()
        if raw_loops:
            for cl in raw_loops:
                for c in _iter_curvas_en_curveloop(cl):
                    if c is not None:
                        out.append(c)
    except Exception:
        pass
    if out:
        return out
    try:
        eloops = face.EdgeLoops
        if eloops is None:
            return out
        for loop in eloops:
            for edge in loop:
                try:
                    c = edge.AsCurve()
                    if c is not None:
                        out.append(c)
                except Exception:
                    pass
    except Exception:
        pass
    return out


def arista_horizontal_mas_larga_cara_lateral(face, transform):
    """
    Tramo **horizontal** (misma cota Z en extremos) de mayor longitud en planta (XY).

    Acepta **Line**, **Arc** u otra ``Curve`` con ``GetEndPoint`` (cuerda recta para la barra).

    Returns:
        ``Line`` en coordenadas de documento, o None.
    """
    if face is None:
        return None
    curvas = _curvas_borde_desde_face(face)
    if not curvas:
        return None
    tol_z = max(0.02 / 304.8, 1e-6)
    best = None
    best_len = -1.0
    for c in curvas:
        if not isinstance(c, Curve):
            continue
        line, ln = _linea_corda_horizontal_desde_curva(c, transform, tol_z)
        if line is None:
            continue
        if ln <= best_len:
            continue
        best_len = ln
        best = line
    return best


def linea_horizontal_cara_lateral_a_cota_z(linea_ref, z_ft):
    """Replica la arista horizontal en otra cota Z (mismas XY)."""
    if linea_ref is None or not isinstance(linea_ref, Line):
        return None
    try:
        p0 = linea_ref.GetEndPoint(0)
        p1 = linea_ref.GetEndPoint(1)
        z = float(z_ft)
        p0n = XYZ(p0.X, p0.Y, z)
        p1n = XYZ(p1.X, p1.Y, z)
        return Line.CreateBound(p0n, p1n)
    except Exception:
        return None


def offset_linea_hacia_interior_desde_cara_inferior_mm(line, n_inferior_saliente, distancia_mm):
    """
    Traslada la ``Line`` una distancia ``distancia_mm`` en la dirección **hacia el interior**
    del sólido desde el plano de la cara inferior: vector ``-n`` si ``n`` es la normal saliente.

    Usado p. ej. para situar la primera barra lateral a una cota fija desde la cara inferior.
    """
    if line is None or not isinstance(line, Line) or n_inferior_saliente is None:
        return line
    try:
        if float(n_inferior_saliente.GetLength()) < 1e-12:
            return line
        inward = n_inferior_saliente.Normalize().Negate()
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        d_ft = UnitUtils.ConvertToInternalUnits(
            float(distancia_mm), UnitTypeId.Millimeters
        )
    except Exception:
        try:
            d_ft = float(distancia_mm) / 304.8
            inward = n_inferior_saliente.Normalize().Negate()
        except Exception:
            return line
    try:
        v = inward.Multiply(d_ft)
        p0 = line.GetEndPoint(0).Add(v)
        p1 = line.GetEndPoint(1).Add(v)
        return Line.CreateBound(p0, p1)
    except Exception:
        return line


_MIN_PATA_U_TRAMO_MM = 5.0
# Patas verticales modeladas (inf.+sup.): altura útil − este recorte (mm) según criterio BIMTools.
_DESCUENTO_LONGITUD_PATA_U_FUNDACION_MM = 150.0


def longitud_pata_u_fundacion_inf_sup_ft(z_min_ft, z_max_ft, menos_mm=None):
    """
    Longitud de cada pata a 90° (pies) cuando se modela la U con inf.+sup. activos:
    ``(z_max − z_min) − menos_mm`` en mm, convertido a pies. Si no es positiva, no aplica.
    """
    if menos_mm is None:
        menos_mm = _DESCUENTO_LONGITUD_PATA_U_FUNDACION_MM
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        m_ft = UnitUtils.ConvertToInternalUnits(float(menos_mm), UnitTypeId.Millimeters)
    except Exception:
        m_ft = float(menos_mm) / 304.8
    try:
        h = float(z_max_ft) - float(z_min_ft)
        span_ft = h - m_ft
    except Exception:
        return None
    if span_ft < 1e-12:
        return None
    try:
        mm_span = span_ft * 304.8
        if float(mm_span) < float(_MIN_PATA_U_TRAMO_MM):
            return None
    except Exception:
        return None
    return float(span_ft)


def largo_gancho_u_tabla_mm(diameter_nominal_mm, concrete_grade=None):
    """
    Largo de pata (mm) por ø nominal — :mod:`bimtools_rebar_hook_lengths`
    (``concrete_grade``: G25/G35/G45 cuando aplica).
    """
    try:
        d = float(diameter_nominal_mm)
    except Exception:
        return None
    if d <= 0.0 or d != d:
        return None
    try:
        from bimtools_rebar_hook_lengths import hook_length_mm_from_nominal_diameter_mm

        return float(hook_length_mm_from_nominal_diameter_mm(d, concrete_grade))
    except Exception:
        return None


def _line_acortar_eje_total_mm_para_cota_forma_revit(line, delta_total_mm):
    """
    Acorta ``Line`` ``delta_total_mm`` (mm) en total, mitad desde cada extremo.
    Compensa repuntes ~ø/2 en la cota del **tramo largo** (p. ej. parámetro B en forma «03»).
    """
    if line is None or not isinstance(line, Line):
        return line
    try:
        dt = float(delta_total_mm)
    except Exception:
        return line
    if dt < 1e-9:
        return line
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        h_ft = 0.5 * float(
            UnitUtils.ConvertToInternalUnits(dt, UnitTypeId.Millimeters)
        )
    except Exception:
        h_ft = 0.5 * float(dt) / 304.8
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        tu = p1.Subtract(p0)
        L = float(tu.GetLength())
        if L < 2.0 * h_ft + 1e-12:
            return line
        tu = tu.Normalize()
        pa = p0.Add(tu.Multiply(h_ft))
        pb = p1.Subtract(tu.Multiply(h_ft))
        return Line.CreateBound(pa, pb)
    except Exception:
        return line


def construir_polilinea_u_fundacion_desde_eje_horizontal(
    linea_eje,
    n_cara_saliente,
    leg_len_ft,
    diameter_nominal_mm=None,
    acortar_eje_central_para_cota_revit=True,
):
    """
    Tres ``Line`` conectadas: pata inicial + eje horizontal + pata final.

    Dirección de cada pata: **+N** (misma que la normal **saliente** ``N`` del marco UVN), de
    forma que los extremos ``q0``/``q1`` queden hacia **dentro** del bloque respecto a la cara
    en la que apoya la malla (Revit / marco UVN: ``−N`` dejaba las patas hacia fuera). Mismo
    vector en ambos extremos; sin mutar ``n_cara_saliente``.

    Args:
        linea_eje: tramo central (eje de barra ya offset).
        n_cara_saliente: ``N`` del marco UVN de la **cara correspondiente** (inferior o superior).
        leg_len_ft: longitud de cada pata (pies).
        diameter_nominal_mm: si se indica y ``acortar_eje_central_para_cota_revit``, el tramo
            central se acorta **ø/2** (mm) en total para alinear la cota del tramo largo con la
            forma en Revit (p. ej. fundación aislada). En **fundación corrida (Wall Foundation)**
            suele desactivarse para cotar el tramo largo = ancho útil con recubrimiento al **eje**
            (p. ej. 50+50 mm), sin descontar otra vez media barra en cota.
        acortar_eje_central_para_cota_revit: si ``False``, no se aplica el acortamiento por
            ``diameter_nominal_mm`` (el diámetro puede seguir pasándose por motivos de API).

    Returns:
        tuple ``(Line, Line, Line)`` o ``None``.
    """
    if linea_eje is None or not isinstance(linea_eje, Line):
        return None
    try:
        leg_ft = float(leg_len_ft)
    except Exception:
        return None
    if leg_ft < 1e-9:
        return None
    line_axis = linea_eje
    if acortar_eje_central_para_cota_revit and diameter_nominal_mm is not None:
        try:
            dk = float(int(round(float(diameter_nominal_mm))))
            if dk > 1e-6:
                sh = _line_acortar_eje_total_mm_para_cota_forma_revit(
                    linea_eje, 0.5 * dk
                )
                if sh is not None:
                    line_axis = sh
        except Exception:
            line_axis = linea_eje
    if n_cara_saliente is None:
        n_unit = XYZ(0.0, 0.0, -1.0)
    else:
        try:
            ln = float(n_cara_saliente.GetLength())
            if ln < 1e-12:
                n_unit = XYZ(0.0, 0.0, -1.0)
            else:
                n_unit = XYZ(
                    float(n_cara_saliente.X) / ln,
                    float(n_cara_saliente.Y) / ln,
                    float(n_cara_saliente.Z) / ln,
                )
        except Exception:
            n_unit = XYZ(0.0, 0.0, -1.0)
    try:
        p0 = line_axis.GetEndPoint(0)
        p1 = line_axis.GetEndPoint(1)
        # ``n_unit`` = normal **saliente** del marco. Patas: **+n_unit** (sentido opuesto al
        # ``−n_unit`` anterior) para que el trazado quede dentro de la fundación.
        vx = float(n_unit.X) * leg_ft
        vy = float(n_unit.Y) * leg_ft
        vz = float(n_unit.Z) * leg_ft
        lv = XYZ(vx, vy, vz)
        # Mismo desplazamiento en **ambos** extremos del eje: q0 = p0−lv y q1 = p1−lv.
        # Con q1 = p1+lv los extremos quedaban en lados opuestos del tramo horizontal (forma Z).
        q0 = p0.Subtract(lv)
        q1 = p1.Subtract(lv)
        ln0 = Line.CreateBound(q0, p0)
        ln1 = Line.CreateBound(p0, p1)
        ln2 = Line.CreateBound(p1, q1)
    except Exception:
        return None
    # #region agent log
    try:
        import json
        import os
        import time

        da = p0.Subtract(q0)
        db = p1.Subtract(q1)
        la = float(da.GetLength())
        lb = float(db.GetLength())
        dot_u = None
        if la > 1e-9 and lb > 1e-9:
            dot_u = (float(da.X) * float(db.X) + float(da.Y) * float(db.Y) + float(da.Z) * float(db.Z)) / (
                la * lb
            )
        _p = os.path.normpath(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "debug-9aea0b.log")
        )
        _line = {
            "sessionId": "9aea0b",
            "location": "geometria_fundacion_cara_inferior:construir_polilinea_u",
            "message": "poly_u_leg_parallelism",
            "data": {"dot_unit_dirs": dot_u},
            "timestamp": int(time.time() * 1000),
            "hypothesisId": "H_geom",
        }
        with open(_p, "a") as f:
            f.write(json.dumps(_line) + "\n")
    except Exception:
        pass
    # #endregion
    return (ln0, ln1, ln2)


def construir_polilinea_fundacion_ganchos_geometricos_desde_eje(
    linea_eje,
    n_cara_saliente,
    leg_len_ft,
    gancho_inicio,
    gancho_fin,
    diameter_nominal_mm=None,
    acortar_eje_central_para_cota_revit=True,
):
    """
    Lista de ``Line`` conectadas que modelan ganchos como tramos rectos (igual criterio de
    patas que :func:`construir_polilinea_u_fundacion_desde_eje_horizontal`), opcional por extremo
    para troceo con empalme (sin ``RebarHookType``).

    ``diameter_nominal_mm``: si ``acortar_eje_central_para_cota_revit``, acorta el tramo recto
    del eje **ø/2** (mm) en total si hay ganchos, para cotas de forma en Revit. En fundación
    corrida suele desactivarse (mismo criterio que la U transversal).

    Returns:
        ``(tramos, linea_central_para_normales)`` o ``(None, None)``. El segundo elemento es
        siempre el tramo recto ``p0``–``p1`` del eje (para ``CreateFromCurves*``); si solo hay
        una curva, coincide con ``linea_eje``.
    """
    if linea_eje is None or not isinstance(linea_eje, Line):
        return None, None
    try:
        gi = bool(gancho_inicio)
        gf = bool(gancho_fin)
    except Exception:
        gi = gf = False
    try:
        leg_ft = float(leg_len_ft)
    except Exception:
        leg_ft = 0.0
    if leg_ft < 1e-9:
        gi = False
        gf = False
    if not gi and not gf:
        try:
            if float(linea_eje.Length) < 1e-12:
                return None, None
        except Exception:
            return None, None
        return [linea_eje], linea_eje
    if gi and gf:
        tri = construir_polilinea_u_fundacion_desde_eje_horizontal(
            linea_eje,
            n_cara_saliente,
            leg_ft,
            diameter_nominal_mm,
            acortar_eje_central_para_cota_revit,
        )
        if tri is None:
            return None, None
        return [tri[0], tri[1], tri[2]], tri[1]
    seg_eje = linea_eje
    if (
        acortar_eje_central_para_cota_revit
        and (gi or gf)
        and diameter_nominal_mm is not None
    ):
        try:
            dk = float(int(round(float(diameter_nominal_mm))))
            if dk > 1e-6:
                _sh = _line_acortar_eje_total_mm_para_cota_forma_revit(
                    linea_eje, 0.5 * dk
                )
                if _sh is not None:
                    seg_eje = _sh
        except Exception:
            seg_eje = linea_eje
    if n_cara_saliente is None:
        n_unit = XYZ(0.0, 0.0, -1.0)
    else:
        try:
            ln_n = float(n_cara_saliente.GetLength())
            if ln_n < 1e-12:
                n_unit = XYZ(0.0, 0.0, -1.0)
            else:
                n_unit = XYZ(
                    float(n_cara_saliente.X) / ln_n,
                    float(n_cara_saliente.Y) / ln_n,
                    float(n_cara_saliente.Z) / ln_n,
                )
        except Exception:
            n_unit = XYZ(0.0, 0.0, -1.0)
    try:
        p0 = seg_eje.GetEndPoint(0)
        p1 = seg_eje.GetEndPoint(1)
        vx = float(n_unit.X) * leg_ft
        vy = float(n_unit.Y) * leg_ft
        vz = float(n_unit.Z) * leg_ft
        lv = XYZ(vx, vy, vz)
        main = Line.CreateBound(p0, p1)
        tramos = []
        if gi:
            q0 = p0.Subtract(lv)
            tramos.append(Line.CreateBound(q0, p0))
        tramos.append(main)
        if gf:
            q1 = p1.Subtract(lv)
            tramos.append(Line.CreateBound(p1, q1))
    except Exception:
        return None, None
    if not tramos:
        return None, None
    return tramos, main


def _sketch_plane_para_linea_modelo(document, line):
    """``SketchPlane`` que contiene la ``Line`` (para ``NewModelCurve``)."""
    if document is None or line is None or not isinstance(line, Line):
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = p1.Subtract(p0)
        if float(d.GetLength()) < 1e-12:
            return None
        du = d.Normalize()
        if abs(float(du.Z)) < 0.9:
            aux = XYZ.BasisZ
        else:
            aux = XYZ.BasisX
        n = du.CrossProduct(aux)
        if float(n.GetLength()) < 1e-12:
            return None
        pl = Plane.CreateByNormalAndOrigin(n.Normalize(), p0)
        return SketchPlane.Create(document, pl)
    except Exception:
        return None


def crear_model_curves_verificacion_polilinea_u(document, tres_lineas):
    """
    Crea una ``ModelCurve`` por tramo de la polilínea U (pata + eje + pata).

    Sirve para comprobar en 3D la geometría enviada a ``CreateFromCurves``. No abre
    transacción: debe llamarse dentro de la transacción activa del llamador.

    Returns:
        Número de tramos creados (0–3).
    """
    if document is None or not tres_lineas:
        return 0
    try:
        seq = tuple(tres_lineas)
    except Exception:
        return 0
    n_ok = 0
    for seg in seq:
        if seg is None or not isinstance(seg, Line):
            continue
        try:
            if float(seg.Length) < _MIN_DETAIL_LINE_LEN_FT:
                continue
        except Exception:
            continue
        sp = _sketch_plane_para_linea_modelo(document, seg)
        if sp is None:
            continue
        try:
            mc = document.Create.NewModelCurve(seg, sp)
            if mc is not None:
                n_ok += 1
        except Exception:
            continue
    return n_ok


def longitud_array_lateral_altura_fundacion_menos_mm_ft(z_min_ft, z_max_ft, menos_mm):
    """
    Longitud del array (pies) para la armadura lateral: **altura de la fundación**
    (``z_max − z_min``) **menos** ``menos_mm`` (p. ej. 200 mm según parámetro del script).
    """
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        m_ft = UnitUtils.ConvertToInternalUnits(float(menos_mm), UnitTypeId.Millimeters)
    except Exception:
        m_ft = float(menos_mm) / 304.8
    try:
        h = float(z_max_ft) - float(z_min_ft)
        span = h - m_ft
    except Exception:
        return 1e-9
    if span < 1e-12:
        return 1e-9
    try:
        return max(float(span), 1e-6)
    except Exception:
        return 1e-9


def primera_cota_z_armadura_lateral_ft(z_min_ft, z_max_ft, rec_tapa_mm, d_mm):
    """
    Cota Z (pies) del **primer** tramo de armadura lateral (recubrimiento tapas + ø/2).
    Si no hay hueco útil entre tapas, devuelve la cota media del elemento.
    """
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        rec_ft = UnitUtils.ConvertToInternalUnits(float(rec_tapa_mm), UnitTypeId.Millimeters)
    except Exception:
        rec_ft = float(rec_tapa_mm) / 304.8
    try:
        d_half_ft = UnitUtils.ConvertToInternalUnits(float(d_mm) * 0.5, UnitTypeId.Millimeters)
    except Exception:
        d_half_ft = float(d_mm) * 0.5 / 304.8
    z_lo = float(z_min_ft) + rec_ft + d_half_ft
    z_hi = float(z_max_ft) - rec_ft - d_half_ft
    if z_hi < z_lo - 1e-9:
        return 0.5 * (float(z_min_ft) + float(z_max_ft))
    if z_hi <= z_lo + 1e-9:
        return 0.5 * (z_lo + z_hi)
    return z_lo


def longitud_distribucion_vertical_lateral_ft(z_min_ft, z_max_ft, rec_tapa_mm, d_mm):
    """
    Longitud en **pies** del tramo vertical útil (entre cotas internas de recubrimiento)
    para ``RebarShapeDrivenAccessor.SetLayoutAsMaximumSpacing`` en armadura lateral:
    un solo ``Rebar`` por arista con distribución en **altura** (misma regla que las cotas
    de :func:`cotas_z_capas_maximum_spacing_lateral`).
    """
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        rec_ft = UnitUtils.ConvertToInternalUnits(float(rec_tapa_mm), UnitTypeId.Millimeters)
    except Exception:
        rec_ft = float(rec_tapa_mm) / 304.8
    try:
        d_half_ft = UnitUtils.ConvertToInternalUnits(float(d_mm) * 0.5, UnitTypeId.Millimeters)
    except Exception:
        d_half_ft = float(d_mm) * 0.5 / 304.8
    z_lo = float(z_min_ft) + rec_ft + d_half_ft
    z_hi = float(z_max_ft) - rec_ft - d_half_ft
    span = z_hi - z_lo
    if span < 1e-12:
        # Sin luz útil: un solo tramo; ``aplicar_layout_maximum_spacing_rebar`` usará SetLayoutAsSingle.
        return 1e-9
    try:
        return max(float(span), 1e-6)
    except Exception:
        return 1.0


def cotas_z_capas_maximum_spacing_lateral(z_min_ft, z_max_ft, rec_tapa_mm, sep_mm, d_mm):
    """
    Cotas Z (pies) para barras horizontales en cara lateral: primera capa tras recubrimiento
    de tapas + radio, luego paso ``sep_mm`` (regla tipo *maximum spacing* en altura).
    """
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId

        rec_ft = UnitUtils.ConvertToInternalUnits(float(rec_tapa_mm), UnitTypeId.Millimeters)
    except Exception:
        rec_ft = float(rec_tapa_mm) / 304.8
    try:
        d_half_ft = UnitUtils.ConvertToInternalUnits(float(d_mm) * 0.5, UnitTypeId.Millimeters)
    except Exception:
        d_half_ft = float(d_mm) * 0.5 / 304.8
    try:
        sep_ft = UnitUtils.ConvertToInternalUnits(float(sep_mm), UnitTypeId.Millimeters)
    except Exception:
        sep_ft = float(sep_mm) / 304.8
    z_lo = float(z_min_ft) + rec_ft + d_half_ft
    z_hi = float(z_max_ft) - rec_ft - d_half_ft
    # Zapata baja: sin hueco útil entre tapas, una sola capa a media altura.
    if z_hi < z_lo - 1e-9:
        zm = 0.5 * (float(z_min_ft) + float(z_max_ft))
        return [zm]
    if z_hi <= z_lo + 1e-9:
        return [(z_lo + z_hi) * 0.5]
    out = []
    z = z_lo
    while z <= z_hi + 1e-9:
        out.append(z)
        z += sep_ft
        if sep_ft < 1e-12:
            break
    if not out:
        out.append((z_lo + z_hi) * 0.5)
    return out
