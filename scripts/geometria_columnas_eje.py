# -*- coding: utf-8 -*-
"""
Eje de columnas: **corte** con plano ⟂ al eje en el punto medio de ``LocationCurve``
(intersección aristas–plano → contorno); una línea paralela al eje por vértice o muestras
perimetrales; si falla, **proyección** a cada cara planar. Sección con **4 vértices** y
``num_curvas_eje = 4+4*k``: índice **1 = esquina inferior-izquierda** y avance **antihorario**
por el perímetro (p. ej. 12 barras: 2 puntos entre esquinas por arista). Con **una** columna,
``ModelCurve`` recibe **Marca** ``01``…``N`` y **Comentarios** ``BIMTools_EjeCol_Impar`` / ``Par``.
Recubrimiento en esquinas: desplazamiento en dos ejes locales del plano de corte; en aristas,
offset radial al centroide. Fundaciones; recorte;
fusión; **empotramiento** (solo ``selection_set`` del usuario):
La **sonda solo sale de los extremos** de cada tramo fusionado (``Line``): desde cada nodo,
tangente saliente opuesta al interior del tramo; **no** se disparan sondas en puntos interiores
del eje. Sobre ese rayo de **500 mm** se muestrea el tramo (varios puntos + esfera **2 mm**;
revolución API) ∩ sólidos seleccionados vía booleana. **Colisión** → anular la prueba, estirón según
**tabla anclaje/empalme** (según Ø nominal: UI **CmbSupDiam** / inferior si aplica, o respaldo
``_EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM``) en la tangente ``du``.
**Sin colisión** → ``p_ext`` (0 estirón). **``du``** = tangente unitaria. ``ModelCurve`` en transacción;
``SketchPlane`` por extremo de eje de cada columna (normal = tangente), ordenados por Z y nombrados.
Pilar **superior** apilado: alargue vertical hacia la junta con la **misma tabla mm/Ø** que el
estirón de empotramiento (UI o ``_EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM``).
"""

import clr
import math

from Autodesk.Revit.DB import (
    Arc,
    BuiltInCategory,
    BuiltInParameter,
    CurveElement,
    CurveLoop,
    IntersectionResultArray,
    ElementId,
    Frame,
    GeometryCreationUtilities,
    GeometryInstance,
    JoinGeometryUtils,
    Line,
    LocationCurve,
    LocationPoint,
    ModelCurve,
    Options,
    PlanarFace,
    Plane,
    SetComparisonResult,
    SketchPlane,
    Solid,
    Transaction,
    UV,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarHookOrientation,
    RebarStyle,
)

from System.Collections.Generic import List

from geometria_viga_cara_superior_detalle import (
    _MIN_LINE_LEN_FT,
    _TOL_SPLIT_FT,
    _iter_planar_faces_elemento,
    _plano_desde_face,
    _unificar_lineas_colineales,
)
from evaluacion_curva_puntos_obstaculos import _punto_en_volumen_solido
from geometria_colision_vigas import solidos_intersectan_por_booleana
from enfierrado_shaft_hashtag import (
    _nearest_tabulated_bar_diameter_mm,
    extension_mm_por_diametro_nominal_mm,
)

# Distancia máxima (pies) de un punto a la recta de referencia para considerar la misma recta infinita.
_TOL_COPLANAR_LINE_FT = max(_TOL_SPLIT_FT, 1.0 / 304.8)  # ≥ ~1 mm
# Coincidencia entre puntos medios de tramos croquis (deduplicación previa al troceo).
_TOL_DEDUPE_PUNTO_MEDIO_FT = _TOL_COPLANAR_LINE_FT
# Cara **lateral** del pilar: |normal_cara·eje_columna| ≤ esto (tapas tienen |·| ≈ 1).
_TOL_LATERAL_FACE_AXIS_DOT = 0.35
# Si |tangente·n_ext| < esto, la barra yace ~en la cara: normal CreateFromCurves = **−n_ext** (hacia interior).
_TOL_REBAR_U_DOT_NEXT_PERP_NFACE = 0.92

# Recorte a lo largo del eje (orden Revit: extremo 0 → 1). El inicio 50 mm solo si hay fundación unida.
_RECORTE_INICIO_CON_FUNDACION_MM = 50.0
_RECORTE_FIN_MM = 25.0

# Empotramiento: sonda diagnóstica y estirón tras colisión (mm).
_EMPOTRAMIENTO_PRUEBA_MM = 500.0
# Diámetro nominal (mm) de respaldo si no llega Ø desde la UI (tabla anclaje/empotramiento).
_EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM = 16.0
# Bbox de Revit por columna + esta holgura mínima (pies) al comprobar si ``pt`` está
# dentro; evita inflar la zona de exclusión del host (no usar mm extra grandes).
_TOL_BBOX_PT_EN_COLUMNA_FT = 0.001  # ~0.3 mm

_TOL_XY_APILADO_MM = 40.0
_MIN_DOT_Z_TRAMO_VERTICAL = 0.82
# Máx. holgura vertical (mm) entre techo del pilar inferior y base del superior para considerar apilado.
_GAP_VERTICAL_APILADO_MM = 80.0

# Esfera de tolerancia en el punto probe (radio mm). Revit no expone CreateSolidBySphere;
# se crea vía ``GeometryCreationUtilities.CreateRevolvedGeometry`` (perfil semicircular).
_PROBE_SPHERE_RADIUS_MM = 2.0
# Umbral volumen (pies³) para intersección booleana (esfera pequeña: usar umbral bajo).
_TOL_VOLUMEN_INTERSECCION_PROBE_FT3 = 1e-12
# Muestras a lo largo de la sonda (p_ext → p_probe): la punta sola puede quedar en aire y fallar
# aunque el tramo corte losa/encuentro; debe detectarse colisión en todo el tramo.
_EMPOTRAMIENTO_SONDA_MUESTRAS = 16
# Márgenes mínimos al repartir curvas en el ancho de la cara (mm).
_DISTRIB_CURVAS_MARGEN_MIN_MM = 25.0
# Caras laterales: normal más horizontal que vertical (|nz| bajo).
_CARA_LATERAL_MAX_ABS_NZ = 0.85
# Dedup de vértices de sección al intersecar aristas con el plano de corte (~2 mm).
_TOL_VERTICES_DEDUP_FT = 2.0 / 304.8
# Offset hacia el interior de cara para ModelLines de **2.ª capa y siguientes** (la 1.ª usa el de UI, ~245 mm).
_MODELO_LINE_OFFSET_CARA_CAPA_INTERIOR_MM = 100.0
# ``ModelCurve`` generada por copia en cadena (4·m): capa ≥2 desde geometría de capa 0 + −**FaceNormal**.
_COMMENT_MODEL_LINE_EJE_CADENA = u"BIMTools_EjeCol_Cadena"
# Segmento modelo que indica la **normal del plano** que soporta la curva de eje (SketchPlane ⟂ tangente).
_MARCADOR_NORMAL_CURVA_EJE_MM = 300.0
# Recubrimiento mínimo cara de hormigón (mm) en la fórmula de desplaz. transversal 4 barras/cara.
_RECUBRIMIENTO_CARA_BASE_MM = 25.0
# Descuento sobre el ancho de cara (mm) para el paso entre las dos barras en modo 8 barras/cara.
_OCHO_BARRAS_DESCUENTO_ANCHO_MM = 50.0


def _mm_to_ft(mm):
    try:
        return float(mm) / 304.8
    except Exception:
        return 0.0


def _mm_traslacion_ancho_cara_cuatro_barras(width_face_mm, diam_estribo_mm, diam_longitudinal_mm):
    """
    Desplazamiento (mm) a lo largo del ancho de la cara: **mitad del ancho** menos 25 mm,
    Ø estribo y **mitad** del Ø longitudinal (barra en esa cara).
    """
    try:
        half = 0.5 * float(width_face_mm)
    except Exception:
        return 0.0
    try:
        d_e = float(diam_estribo_mm or 0.0)
    except Exception:
        d_e = 0.0
    try:
        d_l = float(diam_longitudinal_mm or 0.0)
    except Exception:
        d_l = 0.0
    try:
        return max(
            0.0,
            half - float(_RECUBRIMIENTO_CARA_BASE_MM) - d_e - 0.5 * d_l,
        )
    except Exception:
        return 0.0


def _mm_largo_calculado_copia_por_ancho_cara(
    width_face_mm, diam_estribo_mm, diam_longitudinal_mm
):
    """
    **Largo calculado** común a 8 y 12 barras (mm): ancho de cara asociado menos 50,
    2·Ø estribo y Ø longitudinal. Sobre este valor: en **8** barras se usa **mitad**; en **12**
    se **divide en 3** para el paso entre copias encadenadas.
    """
    try:
        w = float(width_face_mm)
    except Exception:
        return 0.0
    try:
        d_e = float(diam_estribo_mm or 0.0)
    except Exception:
        d_e = 0.0
    try:
        d_l = float(diam_longitudinal_mm or 0.0)
    except Exception:
        d_l = 0.0
    try:
        return max(
            0.0,
            w
            - float(_OCHO_BARRAS_DESCUENTO_ANCHO_MM)
            - 2.0 * d_e
            - d_l,
        )
    except Exception:
        return 0.0


def _mm_offset_normal_segunda_capa_mm(
    width_face_mm,
    k_barras_cara,
    diam_estribo_mm,
    diam_longitudinal_mm,
    es_cuatro_mult_malla,
    k_cuatro_mult,
    indice_anillo_interior=1,
):
    """
    Offset (mm) hacia el interior según cara para anillos **interiores** (antes de troceo):
    ``25 + Ø estribo + Ø long./2 + indice_anillo_interior × paso``;
    ``paso`` = distancia entre líneas del 1.er anillo en esa cara (**N/m** en **4·m**,
    ``N/(k-1)`` genérico). ``indice_anillo_interior`` = 1 para la 2.ª capa, 2 para la 3.ª, etc.
    """
    try:
        base = float(_RECUBRIMIENTO_CARA_BASE_MM)
    except Exception:
        base = 25.0
    try:
        d_e = float(diam_estribo_mm or 0.0)
    except Exception:
        d_e = 0.0
    try:
        d_l = float(diam_longitudinal_mm or 0.0)
    except Exception:
        d_l = 0.0
    rec = base + d_e + 0.5 * d_l
    try:
        kint = int(k_barras_cara)
    except Exception:
        kint = 1
    paso = 0.0
    try:
        if es_cuatro_mult_malla and k_cuatro_mult is not None and kint == int(k_cuatro_mult):
            Lc = _mm_largo_calculado_copia_por_ancho_cara(
                width_face_mm, diam_estribo_mm, diam_longitudinal_mm
            )
            km = int(k_cuatro_mult)
            if km > 0:
                paso = float(Lc) / float(km)
        elif kint > 1:
            Lc = _mm_largo_calculado_copia_por_ancho_cara(
                width_face_mm, diam_estribo_mm, diam_longitudinal_mm
            )
            paso = float(Lc) / float(kint - 1)
    except Exception:
        paso = 0.0
    try:
        mult = max(1, int(indice_anillo_interior))
    except Exception:
        mult = 1
    try:
        return max(0.0, float(rec) + float(paso) * float(mult))
    except Exception:
        return max(0.0, rec)


def _traslacion_linea_mm_u_dir(linea, u_dir, mm_delta):
    """Traslación de ``linea`` en dirección ``±u_dir`` según el signo de ``mm_delta`` (mm)."""
    if linea is None or u_dir is None:
        return None
    try:
        lg = float(u_dir.GetLength())
        if lg < 1e-12:
            return None
        sc = _mm_to_ft(float(mm_delta)) / lg
        du = XYZ(
            float(u_dir.X) * sc,
            float(u_dir.Y) * sc,
            float(u_dir.Z) * sc,
        )
        return _linea_trasladada_por_vector(linea, du)
    except Exception:
        return None


def _linea_trasladada_por_vector(linea, delta_xyz):
    """Traslación rígida de un ``Line``; si falla devuelve ``linea``."""
    if linea is None or delta_xyz is None:
        return linea
    try:
        p0 = linea.GetEndPoint(0) + delta_xyz
        p1 = linea.GetEndPoint(1) + delta_xyz
        if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
            return linea
        return Line.CreateBound(p0, p1)
    except Exception:
        return linea


def _mm_desplazamiento_plano_par_tabla_empalme(diam_nominal_mm):
    """
    Distancia **en mm** para trasladar el plano de troceo de **pares** (+normal), según la tabla
    anclaje/empalme por Ø (mismas filas que ``EXTENSION_MM_BY_BAR_DIAMETER_MM`` en
    ``enfierrado_shaft_hashtag``: 8→570, 10→710, …, 36→3210 mm; interpolación entre filas).

    ``diam_nominal_mm`` (UI / ``empotramiento_diam_nominal_mm``) elige la fila vía
    ``_mm_extension_tabla_anclaje_desde_diametro_nominal``.
    """
    return _mm_extension_tabla_anclaje_desde_diametro_nominal(diam_nominal_mm)


def _mm_estiron_post_troceo_linea_por_diametro(diam_nominal_mm):
    """
    Estirón (mm) del **primer** tramo tras troceo (``p0``→corte), desde el extremo del corte hacia
    el segundo tramo. Misma lógica **Ø → mm** que ``EXTENSION_MM_BY_BAR_DIAMETER_MM``
    (8→570, 10→710, …, 36→3210; interpolación entre filas): ``_mm_extension_tabla_anclaje_desde_diametro_nominal``.
    El **último** tramo (``corte``→``p1``) no se modifica.

    ``diam_nominal_mm``: UI / ``empotramiento_diam_nominal_mm`` (misma entrada que desplaz. plano par).
    """
    return _mm_extension_tabla_anclaje_desde_diametro_nominal(diam_nominal_mm)


def _line_extender_extremo_en_direccion_tramo(ln, extremo_idx, delta_ft):
    """
    Prolonga ``ln`` desde ``GetEndPoint(extremo_idx)``: índice 1 a lo largo de la tangente ``p1-p0``;
    índice 0 en sentido contrario. ``delta_ft`` en pies internos.

    No usa ``XYZ.Normalize()`` como asignación: en IronPython puede quedar ``None`` (void .NET)
    y silenciar el estirón vía ``except``.
    """
    if ln is None or delta_ft <= 1e-12:
        return ln
    try:
        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
        dvec = p1 - p0
        L = float(dvec.GetLength())
        if L < _MIN_LINE_LEN_FT:
            return ln
        invL = 1.0 / L
        udir = XYZ(float(dvec.X) * invL, float(dvec.Y) * invL, float(dvec.Z) * invL)
        dlt = float(delta_ft)
        i = int(extremo_idx)
        if i == 0:
            out = Line.CreateBound(p0 - udir.Multiply(dlt), p1)
        elif i == 1:
            out = Line.CreateBound(p0, p1 + udir.Multiply(dlt))
        else:
            return ln
        if float(out.Length) < _MIN_LINE_LEN_FT:
            return ln
        return out
    except Exception:
        return ln


def _line_primer_tramo_con_estiron_post_troceo(ln_primero, ln_segundo, diam_nominal_mm):
    """
    Si hubo troceo en dos tramos (``ln_segundo`` no ``None``), alarga ``ln_primero`` desde el
    extremo 1 (corte) en dirección al segundo tramo. ``ln_segundo`` queda intacto.
    """
    if ln_primero is None or ln_segundo is None:
        return ln_primero
    ft = _mm_to_ft(_mm_estiron_post_troceo_linea_por_diametro(diam_nominal_mm))
    if ft <= 1e-12:
        return ln_primero
    return _line_extender_extremo_en_direccion_tramo(ln_primero, 1, ft)


def _plano_desplazado_seg_normal(plane, desplazamiento_ft):
    """Misma orientación; nuevo ``Origin`` = anterior + ``desplazamiento_ft`` · normal (XYZ nuevo)."""
    if plane is None:
        return None
    try:
        n = plane.Normal.Normalize()
        o = plane.Origin
        df = float(desplazamiento_ft)
        if abs(df) < 1e-12:
            return Plane.CreateByNormalAndOrigin(
                n, XYZ(float(o.X), float(o.Y), float(o.Z))
            )
        o2 = XYZ(
            float(o.X) + float(n.X) * df,
            float(o.Y) + float(n.Y) * df,
            float(o.Z) + float(n.Z) * df,
        )
        return Plane.CreateByNormalAndOrigin(n, o2)
    except Exception:
        return None


def _desplaza_punto_interior_desde_base_en_segmento(
    p0, p1, pt_base, delta_ft, min_end
):
    """Desplaza ``pt_base`` sobre la recta ``p0-p1`` ±``delta_ft`` (parámetro arco) hasta quedar interior."""
    if p0 is None or p1 is None or pt_base is None:
        return None
    try:
        d = p1 - p0
        L = float(d.GetLength())
        md = float(min_end)
        if L < 2.0 * md + 1e-12:
            return None
        df = float(delta_ft)
        if df < 1e-12:
            return None
        uu = d.Multiply(1.0 / L)
        t0 = float((pt_base - p0).DotProduct(uu))
        for sgn in (1.0, -1.0):
            t = t0 + sgn * df
            if t > md and t < L - md:
                return p0 + d.Multiply(t / L)
        return None
    except Exception:
        return None


def _ajustar_punto_troceo_par_si_coincide_con_impar(
    p0, p1, pt_par, pt_impar, diam_ft, min_end
):
    """
    Si el corte con el plano trasladado coincide con el del plano base (p. ej. respaldo por
    **Z** con ``o.Z`` igual cuando la normal es horizontal y el pilar es vertical), desplazar
    el punto de troceo **Ø** a lo largo del tramo para materializar el offset del plano.
    """
    if pt_par is None:
        return None
    try:
        df = float(diam_ft)
    except Exception:
        return pt_par
    if df < 1e-12:
        return pt_par
    tol = max(1e-4 / 304.8, float(min_end) * 0.02)
    if pt_impar is not None and float(pt_par.DistanceTo(pt_impar)) > tol:
        return pt_par
    base = pt_impar if pt_impar is not None else pt_par
    cand = _desplaza_punto_interior_desde_base_en_segmento(
        p0, p1, base, df, min_end
    )
    return cand if cand is not None else pt_par


def _mm_extension_tabla_anclaje_desde_diametro_nominal(diam_nominal_mm):
    """
    Longitud (mm) según ``EXTENSION_MM_BY_BAR_DIAMETER_MM``: empotramiento tras sonda y junta
    pilar superior comparten esta función (mismo Ø → mismo valor tabulado o interpolado).
    """
    d_in = _EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM
    if diam_nominal_mm is not None:
        try:
            d_in = float(diam_nominal_mm)
        except Exception:
            d_in = _EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM
    if d_in <= 1e-9:
        d_in = _EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM
    try:
        d_tab = _nearest_tabulated_bar_diameter_mm(d_in)
        ext_mm, _ = extension_mm_por_diametro_nominal_mm(d_tab)
        return max(0.0, float(ext_mm))
    except Exception:
        try:
            d_tab = _nearest_tabulated_bar_diameter_mm(_EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM)
            ext_mm, _ = extension_mm_por_diametro_nominal_mm(d_tab)
            return max(0.0, float(ext_mm))
        except Exception:
            return 0.0


def _es_cara_lateral_columna(face, max_abs_nz=_CARA_LATERAL_MAX_ABS_NZ):
    """``True`` si la normal es mayormente horizontal (cara vertical típica de pilar)."""
    if face is None or not isinstance(face, PlanarFace):
        return False
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return False
        nz = abs(float(n.Z))
        return nz < float(max_abs_nz)
    except Exception:
        return False


def _ordenar_caras_por_normal_xy(faces):
    """Orden estable por ángulo del vector normal en planta (emparejar opuestos i y i+F/2)."""

    def ang(f):
        try:
            n = f.FaceNormal
            return math.atan2(float(n.Y), float(n.X))
        except Exception:
            return 0.0

    return sorted(faces or [], key=ang)


def _conteos_por_cara_parejas_opuestas(total_curvas, num_caras):
    """
    Reparto equitativo de ``total_curvas`` entre ``num_caras``. Si ``num_caras`` es par,
    reparte primero entre parejas de caras opuestas (0↔F/2, 1↔F/2+1); dentro de cada par,
    divide lo más equilibrado posible. Si es impar, ``divmod`` clásico.
    """
    F = int(num_caras)
    if F <= 0:
        return []
    try:
        N = max(0, int(total_curvas))
    except Exception:
        N = 0
    if F % 2 != 0:
        base, rem = divmod(N, F)
        return [base + (1 if i < rem else 0) for i in range(F)]
    n_pairs = F // 2
    if n_pairs == 0:
        return [0] * F
    counts = [0] * F
    base_pair, rem_pair = divmod(N, n_pairs)
    for p in range(n_pairs):
        T = base_pair + (1 if p < rem_pair else 0)
        a = T // 2
        b = T - a
        counts[p] = a
        counts[p + n_pairs] = b
    return counts


def _span_bbox_along_u(elemento, u_vec, origin):
    """Longitud (pies) del soporte 1D de la caja del elemento sobre ``u`` unitario desde ``origin``."""
    if elemento is None or u_vec is None or origin is None:
        return 0.0
    try:
        bb = elemento.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return 0.0
    try:
        u = u_vec.Normalize()
    except Exception:
        return 0.0
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
    dots = []
    for c in corners:
        try:
            dots.append(float((c - origin).DotProduct(u)))
        except Exception:
            continue
    if not dots:
        return 0.0
    return max(dots) - min(dots)


def _span_cara_planar_along_u(face, u_vec, origin):
    """
    Extensión (pies) del **borde** de una ``PlanarFace`` proyectado sobre ``u`` desde
    ``origin``. ``u`` debe ser unitario y coplanar con la cara (p. ej. perpendicular al eje
    en la cara). Recorre los bordes del contorno (incl. punto medio de cada curva) para arcos.
    """
    if face is None or not isinstance(face, PlanarFace) or u_vec is None or origin is None:
        return 0.0
    try:
        u = u_vec.Normalize()
    except Exception:
        return 0.0
    dots = []
    try:
        eloops = face.EdgeLoops
        if eloops is None:
            return 0.0
        n_loops = int(eloops.Size)
    except Exception:
        return 0.0
    for li in range(n_loops):
        try:
            loop = eloops.get_Item(li)
            ne = int(loop.Size)
        except Exception:
            continue
        for ej in range(ne):
            try:
                edge = loop.get_Item(ej)
                c = edge.AsCurve()
            except Exception:
                continue
            if c is None or not c.IsBound:
                continue
            for par in (0.0, 0.5, 1.0):
                try:
                    p = c.Evaluate(par, True)
                    dots.append(float((p - origin).DotProduct(u)))
                except Exception:
                    try:
                        if par <= 1e-9:
                            p = c.GetEndPoint(0)
                        elif par >= 1.0 - 1e-9:
                            p = c.GetEndPoint(1)
                        else:
                            continue
                        dots.append(float((p - origin).DotProduct(u)))
                    except Exception:
                        continue
    if len(dots) < 2:
        return 0.0
    return max(dots) - min(dots)


def _vector_ancho_en_cara(face, axis_dir):
    """
    Vector unitario en el plano de la cara, perpendicular al eje proyectado (separación entre curvas).
    ``axis_dir``: dirección de la línea de eje en la cara (normalizada).
    """
    if face is None or not isinstance(face, PlanarFace) or axis_dir is None:
        return None
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        ax = axis_dir.Normalize()
        u = ax.CrossProduct(n)
        if u.GetLength() < 1e-9:
            u = n.CrossProduct(ax)
        if u.GetLength() < 1e-9:
            return None
        return u.Normalize()
    except Exception:
        return None


def _offsets_equidistantes_en_ancho_ft(k, width_ft, margin_ft):
    """
    Desplazamientos (pies) a lo largo de ``u`` para ``k`` curvas paralelas, centradas,
    con ``margin_ft`` libre en ambos lados del ancho útil ``width_ft``.

    - ``k`` **impar**: extremos en ±``avail/2`` e includes **0** (eje proyectado en la cara).
    - ``k`` **par** (p. ej. 2 barras por cara en 8 en total): puntos centrados en la franja
      simétrica y **una** posición se fija en **0** para no omitir la curva proyectada inicial.
    """
    try:
        k = int(k)
    except Exception:
        k = 0
    if k <= 0:
        return []
    try:
        w = float(width_ft)
        m = max(0.0, float(margin_ft))
    except Exception:
        return [0.0] * max(1, k)
    avail = w - 2.0 * m
    if avail < _MIN_LINE_LEN_FT:
        return [0.0] * k
    if k == 1:
        return [0.0]
    if k % 2 == 0:
        base = [
            -0.5 * avail + (float(i) + 0.5) * avail / float(k) for i in range(k)
        ]
        try:
            i0 = min(range(k), key=lambda i: abs(float(base[i])))
            base = list(base)
            base[i0] = 0.0
            return base
        except Exception:
            return base
    step = avail / float(k - 1)
    return [-0.5 * avail + float(i) * step for i in range(k)]


def _lineas_distribuidas_paralelas(line_base, u_dir, offsets_ft):
    """Copias paralelas de ``line_base`` desplazadas ``offsets_ft[i] * u_dir``."""
    if line_base is None or u_dir is None:
        return []
    try:
        p0 = line_base.GetEndPoint(0)
        p1 = line_base.GetEndPoint(1)
        u = u_dir.Normalize()
    except Exception:
        return []
    out = []
    for t in offsets_ft or [0.0]:
        try:
            dt = float(t)
            dlt = u.Multiply(dt)
            ln = Line.CreateBound(p0 + dlt, p1 + dlt)
            if ln is not None and ln.Length >= _MIN_LINE_LEN_FT:
                out.append(ln)
        except Exception:
            continue
    return out if out else [line_base]


def _margen_distribucion_mm_desde_offset(offset_mm):
    try:
        o = float(offset_mm or 0.0)
    except Exception:
        o = 0.0
    base = float(_DISTRIB_CURVAS_MARGEN_MIN_MM)
    if o <= 1e-9:
        return base
    return max(base, min(o * 0.5, 150.0))


def _aplicar_offset_interior_cara(linea, face, offset_mm):
    """
    Desplaza la línea en paralelo hacia el **interior** del volumen respecto a la
    ``PlanarFace`` de referencia (``FaceNormal`` exterior → movimiento ``-n``).
    ``offset_mm`` debe ser el total en mm (25 + Ø estribo + Ø long / 2 cuando aplica).
    """
    if linea is None or face is None or offset_mm <= 1e-9:
        return linea
    if not isinstance(face, PlanarFace):
        return linea
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return linea
        n = n.Normalize()
        ft = _mm_to_ft(offset_mm)
        inward = n.Multiply(-ft)
        p0 = linea.GetEndPoint(0) + inward
        p1 = linea.GetEndPoint(1) + inward
        if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
            return linea
        return Line.CreateBound(p0, p1)
    except Exception:
        return linea


def _face_basis_x_unit_tuple(face):
    """
    Eje **X** tangente de la cara en UV medio (``ComputeDerivatives``), como tupla unitaria.
    Es el vector de referencia pedido para el plano de la ``Rebar``.
    """
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        bbuv = face.GetBoundingBox()
        if bbuv is None:
            return None
        uu = 0.5 * (float(bbuv.Min.U) + float(bbuv.Max.U))
        vv = 0.5 * (float(bbuv.Min.V) + float(bbuv.Max.V))
        tr = face.ComputeDerivatives(UV(uu, vv))
        if tr is None:
            return None
        bx = tr.BasisX
        if bx is None or bx.GetLength() < 1e-12:
            return None
        bx = bx.Normalize()
        return (float(bx.X), float(bx.Y), float(bx.Z))
    except Exception:
        return None


def _face_basis_y_unit_tuple(face):
    """Eje **V** (∂/∂V = ``BasisY`` del ``Transform``) en UV medio, unitario, como tupla."""
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        bbuv = face.GetBoundingBox()
        if bbuv is None:
            return None
        uu = 0.5 * (float(bbuv.Min.U) + float(bbuv.Max.U))
        vv = 0.5 * (float(bbuv.Min.V) + float(bbuv.Max.V))
        tr = face.ComputeDerivatives(UV(uu, vv))
        if tr is None:
            return None
        by = tr.BasisY
        if by is None or by.GetLength() < 1e-12:
            return None
        by = by.Normalize()
        return (float(by.X), float(by.Y), float(by.Z))
    except Exception:
        return None


def _normal_plano_rebar_createfromcurves_desde_tangente_y_eje_referencia(u, ref_axis):
    """
    Normal al plano de ``CreateFromCurves``: ``u`` = tangente del tramo; ``ref_axis`` = referencia
    en el plano de la cara (p. ej. ``BasisU`` de la cara lateral). Proyección ⟂ ``u`` y
    ``n = u × ref_proj`` (regla mano derecha).
    """
    if u is None or ref_axis is None:
        return None
    try:
        if ref_axis.GetLength() < 1e-12:
            return None
        ref = ref_axis.Normalize()
        d = float(u.DotProduct(ref))
        ref_proj = ref - u.Multiply(d)
        if ref_proj.GetLength() < 1e-8:
            return None
        ref_proj = ref_proj.Normalize()
        n = u.CrossProduct(ref_proj)
        if n.GetLength() < 1e-8:
            return None
        return n.Normalize()
    except Exception:
        return None


def _columna_eje_unitario(columna):
    """Dirección del eje del pilar (``LocationCurve``/vertical desde ``LocationPoint``), unitaria."""
    crv = _curva_eje_para_proyeccion(columna)
    if crv is None:
        return XYZ.BasisZ
    t = _tangente_eje_curva_normalizada(crv)
    if t is not None:
        return t
    try:
        d = crv.GetEndPoint(1) - crv.GetEndPoint(0)
        if d.GetLength() > 1e-12:
            return d.Normalize()
    except Exception:
        pass
    return XYZ.BasisZ


def _lateral_face_anchors_from_columna(columna, tol_axis_dot=None):
    """
    Caras **laterales** del pilar: normal casi ⟂ al eje. Cada entrada:
    ``{u'face', u'center', u'host'}`` con ``center`` = ``Evaluate(UV`` medio del bbox de la cara).
    """
    if columna is None:
        return []
    if tol_axis_dot is None:
        tol_axis_dot = float(_TOL_LATERAL_FACE_AXIS_DOT)
    ax = _columna_eje_unitario(columna)
    out = []
    for face in _iter_planar_faces_elemento(columna):
        if face is None or not isinstance(face, PlanarFace):
            continue
        try:
            fn = face.FaceNormal
            if fn is None or fn.GetLength() < 1e-12:
                continue
            fn = fn.Normalize()
            if abs(float(fn.DotProduct(ax))) > tol_axis_dot:
                continue
            bbuv = face.GetBoundingBox()
            if bbuv is None:
                continue
            uu = 0.5 * (float(bbuv.Min.U) + float(bbuv.Max.U))
            vv = 0.5 * (float(bbuv.Min.V) + float(bbuv.Max.V))
            cen = face.Evaluate(UV(uu, vv))
            if cen is None:
                continue
            out.append(
                {u"face": face, u"center": cen, u"host": columna}
            )
        except Exception:
            continue
    return out


def _lateral_face_anchors_from_columnas(columnas):
    acc = []
    for col in columnas or []:
        acc.extend(_lateral_face_anchors_from_columna(col))
    return acc


def _match_punto_a_cara_lateral_mas_cercana(pm, anchors):
    """
    Punto 3D (p. ej. medio de tramo **croquis sin trocear**) vs centros de caras laterales:
    **mínima distancia** euclídea.
    """
    if not anchors or pm is None:
        return None
    best = None
    best_d = None
    for a in anchors:
        try:
            c = a.get(u"center")
            if c is None:
                continue
            d = float(pm.DistanceTo(c))
        except Exception:
            continue
        if best_d is None or d < best_d - 1e-12:
            best_d = d
            best = a
    return best


def _enriquecer_line_metas_croquis_cara_lateral_cercana(
    lineas_croquis,
    line_metas,
    lateral_face_anchors,
):
    """
    **Previo al troceo:** punto medio del croquis → cara lateral más cercana; en la meta:
    ``face_basis_x``, ``face_basis_y`` (marco UV), ``n_ext``, ``face_center_xyz``, ``host_column_id``.
    La normal CreateFromCurves prioriza **−n_ext** cuando la barra está en la cara (|u·n| bajo).
    """
    if not lineas_croquis:
        return line_metas
    if not lateral_face_anchors:
        return line_metas
    n = len(lineas_croquis)
    out = []
    for i in range(n):
        base = {}
        if line_metas is not None and i < len(line_metas):
            m = line_metas[i]
            if m is not None and isinstance(m, dict):
                base = dict(m)
        ln = lineas_croquis[i]
        if ln is not None and getattr(ln, "IsBound", False):
            pm = _punto_medio_linea(ln)
            if pm is not None:
                ma = _match_punto_a_cara_lateral_mas_cercana(
                    pm,
                    lateral_face_anchors,
                )
                if ma is not None:
                    fc = ma.get(u"face")
                    ho = ma.get(u"host")
                    c0 = ma.get(u"center")
                    if c0 is not None:
                        try:
                            base[u"face_center_xyz"] = (
                                float(c0.X),
                                float(c0.Y),
                                float(c0.Z),
                            )
                        except Exception:
                            pass
                    if fc is not None:
                        fbx = _face_basis_x_unit_tuple(fc)
                        if fbx is not None:
                            base[u"face_basis_x"] = fbx
                        fby = _face_basis_y_unit_tuple(fc)
                        if fby is not None:
                            base[u"face_basis_y"] = fby
                        nex = _face_normal_unit_xyz_tuple(fc)
                        if nex is not None:
                            base[u"n_ext"] = nex
                    if ho is not None:
                        try:
                            base[u"host_column_id"] = ho.Id
                        except Exception:
                            pass
        out.append(base if base else None)
    return out


def _face_normal_unit_xyz_tuple(face):
    """Normal unitaria exterior de la ``PlanarFace`` como tupla (x,y,z), o ``None``."""
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        return (float(n.X), float(n.Y), float(n.Z))
    except Exception:
        return None


def _meta_linea_eje_cadena(es_copia_cadena, face, columna=None):
    """Metadatos paralelos a cada ``Line`` de eje (identificación cadena 4·m / normal de cara)."""
    d = {
        u"es_copia_cadena": bool(es_copia_cadena),
        u"n_ext": _face_normal_unit_xyz_tuple(face),
    }
    fbx = _face_basis_x_unit_tuple(face)
    if fbx is not None:
        d[u"face_basis_x"] = fbx
    if columna is not None:
        try:
            d[u"host_column_id"] = columna.Id
        except Exception:
            pass
    return d


def _linea_eje_medio_cerca_vertice_seccion(linea, plane, verts4, tol_ft):
    if linea is None or plane is None or not verts4:
        return False
    try:
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        pm = p0 + (p1 - p0).Multiply(0.5)
        pm2 = _project_point_to_plane(pm, plane)
        if pm2 is None:
            return False
        tf = float(tol_ft)
        for q in verts4:
            if q is None:
                continue
            q2 = _project_point_to_plane(q, plane)
            if q2 is None:
                continue
            if float(pm2.DistanceTo(q2)) <= tf:
                return True
        return False
    except Exception:
        return False


def _filtrar_lineas_capa_sin_vertices_seccion(
    columna, lineas, opts_geom, tol_ft=None, metas_lineas=None
):
    """
    Quita líneas cuyo punto medio (proyectado al plano de sección) coincide con un vértice
    del cuadrilátero de sección (p. ej. capas de refuerzo interiores sin barra en esquina).
    Si ``metas_lineas`` se informa, devuelve ``(lineas, metas)`` con las mismas entradas eliminadas.
    """
    if columna is None or not lineas or opts_geom is None:
        out0 = list(lineas or [])
        if metas_lineas is not None:
            m0 = list(metas_lineas or [])
            if len(m0) < len(out0):
                m0.extend([None] * (len(out0) - len(m0)))
            return out0, m0[: len(out0)]
        return out0
    lineas = list(lineas or [])

    def _devolver_sin_filtrar():
        if metas_lineas is not None:
            m = list(metas_lineas)
            if len(m) < len(lineas):
                m.extend([None] * (len(lineas) - len(m)))
            return lineas, m[: len(lineas)]
        return lineas

    crv = _curva_eje_para_proyeccion(columna)
    if crv is None:
        return _devolver_sin_filtrar()
    plane = _plano_corte_perpendicular_eje_medio(crv)
    axis = _tangente_eje_curva_normalizada(crv)
    if plane is None or axis is None:
        return _devolver_sin_filtrar()
    all_pts = []
    for solid in _iter_solids_elemento(columna, opts_geom):
        for p in _puntos_poligono_seccion_desde_solido(solid, plane):
            all_pts.append(p)
    all_pts = _dedupe_points_xyz(all_pts, _TOL_VERTICES_DEDUP_FT)
    if len(all_pts) < 2:
        return _devolver_sin_filtrar()
    orden = _ordenar_vertices_poligono_en_plano(all_pts, plane)
    if len(orden) != 4:
        return _devolver_sin_filtrar()
    try:
        orden = _cuadrilatero_bl_ccw(orden, plane, axis)
    except Exception:
        pass
    try:
        tf = (
            float(tol_ft)
            if tol_ft is not None
            else max(_TOL_VERTICES_DEDUP_FT, 4.0 / 304.8)
        )
    except Exception:
        tf = max(_TOL_VERTICES_DEDUP_FT, 4.0 / 304.8)
    out = []
    out_m = [] if metas_lineas is not None else None
    for idx, ln in enumerate(lineas):
        if ln is None:
            continue
        if _linea_eje_medio_cerca_vertice_seccion(ln, plane, orden, tf):
            continue
        out.append(ln)
        if out_m is not None:
            out_m.append(
                metas_lineas[idx] if idx < len(metas_lineas) else None
            )
    if out_m is not None:
        return out, out_m
    return out


def _recortar_extremos_linea(linea, mm_inicio, mm_fin):
    """
    Acorta el segmento desde el extremo 0 (inicio) y el extremo 1 (término),
    medidos en mm a lo largo de la dirección del tramo.
    Si el tramo resultante sería demasiado corto, devuelve ``None``.
    """
    if linea is None:
        return None
    try:
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        dvec = p1 - p0
        L = float(dvec.GetLength())
        if L < 1e-12:
            return None
        d = dvec.Normalize()
        a = max(0.0, _mm_to_ft(mm_inicio))
        b = max(0.0, _mm_to_ft(mm_fin))
        if L <= a + b + _MIN_LINE_LEN_FT:
            return None
        p0n = p0 + d.Multiply(a)
        p1n = p1 - d.Multiply(b)
        if p0n.DistanceTo(p1n) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(p0n, p1n)
    except Exception:
        return None


def _punto_en_bbox(pt, bb, tol_ft=0.02, expand_ft=0.0):
    if pt is None or bb is None:
        return False
    try:
        m = float(tol_ft) + float(expand_ft or 0.0)
        return (
            bb.Min.X - m <= pt.X <= bb.Max.X + m
            and bb.Min.Y - m <= pt.Y <= bb.Max.Y + m
            and bb.Min.Z - m <= pt.Z <= bb.Max.Z + m
        )
    except Exception:
        return False


def _geometry_options_empotramiento():
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


def _iter_solids_elemento(elemento, opts):
    if elemento is None:
        return
    try:
        ge = elemento.get_Geometry(opts)
    except Exception:
        return
    if ge is None:
        return
    for obj in ge:
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if float(obj.Volume) < 1e-9:
                    continue
            except Exception:
                continue
            yield obj
        elif isinstance(obj, GeometryInstance):
            try:
                sub = obj.GetInstanceGeometry()
            except Exception:
                continue
            if sub is None:
                continue
            for g2 in sub:
                if isinstance(g2, Solid):
                    try:
                        if float(g2.Volume) < 1e-9:
                            continue
                    except Exception:
                        continue
                    yield g2


def _element_ids_columnas_envolviendo_punto(columnas, pt, expand_ft=0.0):
    """Columnas cuyo bbox contiene ``pt`` (no se usan sus sólidos en la colisión de empotramiento)."""
    excl = []
    tol = float(_TOL_BBOX_PT_EN_COLUMNA_FT)
    for c in columnas or []:
        if c is None:
            continue
        try:
            bb = c.get_BoundingBox(None)
        except Exception:
            bb = None
        if _punto_en_bbox(pt, bb, tol, expand_ft=expand_ft):
            try:
                excl.append(c.Id)
            except Exception:
                pass
    return excl


def _excl_columnas_host_empotramiento(columnas, p_ext):
    """
    Excluye sólidos de **como mucho un** pilar cuyo bbox contiene el nodo ``p_ext``.
    Si ``p_ext`` pertenece a la junta entre dos pilares (varios bbox), **no** se excluye
    ninguno: la colisión debe evaluarse contra todos los ids de la selección, incluido
    el pilar superior vecino (evita falsos «sin colisión» de la sonda 500 mm).
    """
    ids = _element_ids_columnas_envolviendo_punto(columnas, p_ext)
    if len(ids) != 1:
        return []
    return ids


def _id_en_conjunto_excluido(eid, excl_list):
    for x in excl_list or []:
        try:
            if int(x.IntegerValue) == int(eid.IntegerValue):
                return True
        except Exception:
            try:
                if x == eid:
                    return True
            except Exception:
                pass
    return False


def _solid_esfera_tolerancia_en_centro(centre, radius_ft):
    """
    Esfera en memoria (Building Coder): perfil semicircular + revolución 2π.
    ``radius_ft`` en unidades internas de Revit (pies).
    """
    if centre is None:
        return None
    try:
        r = float(radius_ft)
    except Exception:
        return None
    if r <= 1e-12:
        return None
    try:
        frame = Frame(centre, XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ)
        p1 = centre - XYZ.BasisZ.Multiply(r)
        p2 = centre + XYZ.BasisZ.Multiply(r)
        p_mid = centre + XYZ.BasisX.Multiply(r)
        arc = Arc.Create(p1, p2, p_mid)
        ln = Line.CreateBound(arc.GetEndPoint(1), arc.GetEndPoint(0))
        loop = CurveLoop()
        loop.Append(arc)
        loop.Append(ln)
        loops = List[CurveLoop]()
        loops.Add(loop)
        return GeometryCreationUtilities.CreateRevolvedGeometry(
            frame, loops, 0.0, 2.0 * math.pi
        )
    except Exception:
        return None


def _probe_esfera_en_seleccion_exhaustiva(document, ids_seleccion, exclude_elem_ids, p_probe):
    """
    Solo ``ids_seleccion`` del usuario (no se añaden otros elementos del modelo).
    Esfera de radio ``_PROBE_SPHERE_RADIUS_MM`` en ``p_probe``; intersección booleana con
    cada sólido de esos elementos (salvo ids en ``exclude_elem_ids``). Barrido completo.
    Si la esfera no se crea, respaldo punto-volumen en los mismos candidatos.
    """
    if document is None or p_probe is None:
        return False
    r_ft = _mm_to_ft(_PROBE_SPHERE_RADIUS_MM)
    sphere = _solid_esfera_tolerancia_en_centro(p_probe, r_ft)
    opts = _geometry_options_empotramiento()
    hit = False
    for eid in ids_seleccion or []:
        if _id_en_conjunto_excluido(eid, exclude_elem_ids):
            continue
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        for s in _iter_solids_elemento(el, opts):
            if sphere is not None:
                try:
                    if solidos_intersectan_por_booleana(
                        sphere,
                        s,
                        tol_volumen=_TOL_VOLUMEN_INTERSECCION_PROBE_FT3,
                    ):
                        hit = True
                except Exception:
                    pass
            else:
                try:
                    if _punto_en_volumen_solido(s, p_probe):
                        hit = True
                except Exception:
                    pass
    return hit


def _sonda_tramo_toca_seleccion(
    document, ids_seleccion, exclude_elem_ids, p_ext, du_unit, len_ft, n_muestras
):
    """
    Muestreo del rayo de prueba **desde un extremo** ``p_ext`` (nodo de la curva), no desde el
    interior del eje. ``True`` si en algún ``t ∈ [0, len_ft]`` la esfera de tolerancia toca un
    sólido de la selección.
    """
    try:
        L = float(len_ft)
    except Exception:
        return False
    if p_ext is None or du_unit is None or L <= 1e-12:
        return False
    try:
        du = du_unit.Normalize()
    except Exception:
        return False
    n = int(n_muestras)
    if n < 2:
        n = 2
    for i in range(n + 1):
        try:
            t = float(i) / float(n)
        except Exception:
            t = 0.0
        p = p_ext + du.Multiply(L * t)
        if _probe_esfera_en_seleccion_exhaustiva(
            document, ids_seleccion, exclude_elem_ids, p
        ):
            return True
    return False


def _nuevo_extremo_empotramiento(
    p_ext,
    dir_saliente_unit,
    document,
    ids_seleccion,
    columnas,
    mm_prueba,
    diam_nominal_mm=None,
):
    """
    Un **solo extremo** del tramo fusionado: ``p_ext`` es nodo inicial o final de la ``Line``; ``du``
    apunta **hacia fuera** del tramo (no hay sondas desde el interior del eje).

    Colisión **solo** contra ``ids_seleccion``. Sobre el rayo 0…``mm_prueba`` se muestrea el tramo;
    ningún tramo de prueba es geometría final.

    - **Sin colisión:** ``p_ext`` (extensión 0).
    - **Con colisión:** ``p_ext + du × L_ft`` con ``L`` desde la tabla anclaje/empalme
      (``extension_mm_por_diametro_nominal_mm``). ``diam_nominal_mm`` es el Ø (mm) para elegir fila;
      si es ``None`` o no positivo, se usa ``_EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM``.
    """
    if p_ext is None or dir_saliente_unit is None:
        return p_ext
    try:
        du = dir_saliente_unit.Normalize()
    except Exception:
        return p_ext
    excl = _excl_columnas_host_empotramiento(columnas, p_ext)
    # Sonda 500 mm: varias muestras en el tramo (no solo la punta; evita falsos «sin colisión»).
    L_try = _mm_to_ft(mm_prueba)
    choca = _sonda_tramo_toca_seleccion(
        document,
        ids_seleccion,
        excl,
        p_ext,
        du,
        L_try,
        _EMPOTRAMIENTO_SONDA_MUESTRAS,
    )
    if not choca:
        return p_ext
    anchorage_stretch_mm = _mm_extension_tabla_anclaje_desde_diametro_nominal(diam_nominal_mm)
    if anchorage_stretch_mm <= 1e-9:
        return p_ext
    L_anchorage_ft = _mm_to_ft(anchorage_stretch_mm)
    return p_ext + du.Multiply(L_anchorage_ft)


def aplicar_empotramiento_lineas_unificadas(
    document,
    lineas,
    ids_seleccion,
    columnas,
    mm_prueba,
    diam_nominal_mm=None,
    skip_indices_0based=None,
    line_metas=None,
):
    """
    Por cada tramo fusionado: la sonda **solo** se aplica en **los dos extremos** de la curva
    (``GetEndPoint(0/1)``), en dirección tangente saliente; luego ``Line.CreateBound`` con los
    nodos ajustados. Colisión solo en selección; estirón fijo si el rayo de prueba toca sólido.
    ``diam_nominal_mm``: Ø (mm) para la tabla de estirón (p. ej. desde combo de barra en UI).
    ``skip_indices_0based``: índices (0-based) que se dejan **sin** sonda ni estirón (copia de la
    ``Line`` tal cual).
    ``line_metas``: lista paralela opcional; si se informa, el retorno es ``(lineas, line_metas)``.
    """
    salida = []
    salida_meta = [] if line_metas is not None else None
    skip = skip_indices_0based or set()

    def _push_meta(k):
        if salida_meta is not None:
            salida_meta.append(
                line_metas[k] if k < len(line_metas) else None
            )

    for k, ln in enumerate(lineas or []):
        if ln is None:
            continue
        if k in skip:
            salida.append(ln)
            _push_meta(k)
            continue
        try:
            p0 = ln.GetEndPoint(0)
            p1 = ln.GetEndPoint(1)
            dvec = p1 - p0
            if dvec.GetLength() < _MIN_LINE_LEN_FT:
                continue
            d = dvec.Normalize()
            dir_desde_p0 = d.Multiply(-1.0)
            dir_desde_p1 = d
            p0n = _nuevo_extremo_empotramiento(
                p0,
                dir_desde_p0,
                document,
                ids_seleccion,
                columnas,
                mm_prueba,
                diam_nominal_mm=diam_nominal_mm,
            )
            p1n = _nuevo_extremo_empotramiento(
                p1,
                dir_desde_p1,
                document,
                ids_seleccion,
                columnas,
                mm_prueba,
                diam_nominal_mm=diam_nominal_mm,
            )
            if p0n.DistanceTo(p1n) < _MIN_LINE_LEN_FT:
                continue
            salida.append(Line.CreateBound(p0n, p1n))
            _push_meta(k)
        except Exception:
            continue
    if salida_meta is not None:
        return salida, salida_meta
    return salida


def filtrar_solo_structural_columns(document, element_ids):
    """Lista de ``FamilyInstance`` / columnas estructurales desde ids de selección."""
    if document is None or not element_ids:
        return []
    out = []
    bic = int(BuiltInCategory.OST_StructuralColumns)
    for eid in element_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            cat = el.Category
            if cat is None or int(cat.Id.IntegerValue) != bic:
                continue
        except Exception:
            continue
        out.append(el)
    return out


def _curva_desde_location_columna(columna):
    """``LocationCurve`` del elemento."""
    if columna is None:
        return None
    loc = getattr(columna, "Location", None)
    if not isinstance(loc, LocationCurve):
        return None
    try:
        crv = loc.Curve
    except Exception:
        crv = None
    if crv is None or not crv.IsBound:
        return None
    return crv


def _linea_vertical_desde_location_point(columna):
    """
    Muchas columnas estructurales solo tienen ``LocationPoint``; se construye un
    tramo vertical por el punto y el ``BoundingBox`` de la instancia.
    """
    if columna is None:
        return None
    loc = getattr(columna, "Location", None)
    if not isinstance(loc, LocationPoint):
        return None
    try:
        pt = loc.Point
    except Exception:
        return None
    bb = None
    try:
        bb = columna.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return None
    try:
        z0 = float(bb.Min.Z)
        z1 = float(bb.Max.Z)
        if z1 - z0 < _MIN_LINE_LEN_FT:
            z1 = z0 + max(_MIN_LINE_LEN_FT * 2.0, 1.0 / 304.8)
        return Line.CreateBound(
            XYZ(pt.X, pt.Y, z0),
            XYZ(pt.X, pt.Y, z1),
        )
    except Exception:
        return None


def _curva_location_miembro_empalme(elemento):
    """
    ``Curve`` de estructura (viga / columna) desde ``LocationCurve``; si solo hay ``LocationPoint``,
    línea vertical por bbox (mismo criterio que columna).
    """
    if elemento is None:
        return None
    loc = getattr(elemento, "Location", None)
    if isinstance(loc, LocationCurve):
        try:
            c = loc.Curve
            if c is not None and c.IsBound:
                return c
        except Exception:
            pass
        return None
    if isinstance(loc, LocationPoint):
        return _linea_vertical_desde_location_point(elemento)
    return None


def _origen_y_normal_plano_empalme_desde_location_curve(crv):
    """
    Punto y vector para ``Plane.CreateByNormalAndOrigin`` según la ``LocationCurve``:

    - **Origen:** ``GetEndPoint(0)`` (inicio del tramo en Revit).
    - **Normal:** si ``crv`` es ``Line``, ``Line.Direction`` (vector de creación del eje); si no,
      tangente en el inicio o ``GetEndPoint(1) - GetEndPoint(0)`` normalizado.
    """
    if crv is None or not crv.IsBound:
        return None, None
    try:
        origin = crv.GetEndPoint(0)
    except Exception:
        return None, None
    n = None
    if isinstance(crv, Line):
        try:
            d = crv.Direction
            if d is not None and d.GetLength() > 1e-12:
                ln = float(d.GetLength())
                n = d.Multiply(1.0 / ln)
        except Exception:
            n = None
    if n is None:
        n = _tangente_unitaria_en_extremo_curva(crv, True)
    if n is None:
        try:
            d = crv.GetEndPoint(1) - crv.GetEndPoint(0)
            Ld = float(d.GetLength())
            if Ld < 1e-12:
                return None, None
            n = d.Multiply(1.0 / Ld)
        except Exception:
            return None, None
    if n is None or n.GetLength() < 1e-12:
        return None, None
    return origin, n


def _curva_eje_para_proyeccion(columna):
    """Prioriza ``LocationCurve``; si no hay, línea vertical desde ``LocationPoint``."""
    c = _curva_desde_location_columna(columna)
    if c is not None:
        return c
    return _linea_vertical_desde_location_point(columna)


def _tangente_eje_curva_normalizada(crv):
    """Tangente unitaria en el punto medio parametral de la curva de ubicación."""
    if crv is None or not crv.IsBound:
        return None
    try:
        tr = crv.ComputeDerivatives(0.5, True)
        v = tr.BasisX
        if v is None or v.GetLength() < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _plano_corte_perpendicular_eje_medio(crv):
    """
    Plano ⟂ al eje en el punto medio ``Evaluate(0.5, True)``; la normal del plano es la tangente.
    """
    if crv is None or not crv.IsBound:
        return None
    try:
        mid = crv.Evaluate(0.5, True)
        n = _tangente_eje_curva_normalizada(crv)
        if mid is None or n is None:
            return None
        return Plane.CreateByNormalAndOrigin(n, mid)
    except Exception:
        return None


def _interseccion_segmento_plano(p0, p1, plane):
    """Intersección del segmento ``p0-p1`` con ``Plane``; solo si cae dentro del segmento."""
    if plane is None or p0 is None or p1 is None:
        return None
    try:
        n = plane.Normal.Normalize()
        o = plane.Origin
        d = p1 - p0
        den = float(n.DotProduct(d))
        if abs(den) < 1e-12:
            return None
        t = float(n.DotProduct(o - p0)) / den
        if t < -1e-6 or t > 1.0 + 1e-6:
            return None
        return p0 + d.Multiply(t)
    except Exception:
        return None


def _interseccion_segmento_plano_interior(p0, p1, plane, min_dist_from_end_ft):
    """
    Como ``_interseccion_segmento_plano`` pero solo si el corte está alejado de ambos
    extremos al menos ``min_dist_from_end_ft`` (evita tramos degenerados al segmentar).
    """
    if plane is None or p0 is None or p1 is None:
        return None
    try:
        d = p1 - p0
        L = float(d.GetLength())
        if L < 2.0 * float(min_dist_from_end_ft) + 1e-12:
            return None
        n = plane.Normal.Normalize()
        o = plane.Origin
        den = float(n.DotProduct(d))
        if abs(den) < 1e-12:
            return None
        t = float(n.DotProduct(o - p0)) / den
        eps = float(min_dist_from_end_ft) / L
        if t <= eps or t >= 1.0 - eps:
            return None
        return p0 + d.Multiply(t)
    except Exception:
        return None


def _punto_interior_valido_en_segmento(p0, p1, pt, min_dist_from_end_ft):
    """``pt`` debe caer en el segmento abierto alejado de ambos extremos (pies)."""
    if p0 is None or p1 is None or pt is None:
        return None
    try:
        d = p1 - p0
        L = float(d.GetLength())
        md = float(min_dist_from_end_ft)
        if L < 2.0 * md + 1e-12:
            return None
        u = d.Multiply(1.0 / L)
        t_len = float((pt - p0).DotProduct(u))
        if t_len <= md or t_len >= L - md:
            return None
        if t_len < -1e-6 or t_len > L + 1e-6:
            return None
        return pt
    except Exception:
        return None


def _punto_troceo_segmento_plano_empalme(p0, p1, plane, min_dist_from_end_ft):
    """
    Punto interior del segmento ``p0-p1`` para trocear con el plano de empalme.

    - **Transversal** (`n·(p1-p0)` no ~0): intersección estándar del segmento con el plano.
    - **Paralelo** (típico: **líneas verticales** de pilar y plano **vertical** con normal
      horizontal): respaldo por **Z** del ``Origin`` del plano. Al trasladar el plano solo en
      **XY**, ``o.Z`` no cambia y el corte **par** coincidiría con el **impar**: el llamador
      debe aplicar ``_ajustar_punto_troceo_par_si_coincide_con_impar`` con el Ø de barra.
    """
    if plane is None or p0 is None or p1 is None:
        return None
    try:
        d = p1 - p0
        L = float(d.GetLength())
        md = float(min_dist_from_end_ft)
        if L < 2.0 * md + 1e-12:
            return None
        try:
            ln_bound = Line.CreateBound(p0, p1)
            arr = IntersectionResultArray()
            r = ln_bound.Intersect(plane, arr)
            if arr is not None and arr.Size > 0:
                if r in (
                    SetComparisonResult.Overlap,
                    SetComparisonResult.Subset,
                    SetComparisonResult.Superset,
                ):
                    pt_hit = arr.get_Item(0).XYZPoint
                    pt_ok = _punto_interior_valido_en_segmento(p0, p1, pt_hit, md)
                    if pt_ok is not None:
                        return pt_ok
        except Exception:
            pass
        n = plane.Normal.Normalize()
        o = plane.Origin
        eps = md / L
        dist0 = float(n.DotProduct(p0 - o))
        dist1 = float(n.DotProduct(p1 - o))
        denom = dist1 - dist0
        norm_q = max(abs(dist0), abs(dist1), 1e-12)
        if dist0 * dist1 < -1e-9 * max(norm_q, L * 1e-9):
            if abs(denom) > 1e-12:
                t_param = -dist0 / denom
                if t_param > eps and t_param < 1.0 - eps:
                    return p0 + d.Multiply(t_param)
        tol_cop = max(md * 0.25, 1e-4 / 304.8, L * 1e-9)
        if abs(dist0) <= tol_cop and abs(dist1) <= tol_cop:
            u = d.Multiply(1.0 / L)
            t_along = float((o - p0).DotProduct(u))
            if t_along > md and t_along < L - md:
                return p0 + d.Multiply(t_along / L)
        den = float(n.DotProduct(d))
        par_tol = max(1e-11, 1e-9 * max(L, 1.0))
        if abs(den) > par_tol:
            t = float(n.DotProduct(o - p0)) / den
            if t > eps and t < 1.0 - eps:
                return p0 + d.Multiply(t)
        dz = float(p1.Z - p0.Z)
        vert_dot = abs(dz) / L if L > 1e-12 else 0.0
        horiz_n = abs(float(n.DotProduct(XYZ.BasisZ)))
        if vert_dot < 0.45 or horiz_n > 0.98:
            return None
        if abs(dz) < 1e-12:
            return None
        z_cut = float(o.Z)
        t = (z_cut - float(p0.Z)) / dz
        if t <= eps or t >= 1.0 - eps:
            return None
        return p0 + d.Multiply(t)
    except Exception:
        return None


def _interseccion_curva_borde_con_plano(crv, plane):
    """Intersección de una arista (``Line``, ``Arc``, …) con el plano de corte."""
    if crv is None or not crv.IsBound or plane is None:
        return None
    try:
        arr = IntersectionResultArray()
        r = crv.Intersect(plane, arr)
        if (
            r == SetComparisonResult.Overlap
            and arr is not None
            and arr.Size > 0
        ):
            return arr.get_Item(0).XYZPoint
    except Exception:
        pass
    return _interseccion_segmento_plano(
        crv.GetEndPoint(0), crv.GetEndPoint(1), plane
    )


def _dedupe_points_xyz(pts, tol_ft):
    if not pts:
        return []
    try:
        tol = float(tol_ft)
    except Exception:
        tol = _TOL_VERTICES_DEDUP_FT
    out = []
    for p in pts:
        if p is None:
            continue
        dup = False
        for q in out:
            try:
                if p.DistanceTo(q) <= tol:
                    dup = True
                    break
            except Exception:
                continue
        if not dup:
            out.append(p)
    return out


def _puntos_poligono_seccion_desde_solido(solid, plane):
    """Vértices (3D) donde las aristas del sólido cortan el plano."""
    pts = []
    if solid is None or plane is None:
        return pts
    try:
        edges = solid.Edges
    except Exception:
        edges = None
    if edges is None:
        return pts
    try:
        for edge in edges:
            try:
                crv = edge.AsCurve()
            except Exception:
                crv = None
            if crv is None or not crv.IsBound:
                continue
            hit = _interseccion_curva_borde_con_plano(crv, plane)
            if hit is not None:
                pts.append(hit)
    except Exception:
        pass
    return _dedupe_points_xyz(pts, _TOL_VERTICES_DEDUP_FT)


def _ordenar_vertices_poligono_en_plano(pts, plane):
    """Orden angular en el plano alrededor del centroide 2D (contorno exterior coherente)."""
    if not pts or plane is None:
        return []
    try:
        o = plane.Origin
        n = plane.Normal.Normalize()
    except Exception:
        return pts
    try:
        if abs(float(n.DotProduct(XYZ.BasisZ))) < 0.9:
            u = n.CrossProduct(XYZ.BasisZ).Normalize()
        else:
            u = n.CrossProduct(XYZ.BasisX).Normalize()
        v = n.CrossProduct(u).Normalize()
    except Exception:
        return pts
    xy = []
    for p in pts:
        try:
            w = p - o
            xy.append((float(w.DotProduct(u)), float(w.DotProduct(v)), p))
        except Exception:
            continue
    if len(xy) < 3:
        return [t[2] for t in xy]
    cx = sum(t[0] for t in xy) / len(xy)
    cy = sum(t[1] for t in xy) / len(xy)

    def ang_key(item):
        return math.atan2(item[1] - cy, item[0] - cx)

    xy.sort(key=ang_key)
    return [t[2] for t in xy]


def _bases_seccion_abajo_y_derecha(plane, axis_columna):
    """
    En el plano de corte: ``e_down`` ≈ gravedad proyectada (hacia zapata), ``e_right`` = n×e_down.
    Sirve para fijar esquina inferior-izquierda y el sentido antihorario como en croquis de taller.
    """
    if plane is None:
        return None, None
    try:
        n = plane.Normal.Normalize()
    except Exception:
        return None, None
    g_down = XYZ(0.0, 0.0, -1.0)
    try:
        e_d = g_down - n.Multiply(float(g_down.DotProduct(n)))
        if e_d.GetLength() < 1e-9:
            ax = axis_columna.Normalize() if axis_columna is not None else XYZ.BasisZ
            e_d = ax - n.Multiply(float(ax.DotProduct(n)))
            e_d = e_d.Multiply(-1.0)
    except Exception:
        e_d = None
    if e_d is None or e_d.GetLength() < 1e-9:
        return None, None
    try:
        e_d = e_d.Normalize()
        e_r = n.CrossProduct(e_d)
        if e_r.GetLength() < 1e-9:
            return None, None
        e_r = e_r.Normalize()
        return e_d, e_r
    except Exception:
        return None, None


def _firmar_area_poligono_en_plano(orden, plane):
    """Área firmada (u,v) del polígono en la base de ``_ordenar_vertices_poligono_en_plano``. >0 → antihorario."""
    if not orden or plane is None or len(orden) < 3:
        return 0.0
    try:
        o = plane.Origin
        n = plane.Normal.Normalize()
    except Exception:
        return 0.0
    try:
        if abs(float(n.DotProduct(XYZ.BasisZ))) < 0.9:
            u = n.CrossProduct(XYZ.BasisZ).Normalize()
        else:
            u = n.CrossProduct(XYZ.BasisX).Normalize()
        v = n.CrossProduct(u).Normalize()
    except Exception:
        return 0.0
    c = _centroide_puntos_xyz(orden)
    if c is None:
        return 0.0
    A = 0.0
    m = len(orden)
    for i in range(m):
        p0 = orden[i] - c
        p1 = orden[(i + 1) % m] - c
        try:
            x0, y0 = float(p0.DotProduct(u)), float(p0.DotProduct(v))
            x1, y1 = float(p1.DotProduct(u)), float(p1.DotProduct(v))
            A += x0 * y1 - x1 * y0
        except Exception:
            continue
    return A


def _cuadrilatero_bl_ccw(orden4, plane, axis_columna):
    """
    Cuadrilátero con primer vértice = esquina inferior-izquierda (criterio gravedad + «izquierda»)
    y recorrido **antihorario** (1→2→… como perímetro de croquis: fondo izq. → fondo der. → …).
    """
    if not orden4 or len(orden4) != 4 or plane is None:
        return list(orden4) if orden4 else []
    e_d, e_r = _bases_seccion_abajo_y_derecha(plane, axis_columna)
    c = _centroide_puntos_xyz(orden4)
    if c is None or e_d is None:
        return list(orden4)
    best_i = 0
    best_key = None
    for i, p in enumerate(orden4):
        if p is None:
            continue
        try:
            w = p - c
            sd = float(w.DotProduct(e_d))
            sr = float(w.DotProduct(e_r))
            key = (-sd, sr)
        except Exception:
            continue
        if best_key is None or key < best_key:
            best_key = key
            best_i = i
    out = orden4[best_i:] + orden4[:best_i]
    if _firmar_area_poligono_en_plano(out, plane) < 0.0:
        out = [out[0], out[3], out[2], out[1]]
    return out


def _centroide_puntos_xyz(pts):
    if not pts:
        return None
    try:
        sx = sy = sz = 0.0
        n = 0
        for p in pts:
            if p is None:
                continue
            sx += float(p.X)
            sy += float(p.Y)
            sz += float(p.Z)
            n += 1
        if n == 0:
            return None
        inv = 1.0 / float(n)
        return XYZ(sx * inv, sy * inv, sz * inv)
    except Exception:
        return None


def _project_point_to_plane(pt, plane):
    """Proyección ortogonal de un ``XYZ`` al plano."""
    if pt is None or plane is None:
        return None
    try:
        n = plane.Normal
        o = plane.Origin
        if n is None or o is None:
            return None
        n = n.Normalize()
        v = pt - o
        dist = float(v.DotProduct(n))
        return pt - n.Multiply(dist)
    except Exception:
        return None


def _project_curve_to_plane(curve, plane):
    """Proyecta ``Line`` / ``Arc`` al plano; otras curvas acotadas → cuerda."""
    if curve is None or not curve.IsBound or plane is None:
        return None
    try:
        if isinstance(curve, Line):
            p0 = _project_point_to_plane(curve.GetEndPoint(0), plane)
            p1 = _project_point_to_plane(curve.GetEndPoint(1), plane)
            if p0 is None or p1 is None:
                return None
            if p0.DistanceTo(p1) < 1e-9:
                return None
            return Line.CreateBound(p0, p1)
        if isinstance(curve, Arc):
            p0 = _project_point_to_plane(curve.GetEndPoint(0), plane)
            p1 = _project_point_to_plane(curve.GetEndPoint(1), plane)
            c = _project_point_to_plane(curve.Center, plane)
            if p0 is None or p1 is None or c is None:
                return None
            if p0.DistanceTo(p1) < 1e-9:
                return None
            try:
                return Arc.Create(p0, p1, c)
            except Exception:
                return Line.CreateBound(p0, p1)
        p0 = _project_point_to_plane(curve.GetEndPoint(0), plane)
        p1 = _project_point_to_plane(curve.GetEndPoint(1), plane)
        if p0 is None or p1 is None or p0.DistanceTo(p1) < 1e-9:
            return None
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def _linea_para_colinealidad_desde_curva_proyectada(crv):
    """
    Reduce la curva proyectada a un ``Line`` (recta o cuerda) para agrupar y fusionar.
    """
    if crv is None or not crv.IsBound:
        return None
    try:
        if isinstance(crv, Line):
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
                return None
            return crv
        if isinstance(crv, Arc):
            return Line.CreateBound(crv.GetEndPoint(0), crv.GetEndPoint(1))
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def fundaciones_estructurales_unidas(document, columna):
    """
    Fundaciones (categoría ``OST_StructuralFoundation``) con geometría unida a la
    columna vía *Join Geometry* (:class:`JoinGeometryUtils`).
    Incluye fundaciones aisladas u otros tipos de cimentación del mismo host unido.
    """
    if document is None or columna is None:
        return []
    try:
        joined_ids = JoinGeometryUtils.GetJoinedElements(document, columna)
    except Exception:
        return []
    out = []
    bic_f = int(BuiltInCategory.OST_StructuralFoundation)
    for jid in joined_ids:
        try:
            el = document.GetElement(jid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            cat = el.Category
            if cat is None or int(cat.Id.IntegerValue) != bic_f:
                continue
        except Exception:
            continue
        out.append(el)
    return out


def _esquinas_bbox(bb):
    if bb is None:
        return []
    try:
        mn = bb.Min
        mx = bb.Max
    except Exception:
        return []
    pts = []
    for x in (mn.X, mx.X):
        for y in (mn.Y, mx.Y):
            for z in (mn.Z, mx.Z):
                try:
                    pts.append(XYZ(x, y, z))
                except Exception:
                    continue
    return pts


def _eid_iguales(eid_a, eid_b):
    try:
        return int(eid_a.IntegerValue) == int(eid_b.IntegerValue)
    except Exception:
        try:
            return eid_a == eid_b
        except Exception:
            return False


def _columna_es_superior_apilada(columna, todas_columnas):
    """
    True si ``columna`` tiene otro pilar estructural debajo (misma XY en planta)
    y su base está alineada con el apilado típico (losas/junta).
    """
    if columna is None or not todas_columnas or len(todas_columnas) < 2:
        return False
    try:
        bb = columna.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return False
    try:
        cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
        cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
        z_min_self = float(bb.Min.Z)
    except Exception:
        return False
    tol_xy = _mm_to_ft(_TOL_XY_APILADO_MM)
    gap = _mm_to_ft(_GAP_VERTICAL_APILADO_MM)
    for otra in todas_columnas:
        if otra is None or _eid_iguales(otra.Id, columna.Id):
            continue
        try:
            bb2 = otra.get_BoundingBox(None)
        except Exception:
            bb2 = None
        if bb2 is None:
            continue
        try:
            cx2 = 0.5 * (float(bb2.Min.X) + float(bb2.Max.X))
            cy2 = 0.5 * (float(bb2.Min.Y) + float(bb2.Max.Y))
            z_max_otra = float(bb2.Max.Z)
        except Exception:
            continue
        if math.hypot(cx - cx2, cy - cy2) > tol_xy:
            continue
        if z_max_otra <= z_min_self + gap:
            return True
    return False


def _extender_extremo_inferior_tramo_vertical(linea, delta_ft):
    """
    Alarga el extremo con menor Z en ``delta_ft`` (pies) a lo largo del eje del tramo,
    hacia la junta (dirección opuesta al ascenso del tramo). Solo tramos casi verticales.
    """
    if linea is None or delta_ft <= 1e-12:
        return linea
    try:
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        dvec = p1 - p0
        L = float(dvec.GetLength())
        if L < _MIN_LINE_LEN_FT:
            return linea
        ax = dvec.Normalize()
        if abs(float(ax.DotProduct(XYZ.BasisZ))) < _MIN_DOT_Z_TRAMO_VERTICAL:
            return linea
        if float(p0.Z) <= float(p1.Z):
            p_lo, p_hi = p0, p1
        else:
            p_lo, p_hi = p1, p0
        ax_up = (p_hi - p_lo).Normalize()
        p_lo_new = p_lo - ax_up.Multiply(float(delta_ft))
        return Line.CreateBound(p_lo_new, p_hi)
    except Exception:
        return linea


def extender_linea_segun_bboxes(linea, elementos):
    """
    Estira el segmento sobre su misma recta infinita para que abarque las
    proyecciones de los bounding boxes de ``elementos`` (altura/volumen de fundación).
    """
    if linea is None:
        return None
    try:
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        dvec = p1 - p0
        lg = float(dvec.GetLength())
        if lg < 1e-12:
            return linea
        d = dvec.Normalize()
    except Exception:
        return linea
    ts = []
    for i in (0, 1):
        try:
            ts.append(float((linea.GetEndPoint(i) - p0).DotProduct(d)))
        except Exception:
            continue
    for el in elementos or []:
        if el is None:
            continue
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
        for pt in _esquinas_bbox(bb):
            try:
                ts.append(float((pt - p0).DotProduct(d)))
            except Exception:
                continue
    if len(ts) < 2:
        return linea
    tmin = min(ts)
    tmax = max(ts)
    if tmax - tmin < _MIN_LINE_LEN_FT:
        return linea
    try:
        return Line.CreateBound(
            p0 + d.Multiply(tmin),
            p0 + d.Multiply(tmax),
        )
    except Exception:
        return linea


def lineas_eje_columna_extendida_fundaciones(
    document,
    columna,
    offset_mm=0.0,
    todas_columnas=None,
    num_curvas_eje=None,
    diam_nominal_mm=None,
    diam_estribo_mm=None,
    segunda_capa_anillo=False,
    indice_anillo_interior=1,
):
    """
    Para **cada** cara planar de la columna: proyectar el eje al plano de la cara,
    offset hacia el interior, extensión por fundaciones y recorte. Si no hay caras
    en la geometría, un único tramo desde la cuerda del eje (sin offset por cara).

    Si ``num_curvas_eje`` es un entero ≥ 1: se reparte ese **total** entre las **caras
    laterales** (normal mayormente horizontal), equilibrando parejas opuestas en el orden
    de normales en planta; en cada cara se generan ``k`` líneas paralelas al eje proyectado,
    equidistantes en el **ancho de esa cara** (borde de la ``PlanarFace`` proyectado sobre la
    dirección en cara; si no aplica, respaldo al bbox de la columna) con márgenes. Si ``num_curvas_eje``
    es ``None``, se mantiene **una** línea por cada cara planar (comportamiento anterior).

    Si ``columna`` es **superior** en un apilado (otro pilar debajo, misma XY), cada
    tramo casi vertical se alarga hacia la junta según ``diam_nominal_mm`` y la tabla
    anclaje/empalme (mismo criterio que empotramiento por sonda).

    Con **4·m** curvas (``m`` entero ≥ 1), **4** caras laterales en planta y reparto uniforme
    (``m`` barras por cara): la **1.ª** en cada cara como en 4 barras
    (``mitad_ancho - 25 - Ø estribo - Ø long./2`` mm en **+u**). Sea
    ``N = ancho - 50 - 2·Ø estribo - Ø longitudinal``. Cada copia encadenada en **−u** se desplaza
    ``N/m`` mm respecto a la anterior (p. ej. **m=2** → ``N/2``, **m=3** → ``N/3``, **m=4** → ``N/4``).
    Otros repartos: genérico.

    ``segunda_capa_anillo``: si ``True``, mismo flujo que la 1.ª capa (proyección, offset normal,
    fundaciones, reparto **u**), con offset hacia el interior
    ``25 + Ø estribo + Ø long./2 + indice_anillo_interior × paso``.
    ``indice_anillo_interior``: 1 = 2.ª capa, 2 = 3.ª capa, etc.
    ``offset_mm`` sigue usándose para márgenes en reparto genérico (típ. valor UI del 1.er anillo).

    Retorno: ``(lineas, metas)``; en ``metas`` cada dict tiene ``es_copia_cadena`` y ``n_ext``.
    """
    if document is None or columna is None:
        return [], []
    cols_ref = list(todas_columnas) if todas_columnas else [columna]
    aplica_ext_sup = _columna_es_superior_apilada(columna, cols_ref)
    delta_junta_ft = _mm_to_ft(
        _mm_extension_tabla_anclaje_desde_diametro_nominal(diam_nominal_mm)
    )
    crv = _curva_eje_para_proyeccion(columna)
    if crv is None:
        return [], []
    funds = fundaciones_estructurales_unidas(document, columna)
    salida = []
    metas = []
    cuatro_mult_malla = False
    k_cuatro_mult = None
    faces = []
    try:
        faces = list(_iter_planar_faces_elemento(columna))
    except Exception:
        faces = []
    faces_loop = []
    counts = []
    if faces:
        use_distrib = False
        try:
            Ntot = int(num_curvas_eje) if num_curvas_eje is not None else 0
        except Exception:
            Ntot = 0
        if num_curvas_eje is not None and Ntot >= 1:
            use_distrib = True
        if use_distrib:
            lateral = [f for f in faces if _es_cara_lateral_columna(f)]
            faces_use = lateral if lateral else faces
            faces_loop = _ordenar_caras_por_normal_xy(faces_use)
            F = len(faces_loop)
            counts = _conteos_por_cara_parejas_opuestas(Ntot, F) if F else []
            if F == 4 and len(counts) == 4:
                try:
                    k0 = int(counts[0])
                    if all(int(counts[i]) == k0 for i in range(4)):
                        if Ntot == 4 * k0 and Ntot >= 4:
                            cuatro_mult_malla = True
                            k_cuatro_mult = k0
                except Exception:
                    pass
        else:
            faces_loop = list(faces)
            counts = [1] * len(faces_loop)
    if faces_loop:
        for idx, face in enumerate(faces_loop):
            if not isinstance(face, PlanarFace):
                continue

            def _append_out(line_obj, es_cadena=False, fmeta=face):
                if line_obj is None:
                    return
                salida.append(line_obj)
                metas.append(_meta_linea_eje_cadena(es_cadena, fmeta, columna))

            try:
                k = int(counts[idx]) if idx < len(counts) else 1
            except Exception:
                k = 1
            if k <= 0:
                continue
            plane = _plano_desde_face(face)
            if plane is None:
                continue
            proj = _project_curve_to_plane(crv, plane)
            ln = _linea_para_colinealidad_desde_curva_proyectada(proj)
            if ln is None:
                continue
            off_aplicar = 0.0
            if segunda_capa_anillo:
                axis0 = _direccion_normalizada(ln)
                u0 = _vector_ancho_en_cara(face, axis0)
                try:
                    p0a = ln.GetEndPoint(0)
                    p1a = ln.GetEndPoint(1)
                    mid0 = p0a + (p1a - p0a).Multiply(0.5)
                except Exception:
                    mid0 = None
                if u0 is not None and mid0 is not None:
                    w_ft0 = _span_cara_planar_along_u(face, u0, mid0)
                    if w_ft0 < _MIN_LINE_LEN_FT:
                        w_ft0 = _span_bbox_along_u(columna, u0, mid0)
                    w_mm_pre = float(w_ft0) * 304.8
                else:
                    w_mm_pre = 0.0
                off_aplicar = _mm_offset_normal_segunda_capa_mm(
                    w_mm_pre,
                    k,
                    diam_estribo_mm,
                    diam_nominal_mm,
                    cuatro_mult_malla,
                    k_cuatro_mult,
                    indice_anillo_interior=indice_anillo_interior,
                )
            else:
                try:
                    off_aplicar = float(offset_mm) if float(offset_mm) > 1e-9 else 0.0
                except Exception:
                    off_aplicar = 0.0
            if off_aplicar > 1e-9:
                ln = _aplicar_offset_interior_cara(ln, face, off_aplicar)
            if ln is None:
                continue
            out = ln
            if funds:
                ext = extender_linea_segun_bboxes(ln, funds)
                if ext is not None:
                    out = ext
            if funds:
                out = _recortar_extremos_linea(
                    out, _RECORTE_INICIO_CON_FUNDACION_MM, _RECORTE_FIN_MM
                )
            else:
                out = _recortar_extremos_linea(out, 0.0, _RECORTE_FIN_MM)
            if out is None:
                continue
            if aplica_ext_sup:
                out = _extender_extremo_inferior_tramo_vertical(out, delta_junta_ft)
            if out is None:
                continue
            axis_dir = _direccion_normalizada(out)
            u_dir = _vector_ancho_en_cara(face, axis_dir)
            try:
                p0 = out.GetEndPoint(0)
                p1 = out.GetEndPoint(1)
                mid = p0 + (p1 - p0).Multiply(0.5)
            except Exception:
                mid = None
            if u_dir is not None and mid is not None:
                w_ft_u = _span_cara_planar_along_u(face, u_dir, mid)
                if w_ft_u < _MIN_LINE_LEN_FT:
                    w_ft_u = _span_bbox_along_u(columna, u_dir, mid)
                w_mm_u = float(w_ft_u) * 304.8
            else:
                w_mm_u = 0.0
            if (
                cuatro_mult_malla
                and k_cuatro_mult is not None
                and k == k_cuatro_mult
                and u_dir is not None
                and mid is not None
            ):
                d4 = _mm_traslacion_ancho_cara_cuatro_barras(
                    w_mm_u, diam_estribo_mm, diam_nominal_mm
                )
                L_copia = _mm_largo_calculado_copia_por_ancho_cara(
                    w_mm_u, diam_estribo_mm, diam_nominal_mm
                )
                cur = out
                if d4 > 1e-6:
                    t0 = _traslacion_linea_mm_u_dir(out, u_dir, d4)
                    if t0 is not None:
                        cur = t0
                _append_out(cur, False)
                try:
                    kcm = int(k_cuatro_mult)
                except Exception:
                    kcm = 1
                try:
                    step_seg = float(L_copia) / float(kcm) if kcm > 0 else 0.0
                except Exception:
                    step_seg = 0.0
                for _j in range(1, kcm):
                    nxt = cur
                    if step_seg > 1e-9:
                        tj = _traslacion_linea_mm_u_dir(cur, u_dir, -step_seg)
                        if tj is not None:
                            nxt = tj
                    _append_out(nxt, True)
                    cur = nxt
                continue
            if u_dir is None or mid is None or k == 1:
                _append_out(out, False)
                continue
            width_ft = _span_cara_planar_along_u(face, u_dir, mid)
            if width_ft < _MIN_LINE_LEN_FT:
                width_ft = _span_bbox_along_u(columna, u_dir, mid)
            margin_ft = _mm_to_ft(_margen_distribucion_mm_desde_offset(offset_mm))
            offs_ft = _offsets_equidistantes_en_ancho_ft(k, width_ft, margin_ft)
            for piece in _lineas_distribuidas_paralelas(out, u_dir, offs_ft):
                _append_out(piece, False)
        return salida, metas
    ln = _linea_para_colinealidad_desde_curva_proyectada(crv)
    if ln is None:
        return [], []
    out = ln
    if funds:
        ext = extender_linea_segun_bboxes(ln, funds)
        if ext is not None:
            out = ext
    if funds:
        out = _recortar_extremos_linea(
            out, _RECORTE_INICIO_CON_FUNDACION_MM, _RECORTE_FIN_MM
        )
    else:
        out = _recortar_extremos_linea(out, 0.0, _RECORTE_FIN_MM)
    if out is None:
        return [], []
    if aplica_ext_sup:
        out = _extender_extremo_inferior_tramo_vertical(out, delta_junta_ft)
    if out is None:
        return [], []
    return [out], [_meta_linea_eje_cadena(False, None, columna)]


def _direccion_normalizada(line):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < _MIN_LINE_LEN_FT:
            return None
        return d.Normalize()
    except Exception:
        return None


def _dist_punto_a_recta_infinita(punto, linea):
    """Distancia mínima de un punto a la recta definida por ``linea`` (segmento)."""
    try:
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        d = p1 - p0
        L = float(d.GetLength())
        if L < 1e-12:
            return float("inf")
        ax = d.Normalize()
        v = punto - p0
        t = float(v.DotProduct(ax))
        proj = p0 + ax.Multiply(t)
        return float(punto.DistanceTo(proj))
    except Exception:
        return float("inf")


def misma_recta_infinita(linea_a, linea_b, tol_ft=_TOL_COPLANAR_LINE_FT):
    """
    ``True`` si ambos segmentos yacen en la misma recta infinita (misma dirección
    o opuesta y distancia entre rectas ≤ ``tol_ft``).
    """
    da = _direccion_normalizada(linea_a)
    db = _direccion_normalizada(linea_b)
    if da is None or db is None:
        return False
    try:
        if abs(abs(float(da.DotProduct(db))) - 1.0) > 1.0e-4:
            return False
    except Exception:
        return False
    d0 = _dist_punto_a_recta_infinita(linea_b.GetEndPoint(0), linea_a)
    d1 = _dist_punto_a_recta_infinita(linea_b.GetEndPoint(1), linea_a)
    return d0 <= tol_ft and d1 <= tol_ft


def _punto_medio_linea(ln):
    """Punto al centro del tramo acotado (``Evaluate(0.5)`` o promedio de extremos)."""
    if ln is None:
        return None
    try:
        if getattr(ln, "IsBound", False):
            try:
                return ln.Evaluate(0.5, True)
            except Exception:
                pass
        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
        return XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
    except Exception:
        return None


def _dedupe_lineas_mismo_punto_medio(lineas, metas, tol_ft=None):
    """
    Elimina tramos cuyo punto medio coincide (3D) con el de un tramo ya conservado.
    Conserva el primer ejemplar; reindexa implícitamente al devolver listas compactas.

    Retorna ``(lineas, metas, n_eliminados_punto_medio)``.
    """
    if not lineas:
        return [], list(metas or []), 0
    try:
        tf = (
            float(tol_ft)
            if tol_ft is not None
            else float(_TOL_DEDUPE_PUNTO_MEDIO_FT)
        )
    except Exception:
        tf = float(_TOL_DEDUPE_PUNTO_MEDIO_FT)
    m_in = list(metas or [])
    if len(m_in) < len(lineas):
        m_in.extend([None] * (len(lineas) - len(m_in)))
    out_l = []
    out_m = []
    centros_kept = []
    n_dup = 0
    for i, ln in enumerate(lineas):
        if ln is None:
            continue
        pm = _punto_medio_linea(ln)
        if pm is not None:
            es_dup = False
            for c_prev in centros_kept:
                if c_prev is None:
                    continue
                try:
                    if pm.DistanceTo(c_prev) <= tf:
                        es_dup = True
                        break
                except Exception:
                    continue
            if es_dup:
                n_dup += 1
                continue
        out_l.append(ln)
        out_m.append(m_in[i] if i < len(m_in) else None)
        centros_kept.append(pm)
    return out_l, out_m, n_dup


def _segmentos_coinciden_misma_posicion(ln_a, ln_b, tol_ft):
    """``True`` si ambos ``Line`` comparten los mismos extremos (mismo orden o invertido)."""
    if ln_a is None or ln_b is None:
        return False
    try:
        p0a, p1a = ln_a.GetEndPoint(0), ln_a.GetEndPoint(1)
        p0b, p1b = ln_b.GetEndPoint(0), ln_b.GetEndPoint(1)
        tf = float(tol_ft)
    except Exception:
        return False
    if p0a.DistanceTo(p0b) <= tf and p1a.DistanceTo(p1b) <= tf:
        return True
    if p0a.DistanceTo(p1b) <= tf and p1a.DistanceTo(p0b) <= tf:
        return True
    return False


def _segmento_contenido_en_otro(inner, outer, tol_ft):
    """
    ``True`` si el segmento ``inner`` es colineal con ``outer`` y sus extremos caen
    sobre el tramo acotado de ``outer`` (totalmente contenido, salvo tolerancia).
    """
    if inner is None or outer is None:
        return False
    try:
        tf = float(tol_ft)
    except Exception:
        tf = _TOL_COPLANAR_LINE_FT
    try:
        if not misma_recta_infinita(inner, outer, tf):
            return False
        p0o = outer.GetEndPoint(0)
        p1o = outer.GetEndPoint(1)
        v = p1o - p0o
        vv = v.DotProduct(v)
        if vv < tf * tf:
            return False
        len_o = vv ** 0.5

        def t_on_outer(pt):
            return (pt - p0o).DotProduct(v) / vv

        pa0 = inner.GetEndPoint(0)
        pa1 = inner.GetEndPoint(1)
        t0 = t_on_outer(pa0)
        t1 = t_on_outer(pa1)
        lo = min(t0, t1)
        hi = max(t0, t1)
        eps_t = max(1e-9, 2.0 * tf / len_o if len_o > 1e-12 else 1e-9)
        return lo >= -eps_t and hi <= 1.0 + eps_t
    except Exception:
        return False


def _filtrar_subsegmentos_contenidos_colineales(lineas, metas, tol_ft=None):
    """
    Quita cada tramo que esté totalmente contenido en otro más largo (misma recta).
    No sustituye a :func:`_dedupe_lineas_misma_posicion` (mismos extremos).

    Retorna ``(lineas, metas, n_eliminados_subsegmento)``.
    """
    if not lineas or len(lineas) <= 1:
        return list(lineas or []), list(metas or []), 0
    try:
        tf = (
            float(tol_ft)
            if tol_ft is not None
            else max(_TOL_COPLANAR_LINE_FT, 1.0 / 304.8)
        )
    except Exception:
        tf = max(_TOL_COPLANAR_LINE_FT, 1.0 / 304.8)
    m_in = list(metas or [])
    if len(m_in) < len(lineas):
        m_in.extend([None] * (len(lineas) - len(m_in)))
    n = len(lineas)
    remove = set()
    for i in range(n):
        ln_i = lineas[i]
        if ln_i is None:
            continue
        try:
            len_i = float(ln_i.Length)
        except Exception:
            continue
        for j in range(n):
            if i == j:
                continue
            ln_j = lineas[j]
            if ln_j is None:
                continue
            try:
                len_j = float(ln_j.Length)
            except Exception:
                continue
            if len_i > len_j + 2.0 * tf:
                continue
            if not _segmento_contenido_en_otro(ln_i, ln_j, tf):
                continue
            remove.add(i)
            break
    out_l = []
    out_m = []
    for i, ln in enumerate(lineas):
        if i in remove:
            continue
        if ln is None:
            continue
        out_l.append(ln)
        out_m.append(m_in[i] if i < len(m_in) else None)
    try:
        n_sub = int(len(remove))
    except Exception:
        n_sub = 0
    return out_l, out_m, n_sub


def _preparar_lineas_croquis_previo_troceo(lineas, metas=None, tol_ft=None):
    """
    Tras empotramiento y antes del troceo: quita ``None``, deduplica misma posición,
    deduplica por coincidencia del **punto medio** del tramo, elimina subsegmentos
    contenidos en otro tramo colineal y deja listas compactas
    (reindexadas 0…n−1 para impar/par del croquis).

    Retorna ``(lineas, metas, stats)`` con ``stats`` (unicode keys): ``tramos_entrada``,
    ``omitidos_none``, ``dup_misma_posicion``, ``dup_punto_medio``,
    ``dup_subsegmento_contenido``, ``tramos_salida``.
    """
    vac_stats = {
        u"tramos_entrada": 0,
        u"omitidos_none": 0,
        u"dup_misma_posicion": 0,
        u"dup_punto_medio": 0,
        u"dup_subsegmento_contenido": 0,
        u"tramos_salida": 0,
    }
    if not lineas:
        return [], list(metas or []), vac_stats
    try:
        tf = (
            float(tol_ft)
            if tol_ft is not None
            else max(_TOL_COPLANAR_LINE_FT, 1.0 / 304.8)
        )
    except Exception:
        tf = max(_TOL_COPLANAR_LINE_FT, 1.0 / 304.8)
    try:
        omit_none = sum(1 for _ln in lineas if _ln is None)
    except Exception:
        omit_none = 0
    m_in = list(metas or [])
    if len(m_in) < len(lineas):
        m_in.extend([None] * (len(lineas) - len(m_in)))
    compact_l = []
    compact_m = []
    for i, ln in enumerate(lineas):
        if ln is None:
            continue
        compact_l.append(ln)
        compact_m.append(m_in[i] if i < len(m_in) else None)
    if not compact_l:
        st = dict(vac_stats)
        st[u"tramos_entrada"] = len(lineas)
        st[u"omitidos_none"] = omit_none
        return [], [], st
    compact_l, compact_m, n_dup_pos = _dedupe_lineas_misma_posicion(
        compact_l, compact_m, tf
    )
    compact_l, compact_m, n_dup_mid = _dedupe_lineas_mismo_punto_medio(
        compact_l, compact_m, tf
    )
    compact_l, compact_m, n_dup_sub = _filtrar_subsegmentos_contenidos_colineales(
        compact_l, compact_m, tf
    )
    stats = {
        u"tramos_entrada": len(lineas),
        u"omitidos_none": omit_none,
        u"dup_misma_posicion": int(n_dup_pos),
        u"dup_punto_medio": int(n_dup_mid),
        u"dup_subsegmento_contenido": int(n_dup_sub),
        u"tramos_salida": len(compact_l),
    }
    return compact_l, compact_m, stats


def _meta_anexar_capa_prep(m, cap_idx):
    """Marca la meta con el índice de lote/capa para poder repartir tras deduplicar global."""
    try:
        ci = int(cap_idx)
    except Exception:
        ci = 0
    if m is None:
        return {u"_prep_capa": ci}
    if isinstance(m, dict):
        d = dict(m)
        d[u"_prep_capa"] = ci
        return d
    return {u"_prep_capa": ci, u"_prep_meta_obj": m}


def _meta_quitar_marcas_prep(m):
    """Quita marcas internas de preparación antes de crear ``ModelCurve``."""
    if not isinstance(m, dict):
        return m
    d = dict(m)
    d.pop(u"_prep_capa", None)
    mo = d.pop(u"_prep_meta_obj", None)
    if mo is not None and not d:
        return mo
    if not d:
        return None
    return d


def _lotes_croquis_previo_troceo_todos_lotes(lotes_post_empotramiento):
    """
    Une **todos** los tramos croquis de **todos** los lotes (p. ej. cada capa),
    ejecuta :func:`_preparar_lineas_croquis_previo_troceo` una sola vez y vuelve a
    armar los lotes por capa, en orden de índice de lote (reindexación por capa).
    """
    vac_stats = {
        u"tramos_entrada": 0,
        u"omitidos_none": 0,
        u"dup_misma_posicion": 0,
        u"dup_punto_medio": 0,
        u"dup_subsegmento_contenido": 0,
        u"tramos_salida": 0,
    }
    flat_l = []
    flat_m = []
    for cap_i, (lns, ms, _tag) in enumerate(lotes_post_empotramiento or []):
        lns = list(lns or [])
        m_in = list(ms or [])
        if len(m_in) < len(lns):
            m_in.extend([None] * (len(lns) - len(m_in)))
        for j, ln in enumerate(lns):
            if ln is None:
                continue
            m0 = m_in[j] if j < len(m_in) else None
            flat_l.append(ln)
            flat_m.append(_meta_anexar_capa_prep(m0, cap_i))
    if not flat_l:
        return [], 0, vac_stats
    nl, nm, st_prep = _preparar_lineas_croquis_previo_troceo(flat_l, flat_m)
    buckets = {}
    for ln, m in zip(nl or [], nm or []):
        if ln is None:
            continue
        ci = 0
        if isinstance(m, dict):
            try:
                ci = int(m.get(u"_prep_capa", 0))
            except Exception:
                ci = 0
        m_cl = _meta_quitar_marcas_prep(m)
        if ci not in buckets:
            buckets[ci] = [[], []]
        buckets[ci][0].append(ln)
        buckets[ci][1].append(m_cl)
    lotes_compact = []
    n_tramos_croquis = 0
    for cap_i in sorted(buckets.keys()):
        lns_b, ms_b = buckets[cap_i]
        if not lns_b:
            continue
        n_tramos_croquis += len(lns_b)
        lotes_compact.append((lns_b, ms_b, 0))
    return lotes_compact, n_tramos_croquis, st_prep


def _dedupe_lineas_misma_posicion(lineas, metas, tol_ft=None):
    """Elimina tramos duplicados (misma posición); conserva el primero y su meta.

    Retorna ``(lineas, metas, n_eliminados_misma_posicion)``.
    """
    if not lineas:
        return [], list(metas or []), 0
    try:
        tf = (
            float(tol_ft)
            if tol_ft is not None
            else max(_TOL_COPLANAR_LINE_FT, 1.0 / 304.8)
        )
    except Exception:
        tf = max(_TOL_COPLANAR_LINE_FT, 1.0 / 304.8)
    m_in = list(metas or [])
    if len(m_in) < len(lineas):
        m_in.extend([None] * (len(lineas) - len(m_in)))
    out_l = []
    out_m = []
    n_dup = 0
    for i, ln in enumerate(lineas):
        if ln is None:
            continue
        duplicado = False
        for prev in out_l:
            try:
                if _segmentos_coinciden_misma_posicion(ln, prev, tf):
                    duplicado = True
                    break
            except Exception:
                continue
        if duplicado:
            n_dup += 1
            continue
        out_l.append(ln)
        out_m.append(m_in[i] if i < len(m_in) else None)
    return out_l, out_m, n_dup


def _union_find_merge(indices_pairs, n):
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[ri] = rj

    for i, j in indices_pairs:
        union(i, j)
    groups = {}
    for i in range(n):
        r = find(i)
        groups.setdefault(r, []).append(i)
    return list(groups.values())


def agrupar_indices_colineales(lineas, tol_ft=_TOL_COPLANAR_LINE_FT):
    """Índices de ``lineas`` que comparten la misma recta infinita."""
    n = len(lineas)
    if n == 0:
        return []
    pairs = []
    for i in range(n):
        for j in range(i + 1, n):
            try:
                if misma_recta_infinita(lineas[i], lineas[j], tol_ft):
                    pairs.append((i, j))
            except Exception:
                continue
    return _union_find_merge(pairs, n)


def fusionar_lineas_colineales(lineas):
    """
    A partir de tramos ya agrupados en la misma recta, devuelve un único ``Line``
    que cubre el mayor alcance sobre el eje común.
    """
    if not lineas:
        return None
    ref = lineas[0]
    return _unificar_lineas_colineales(lineas, ref)


def _cn_capas_desde_argumento(capas_num_curvas):
    """Lista de enteros ≥ 1 desde ``capas_num_curvas`` (vacía si no aplica multicapa)."""
    _cn = []
    if not capas_num_curvas:
        return _cn
    try:
        for x in list(capas_num_curvas):
            try:
                xi = int(x)
            except Exception:
                continue
            if xi >= 1:
                _cn.append(xi)
    except Exception:
        return []
    return _cn


def lineas_eje_fusionadas_desde_columnas(
    document,
    columnas,
    offset_mm=0.0,
    num_curvas_eje=None,
    diam_nominal_mm=None,
    diam_estribo_mm=None,
    fusionar_colineales=True,
    capas_num_curvas=None,
    incremento_offset_capas_mm=None,
    solo_indice_capa=None,
):
    """
    Por columna: eje proyectado a **cada** cara planar, offset, fundaciones y recorte;
    opcionalmente ``num_curvas_eje`` curvas repartidas por caras laterales.

    Si ``capas_num_curvas`` es una lista no vacía, la **1.ª capa** usa ``offset_mm`` (típ. ~245 mm
    desde UI) y ``capas_num_curvas[0]`` barras.
    Las **siguientes** capas usan la **misma** generación que la 1.ª (proyección a caras, fundaciones,
    reparto en **u**) con offset normal ``25 + Ø estribo + Ø long./2 + paso`` entre líneas del 1.er
    anillo ``i`` (:func:`_mm_offset_normal_segunda_capa_mm` con ``indice_anillo_interior=i``).
    ``incremento_offset_capas_mm`` se conserva en la API por compatibilidad; **no** entra en la
    geometría entre anillos. En capas ≥ 2 se filtran
    vértices de sección (:func:`_filtrar_lineas_capa_sin_vertices_seccion`) y se deduplican tramos
    (:func:`_dedupe_lineas_misma_posicion`). Solo el **primer** entero de la lista define
    el número de barras; el resto de entradas solo fija cuántos anillos hay (p. ej. duplicados en UI columnas).

    Si ``fusionar_colineales`` es ``True`` (predeterminado), agrupa tramos en la misma recta
    infinita y fusiona en un solo ``Line``. Si es ``False``, se conserva **un** ``Line`` por
    tramo generado (p. ej. antes de troceo por plano de empalme: no unir barras paralelas ni
    segmentos que deban quedar independientes).

    ``diam_nominal_mm``: Ø (mm) longitudinal para junta apilada y traslación 4 barras/cara.

    ``diam_estribo_mm``: Ø (mm) estribo para la traslación transversal con ``num_curvas_eje=4``.

    ``solo_indice_capa``: si no es ``None`` y hay multicapa, solo se genera ese índice de anillo
    (0 = 1.ª capa, 1 = 2.ª, …) para troceo/empotramiento **por capa** en ``ejecutar_model_lines_eje_columnas``.

    Retorno: ``(lineas, metas)`` listas paralelas. Si en una sola llamada se emiten **varias** capas
    (``solo_indice_capa`` es ``None`` y hay más de un anillo), no se aplica fusión colineal para no
    mezclar tramos de distintos anillos. Si se pide **una** capa con ``solo_indice_capa``, la fusión
    sigue el mismo criterio que con un solo anillo en la lista. Tras fusión, ``metas`` queda en
    ``None`` por tramo fusionado.
    """
    segs = []
    seg_metas = []
    _cn = _cn_capas_desde_argumento(capas_num_curvas)
    use_capas = len(_cn) > 0
    n_capas_req = len(_cn) if use_capas else 1
    if use_capas:
        opts_geom_capas = _geometry_options_empotramiento()
        for col in columnas or []:
            try:
                n0 = int(_cn[0])
            except Exception:
                continue
            if n0 < 1:
                continue
            for capa_idx in range(n_capas_req):
                if solo_indice_capa is not None:
                    try:
                        if int(capa_idx) != int(solo_indice_capa):
                            continue
                    except Exception:
                        continue
                off_ui = float(offset_mm)
                if capa_idx == 0:
                    capa_lineas, capa_metas_out = lineas_eje_columna_extendida_fundaciones(
                        document,
                        col,
                        offset_mm=off_ui,
                        todas_columnas=columnas,
                        num_curvas_eje=n0,
                        diam_nominal_mm=diam_nominal_mm,
                        diam_estribo_mm=diam_estribo_mm,
                        segunda_capa_anillo=False,
                    )
                else:
                    capa_lineas, capa_metas_out = lineas_eje_columna_extendida_fundaciones(
                        document,
                        col,
                        offset_mm=off_ui,
                        todas_columnas=columnas,
                        num_curvas_eje=n0,
                        diam_nominal_mm=diam_nominal_mm,
                        diam_estribo_mm=diam_estribo_mm,
                        segunda_capa_anillo=True,
                        indice_anillo_interior=capa_idx,
                    )
                    capa_lineas, capa_metas_out = (
                        _filtrar_lineas_capa_sin_vertices_seccion(
                            col,
                            capa_lineas,
                            opts_geom_capas,
                            metas_lineas=capa_metas_out,
                        )
                    )
                    capa_lineas, capa_metas_out, _dup_capa = (
                        _dedupe_lineas_misma_posicion(
                            capa_lineas, capa_metas_out
                        )
                    )
                for ln, m in zip(capa_lineas or [], capa_metas_out or []):
                    if ln is not None:
                        segs.append(ln)
                        seg_metas.append(m)
    else:
        for col in columnas or []:
            lns, ms = lineas_eje_columna_extendida_fundaciones(
                document,
                col,
                offset_mm=offset_mm,
                todas_columnas=columnas,
                num_curvas_eje=num_curvas_eje,
                diam_nominal_mm=diam_nominal_mm,
                diam_estribo_mm=diam_estribo_mm,
            )
            for ln, m in zip(lns or [], ms or []):
                if ln is not None:
                    segs.append(ln)
                    seg_metas.append(m)
    if not segs:
        return [], []
    fusionar_ok = bool(fusionar_colineales)
    if use_capas and n_capas_req > 1 and solo_indice_capa is None:
        fusionar_ok = False
    if not fusionar_ok:
        return segs, seg_metas
    grupos_idx = agrupar_indices_colineales(segs)
    salida = []
    for idxs in grupos_idx:
        grupo = [segs[k] for k in idxs]
        uni = fusionar_lineas_colineales(grupo)
        if uni is not None:
            salida.append(uni)
    return salida, [None] * len(salida)


def estimar_largo_max_mm_eje_columnas_fallback_ubicacion(document, element_ids):
    """
    Largo máximo (mm) fusionando **solo** el eje de ubicación de cada columna (sin corte/plano
    ni offset). Detecta pilares colineales apilados aunque falle la geometría de caras.
    """
    if document is None or not element_ids:
        return None
    cols = filtrar_solo_structural_columns(document, element_ids)
    if not cols:
        return None
    segs = []
    for col in cols:
        crv = _curva_eje_para_proyeccion(col)
        if crv is None or not crv.IsBound:
            continue
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
                continue
            segs.append(Line.CreateBound(p0, p1))
        except Exception:
            continue
    if not segs:
        return None
    grupos_idx = agrupar_indices_colineales(segs)
    max_mm = 0.0
    for idxs in grupos_idx:
        grupo = [segs[k] for k in idxs]
        uni = fusionar_lineas_colineales(grupo)
        if uni is None:
            continue
        try:
            max_mm = max(max_mm, float(uni.Length) * 304.8)
        except Exception:
            try:
                p0 = uni.GetEndPoint(0)
                p1 = uni.GetEndPoint(1)
                max_mm = max(max_mm, float(p0.DistanceTo(p1)) * 304.8)
            except Exception:
                continue
    return max_mm if max_mm > 1e-6 else None


def _estimar_largo_max_mm_bbox_columnas(cols):
    """Altura Z del bbox en mm (respaldo si geometría de eje devuelve poco o nada)."""
    max_mm = 0.0
    for col in cols or []:
        if col is None:
            continue
        try:
            bb = col.get_BoundingBox(None)
            if bb is None:
                continue
            h_ft = abs(float(bb.Max.Z) - float(bb.Min.Z))
            if h_ft < 1e-9:
                continue
            max_mm = max(max_mm, h_ft * 304.8)
        except Exception:
            continue
    return max_mm if max_mm > 1e-6 else None


def estimar_largo_max_mm_eje_columnas_fusionado(
    document, element_ids, offset_mm=0.0, num_curvas_eje=None
):
    """
    Longitud máxima (mm) entre tramos fusionados del eje de columnas, misma construcción que
    ``ejecutar_model_lines_eje_columnas`` (sin extensiones longitudinales ni empotramiento).
    Combina con el eje de **ubicación** (fallback) para no perder el umbral 12 m si el corte
    medio no devuelve líneas pero el pilar es alto o hay apilados colineales.
    ``None`` si no hay columnas estructurales.
    """
    if document is None or not element_ids:
        return None
    cols = filtrar_solo_structural_columns(document, element_ids)
    if not cols:
        return None
    L_fb = estimar_largo_max_mm_eje_columnas_fallback_ubicacion(document, element_ids)
    lineas, _ = lineas_eje_fusionadas_desde_columnas(
        document,
        cols,
        offset_mm=float(offset_mm or 0.0),
        num_curvas_eje=num_curvas_eje,
        diam_nominal_mm=None,
    )
    max_mm = 0.0
    if lineas:
        for ln in lineas:
            if ln is None:
                continue
            try:
                L_ft = float(ln.Length)
                max_mm = max(max_mm, L_ft * 304.8)
            except Exception:
                try:
                    p0 = ln.GetEndPoint(0)
                    p1 = ln.GetEndPoint(1)
                    max_mm = max(max_mm, float(p0.DistanceTo(p1)) * 304.8)
                except Exception:
                    continue
    L_bb = _estimar_largo_max_mm_bbox_columnas(cols)
    L_geom = max_mm if max_mm > 1e-6 else None
    vals = []
    for v in (L_geom, L_fb, L_bb):
        if v is None:
            continue
        try:
            vals.append(float(v))
        except Exception:
            continue
    if not vals:
        return None
    return max(vals)


def _tangente_unitaria_en_extremo_curva(crv, use_start):
    """Tangente unitaria en ``u=0`` o ``u=1`` (``ComputeDerivatives``)."""
    if crv is None or not crv.IsBound:
        return None
    try:
        u = 0.0 if use_start else 1.0
        tr = crv.ComputeDerivatives(u, True)
        v = tr.BasisX
        if v is None or v.GetLength() < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def crear_sketch_planes_extremos_columnas_ordenados(document, columnas):
    """
    Crea un ``SketchPlane`` por cada extremo del eje de cada columna (``LocationCurve`` o línea
    sintética desde ``LocationPoint``). Normal del plano = tangente al eje en ese extremo;
    origen = punto del extremo.

    Todos los planos se ordenan por elevación (coordenada **Z** del origen, luego X, Y) de más
    bajo a más alto y se nombran ``BIMTools_EjeCol_001`` … para referencia en operaciones posteriores.

    Returns:
        Lista de ``ElementId`` en ese orden de numeración (1 = más bajo).
    """
    ids_out = []
    if document is None or not columnas:
        return ids_out
    entries = []
    for col in columnas:
        if col is None:
            continue
        crv = _curva_eje_para_proyeccion(col)
        if crv is None:
            continue
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
        except Exception:
            continue
        n0 = _tangente_unitaria_en_extremo_curva(crv, True)
        n1 = _tangente_unitaria_en_extremo_curva(crv, False)
        if n0 is None or n1 is None:
            try:
                d = p1 - p0
                if d.GetLength() < 1e-12:
                    continue
                du = d.Normalize()
                n0 = n1 = du
            except Exception:
                continue
        entries.append((float(p0.Z), float(p0.X), float(p0.Y), p0, n0))
        entries.append((float(p1.Z), float(p1.X), float(p1.Y), p1, n1))
    entries.sort(key=lambda t: (t[0], t[1], t[2]))
    idx = 0
    for _z, _x, _y, p, n in entries:
        idx += 1
        try:
            pl = Plane.CreateByNormalAndOrigin(n, p)
            sp = SketchPlane.Create(document, pl)
            if sp is None:
                continue
            try:
                sp.Name = u"BIMTools_EjeCol_{0:03d}".format(idx)
            except Exception:
                pass
            ids_out.append(sp.Id)
        except Exception:
            continue
    return ids_out


def crear_sketch_planes_empalme_desde_location_curve(
    document,
    empalme_element_ids,
    crear_marcador_normal_primer_plano=True,
    crear_marcador_normal_cada_plano=False,
):
    """
    Por cada elemento de empalme (viga / columna): ``SketchPlane`` con **origen** en
    ``GetEndPoint(0)`` de la ``LocationCurve`` y **normal** = ``Line.Direction`` si el tramo es
    ``Line`` (vector del eje definido por la API); en otro caso, tangente en el inicio. El plano
    es ⟂ a ese eje y pasa por el arranque del miembro.

    ``crear_marcador_normal_primer_plano``: si ``False``, no se crea la ``ModelLine`` marcador
    (Marca ``N``) del primer plano — p. ej. al colocar solo ``Rebar`` sin líneas de modelo.
    Ignorado si ``crear_marcador_normal_cada_plano`` es ``True``.

    ``crear_marcador_normal_cada_plano``: si ``True``, una ``ModelLine`` por plano desde el origen
    en dirección de la **normal** del plano (flecha de verificación); marcas ``N001``, ``N002``, …

    Returns:
        ``(ids_sketch_planes, planes_empalme_list, ids_marcadores_normal)``: lista de ``Plane``
        (uno por elemento válido, **mismo orden** que los ``SketchPlane`` creados).
    """
    ids_out = []
    planes_list = []
    marker_ids = []
    if document is None or not empalme_element_ids:
        return ids_out, planes_list, marker_ids
    idx_nom = 0
    first_done = False
    for eid in empalme_element_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        crv = _curva_location_miembro_empalme(el)
        if crv is None:
            continue
        origin, n = _origen_y_normal_plano_empalme_desde_location_curve(crv)
        if origin is None or n is None:
            continue
        try:
            pl = Plane.CreateByNormalAndOrigin(n, origin)
            sp = SketchPlane.Create(document, pl)
            if sp is None:
                continue
            planes_list.append(pl)
            if crear_marcador_normal_cada_plano:
                m_id = crear_marcador_normal_plano_empalme(
                    document,
                    origin,
                    n,
                    mark_text=u"N{0:03d}".format(idx_nom + 1),
                )
                if m_id is not None:
                    marker_ids.append(m_id)
            else:
                if not first_done:
                    first_done = True
                    if crear_marcador_normal_primer_plano:
                        m_id = crear_marcador_normal_plano_empalme(document, origin, n)
                        if m_id is not None:
                            marker_ids.append(m_id)
            idx_nom += 1
            try:
                sp.Name = u"BIMTools_Empalme_{0:03d}".format(idx_nom)
            except Exception:
                pass
            ids_out.append(sp.Id)
        except Exception:
            continue
    return ids_out, planes_list, marker_ids


def _indice_marca_perimetral_desde_elemento(el):
    """Primer bloque numérico inicial de **Marca** (``01`` → ``1``); ``None`` si no aplica."""
    if el is None:
        return None
    try:
        pm = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if pm is None:
            return None
        s = pm.AsString()
        if not s:
            s = pm.AsValueString()
        if not s:
            return None
        s = s.strip()
        n = 0
        while n < len(s) and s[n].isdigit():
            n += 1
        if n == 0:
            return None
        return int(s[:n])
    except Exception:
        return None


def _es_model_curve_impar_sin_segmentar(el):
    """``True`` si el tramo es impar operable (comentario base, o marca impar sin troceo previo)."""
    if el is None:
        return False
    try:
        s = u""
        p_c = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p_c is not None:
            raw = p_c.AsString()
            if not raw:
                raw = p_c.AsValueString()
            s = (raw or u"").strip()
        if s:
            if (
                u"BIMTools_EjeCol_ImparMas" in s
                or u"BIMTools_EjeCol_ImparMenos" in s
            ):
                return False
            if (
                u"BIMTools_EjeCol_ParMas" in s
                or u"BIMTools_EjeCol_ParMenos" in s
            ):
                return False
            if u"BIMTools_NormalPlanoEmpalme" in s:
                return False
            if s == u"BIMTools_EjeCol_Par":
                return False
            if s == u"BIMTools_EjeCol_Impar":
                return True
        idx = _indice_marca_perimetral_desde_elemento(el)
        if idx is not None and (idx % 2 == 1):
            if not s:
                return True
            if s == u"BIMTools_EjeCol_Par":
                return False
        return False
    except Exception:
        return False


def _es_model_curve_par_sin_segmentar(el):
    """``True`` si el tramo es par base (antes de ParMas/Menos o troceo por plano desplazado)."""
    if el is None:
        return False
    try:
        s = u""
        p_c = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p_c is not None:
            raw = p_c.AsString()
            if not raw:
                raw = p_c.AsValueString()
            s = (raw or u"").strip()
        if s:
            if (
                u"BIMTools_EjeCol_ParMas" in s
                or u"BIMTools_EjeCol_ParMenos" in s
            ):
                return False
            if u"BIMTools_NormalPlanoEmpalme" in s:
                return False
            if s == u"BIMTools_EjeCol_Impar":
                return False
            if s == u"BIMTools_EjeCol_Par":
                return True
        idx = _indice_marca_perimetral_desde_elemento(el)
        if idx is not None and (idx % 2 == 0):
            if not s:
                return True
            if s == u"BIMTools_EjeCol_Impar":
                return False
        return False
    except Exception:
        return False


def _comentario_lado_respecto_plano(
    ln, plane_origin, plane_normal_unit, es_par=False
):
    """``ImparMas``/``Menos`` o ``ParMas``/``Menos`` según el lado del punto medio del tramo."""
    tol_ft = max(1e-4, _MIN_LINE_LEN_FT * 0.01)
    if es_par:
        tag_base = u"BIMTools_EjeCol_Par"
        tag_mas = u"BIMTools_EjeCol_ParMas"
        tag_menos = u"BIMTools_EjeCol_ParMenos"
    else:
        tag_base = u"BIMTools_EjeCol_Impar"
        tag_mas = u"BIMTools_EjeCol_ImparMas"
        tag_menos = u"BIMTools_EjeCol_ImparMenos"
    if ln is None or plane_origin is None or plane_normal_unit is None:
        return tag_base
    try:
        n = plane_normal_unit.Normalize()
        o = plane_origin
        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
        mid = p0 + (p1 - p0).Multiply(0.5)
        side = float((mid - o).DotProduct(n))
    except Exception:
        return tag_base
    if side > tol_ft:
        return tag_mas
    if side < -tol_ft:
        return tag_menos
    return tag_base


def _origen_y_normal_unidad_plano(plane):
    """``(Origin, normal unit)`` sin depender de asignar ``Normalize()`` a variable en IronPython."""
    if plane is None:
        return None, None
    try:
        o = plane.Origin
        raw = plane.Normal
        Ln = float(raw.GetLength())
        if Ln < 1e-12:
            return None, None
        n = XYZ(float(raw.X) / Ln, float(raw.Y) / Ln, float(raw.Z) / Ln)
        return o, n
    except Exception:
        return None, None


def _sufijo_marca_troceo(indice_tramo):
    """Sufijo estable para Marca tras multicorte: ``a``…``z``, luego ``t0``, ``t1``, …"""
    jj = int(indice_tramo)
    az = u"abcdefghijklmnopqrstuvwxyz"
    if 0 <= jj < len(az):
        return az[jj : jj + 1]
    return u"t{0}".format(jj)


def _cortes_interior_ordenados_multiples_planos(
    p0,
    p1,
    planes_corte,
    planes_impar_ref,
    es_par,
    off_ft,
    min_end,
    stats=None,
):
    """
    Intersecciones interiores del segmento ``p0``–``p1`` con cada plano en ``planes_corte``,
    ordenadas de ``p0`` a ``p1`` y sin duplicados cercanos. Si ``es_par``, corrige cada corte
    frente a ``planes_impar_ref[i]`` cuando aplica.
    """
    if p0 is None or p1 is None or not planes_corte:
        return []
    d = p1 - p0
    L = float(d.GetLength())
    if L < 2.0 * float(min_end) + 1e-12:
        return []
    invL = 1.0 / L
    u = XYZ(float(d.X) * invL, float(d.Y) * invL, float(d.Z) * invL)
    n_pc = len(planes_corte)
    n_ir = len(planes_impar_ref) if planes_impar_ref else 0
    try:
        off_ok = float(off_ft) if off_ft is not None else 0.0
    except Exception:
        off_ok = 0.0
    candidatos = []
    for i in range(n_pc):
        pl = planes_corte[i]
        if pl is None:
            continue
        pt = _punto_troceo_segmento_plano_empalme(p0, p1, pl, min_end)
        if (
            es_par
            and planes_impar_ref
            and i < n_ir
            and off_ok > 1e-12
        ):
            pl_im = planes_impar_ref[i]
            if pl_im is not None:
                pt_im = _punto_troceo_segmento_plano_empalme(
                    p0, p1, pl_im, min_end
                )
                if pt is not None:
                    pt_prev = pt
                    pt = _ajustar_punto_troceo_par_si_coincide_con_impar(
                        p0, p1, pt, pt_im, off_ok, min_end
                    )
                    if (
                        stats is not None
                        and pt is not None
                        and pt_prev is not None
                        and float(pt.DistanceTo(pt_prev))
                        > max(1e-6, float(min_end) * 0.01)
                    ):
                        stats[u"par_troceo_ajustado_tras_coincidir_impar"] += 1
        if pt is None:
            continue
        t = float((pt - p0).DotProduct(u)) / L
        if t * L <= float(min_end) or (1.0 - t) * L <= float(min_end):
            continue
        candidatos.append((t, pt, i))
    candidatos.sort(key=lambda x: x[0])
    tol_t = max(float(min_end) / L, 1e-9)
    out = []
    for t, pt, idx in candidatos:
        if not out:
            out.append((t, pt, idx))
        elif t - out[-1][0] > tol_t:
            out.append((t, pt, idx))
    return out


def _segmentos_y_comentarios_desde_cortes_ordenados(
    p0, p1, cuts_ordered, planes_comentario, es_par
):
    """Parte ``p0``–``p1`` en ``Line`` según cortes; comentarios según plano índice de cada corte."""
    if not cuts_ordered or not planes_comentario:
        return None, None
    lines_out = []
    coms_out = []
    q0 = p0
    n_planes = len(planes_comentario)
    for _j, (_t, pt, pidx) in enumerate(cuts_ordered):
        if pidx < 0 or pidx >= n_planes:
            pl_ref = planes_comentario[0]
        else:
            pl_ref = planes_comentario[pidx]
        o, n = _origen_y_normal_unidad_plano(pl_ref)
        try:
            ln = Line.CreateBound(q0, pt)
        except Exception:
            return None, None
        if float(ln.Length) < _MIN_LINE_LEN_FT:
            return None, None
        com = _comentario_lado_respecto_plano(ln, o, n, es_par=es_par)
        lines_out.append(ln)
        coms_out.append(com)
        q0 = pt
    pidx_last = int(cuts_ordered[-1][2])
    if pidx_last < 0 or pidx_last >= n_planes:
        pl_last = planes_comentario[0]
    else:
        pl_last = planes_comentario[pidx_last]
    o, n = _origen_y_normal_unidad_plano(pl_last)
    try:
        ln_last = Line.CreateBound(q0, p1)
    except Exception:
        return None, None
    if float(ln_last.Length) < _MIN_LINE_LEN_FT:
        return None, None
    com_last = _comentario_lado_respecto_plano(
        ln_last, o, n, es_par=es_par
    )
    lines_out.append(ln_last)
    coms_out.append(com_last)
    return lines_out, coms_out


def _aplicar_estiron_cadena_tramos_salvo_ultimo(segments, diam_mm, stats=None):
    """Estira cada tramo hacia el siguiente; el último no se modifica."""
    if not segments or len(segments) < 2:
        return segments
    out = []
    for i, seg in enumerate(segments):
        if i < len(segments) - 1:
            try:
                antes = float(seg.Length)
            except Exception:
                antes = 0.0
            nxt = segments[i + 1]
            stretched = _line_primer_tramo_con_estiron_post_troceo(
                seg, nxt, diam_mm
            )
            if stats is not None:
                try:
                    if float(stretched.Length) > antes + 1e-9:
                        stats[u"estiron_post_troceo_prim_tramo"] += 1
                except Exception:
                    pass
            out.append(stretched)
        else:
            out.append(seg)
    return out


def _lineas_y_marcas_previas_troceo_plano_empalme(
    lineas,
    planes_impar_list,
    planes_par_list,
    diam_mm_troceo_par=None,
    mark_index_base_0=0,
    line_metas=None,
):
    """
    Por cada **elemento** elegido en empalmes, un plano divisor (``planes_impar_list`` en orden).
    Los **impares** del croquis cortan contra todos esos planos; los **pares** contra la lista
    trasladada (+ tabla mm). Se ordenan los cortes a lo largo del tramo; pueden generarse más de
    dos ``Line`` por tramo croquis. Marcas ``01a``, ``01b``, ``01c``, … Estirón en todos los
    tramos salvo el último de cada cadena.
    ``mark_index_base_0``: suma al prefijo numérico de Marca (p. ej. encadenar lotes en una misma
    corrida). El criterio **impar/par** del troceo usa solo el índice local ``k`` del lote. En
    ``ejecutar_model_lines_eje_columnas`` con varias capas se pasa ``0`` por anillo para que marcas
    y paridad coincidan en cada anillo.

    Returns:
        ``(lineas_flat, marcas_tuplas, stats, metas_por_segmento)``. La última lista
        es paralela a ``lineas_flat`` (meta del tramo de croquis de origen repetida
        en cada trozo); si ``line_metas`` es ``None``, son ``None`` por elemento.
    """
    flat = []
    marks = []
    if not lineas or not planes_impar_list:
        out_lines = list(lineas or [])
        if line_metas is None:
            mf = [None] * len(out_lines)
        else:
            mf = []
            for i in range(len(out_lines)):
                mf.append(line_metas[i] if i < len(line_metas) else None)
        return out_lines, None, None, mf
    stats = {
        u"lineas_in": 0,
        u"lineas_out": 0,
        u"pares": 0,
        u"impares": 0,
        u"partidos_api": 0,
        u"sin_corte_plano": 0,
        u"trozos_degenerados": 0,
        u"err_extremos": 0,
        u"err_create_bound": 0,
        u"partidos_api_par": 0,
        u"sin_corte_plano_par": 0,
        u"trozos_degenerados_par": 0,
        u"err_extremos_par": 0,
        u"err_create_bound_par": 0,
        u"par_troceo_ajustado_tras_coincidir_impar": 0,
        u"estiron_post_troceo_prim_tramo": 0,
        u"mm_estiron_ref": 0.0,
        u"n_planos_empalme": len(planes_impar_list),
    }
    try:
        stats[u"mm_estiron_ref"] = float(
            _mm_estiron_post_troceo_linea_por_diametro(diam_mm_troceo_par)
        )
    except Exception:
        stats[u"mm_estiron_ref"] = 0.0
    for _ln in lineas:
        if _ln is not None:
            stats[u"lineas_in"] += 1
    plist_par = planes_par_list
    if not plist_par or len(plist_par) != len(planes_impar_list):
        d_mm = _mm_desplazamiento_plano_par_tabla_empalme(diam_mm_troceo_par)
        dft = _mm_to_ft(float(d_mm))
        plist_par = []
        for p in planes_impar_list:
            pq = _plano_desplazado_seg_normal(p, dft)
            plist_par.append(pq if pq is not None else p)
    d_mm = _mm_desplazamiento_plano_par_tabla_empalme(diam_mm_troceo_par)
    off_ft = _mm_to_ft(float(d_mm))
    min_end = max(_MIN_LINE_LEN_FT, 1.0 / 304.8)
    try:
        mib = int(mark_index_base_0)
    except Exception:
        mib = 0
    k = 0
    metas_flat = []
    for idx, ln in enumerate(lineas):
        if ln is None:
            continue
        k += 1
        src_meta = None
        if line_metas is not None and idx < len(line_metas):
            src_meta = line_metas[idx]
        gr = u"{0:02d}".format(k + mib)
        es_par_croquis = k % 2 == 0
        if es_par_croquis:
            stats[u"pares"] += 1
            if not plist_par:
                flat.append(ln)
                marks.append((gr, u"BIMTools_EjeCol_Par"))
                metas_flat.append(src_meta)
                continue
            try:
                p0 = ln.GetEndPoint(0)
                p1 = ln.GetEndPoint(1)
            except Exception:
                stats[u"err_extremos_par"] += 1
                flat.append(ln)
                marks.append((gr, u"BIMTools_EjeCol_Par"))
                metas_flat.append(src_meta)
                continue
            cuts = _cortes_interior_ordenados_multiples_planos(
                p0,
                p1,
                plist_par,
                planes_impar_list,
                True,
                off_ft,
                min_end,
                stats,
            )
            if not cuts:
                stats[u"sin_corte_plano_par"] += 1
                po, nu = _origen_y_normal_unidad_plano(plist_par[0])
                com = _comentario_lado_respecto_plano(
                    Line.CreateBound(p0, p1),
                    po,
                    nu,
                    es_par=True,
                )
                flat.append(ln)
                marks.append((gr, com))
                metas_flat.append(src_meta)
                continue
            segs, coms = _segmentos_y_comentarios_desde_cortes_ordenados(
                p0, p1, cuts, plist_par, True
            )
            if segs is None:
                stats[u"trozos_degenerados_par"] += 1
                po, nu = _origen_y_normal_unidad_plano(plist_par[0])
                com = _comentario_lado_respecto_plano(
                    Line.CreateBound(p0, p1),
                    po,
                    nu,
                    es_par=True,
                )
                flat.append(ln)
                marks.append((gr, com))
                metas_flat.append(src_meta)
                continue
            segs = _aplicar_estiron_cadena_tramos_salvo_ultimo(
                segs, diam_mm_troceo_par, stats
            )
            stats[u"partidos_api_par"] += 1
            for j, (ln_s, com) in enumerate(zip(segs, coms)):
                flat.append(ln_s)
                marks.append((gr + _sufijo_marca_troceo(j), com))
                metas_flat.append(src_meta)
            continue
        stats[u"impares"] += 1
        try:
            p0 = ln.GetEndPoint(0)
            p1 = ln.GetEndPoint(1)
        except Exception:
            stats[u"err_extremos"] += 1
            flat.append(ln)
            marks.append((gr, u"BIMTools_EjeCol_Impar"))
            metas_flat.append(src_meta)
            continue
        cuts = _cortes_interior_ordenados_multiples_planos(
            p0,
            p1,
            planes_impar_list,
            None,
            False,
            0.0,
            min_end,
            None,
        )
        if not cuts:
            stats[u"sin_corte_plano"] += 1
            po, nu = _origen_y_normal_unidad_plano(planes_impar_list[0])
            com = _comentario_lado_respecto_plano(
                Line.CreateBound(p0, p1),
                po,
                nu,
                es_par=False,
            )
            flat.append(ln)
            marks.append((gr, com))
            metas_flat.append(src_meta)
            continue
        segs, coms = _segmentos_y_comentarios_desde_cortes_ordenados(
            p0, p1, cuts, planes_impar_list, False
        )
        if segs is None:
            stats[u"trozos_degenerados"] += 1
            po, nu = _origen_y_normal_unidad_plano(planes_impar_list[0])
            com = _comentario_lado_respecto_plano(
                Line.CreateBound(p0, p1),
                po,
                nu,
                es_par=False,
            )
            flat.append(ln)
            marks.append((gr, com))
            metas_flat.append(src_meta)
            continue
        segs = _aplicar_estiron_cadena_tramos_salvo_ultimo(
            segs, diam_mm_troceo_par, stats
        )
        stats[u"partidos_api"] += 1
        for j, (ln_s, com) in enumerate(zip(segs, coms)):
            flat.append(ln_s)
            marks.append((gr + _sufijo_marca_troceo(j), com))
            metas_flat.append(src_meta)
    stats[u"lineas_out"] = len(flat)
    return flat, marks, stats, metas_flat


def _mensaje_diagnostico_troceo_line_api(
    stats, n_model_ok, n_model_fall, uso_troceo_post_modelcurve
):
    """
    Texto para barra de estado: contrasta troceo a nivel ``Line`` (API) vs fallos al crear
    ``ModelCurve`` (curva / SketchPlane / proyección).
    """
    if not stats:
        if uso_troceo_post_modelcurve:
            return (
                u"[Troceo] Solo troceo sobre ModelCurve (no hubo troceo previo de Line API). "
                u"ModelCurve: {0} creadas, {1} fallidas.".format(n_model_ok, n_model_fall)
            )
        return u""
    s = stats
    troceo_ok = int(s.get(u"partidos_api", 0) or 0) > 0
    frag = [
        u"[Troceo Line API]",
        u"croquis {0} tramo(s) → {1} objetos Line listos para ModelCurve.".format(
            s.get(u"lineas_in", 0),
            s.get(u"lineas_out", 0),
        ),
        u"Impares: {0}; croquis con troceo (≥1 corte): {1}.".format(
            s.get(u"impares", 0),
            s.get(u"partidos_api", 0),
        ),
        u"Pares (pos. croquis): {0}; croquis con troceo: {1}.".format(
            s.get(u"pares", 0),
            s.get(u"partidos_api_par", 0),
        ),
    ]
    try:
        npl = int(s.get(u"n_planos_empalme", 0) or 0)
    except Exception:
        npl = 0
    if npl > 0:
        frag.append(u"Planos divisores empalme: {0}.".format(npl))
    det = []
    if int(s.get(u"sin_corte_plano", 0) or 0) > 0:
        det.append(
            u"sin punto de corte interior con el plano: {0}".format(
                s[u"sin_corte_plano"]
            )
        )
    if int(s.get(u"trozos_degenerados", 0) or 0) > 0:
        det.append(
            u"tramos cortos tras corte (< umbral): {0}".format(
                s[u"trozos_degenerados"]
            )
        )
    if int(s.get(u"err_extremos", 0) or 0) > 0:
        det.append(
            u"error leyendo GetEndPoint: {0}".format(s[u"err_extremos"])
        )
    if int(s.get(u"err_create_bound", 0) or 0) > 0:
        det.append(
            u"error Line.CreateBound: {0}".format(s[u"err_create_bound"])
        )
    if int(s.get(u"sin_corte_plano_par", 0) or 0) > 0:
        det.append(
            u"par: sin corte plano: {0}".format(s[u"sin_corte_plano_par"])
        )
    if int(s.get(u"trozos_degenerados_par", 0) or 0) > 0:
        det.append(
            u"par: tramos cortos: {0}".format(s[u"trozos_degenerados_par"])
        )
    if int(s.get(u"par_troceo_ajustado_tras_coincidir_impar", 0) or 0) > 0:
        det.append(
            u"par: troceo corr. (mismo nivel que impar) ±Ø eje: {0}".format(
                s[u"par_troceo_ajustado_tras_coincidir_impar"]
            )
        )
    n_es = int(s.get(u"estiron_post_troceo_prim_tramo", 0) or 0)
    try:
        mm_rf = float(s.get(u"mm_estiron_ref", 0.0) or 0.0)
    except Exception:
        mm_rf = 0.0
    det.append(
        u"estiron (salvo últ. tramo/croquis): {0} (~{1:.0f} mm tabla)".format(
            n_es, mm_rf
        )
    )
    if det:
        frag.append(u"— " + u"; ".join(det) + u".")
    imp = int(s.get(u"impares", 0) or 0)
    part = int(s.get(u"partidos_api", 0) or 0)
    if troceo_ok and part == imp and imp > 0:
        frag.append(
            u"Conclusión: todos los impares tuvieron troceo; si alguna ModelLine no se partió, "
            u"revisa fallidas (proyección/SketchPlane)."
        )
    elif part == 0 and imp > 0:
        frag.append(
            u"Conclusión: ningún impar se partió a nivel Line; el plano no corta los tramos o "
            u"cuentan los ítems anteriores."
        )
    frag.append(
        u"ModelCurve: {0} creadas, {1} fallidas (SketchPlane / proyección / API).".format(
            n_model_ok,
            n_model_fall,
        )
    )
    return u" ".join(frag)


def _motivo_troceo_line_no_stats(
    n_columnas,
    n_empalme_ids,
    hay_plano_empalme,
    troceo_line_omitido_origen,
    uso_troceo_post_mc,
):
    """Explica por qué no hay ``stats`` de troceo Line (texto para TaskDialog)."""
    if troceo_line_omitido_origen:
        return u"Plano de empalme sin Origin/Normal válidos."
    if uso_troceo_post_mc:
        return u"No se usó troceo previo de Line; solo troceo sobre ModelCurve."
    if n_empalme_ids < 1:
        return u"No hay elementos de empalme elegidos (botón vigas/columnas empalme)."
    if not hay_plano_empalme:
        return u"No se pudo obtener plano divisor desde los elementos de empalme."
    return (
        u"No se aplicó troceo previo (condición no cumplida). "
        u"Columnas: {0}; IDs empalme: {1}.".format(n_columnas, n_empalme_ids)
    )


def _task_dialog_resumen_troceo_line(
    stats_troceo_line,
    n_tramos_croquis,
    n_model_curves,
    n_fallidas,
    troceo_line_omitido_origen,
    uso_troceo_post_mc,
    n_columnas,
    n_empalme_ids,
    hay_plano_empalme,
    stats_prep_croquis=None,
    rebar_errores_muestra=None,
):
    """Cuadro modal con el conteo de ``Line`` tras el troceo (visible aunque falle la barra de estado)."""
    try:
        from Autodesk.Revit.UI import TaskDialog
    except Exception:
        return
    if stats_troceo_line is not None:
        ln_out = stats_troceo_line.get(u"lineas_out", 0)
        ln_in = stats_troceo_line.get(u"lineas_in", n_tramos_croquis)
        line_tras = u"{0}".format(ln_out)
    else:
        ln_in = n_tramos_croquis
        line_tras = u"n/a (sin troceo previo; se usan {0} Line igual que el croquis)".format(
            n_tramos_croquis
        )
    lines = [
        u"Conteo tras troceo (Line croquis → Structural Rebar):",
        u"",
        u"• Line listas tras troceo: {0}".format(line_tras),
        u"• Tramos Line del croquis (entrada): {0} (lineas_in en stats: {1}).".format(
            n_tramos_croquis,
            ln_in,
        ),
    ]
    if stats_prep_croquis:
        try:
            te = int(stats_prep_croquis.get(u"tramos_entrada") or 0)
            nn = int(stats_prep_croquis.get(u"omitidos_none") or 0)
            dpos = int(stats_prep_croquis.get(u"dup_misma_posicion") or 0)
            dmid = int(stats_prep_croquis.get(u"dup_punto_medio") or 0)
            dsub = int(stats_prep_croquis.get(u"dup_subsegmento_contenido") or 0)
            ts = int(stats_prep_croquis.get(u"tramos_salida") or 0)
        except Exception:
            te = nn = dpos = dmid = dsub = ts = 0
        lines.append(
            u"• Croquis previo troceo (todos los lotes/capas, una sola evaluación): "
            u"entraron {0} tramo(s)".format(te)
            + (u" ({0} None omit.)".format(nn) if nn else u"")
            + u"; misma posición: {0}; punto medio: {1}; subseg. contenido: {2}; quedan {3}.".format(
                dpos, dmid, dsub, ts
            )
        )
        if dpos == 0 and dmid == 0 and dsub == 0:
            lines.append(
                u"  (Sin duplicados geométricos detectados en ese paso; revisa marcas/vista si una "
                u"barra parece fuera de lugar.)"
            )
    lines.append(
        u"• Columnas en selección: {0}. IDs empalme: {1}. Plano empalme: {2}.".format(
            n_columnas,
            n_empalme_ids,
            u"sí" if hay_plano_empalme else u"no",
        )
    )
    if stats_troceo_line is not None:
        st = stats_troceo_line
        try:
            npl_e = int(st.get(u"n_planos_empalme", 0) or 0)
        except Exception:
            npl_e = 0
        if npl_e > 0:
            lines.append(
                u"• Planos divisores empalme (cortes posibles por tramo): {0}.".format(
                    npl_e
                )
            )
        lines.append(
            u"• Impares con troceo (≥1 corte): {0} de {1}.".format(
                st.get(u"partidos_api", 0),
                st.get(u"impares", 0),
            )
        )
        lines.append(
            u"• Pares con troceo: {0} de {1} pares.".format(
                st.get(u"partidos_api_par", 0),
                st.get(u"pares", 0),
            )
        )
        if int(st.get(u"sin_corte_plano", 0) or 0) > 0:
            lines.append(
                u"• Sin punto de corte con el plano: {0}.".format(
                    st[u"sin_corte_plano"]
                )
            )
        if int(st.get(u"trozos_degenerados", 0) or 0) > 0:
            lines.append(
                u"• Tramos cortos tras corte: {0}.".format(st[u"trozos_degenerados"])
            )
        n_est = int(st.get(u"estiron_post_troceo_prim_tramo", 0) or 0)
        mm_ref = st.get(u"mm_estiron_ref", 0.0)
        try:
            mm_ref = float(mm_ref)
        except Exception:
            mm_ref = 0.0
        lines.append(
            u"• Estirón en tramos (salvo el último de cada croquis): {0}; "
            u"mm ref. ≈ {1:.0f} (tabla Ø→anclaje).".format(n_est, mm_ref)
        )
    else:
        lines.append(
            u"• Troceo previo Line API: no ejecutado. "
            + _motivo_troceo_line_no_stats(
                n_columnas,
                n_empalme_ids,
                hay_plano_empalme,
                troceo_line_omitido_origen,
                uso_troceo_post_mc,
            )
        )
    lines.extend(
        [
            u"",
            u"Structural Rebar creadas: {0}; fallidas: {1}.".format(
                n_model_curves,
                n_fallidas,
            ),
        ]
    )
    if rebar_errores_muestra:
        lines.append(u"")
        lines.append(u"Muestra de errores CreateFromCurves / constructor:")
        for tx in rebar_errores_muestra[:6]:
            if tx:
                lines.append(u"  • {0}".format(tx))
    try:
        TaskDialog.Show(
            u"BIMTools — Troceo Line (eje columnas)",
            u"\n".join(lines),
        )
    except Exception:
        pass


def _crear_model_curve_desde_linea_en_documento(document, ln):
    """Misma proyección que ``crear_model_lines_desde_lineas`` para un tramo."""
    if document is None or ln is None:
        return None
    sp = sketch_plane_para_linea(document, ln)
    if sp is None:
        return None
    ln_use = ln
    try:
        pl = sp.GetPlane()
        if pl is not None:
            p0 = _project_point_to_plane(ln.GetEndPoint(0), pl)
            p1 = _project_point_to_plane(ln.GetEndPoint(1), pl)
            if (
                p0 is not None
                and p1 is not None
                and p0.DistanceTo(p1) >= _MIN_LINE_LEN_FT
            ):
                ln_use = Line.CreateBound(p0, p1)
    except Exception:
        pass
    try:
        mc = document.Create.NewModelCurve(ln_use, sp)
        return mc.Id if mc is not None else None
    except Exception:
        return None


def _aplicar_marca_y_comentario_model_curve(document, eid, mark_str, comentario):
    if document is None or eid is None:
        return
    try:
        el = document.GetElement(eid)
    except Exception:
        el = None
    if el is None:
        return
    try:
        p_m = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if p_m is not None and not p_m.IsReadOnly and mark_str:
            p_m.Set(mark_str)
    except Exception:
        pass
    try:
        p_c = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p_c is not None and not p_c.IsReadOnly and comentario:
            p_c.Set(comentario)
    except Exception:
        pass


def _segmentar_model_curves_por_planos_lista_empalme(
    document,
    model_curve_ids,
    planes_corte,
    planes_impar_ref,
    es_par,
    off_ft,
    diam_nominal_estiron_post_troceo_mm=None,
):
    """
    Trocea cada ``ModelCurve`` elegible contra **todos** los planos en ``planes_corte``
    (ordenados a lo largo del tramo). Con ``es_par``, ``planes_impar_ref`` y ``off_ft`` alimentan
    la corrección par/impar por índice. Estirón en cadena (todos los tramos salvo el último).
    """
    if document is None or not model_curve_ids:
        return list(model_curve_ids or [])
    if not planes_corte:
        return list(model_curve_ids)
    min_end = max(_MIN_LINE_LEN_FT, 1.0 / 304.8)
    try:
        off_use = float(off_ft) if off_ft is not None else 0.0
    except Exception:
        off_use = 0.0
    out_ids = []
    for eid in model_curve_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if not isinstance(el, CurveElement):
            try:
                if getattr(el, "GeometryCurve", None) is None:
                    out_ids.append(eid)
                    continue
            except Exception:
                out_ids.append(eid)
                continue
        if es_par:
            if not _es_model_curve_par_sin_segmentar(el):
                out_ids.append(eid)
                continue
        else:
            if not _es_model_curve_impar_sin_segmentar(el):
                out_ids.append(eid)
                continue
        crv = el.GeometryCurve
        if crv is None or not crv.IsBound:
            out_ids.append(eid)
            continue
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
        except Exception:
            out_ids.append(eid)
            continue
        mark_base = u""
        try:
            pm = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if pm is not None and pm.AsString():
                mark_base = pm.AsString().strip()
        except Exception:
            pass
        if not mark_base:
            mark_base = u"?"
        po0, nu0 = _origen_y_normal_unidad_plano(planes_corte[0])
        cuts = _cortes_interior_ordenados_multiples_planos(
            p0,
            p1,
            planes_corte,
            planes_impar_ref if planes_impar_ref else None,
            bool(es_par),
            off_use,
            min_end,
            None,
        )
        if not cuts:
            com = _comentario_lado_respecto_plano(
                Line.CreateBound(p0, p1),
                po0,
                nu0,
                es_par=es_par,
            )
            _aplicar_marca_y_comentario_model_curve(document, eid, mark_base, com)
            out_ids.append(eid)
            continue
        segs, coms = _segmentos_y_comentarios_desde_cortes_ordenados(
            p0, p1, cuts, planes_corte, es_par
        )
        if segs is None or coms is None or len(segs) != len(coms):
            com = _comentario_lado_respecto_plano(
                Line.CreateBound(p0, p1),
                po0,
                nu0,
                es_par=es_par,
            )
            _aplicar_marca_y_comentario_model_curve(document, eid, mark_base, com)
            out_ids.append(eid)
            continue
        segs = _aplicar_estiron_cadena_tramos_salvo_ultimo(
            segs, diam_nominal_estiron_post_troceo_mm, None
        )
        created = []
        ok = True
        for ln_s in segs:
            idn = _crear_model_curve_desde_linea_en_documento(document, ln_s)
            if idn is None:
                ok = False
                break
            created.append(idn)
        if not ok or len(created) != len(coms):
            for cid in created:
                try:
                    document.Delete(cid)
                except Exception:
                    pass
            out_ids.append(eid)
            continue
        try:
            document.Delete(eid)
        except Exception:
            pass
        for j, (idn, com) in enumerate(zip(created, coms)):
            _aplicar_marca_y_comentario_model_curve(
                document,
                idn,
                u"{0}{1}".format(mark_base, _sufijo_marca_troceo(j)),
                com,
            )
            out_ids.append(idn)
    return out_ids


def segmentar_model_curves_impares_con_plano_empalme(
    document, model_curve_ids, planes_empalme, diam_nominal_estiron_mm=None
):
    """
    **Divide** cada ``ModelCurve`` ``Impar`` contra **todos** los planos de empalme
    (lista o un solo ``Plane``).

    **Importante:** no aplicar después empotramiento ni fusión; el estirón post-troceo usa la tabla mm/Ø.
    """
    planes_list = (
        planes_empalme
        if isinstance(planes_empalme, list)
        else ([planes_empalme] if planes_empalme is not None else [])
    )
    return _segmentar_model_curves_por_planos_lista_empalme(
        document,
        model_curve_ids,
        planes_list,
        [],
        False,
        0.0,
        diam_nominal_estiron_mm,
    )


def segmentar_model_curves_pares_plano_desplazado_diametro(
    document,
    model_curve_ids,
    planes_par_list,
    planes_impar_list,
    diam_nominal_mm=None,
):
    """
    **Divide** cada ``ModelCurve`` ``Par`` con los planos trasladados (tabla mm), uno por empalme.
    ``planes_impar_list`` debe tener la misma longitud que ``planes_par_list`` para la corrección
    índice a índice.
    """
    pl_par = (
        planes_par_list
        if isinstance(planes_par_list, list)
        else ([planes_par_list] if planes_par_list is not None else [])
    )
    pl_imp = (
        planes_impar_list
        if isinstance(planes_impar_list, list)
        else ([planes_impar_list] if planes_impar_list is not None else [])
    )
    d_ft = _mm_to_ft(_mm_desplazamiento_plano_par_tabla_empalme(diam_nominal_mm))
    return _segmentar_model_curves_por_planos_lista_empalme(
        document,
        model_curve_ids,
        pl_par,
        pl_imp,
        True,
        d_ft,
        diam_nominal_mm,
    )


def _normal_unidad_plano_soporte_para_linea(linea):
    """
    Vector unitario **normal al plano** que contiene la recta (misma elección que ``SketchPlane`` del
    eje: ⟂ a la tangente; pilar vertical → normal mayormente horizontal).
    Simboliza la orientación del croquis de la ``ModelCurve``.
    """
    if linea is None:
        return None
    try:
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < _MIN_LINE_LEN_FT:
            return None
        d = d.Normalize()
        if abs(float(d.DotProduct(XYZ.BasisZ))) < 0.99:
            n = d.CrossProduct(XYZ.BasisZ).Normalize()
        else:
            n = d.CrossProduct(XYZ.BasisX).Normalize()
        if n.GetLength() < 1e-12:
            return None
        return n.Normalize()
    except Exception:
        return None


def sketch_plane_para_linea(document, linea):
    """Plano que contiene la recta (normal ⟂ al tramo)."""
    if document is None or linea is None:
        return None
    try:
        p0 = linea.GetEndPoint(0)
        n = _normal_unidad_plano_soporte_para_linea(linea)
        if n is None:
            return None
        pl = Plane.CreateByNormalAndOrigin(n, p0)
        return SketchPlane.Create(document, pl)
    except Exception:
        return None


def crear_marcador_normal_plano_empalme(
    document, origin, normal_unit, length_mm=500.0, mark_text=None
):
    """
    ``ModelLine`` corta desde ``origin`` en la dirección del vector normal del plano de empalme
    (flecha de referencia para el troceo). Longitud típica 500 mm.

    ``mark_text``: texto de **Marca** del elemento; por defecto ``N``.
    """
    if document is None or origin is None or normal_unit is None:
        return None
    try:
        n = normal_unit.Normalize()
        if n is None or n.GetLength() < 1e-12:
            return None
        L_ft = _mm_to_ft(float(length_mm))
        if L_ft < _MIN_LINE_LEN_FT * 2.0:
            L_ft = _MIN_LINE_LEN_FT * 4.0
        p_end = origin + n.Multiply(L_ft)
        ln = Line.CreateBound(origin, p_end)
        sp = sketch_plane_para_linea(document, ln)
        if sp is None:
            return None
        ln_use = ln
        try:
            pl = sp.GetPlane()
            if pl is not None:
                q0 = _project_point_to_plane(ln.GetEndPoint(0), pl)
                q1 = _project_point_to_plane(ln.GetEndPoint(1), pl)
                if (
                    q0 is not None
                    and q1 is not None
                    and q0.DistanceTo(q1) >= _MIN_LINE_LEN_FT
                ):
                    ln_use = Line.CreateBound(q0, q1)
        except Exception:
            pass
        mc = document.Create.NewModelCurve(ln_use, sp)
        if mc is None:
            return None
        _mark = mark_text if mark_text else u"N"
        try:
            p_m = mc.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p_m is not None and not p_m.IsReadOnly:
                p_m.Set(_mark)
        except Exception:
            pass
        try:
            p_c = mc.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p_c is not None and not p_c.IsReadOnly:
                p_c.Set(
                    u"BIMTools_NormalPlanoEmpalme | n=({0:.3f},{1:.3f},{2:.3f})".format(
                        n.X, n.Y, n.Z
                    )
                )
        except Exception:
            pass
        try:
            mc.Name = (
                u"BIMTools_NormalEmpalme_{0}".format(mark_text)
                if mark_text
                else u"BIMTools_NormalEmpalme"
            )
        except Exception:
            pass
        return mc.Id
    except Exception:
        return None


def crear_marcador_normal_curva_eje(
    document, linea, length_mm=None
):
    """
    ``ModelLine`` corto desde el **punto medio** de ``linea`` en dirección de la **normal del plano
    soporte** de la curva (igual que el ``SketchPlane`` de la barra de eje). Marca ``n`` (minúscula;
    distinta de ``N`` del plano de empalme).
    """
    if document is None or linea is None:
        return None
    try:
        if length_mm is None:
            length_mm = float(_MARCADOR_NORMAL_CURVA_EJE_MM)
        else:
            length_mm = float(length_mm)
    except Exception:
        length_mm = float(_MARCADOR_NORMAL_CURVA_EJE_MM)
    try:
        n = _normal_unidad_plano_soporte_para_linea(linea)
        if n is None:
            return None
        p0 = linea.GetEndPoint(0)
        p1 = linea.GetEndPoint(1)
        origin = p0 + (p1 - p0).Multiply(0.5)
        n = n.Normalize()
        if n is None or n.GetLength() < 1e-12:
            return None
        L_ft = _mm_to_ft(length_mm)
        if L_ft < _MIN_LINE_LEN_FT * 2.0:
            L_ft = _MIN_LINE_LEN_FT * 4.0
        p_end = origin + n.Multiply(L_ft)
        ln = Line.CreateBound(origin, p_end)
        sp = sketch_plane_para_linea(document, ln)
        if sp is None:
            return None
        ln_use = ln
        try:
            pl = sp.GetPlane()
            if pl is not None:
                q0 = _project_point_to_plane(ln.GetEndPoint(0), pl)
                q1 = _project_point_to_plane(ln.GetEndPoint(1), pl)
                if (
                    q0 is not None
                    and q1 is not None
                    and q0.DistanceTo(q1) >= _MIN_LINE_LEN_FT
                ):
                    ln_use = Line.CreateBound(q0, q1)
        except Exception:
            pass
        mc = document.Create.NewModelCurve(ln_use, sp)
        if mc is None:
            return None
        try:
            p_m = mc.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p_m is not None and not p_m.IsReadOnly:
                p_m.Set(u"n")
        except Exception:
            pass
        try:
            p_c = mc.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p_c is not None and not p_c.IsReadOnly:
                p_c.Set(u"BIMTools_NormalCurvaEje")
        except Exception:
            pass
        try:
            mc.Name = u"BIMTools_NormalCurvaEje"
        except Exception:
            pass
        return mc.Id
    except Exception:
        return None


def crear_marcadores_normal_curvas_eje_desde_ids(
    document, model_curve_ids, length_mm=None
):
    """
    Un marcador por cada ``ModelCurve`` en ``model_curve_ids`` cuya geometría sea ``Line``.
    """
    out = []
    if document is None or not model_curve_ids:
        return out
    for eid in model_curve_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            continue
        if el is None:
            continue
        try:
            crv = el.GeometryCurve
        except Exception:
            crv = None
        if crv is None or not crv.IsBound:
            continue
        if not isinstance(crv, Line):
            continue
        mid = crear_marcador_normal_curva_eje(document, crv, length_mm=length_mm)
        if mid is not None:
            out.append(mid)
    return out


def etiquetar_model_curves_indice_y_paridad(
    document,
    element_ids,
    comentario_fijo_por_indice=None,
    indice_marca_base_0=0,
):
    """
    Asigna **Marca** ``01``… ``N`` (orden de creación; con un solo pilar suele coincidir con el
    índice perimetral del croquis) y **Comentarios** ``BIMTools_EjeCol_Impar`` / ``Par`` según
    posición en ese lote (1,3,5… vs 2,4,6…).
    Con varias columnas en modelo, el orden del lote es el de generación. **Igual que en una
    sola capa:** solo se llama desde ``ejecutar_model_lines_eje_columnas`` cuando hay **una**
    columna en la selección. Con **multicapa** y un único pilar, cada anillo usa
    ``indice_marca_base_0 = 0`` (marcas ``01``… reiniciadas por capa).
    Tras troceo, comentarios ``ImparMas``/``Menos`` o ``ParMas``/``Menos`` y marcas ``NNa``/``NNb``.
    ``comentario_fijo_por_indice``: si la entrada ``k`` no es vacía, se usa como **Comentarios**
    en lugar de impar/par (p. ej. barras en cadena 4·m).
    ``indice_marca_base_0``: desplazamiento del número de **Marca** (01…) al etiquetar varios lotes;
    impar/par de comentario sigue el índice local del lote.
    Ignora parámetros de solo lectura.
    """
    if document is None or not element_ids:
        return
    try:
        ibase = int(indice_marca_base_0)
    except Exception:
        ibase = 0
    for k, eid in enumerate(element_ids):
        idx1 = k + 1 + ibase
        if (
            comentario_fijo_por_indice is not None
            and k < len(comentario_fijo_por_indice)
            and comentario_fijo_por_indice[k]
        ):
            grupo = comentario_fijo_por_indice[k]
        else:
            grupo = (
                u"BIMTools_EjeCol_Impar"
                if ((k + 1) % 2 == 1)
                else u"BIMTools_EjeCol_Par"
            )
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            p_m = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            if p_m is not None and not p_m.IsReadOnly:
                p_m.Set(u"{0:02d}".format(idx1))
        except Exception:
            pass
        try:
            p_c = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if p_c is not None and not p_c.IsReadOnly:
                p_c.Set(grupo)
        except Exception:
            pass


def crear_model_lines_desde_lineas(document, lineas):
    """
    Crea una ``ModelCurve`` por cada ``Line``. Los extremos se proyectan al plano del
    ``SketchPlane`` para que Revit acepte la curva (evita fallos por tolerancia).
    """
    creados = []
    if document is None or not lineas:
        return creados
    for ln in lineas:
        eid = _crear_model_curve_desde_linea_en_documento(document, ln)
        if eid is not None:
            creados.append(eid)
    return creados


def _curve_clr_base_rebar():
    """Tipo base CLR ``Curve`` para ``System.Array`` en ``Rebar.CreateFromCurves``."""
    return clr.GetClrType(Line).BaseType


def _vector_normal_plano_rebar_desde_linea_y_meta(ln, meta):
    """
    Normal ``CreateFromCurves`` desde la **meta** de cara (pre-troceo):

    1. Con ``n_ext``: si |u·n_ext| es bajo (barra ~en la cara), **−n_ext** → orientación hacia
       **interior** del pilar (base estable para ganchos hacia el hormigón).
    2. Respaldo: ``face_basis_x`` y ``u × ref_proj``.
    3. Otro ``−n_ext`` si casi ⟂ ``u``; luego ``BasisZ``/``X``.
    """
    if ln is None or not ln.IsBound:
        return None
    try:
        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
        v = p1 - p0
        if v.GetLength() < 1e-12:
            return None
        u = v.Normalize()
    except Exception:
        return None
    n_ext = None
    try:
        if meta and isinstance(meta, dict):
            t = meta.get(u"n_ext")
            if t is not None and len(t) >= 3:
                n_ext = XYZ(float(t[0]), float(t[1]), float(t[2]))
    except Exception:
        n_ext = None
    if n_ext is not None and n_ext.GetLength() > 1e-12:
        try:
            n_ext = n_ext.Normalize()
            ad = abs(float(u.DotProduct(n_ext)))
            if ad < float(_TOL_REBAR_U_DOT_NEXT_PERP_NFACE):
                return n_ext.Negate()
        except Exception:
            pass
    bx = None
    try:
        if meta and isinstance(meta, dict):
            tbx = meta.get(u"face_basis_x")
            if tbx is not None and len(tbx) >= 3:
                bx = XYZ(float(tbx[0]), float(tbx[1]), float(tbx[2]))
    except Exception:
        bx = None
    if bx is not None and bx.GetLength() > 1e-12:
        n = _normal_plano_rebar_createfromcurves_desde_tangente_y_eje_referencia(u, bx)
        if n is not None:
            return n
    if n_ext is not None and n_ext.GetLength() > 1e-12:
        try:
            if abs(float(u.DotProduct(n_ext))) < 0.995:
                return n_ext.Negate()
        except Exception:
            pass
    try:
        cand = u.CrossProduct(XYZ.BasisZ)
        if cand.GetLength() < 1e-8:
            cand = u.CrossProduct(XYZ.BasisX)
        if cand.GetLength() < 1e-8:
            return None
        return cand.Normalize()
    except Exception:
        return None


def _host_columna_para_rebar(document, ln, meta, columnas):
    """
    Host válido para ``Rebar.CreateFromCurves``: prioriza ``meta['host_column_id']``;
    si no, columna cuyo bbox (con holgura) contiene el punto medio del tramo.
    """
    if not columnas:
        return None
    if document is not None and meta and isinstance(meta, dict):
        eid = meta.get(u"host_column_id")
        if eid is not None:
            try:
                el = document.GetElement(eid)
                if el is not None:
                    return el
            except Exception:
                pass
    pm = _punto_medio_linea(ln)
    if pm is None:
        return columnas[0]
    tol = 1.0
    best = None
    best_d = None
    for col in columnas:
        if col is None:
            continue
        try:
            bb = col.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is None:
            continue
        try:
            inside = (
                float(pm.X) >= float(bb.Min.X) - tol
                and float(pm.X) <= float(bb.Max.X) + tol
                and float(pm.Y) >= float(bb.Min.Y) - tol
                and float(pm.Y) <= float(bb.Max.Y) + tol
                and float(pm.Z) >= float(bb.Min.Z) - tol
                and float(pm.Z) <= float(bb.Max.Z) + tol
            )
        except Exception:
            inside = False
        if not inside:
            continue
        try:
            cx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
            cy = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
            cz = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
            d = (float(pm.X) - cx) ** 2 + (float(pm.Y) - cy) ** 2 + (
                float(pm.Z) - cz
            ) ** 2
        except Exception:
            d = 0.0
        if best is None or d < best_d:
            best = col
            best_d = d
    if best is not None:
        return best
    return columnas[0]


def _crear_rebar_eje_desde_linea(
    document,
    host,
    bar_type,
    ln,
    meta=None,
    normal_plano=None,
):
    """
    Una ``Line`` procesada + host → ``Rebar``.
    Si ``normal_plano`` es un ``XYZ`` (normal al plano ``CreateFromCurves``), se usa directamente;
    si no, se calcula desde ``meta``/respaldo.

    Retorna ``(rebar, mensaje_error)``; si OK, ``mensaje_error`` es ``None``.
    """
    import System

    def _ex_txt(ex):
        try:
            return u"{0!s}".format(ex)
        except Exception:
            return u"(sin detalle)"

    if document is None or host is None or bar_type is None or ln is None:
        return None, u"Faltan documento, host, RebarBarType o Line."
    try:
        _cf = Rebar.CreateFromCurves
        if _cf is None:
            return None, u"Rebar.CreateFromCurves no está expuesto en esta API."
    except Exception as ex:
        return None, (
            u"No existe constructor Rebar.CreateFromCurves utilizable: {0}"
        ).format(_ex_txt(ex))
    norm = normal_plano
    if norm is None:
        norm = _vector_normal_plano_rebar_desde_linea_y_meta(ln, meta)
    if norm is None:
        return None, (
            u"Vector normal del plano de barra no calculado (meta face_basis_x del croquis, "
            u"n_ext o tramo degenerado)."
        )
    # Sin ganchos: la API espera RebarHookType o null — no ElementId.InvalidElementId.
    hook_start = None
    hook_end = None
    try:
        ct = _curve_clr_base_rebar()
        arr = System.Array.CreateInstance(ct, 1)
        arr[0] = ln
    except Exception as ex:
        return None, u"Array CLR de curvas: {0}".format(_ex_txt(ex))
    norms = [norm]
    try:
        norms.append(norm.Negate())
    except Exception:
        pass
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    last_ex = None
    for use_existing in (True, False):
        for create_new in (True, False):
            for nvec in norms:
                for so, eo in orient_pairs:
                    try:
                        r = Rebar.CreateFromCurves(
                            document,
                            RebarStyle.Standard,
                            bar_type,
                            hook_start,
                            hook_end,
                            host,
                            nvec,
                            arr,
                            so,
                            eo,
                            use_existing,
                            create_new,
                        )
                        if r:
                            return r, None
                    except Exception as ex:
                        last_ex = ex
                        continue
    if last_ex is not None:
        return None, (
            u"CreateFromCurves rechazó todas las combinaciones. Último error: {0}"
        ).format(_ex_txt(last_ex))
    return None, u"CreateFromCurves: sin excepción pero sin Rebar devuelto (revisar tipo/host)."


def _segmentar_una_linea_troceo_empalme(
    ln,
    planes_corte,
    planes_impar_ref,
    es_par,
    off_ft,
    diam_nominal_estiron_post_troceo_mm,
    mark_base,
):
    """Equivalente Line API a un ciclo de :func:`_segmentar_model_curves_por_planos_lista_empalme`."""
    out = []
    if ln is None or not ln.IsBound:
        return out
    if not planes_corte:
        return [
            (
                ln,
                mark_base,
                u"BIMTools_EjeCol_Par" if es_par else u"BIMTools_EjeCol_Impar",
            )
        ]
    try:
        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
    except Exception:
        return out
    min_end = max(_MIN_LINE_LEN_FT, 1.0 / 304.8)
    try:
        off_use = float(off_ft) if off_ft is not None else 0.0
    except Exception:
        off_use = 0.0
    try:
        po0, nu0 = _origen_y_normal_unidad_plano(planes_corte[0])
    except Exception:
        po0, nu0 = None, None
    if po0 is None or nu0 is None:
        try:
            out.append((ln, mark_base, u"BIMTools_EjeCol_Par" if es_par else u"BIMTools_EjeCol_Impar"))
        except Exception:
            pass
        return out
    cuts = _cortes_interior_ordenados_multiples_planos(
        p0,
        p1,
        planes_corte,
        planes_impar_ref if planes_impar_ref else None,
        bool(es_par),
        off_use,
        min_end,
        None,
    )
    if not cuts:
        com = _comentario_lado_respecto_plano(
            Line.CreateBound(p0, p1),
            po0,
            nu0,
            es_par=es_par,
        )
        out.append((ln, mark_base, com))
        return out
    segs, coms = _segmentos_y_comentarios_desde_cortes_ordenados(
        p0, p1, cuts, planes_corte, es_par
    )
    if segs is None or coms is None or len(segs) != len(coms):
        com = _comentario_lado_respecto_plano(
            Line.CreateBound(p0, p1),
            po0,
            nu0,
            es_par=es_par,
        )
        out.append((ln, mark_base, com))
        return out
    segs = _aplicar_estiron_cadena_tramos_salvo_ultimo(
        segs, diam_nominal_estiron_post_troceo_mm, None
    )
    for j, com in enumerate(coms):
        if j >= len(segs):
            break
        ln_s = segs[j]
        out.append(
            (
                ln_s,
                u"{0}{1}".format(mark_base, _sufijo_marca_troceo(j)),
                com,
            )
        )
    return out


def _lineas_lote_troceo_empalme_dos_pasos(
    lineas_lote,
    planes_emp,
    planes_par,
    diam_mm,
    mark_etiqueta_base_0,
):
    """
    Troceo por planos de empalme solo con ``Line`` (sin ``ModelCurve`` intermedia): primero
    impares, luego pares (igual que el pipeline anterior basado en modelo).
    """
    if not lineas_lote or not planes_emp:
        return []
    try:
        d_ft = _mm_to_ft(_mm_desplazamiento_plano_par_tabla_empalme(diam_mm))
    except Exception:
        d_ft = 0.0
    try:
        mib = int(mark_etiqueta_base_0)
    except Exception:
        mib = 0
    pl_par = planes_par if planes_par else []
    paso1 = []
    idx_line = 0
    for ln in lineas_lote:
        if ln is None:
            continue
        mark_base = u"{0:02d}".format(idx_line + 1 + mib)
        if idx_line % 2 == 0:
            ch = _segmentar_una_linea_troceo_empalme(
                ln,
                planes_emp,
                [],
                False,
                0.0,
                diam_mm,
                mark_base,
            )
            paso1.extend(ch or [(ln, mark_base, u"BIMTools_EjeCol_Impar")])
        else:
            paso1.append((ln, mark_base, u"BIMTools_EjeCol_Par"))
        idx_line += 1
    final = []
    for ln, mark_full, com in paso1:
        if com == u"BIMTools_EjeCol_Par" and pl_par:
            ch = _segmentar_una_linea_troceo_empalme(
                ln,
                pl_par,
                planes_emp,
                True,
                d_ft,
                diam_mm,
                mark_full,
            )
            final.extend(ch or [(ln, mark_full, com)])
        else:
            final.append((ln, mark_full, com))
    return final


def ejecutar_model_lines_eje_columnas(
    document,
    uidocument,
    element_ids,
    offset_mm=0.0,
    num_curvas_eje=None,
    empotramiento_diam_nominal_mm=None,
    diam_estribo_nominal_mm=None,
    empalme_element_ids=None,
    capas_num_curvas=None,
    incremento_offset_capas_mm=None,
    rebar_bar_type=None,
):
    """
    Filtra columnas, fusiona ejes, aplica empotramiento (sonda 500 mm; si hay colisión en la punta,
    estirón tabla anclaje/empalme si colisión de sonda) y crea ``Structural Rebar`` desde las
    curvas procesadas (sin ``ModelLine`` / ``ModelCurve`` de eje ni marcadores normales).

    ``rebar_bar_type``: obligatorio — :class:`RebarBarType` longitudinal (host = primera columna de
    la selección). Sin tipo, el comando no crea elementos.
    ``num_curvas_eje``: total de líneas por columna repartidas en caras laterales; ``None`` = una por cara planar.
    ``capas_num_curvas``: solo ``[0]`` define el número de barras por anillo; el largo de la lista
    es el número de capas. Capas ≥2: mismo algoritmo que la 1.ª con offset normal
    ``25 + Ø estribo + Ø long./2 + i×paso`` con ``i`` = índice de anillo interior (véase
    :func:`_mm_offset_normal_segunda_capa_mm`).
    ``incremento_offset_capas_mm`` se ignora en la geometría del 2.º anillo (API reservada).
    y sin vértices (índice o geométrico).
    ``empotramiento_diam_nominal_mm``: Ø (mm) longitudinal para tabla de estirón (empotramiento y junta
    pilar superior apilado); ``None`` o ≤0 → ``_EMPOTRAMIENTO_DIAM_NOMINAL_TABLA_MM``.

    ``diam_estribo_nominal_mm``: Ø (mm) estribo para la traslación en ancho de cara con **4** barras.
    ``empalme_element_ids``: vigas/columnas elegidas con el botón de empalme; **por cada una**
    se crea un ``SketchPlane`` y un ``Plane`` divisor (origen = ``GetEndPoint(0)``, normal =
    dirección del eje). Todos esos planos cortan cada tramo croquis (impares: base; pares: plano
    trasladado según tabla mm/Ø, índice a índice). Los puntos de corte se ordenan a lo largo del
    tramo → pueden salir **más de dos** ``Line`` por croquis (marcas ``NNa``, ``NNb``, ``NNc``, …).

    Con empalme y **una** columna no se aplica **fusión colineal** entre tramos de eje antes de
    colocar las barras (cada tramo sigue siendo su propia curva hasta ``Rebar``). Tras el troceo
    por plano no hay fusión adicional sobre la geometría.

    Con empalme y **una** columna, las posiciones **impares** (1,3,5,…) se omiten en
    ``aplicar_empotramiento_lineas_unificadas``. Tras el troceo, **todos los tramos salvo el último**
    de cada croquis se alargan (tabla anclaje/desarrollo por Ø, mismos mm que desplaz. plano par).
    No se aplica después otro empotramiento, fusión ni estirón adicional en este comando.

    Sin elementos de empalme, crea ``SketchPlane`` en ambos extremos del eje de cada columna
    seleccionada en modelo (normal = tangente), ordenados por Z: ``BIMTools_EjeCol_NNN``.
    **Con** ``empalme_element_ids`` no vacíos **no** se generan esos planos por columna: los únicos
    ``SketchPlane`` auxiliares del comando son los de empalme (``BIMTools_Empalme_NNN``), que son
    también la referencia del **troceo** (impar/par). **No** se crean ``ModelLine`` de marcador de
    normal en empalme (flujo exclusivo ``Rebar``).

    Retorna ``(n_creados, mensaje, ids_creados, ids_sketch_planes, ids_marcadores_normal)``;
    ``ids_creados`` son ``ElementId`` de ``Rebar``. ``ids_marcadores_normal`` queda vacío (no hay
    marcadores de modelo en este flujo).
    """
    vac = []
    if document is None:
        return 0, u"No hay documento.", vac, vac, vac
    cols = filtrar_solo_structural_columns(document, element_ids or [])
    if not cols:
        return (
            0,
            u"No hay columnas estructurales en la selección; no se creó armadura.",
            vac,
            vac,
            vac,
        )
    if rebar_bar_type is None:
        return (
            0,
            u"No se indicó tipo de barra longitudinal (RebarBarType).",
            vac,
            vac,
            vac,
        )
    try:
        emp_ids = list(empalme_element_ids or [])
    except Exception:
        emp_ids = []
    fusionar_tramos = True
    if emp_ids and len(cols) == 1:
        fusionar_tramos = False
    _cn_capas = _cn_capas_desde_argumento(capas_num_curvas)
    multicapa_varios_anillos = len(_cn_capas) > 1
    lotes_post_empotramiento = []

    if not multicapa_varios_anillos:
        lineas, lineas_metas = lineas_eje_fusionadas_desde_columnas(
            document,
            cols,
            offset_mm=offset_mm,
            num_curvas_eje=num_curvas_eje,
            diam_nominal_mm=empotramiento_diam_nominal_mm,
            diam_estribo_mm=diam_estribo_nominal_mm,
            fusionar_colineales=fusionar_tramos,
            capas_num_curvas=capas_num_curvas,
            incremento_offset_capas_mm=incremento_offset_capas_mm,
            solo_indice_capa=None,
        )
        if not lineas:
            return (
                0,
                u"No se pudo obtener eje (LocationCurve/LocationPoint ni cara) en ninguna columna.",
                vac,
                vac,
                vac,
            )
        skip_emp_impares = None
        if (
            emp_ids
            and len(cols) == 1
            and len(lineas) > 0
        ):
            skip_emp_impares = {j for j in range(len(lineas)) if j % 2 == 0}
        emp_kwargs = dict(
            document=document,
            lineas=lineas,
            ids_seleccion=element_ids or [],
            columnas=cols,
            mm_prueba=_EMPOTRAMIENTO_PRUEBA_MM,
            diam_nominal_mm=empotramiento_diam_nominal_mm,
            skip_indices_0based=skip_emp_impares,
        )
        if lineas_metas is not None and len(lineas_metas) == len(lineas):
            lineas, lineas_metas = aplicar_empotramiento_lineas_unificadas(
                line_metas=lineas_metas, **emp_kwargs
            )
        else:
            lineas = aplicar_empotramiento_lineas_unificadas(**emp_kwargs)
            lineas_metas = None
        if not lineas:
            return (
                0,
                u"Tramo demasiado corto tras evaluar empotramiento.",
                vac,
                vac,
                vac,
            )
        lotes_post_empotramiento = [(lineas, lineas_metas, 0)]
    else:
        for capa_k in range(len(_cn_capas)):
            lineas_k, lineas_metas_k = lineas_eje_fusionadas_desde_columnas(
                document,
                cols,
                offset_mm=offset_mm,
                num_curvas_eje=None,
                diam_nominal_mm=empotramiento_diam_nominal_mm,
                diam_estribo_mm=diam_estribo_nominal_mm,
                fusionar_colineales=fusionar_tramos,
                capas_num_curvas=capas_num_curvas,
                incremento_offset_capas_mm=incremento_offset_capas_mm,
                solo_indice_capa=capa_k,
            )
            if not lineas_k:
                continue
            skip_emp_k = None
            if emp_ids and len(cols) == 1 and len(lineas_k) > 0:
                skip_emp_k = {j for j in range(len(lineas_k)) if j % 2 == 0}
            emp_kwargs_k = dict(
                document=document,
                lineas=lineas_k,
                ids_seleccion=element_ids or [],
                columnas=cols,
                mm_prueba=_EMPOTRAMIENTO_PRUEBA_MM,
                diam_nominal_mm=empotramiento_diam_nominal_mm,
                skip_indices_0based=skip_emp_k,
            )
            if lineas_metas_k is not None and len(lineas_metas_k) == len(lineas_k):
                lineas_k, lineas_metas_k = aplicar_empotramiento_lineas_unificadas(
                    line_metas=lineas_metas_k, **emp_kwargs_k
                )
            else:
                lineas_k = aplicar_empotramiento_lineas_unificadas(**emp_kwargs_k)
                lineas_metas_k = None
            if not lineas_k:
                continue
            lotes_post_empotramiento.append((lineas_k, lineas_metas_k, 0))
        if not lotes_post_empotramiento:
            return (
                0,
                u"No se pudo obtener eje (LocationCurve/LocationPoint ni cara) en ninguna capa.",
                vac,
                vac,
                vac,
            )

    lotes_compact, n_tramos_croquis, stats_prep_croquis = (
        _lotes_croquis_previo_troceo_todos_lotes(lotes_post_empotramiento)
    )
    if not lotes_compact:
        return (
            0,
            u"No quedaron tramos de croquis tras deduplicar (misma posición, punto medio o segmentos contenidos).",
            vac,
            vac,
            vac,
        )
    lotes_post_empotramiento = lotes_compact

    lateral_face_anchors = _lateral_face_anchors_from_columnas(cols)

    ids_creados = []
    ids_planos = []
    ids_marcadores_normal = []
    n_planos_eje = 0
    stats_troceo_line = None
    n_mc_fallidas = 0
    uso_troceo_post_mc = False
    troceo_line_omitido_origen = False
    planes_empalme_list = []
    planes_par_list = []
    rebar_err_muestra = []
    rebar_err_visto = set()
    try:
        with Transaction(document, u"BIMTools — Eje columnas (planos + Rebar)") as t:
            t.Start()
            try:
                ids_planos = []
                n_planos_eje = 0
                if not emp_ids:
                    ids_planos = crear_sketch_planes_extremos_columnas_ordenados(
                        document, cols
                    )
                    n_planos_eje = len(ids_planos)
                if not multicapa_varios_anillos:
                    ids_emp_sp, planes_empalme_list, mk_emp = (
                        crear_sketch_planes_empalme_desde_location_curve(
                            document,
                            emp_ids,
                            crear_marcador_normal_primer_plano=False,
                        )
                    )
                    if ids_emp_sp:
                        ids_planos.extend(ids_emp_sp)
                    if mk_emp:
                        ids_marcadores_normal.extend(mk_emp)
                    planes_par_list = []
                    if planes_empalme_list:
                        d_troceo_mm = _mm_desplazamiento_plano_par_tabla_empalme(
                            empotramiento_diam_nominal_mm
                        )
                        dft_tr = _mm_to_ft(d_troceo_mm)
                        for p in planes_empalme_list:
                            pq = _plano_desplazado_seg_normal(p, dft_tr)
                            planes_par_list.append(pq if pq is not None else p)
                ids_creados = []
                n_mc_fallidas = 0
                mark_etiqueta_base = 0
                for lineas_lote, lineas_metas_lote, _mark_troceo_unused in (
                    lotes_post_empotramiento
                ):
                    pe_capa = planes_empalme_list
                    pp_capa = planes_par_list
                    if multicapa_varios_anillos:
                        pe_capa = []
                        pp_capa = []
                        if emp_ids:
                            ids_emp_sp_k, pe_capa, mk_emp_k = (
                                crear_sketch_planes_empalme_desde_location_curve(
                                    document,
                                    emp_ids,
                                    crear_marcador_normal_primer_plano=False,
                                )
                            )
                            if ids_emp_sp_k:
                                ids_planos.extend(ids_emp_sp_k)
                            if mk_emp_k:
                                ids_marcadores_normal.extend(mk_emp_k)
                            if pe_capa:
                                d_troceo_mm = (
                                    _mm_desplazamiento_plano_par_tabla_empalme(
                                        empotramiento_diam_nominal_mm
                                    )
                                )
                                dft_tr = _mm_to_ft(d_troceo_mm)
                                for p in pe_capa:
                                    pq = _plano_desplazado_seg_normal(p, dft_tr)
                                    pp_capa.append(pq if pq is not None else p)
                    lineas_para_crear = lineas_lote
                    marcas_previas = None
                    metas_para_crear = lineas_metas_lote
                    troceo_omitido_este_lote = False
                    if lateral_face_anchors and lineas_lote:
                        metas_para_crear = (
                            _enriquecer_line_metas_croquis_cara_lateral_cercana(
                                lineas_lote,
                                lineas_metas_lote,
                                lateral_face_anchors,
                            )
                        )
                    if (
                        pe_capa
                        and lineas_lote
                        and emp_ids
                    ):
                        po, nu = _origen_y_normal_unidad_plano(
                            pe_capa[0]
                        )
                        if po is not None and nu is not None:
                            (
                                lineas_para_crear,
                                marcas_previas,
                                st_troceo,
                                metas_para_crear,
                            ) = _lineas_y_marcas_previas_troceo_plano_empalme(
                                lineas_lote,
                                pe_capa,
                                pp_capa,
                                empotramiento_diam_nominal_mm,
                                mark_index_base_0=0,
                                line_metas=lineas_metas_lote,
                            )
                            if stats_troceo_line is None:
                                stats_troceo_line = st_troceo
                        else:
                            troceo_omitido_este_lote = True
                            troceo_line_omitido_origen = True
                    items = []
                    if marcas_previas is not None:
                        _lpc_mc = lineas_para_crear or []
                        _mpc_mc = (
                            list(metas_para_crear)
                            if metas_para_crear is not None
                            else [None] * len(_lpc_mc)
                        )
                        if len(_mpc_mc) < len(_lpc_mc):
                            _mpc_mc = _mpc_mc + [None] * (
                                len(_lpc_mc) - len(_mpc_mc)
                            )
                        for ln, mc, mm in zip(
                            _lpc_mc,
                            marcas_previas or [],
                            _mpc_mc,
                        ):
                            if ln is not None:
                                items.append((ln, mc[0], mc[1], mm))
                    elif (
                        pe_capa
                        and emp_ids
                        and lineas_para_crear
                        and troceo_omitido_este_lote
                    ):
                        mib = 0 if multicapa_varios_anillos else mark_etiqueta_base
                        seg = _lineas_lote_troceo_empalme_dos_pasos(
                            lineas_para_crear,
                            pe_capa,
                            pp_capa,
                            empotramiento_diam_nominal_mm,
                            mib,
                        )
                        for tup in seg or []:
                            if len(tup) >= 3:
                                items.append((tup[0], tup[1], tup[2], None))
                    else:
                        mib = 0 if multicapa_varios_anillos else mark_etiqueta_base
                        lpc = lineas_para_crear or []
                        for idx, ln in enumerate(lpc):
                            if ln is None:
                                continue
                            mark = u"{0:02d}".format(idx + 1 + mib)
                            com = (
                                u"BIMTools_EjeCol_Impar"
                                if (idx % 2 == 0)
                                else u"BIMTools_EjeCol_Par"
                            )
                            mm = None
                            if metas_para_crear is not None and idx < len(
                                metas_para_crear
                            ):
                                mm = metas_para_crear[idx]
                            if (
                                mm
                                and mm.get(u"es_copia_cadena")
                                and len(cols) == 1
                            ):
                                com = _COMMENT_MODEL_LINE_EJE_CADENA
                            items.append((ln, mark, com, mm))
                    batch_ids = []
                    for ln, mk, coment, meta in items:
                        order_hosts = []
                        seen_h = set()
                        ht = _host_columna_para_rebar(document, ln, meta, cols)
                        for h in [ht] + list(cols):
                            if h is None:
                                continue
                            try:
                                hi = int(h.Id.IntegerValue)
                            except Exception:
                                continue
                            if hi in seen_h:
                                continue
                            seen_h.add(hi)
                            order_hosts.append(h)
                        rb = None
                        err_rb = None
                        for htry in order_hosts:
                            rb, err_rb = _crear_rebar_eje_desde_linea(
                                document,
                                htry,
                                rebar_bar_type,
                                ln,
                                meta,
                            )
                            if rb is not None:
                                break
                        if rb is None:
                            n_mc_fallidas += 1
                            if err_rb:
                                try:
                                    key = err_rb.strip()
                                except Exception:
                                    key = u"{0}".format(err_rb)
                                if key and key not in rebar_err_visto:
                                    rebar_err_visto.add(key)
                                    if len(rebar_err_muestra) < 6:
                                        rebar_err_muestra.append(key)
                            continue
                        eid = rb.Id
                        batch_ids.append(eid)
                        _aplicar_marca_y_comentario_model_curve(
                            document, eid, mk, coment
                        )
                    ids_creados.extend(batch_ids)
                    if not multicapa_varios_anillos:
                        mark_etiqueta_base += len(batch_ids)
            except Exception as ex:
                try:
                    t.RollBack()
                except Exception:
                    pass
                return 0, u"Error: {0}".format(ex), vac, vac, vac
            t.Commit()
    except Exception as ex:
        return 0, u"Error (transacción): {0}".format(ex), vac, vac, vac
    n = len(ids_creados)
    np = len(ids_planos)
    try:
        n_emp = int(np - int(n_planos_eje))
    except Exception:
        n_emp = 0
    hay_plano_empalme_creado = n_emp > 0
    try:
        if uidocument is not None and (n or np or ids_marcadores_normal):
            from System.Collections.Generic import List

            combined = (
                list(ids_planos)
                + list(ids_creados)
                + list(ids_marcadores_normal)
            )
            uidocument.Selection.SetElementIds(List[ElementId](combined))
    except Exception:
        pass
    msg_base = u"{0} plano(s) ({1} eje + {2} empalme), {3} Rebar, {4} columna(s).".format(
        np, n_planos_eje, n_emp, n, len(cols)
    )
    extras = []
    if stats_troceo_line is not None:
        extras.append(
            _mensaje_diagnostico_troceo_line_api(
                stats_troceo_line, n, n_mc_fallidas, False
            )
        )
    elif uso_troceo_post_mc:
        extras.append(
            _mensaje_diagnostico_troceo_line_api(None, n, 0, True)
        )
    elif troceo_line_omitido_origen:
        extras.append(
            u"[Troceo Line API] Omitido: no se obtuvo Origin/Normal del plano de empalme."
        )
    msg_final = msg_base
    if extras:
        msg_final = msg_base + u" | " + u" ".join(extras)
    n_emp_ids = len(emp_ids)
    _task_dialog_resumen_troceo_line(
        stats_troceo_line,
        n_tramos_croquis,
        n,
        n_mc_fallidas,
        troceo_line_omitido_origen,
        uso_troceo_post_mc,
        len(cols),
        n_emp_ids,
        hay_plano_empalme_creado,
        stats_prep_croquis,
        rebar_errores_muestra=rebar_err_muestra,
    )
    return (
        n,
        msg_final,
        ids_creados,
        ids_planos,
        list(ids_marcadores_normal),
    )
