# -*- coding: utf-8 -*-
"""
Structural Framing — cara superior e **inferior**: trazo de armadura (offsets, unión de cadenas,
extensiones, troceo por **caras** de obstáculos en extremos (y planos de empalme),
``ModelCurve`` y ``Rebar``. Tras colocar, ``IndependentTag`` (según bandera) con la familia
``EST_A_STRUCTURAL REBAR TAG`` en la vista activa si ``_ETIQUETAR_REBAR_EN_VISTA_ACTIVA`` es ``True``
(por defecto **True** en BIMTools vigas; desactívelo en código si no desea etiquetas automáticas).
Los ganchos usan el tipo nombrado ``Standard - 90 deg.`` (``HOOK_GANCHO_90_STANDARD_NAME``).
**Suples:** desplazamiento −n respecto a las capas de la cara activa; **superior** trocea en
0,25/0,75 y deja trozos impares; **inferior** en **0,10/0,90** y deja solo trozos de numeración **par** (2, 4, …).
Planos según eje y cotas ``STRUCTURAL_ELEVATION_AT_*``. Recorte por pilares solo en extremos libres.

Contrato establecido por ``enfierrado_vigas.ColocarArmaduraVigasStubHandler``:
``crear_detail_lines_largo_cara_superior_en_vista`` → ``(n_model_curves, n_rebar_count, avisos)``.

Depuración: ``_DIBUJAR_MODEL_LINE_PROYECCION_CARA_PARALELA`` dibuja por tramo la guía de laterales
(**cara superior o inferior**) tras la última capa (−n, paso + offset) y el marcador de normal;
ver :func:`_crear_model_line_cara_superior_offset_y_marcador_normal`. En enfierrado, *Model lines*
desactiva ``ModelCurve`` (:func:`crear_detail_lines_largo_cara_superior_en_vista`,
``crear_model_lines=False``).
"""

import math

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilySymbol,
    FilteredElementCollector,
    GeometryInstance,
    IndependentTag,
    Line,
    LocationCurve,
    Options,
    Plane,
    PlanarFace,
    Reference,
    SketchPlane,
    Solid,
    TagMode,
    TagOrientation,
    Transaction,
    UV,
    View3D,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import Rebar, RebarPresentationMode
from bimtools_rebar_hook_lengths import hook_length_mm_from_nominal_diameter_mm
from geometria_colision_vigas import obtener_solidos_elemento
from geometria_fundacion_cara_inferior import (
    _obtener_plano_detalle_vista,
    _proyectar_linea_al_plano_vista,
    _proyectar_punto_al_plano,
)

try:
    from armadura_vigas_capas import (
        _build_collinear_chains_from_elements,
        _read_width_depth_ft,
    )
except Exception:
    _build_collinear_chains_from_elements = None
    _read_width_depth_ft = None

try:
    from barras_bordes_losa_gancho_empotramiento import _rebar_nominal_diameter_mm
except Exception:
    def _rebar_nominal_diameter_mm(bar_type):
        return 16.0

try:
    from enfierrado_shaft_hashtag import lap_mm_para_bar_type
    from enfierrado_shaft_hashtag import _enforce_rebar_hook_types_by_name
    from enfierrado_shaft_hashtag import _sweep_rebar_hook_types_to_name
except Exception:
    lap_mm_para_bar_type = None
    _enforce_rebar_hook_types_by_name = None
    _sweep_rebar_hook_types_to_name = None
try:
    from enfierrado_shaft_hashtag import _rebar_centerline_dominant_curve
    from enfierrado_shaft_hashtag import _rebar_centerline_midpoint_xyz
except Exception:
    _rebar_centerline_dominant_curve = None
    _rebar_centerline_midpoint_xyz = None

try:
    from enfierrado_shaft_hashtag import (
        _create_overlap_dimension_from_detail_refs,
        _get_named_left_right_refs_from_detail_instance,
    )
except Exception:
    _create_overlap_dimension_from_detail_refs = None
    _get_named_left_right_refs_from_detail_instance = None

try:
    from enfierrado_shaft_hashtag import etiquetar_rebars_creados_en_vista
except Exception:
    etiquetar_rebars_creados_en_vista = None
try:
    from enfierrado_shaft_hashtag import _independent_tag_intersects_obstacles
except Exception:
    _independent_tag_intersects_obstacles = None
try:
    from enfierrado_shaft_hashtag import _apply_armadura_largo_total_to_rebars
except Exception:
    _apply_armadura_largo_total_to_rebars = None

try:
    from barras_bordes_losa_gancho_empotramiento import (
        _find_fixed_lap_detail_symbol_id,
    )
except Exception:
    _find_fixed_lap_detail_symbol_id = None

try:
    from lap_detail_link_vigas_schema import set_lap_detail_vigas_rebar_link
except Exception:
    set_lap_detail_vigas_rebar_link = None

try:
    from rebar_fundacion_cara_inferior import aplicar_layout_fixed_number_rebar
    from rebar_fundacion_cara_inferior import crear_rebar_desde_curva_linea_con_ganchos
    from rebar_fundacion_cara_inferior import HOOK_GANCHO_90_STANDARD_NAME
except Exception:
    aplicar_layout_fixed_number_rebar = None
    crear_rebar_desde_curva_linea_con_ganchos = None
    HOOK_GANCHO_90_STANDARD_NAME = u"Standard - 90 deg."

_FRAMING_CAT = int(BuiltInCategory.OST_StructuralFraming)
_OBST_CAT_IDS = frozenset(
    (
        int(BuiltInCategory.OST_StructuralColumns),
        int(BuiltInCategory.OST_Walls),
    )
)

_MIN_FACE_AREA_FT2 = 1e-6
_MIN_LINE_LEN_FT = 1.0 / 304.8 * 5.0

# Recubrimiento normal a la cara (mm) y desplazamiento en planta al eje V.
_OFFSET_NORMAL_MM = 25.0
_OFFSET_V_ALERO_MM = 25.0
# Curva guía laterales (cara sup.): mm en −n **después de la última capa** longitudinal
# (sin incluir el paso entre capas; ese acumulado se suma en código).
_MODELO_LINE_CARA_SUPERIOR_OFFSET_MM = 100.0
# Trazo del marcador de **normal CreateFromCurves** (:func:`_norm_createfromcurves_desde_cara_y_tramo`) en mm.
_MARCADOR_NORMAL_CURVA_MM = 250.0
# Entre capas superiores longitudinales y capa de suple: paso en −n (interior típico), mm.
# Capa k (0-based) desplaza k×este valor; el suple usa N_capas×este valor desde el eje extendido.
_OFFSET_SUPLES_SEGUNDA_CAPA_MM = 50.0
# Fracciones normalizadas sobre el eje inicio→fin (con cotas estructurales si hay) → plano suple.
_SUPLES_LOCATION_FRACCIONES = (0.25, 0.75)
# Suple **inferior**: planos en 0,10 y 0,90; trozos 1-based: conservar solo los **pares**.
_SUPLES_INFERIOR_LOCATION_FRACCIONES = (0.1, 0.9)

# ``SetLayoutAsFixedNumber``: longitud del vector (mm) =
# ``ancho viga − _LAYOUT_ARRAY_SIDE_CLEARANCE_MM − 2×Ø_estribo − Ø_longitudinal``
# (centro–centro extremos; coherente con offsets cara→eje con ½Ø por lado).
_LAYOUT_ARRAY_SIDE_CLEARANCE_MM = 50.0

# Extensión longitudinal a cada lado (m) y recorte final en eje (mm).
_EXTENSION_ENDS_MM = 2000.0
_TRIM_AXIS_ENDS_MM = 25.0

# Zona longitudinal (mm) desde cada extremo de la curva **ya extendida** en la que se
# cuentan intersecciones con **caras** de obstáculos: extensión + margen para ancho del pilar/muro.
_OBST_FACE_ZONE_EXTRA_MM = 1500.0

# Mínimo de intersecciones con caras en esa zona para activar el recorte del tramo exterior.
_MIN_FACE_HITS_EN_EXTREMO = 2

_UMBRAL_LARGO_EMPALMES_MM = 12000.0
# Margen mm al comparar con el umbral (pies→mm, tolerancia numérica).
_UMBRAL_EMPALMES_COMP_EPS_MM = 1.5

# Respaldo si no está disponible ``lap_mm_para_bar_type`` (tabla = borde losa / shaft).
_TRASLAPO_LONG_MULT_DIAM_FALLBACK = 40.0

_TOL_SPLIT_FT = 0.02
_TOL_ON_FACE_FT = 0.06

# Marcador gráfico en punto medio del tramo: tangente (eje X local de la barra) y Z del modelo.
_MARCADOR_EJE_XZ_MM = 350.0
_MARCADOR_EJE_XZ_PARALELO_TOL = 0.985
# ``ModelCurve`` del trazo, marcadores XZ y ``DetailCurve`` en vista (False = solo Rebar).
_DIBUJAR_MODEL_LINE_TRAMO = False
# ``ModelLine`` laterales (cara sup.): tras última capa + _MODELO_LINE_CARA_SUPERIOR_OFFSET_MM; marcador normal.
_DIBUJAR_MODEL_LINE_PROYECCION_CARA_PARALELA = True
_DIBUJAR_DETAIL_CURVE_VISTA = False
_DIBUJAR_MARCADORES_EJE_XZ = False
# Depuración suples: ``ModelLine`` en planos de troceo (off = sin marcadores).
_DIBUJAR_MARCADORES_PLANOS_SUPLES = False
_MARCADOR_PLANO_SUPLES_EN_PLANO_MM = 400.0
_MARCADOR_PLANO_SUPLES_TICK_Z_MM = 200.0
# Detail Item line-based de empalme (misma familia/tipo que borde losa / shaft).
_DIBUJAR_DETAIL_ITEM_TRASLAPE = True

# ``IndependentTag``: familia de etiqueta de armadura (misma que ``enfierrado_shaft_hashtag``).
_ETIQUETAR_REBAR_FAMILIA_NOMBRE = u"EST_A_STRUCTURAL REBAR TAG"
# Respaldo solo si no se puede importar ``etiquetar_rebars_creados_en_vista``: TM_ADDBY_CATEGORY.
# False: no crear etiquetas de Rebar al colocar (incl. tipos «momento»/flexión en la familia de tags).
_ETIQUETAR_REBAR_EN_VISTA_ACTIVA = True
_ETIQUETAR_REBAR_LEADER = True
_ETIQUETAR_REBAR_ORIENT = TagOrientation.Horizontal
_MAX_AVISOS_FALLA_ETIQUETA_REBAR = 8
# Alineación de cabeceras de ``IndependentTag`` en el lote (cara superior o inferior + suples).
_TAG_ALIGN_PERP_EXTRA_FT = 0.04
# Desplazamiento adicional hacia el exterior del hormigón (fibra sup./inf.) tras alinear.
_TAG_OUTSIDE_CONCRETE_EXTRA_FT = 0.35
# Separación iterativa entre cabeceras solapadas: solo a lo largo del eje de viga en vista,
# sin componente ⟂ al eje de alineación (conserva la fila común tras :func:`_alinear_etiquetas_rebar_mismo_lote`).
_TAG_SEPARATE_STEP_MM = 52.0
# Máx. desplazamiento a lo largo del eje de separación desde la cabecera al **iniciar** esta fase
# (tras alinear). Evita líderes kilométricos si hay muchos solapes en la misma zona.
_TAG_SEPARATE_MAX_OFFSET_ALONG_SLIDE_MM = 420.0
# Margen extra en la comprobación AABB (etiquetas deben quedar con algo de aire, no rozando).
_TAG_SEPARATE_CLEARANCE_MM = 32.0
_TAG_SEPARATE_MAX_ITER = 120
# Tras la separación longitudinal (+ tope), solapes residuales: filas extra hacia ``+perp``
# (exterior del hormigón), sin ``snap`` global (no se aplastan de nuevo a una sola línea).
_TAG_SEPARATE_RESIDUAL_ROW_STEP_MM = 40.0
_TAG_SEPARATE_RESIDUAL_MAX_PERP_EXTRA_MM = 300.0
_TAG_SEPARATE_RESIDUAL_MAX_ITER = 140
# Por cada par que sigue solapado en una pasada, hasta N desplazamientos en ``+perp`` (mismo ciclo).
_TAG_SEPARATE_RESIDUAL_SUBSTEPS_PER_PAIR = 18

# Eje del trazo 1.ª capa: ``False`` (por defecto) = proyección de extremos del Location al plano
# de la cara superior por viga + offsets (:func:`_curva_armadura_superior_en_fibra`); ``True`` =
# unificación del Location 3D por cadena (:func:`_curva_armadura_superior_desde_location_unificada`).
_TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA = False
# Misma convención para armadura longitudinal en cara inferior (extremos Location → plano cara inf.).
_TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA = False


def filtrar_solo_structural_framing(document, element_ids):
    """De una lista de ``ElementId``, devuelve solo instancias Structural Framing."""
    if document is None:
        return []
    out = []
    for eid in element_ids or []:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None or not el.IsValidObject:
            continue
        try:
            if el.Category is None:
                continue
            if int(el.Category.Id.IntegerValue) != _FRAMING_CAT:
                continue
        except Exception:
            continue
        out.append(el)
    return out


def filtrar_obstaculos_seleccion_no_framing(document, element_ids):
    """
    Elementos de la selección que actúan como obstáculo (muros y pilares de hormigón
    habituales), excluyendo Structural Framing.
    """
    if document is None:
        return []
    out = []
    for eid in element_ids or []:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None or not el.IsValidObject:
            continue
        try:
            if el.Category is None:
                continue
            cid = int(el.Category.Id.IntegerValue)
        except Exception:
            continue
        if cid == _FRAMING_CAT:
            continue
        if cid in _OBST_CAT_IDS:
            out.append(el)
    return out


def _plano_desde_planar_face(face):
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        bb_uv = face.GetBoundingBox()
        if bb_uv is None:
            return None
        u_mid = 0.5 * (float(bb_uv.Min.U) + float(bb_uv.Max.U))
        v_mid = 0.5 * (float(bb_uv.Min.V) + float(bb_uv.Max.V))
        o = face.Evaluate(UV(u_mid, v_mid))
        if o is None:
            return None
        return Plane.CreateByNormalAndOrigin(n, o)
    except Exception:
        return None


def obtener_cara_superior_framing(elemento):
    """``PlanarFace`` cuya normal tiene mayor componente Z (cara superior típica)."""
    if elemento is None:
        return None
    best_face = None
    best_nz = None
    for solid in obtener_solidos_elemento(elemento):
        if solid is None or solid.Faces is None:
            continue
        try:
            face_enum = solid.Faces
        except Exception:
            continue
        for face in face_enum:
            if not isinstance(face, PlanarFace):
                continue
            try:
                if face.Area < _MIN_FACE_AREA_FT2:
                    continue
            except Exception:
                pass
            try:
                nz = float(face.FaceNormal.Z)
            except Exception:
                continue
            if best_nz is None or nz > best_nz:
                best_nz = nz
                best_face = face
    return best_face


def obtener_cara_inferior_framing(elemento):
    """``PlanarFace`` cuya normal tiene menor componente Z (cara inferior típica)."""
    if elemento is None:
        return None
    best_face = None
    best_nz = None
    for solid in obtener_solidos_elemento(elemento):
        if solid is None or solid.Faces is None:
            continue
        try:
            face_enum = solid.Faces
        except Exception:
            continue
        for face in face_enum:
            if not isinstance(face, PlanarFace):
                continue
            try:
                if face.Area < _MIN_FACE_AREA_FT2:
                    continue
            except Exception:
                pass
            try:
                nz = float(face.FaceNormal.Z)
            except Exception:
                continue
            if best_nz is None or nz < best_nz:
                best_nz = nz
                best_face = face
    return best_face


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


def _line_bound_desde_location_curve(crv):
    """
    ``Line`` entre extremos del ``LocationCurve`` (recta o cuerda si el eje es curvo).
    """
    if crv is None or not crv.IsBound:
        return None
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
        if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def _param_double_optional_bip(elemento, bip):
    try:
        prm = elemento.get_Parameter(bip)
        if prm is None or not prm.HasValue:
            return None
        return float(prm.AsDouble())
    except Exception:
        return None


def _extremos_location_con_cotas_estructurales(elemento):
    """
    Extremos del ``LocationCurve`` con **Z** desde cotas estructurales de inicio/fin
    (pies, mismo espacio que el modelo). Solo se sustituye Z cuando **ambas** cotas
    existen; si no, se mantienen los puntos del eje analítico.
    """
    try:
        crv = _curva_location_framing(elemento)
        if crv is None:
            return None, None
        g0 = crv.GetEndPoint(0)
        g1 = crv.GetEndPoint(1)
    except Exception:
        return None, None
    try:
        z0 = _param_double_optional_bip(
            elemento, BuiltInParameter.STRUCTURAL_ELEVATION_AT_START
        )
        z1 = _param_double_optional_bip(
            elemento, BuiltInParameter.STRUCTURAL_ELEVATION_AT_END
        )
    except Exception:
        return g0, g1
    if z0 is not None and z1 is not None:
        try:
            return (
                XYZ(g0.X, g0.Y, float(z0)),
                XYZ(g1.X, g1.Y, float(z1)),
            )
        except Exception:
            pass
    return g0, g1


def _punto_fraccion_extremos_con_cotas(p0, p1, u_norm):
    """Punto en la línea p0→p1 (vectores 3D con Z ya corregida si aplica)."""
    if p0 is None or p1 is None:
        return None
    try:
        uu = float(u_norm)
        if uu < 0.0 or uu > 1.0:
            return None
        return p0 + (p1 - p0).Multiply(uu)
    except Exception:
        return None


def _plano_division_suple_desde_location_y_cotas_extremos(elemento, u_norm):
    """
    Plano ⟂ eje de la viga para troceo de suples: normal = dirección inicio→fin con
    **cotas estructurales en los extremos**; origen = fracción ``u_norm`` sobre ese
    trazo (no ``Evaluate`` crudo del eje analítico, que puede desfasarse en Z).
    """
    try:
        p0, p1 = _extremos_location_con_cotas_estructurales(elemento)
        if p0 is None or p1 is None:
            return _plano_division_desde_location_curve_fraccion_normalizada(
                elemento, u_norm
            )
        d = p1 - p0
        if d.GetLength() < _MIN_LINE_LEN_FT:
            return _plano_division_desde_location_curve_fraccion_normalizada(
                elemento, u_norm
            )
        u = d.Normalize()
        if u is None or u.GetLength() < 1e-12:
            return _plano_division_desde_location_curve_fraccion_normalizada(
                elemento, u_norm
            )
        origin = _punto_fraccion_extremos_con_cotas(p0, p1, u_norm)
        if origin is None:
            return _plano_division_desde_location_curve_fraccion_normalizada(
                elemento, u_norm
            )
        return Plane.CreateByNormalAndOrigin(u, origin)
    except Exception:
        try:
            return _plano_division_desde_location_curve_fraccion_normalizada(
                elemento, u_norm
            )
        except Exception:
            return None


def _punto_medio_linea_bound(line):
    if line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        return p0 + (p1 - p0).Multiply(0.5)
    except Exception:
        return None


def _distancia_punto_a_segmento_xyz(p0, p1, pt):
    try:
        dv = p1 - p0
        L2 = float(dv.DotProduct(dv))
        if L2 < 1e-24:
            return float(pt.DistanceTo(p0))
        t = float((pt - p0).DotProduct(dv)) / L2
        if t < 0.0:
            t = 0.0
        elif t > 1.0:
            t = 1.0
        closest = p0 + dv.Multiply(t)
        return float(pt.DistanceTo(closest))
    except Exception:
        return 1e30


def _distancia_punto_a_curve_bound(crv, pt):
    """Distancia mínima de ``pt`` a la curva **acotada** (tramo entre extremos)."""
    if crv is None or pt is None:
        return 1e30
    try:
        if isinstance(crv, Line):
            return _distancia_punto_a_segmento_xyz(
                crv.GetEndPoint(0),
                crv.GetEndPoint(1),
                pt,
            )
    except Exception:
        pass
    best = 1e30
    for i in range(33):
        u = float(i) / 32.0
        try:
            q = crv.Evaluate(u, True)
            d = float(pt.DistanceTo(q))
            if d < best:
                best = d
        except Exception:
            continue
    return best


def _punto_en_bbox_elemento_expandido(pt, elemento, margen_ft=0.2):
    """
    True si ``pt`` cae dentro del ``BoundingBox`` del elemento ampliado en ``margen_ft``
    (aproxima “colisión” viga–punto sin depender de API de distancia a ``Solid``).
    """
    if pt is None or elemento is None:
        return False
    try:
        bb = elemento.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return False
    m = float(margen_ft)
    try:
        return (
            bb.Min.X - m <= pt.X <= bb.Max.X + m
            and bb.Min.Y - m <= pt.Y <= bb.Max.Y + m
            and bb.Min.Z - m <= pt.Z <= bb.Max.Z + m
        )
    except Exception:
        return False


def _candidatos_host_empalme(chain, emp_elems):
    """Vigas de la cadena colineal más vigas de troceo/empalme, sin duplicar ``ElementId``."""
    seen = set()
    out = []
    for el in list(chain or []) + list(emp_elems or []):
        if el is None:
            continue
        try:
            eid = int(el.Id.IntegerValue)
        except Exception:
            continue
        if eid in seen:
            continue
        seen.add(eid)
        out.append(el)
    return out


def _host_framing_para_segmento_rebar(pt_mid, candidatos, fallback):
    """
    Elige host ``StructuralFraming`` para el tramo: prioriza vigas cuyo bbox
    envuelve ``pt_mid`` (colisión aproximada); si ninguna, el eje más cercano.
    """
    if pt_mid is None or not candidatos:
        return fallback
    inside = []
    for el in candidatos:
        if _punto_en_bbox_elemento_expandido(pt_mid, el):
            inside.append(el)
    pool = inside if inside else list(candidatos)
    best_el = None
    best_d = None
    for el in pool:
        crv = _curva_location_framing(el)
        d = _distancia_punto_a_curve_bound(crv, pt_mid)
        if best_d is None or d < best_d:
            best_d = d
            best_el = el
    return best_el if best_el is not None else fallback


def linea_largo_sobre_plano_cara_framing(elemento, cara_planar):
    """Proyecta extremos del LocationCurve al plano de una cara plana del framing."""
    if elemento is None or cara_planar is None:
        return None
    crv = _curva_location_framing(elemento)
    if crv is None:
        return None
    plane = _plano_desde_planar_face(cara_planar)
    if plane is None:
        return None
    try:
        p0 = crv.GetEndPoint(0)
        p1 = crv.GetEndPoint(1)
    except Exception:
        return None
    q0 = _proyectar_punto_al_plano(p0, plane)
    q1 = _proyectar_punto_al_plano(p1, plane)
    if q0 is None or q1 is None:
        return None
    try:
        if q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(q0, q1)
    except Exception:
        return None


def linea_largo_sobre_plano_cara_superior(elemento, cara_superior):
    """Proyecta el eje de la viga al plano de la cara superior."""
    return linea_largo_sobre_plano_cara_framing(elemento, cara_superior)


def linea_largo_sobre_plano_cara_inferior(elemento, cara_inferior):
    """Proyecta el eje de la viga al plano de la cara inferior."""
    return linea_largo_sobre_plano_cara_framing(elemento, cara_inferior)


def proyectar_linea_bound_sobre_planar_face(line, planar_face):
    """
    Proyecta ortogonalmente los extremos de ``line`` al plano de ``planar_face``.
    Misma construcción que :func:`linea_largo_sobre_plano_cara_framing`, pero para cualquier
    ``Line`` ya calculada (p. ej. tramo de armadura antes de ``CreateFromCurves``).
    """
    if line is None or planar_face is None:
        return None
    plane = _plano_desde_planar_face(planar_face)
    if plane is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
    except Exception:
        return None
    q0 = _proyectar_punto_al_plano(p0, plane)
    q1 = _proyectar_punto_al_plano(p1, plane)
    if q0 is None or q1 is None:
        return None
    try:
        if q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(q0, q1)
    except Exception:
        return None


def _media_diametro_nominal_rebar_mm(rebar_bar_type):
    """½ Ø nominal (mm) del tipo de barra longitudinal (cara hormigón → eje del acero)."""
    if rebar_bar_type is None:
        return 0.0
    try:
        d = float(_rebar_nominal_diameter_mm(rebar_bar_type) or 0.0)
        return 0.5 * max(0.0, d)
    except Exception:
        return 0.0


def _linea_guia_laterales_cara_superior(
    line_tramo_capa0,
    normal_cara_exterior,
    n_capas_superiores,
    step_mm_entre_capas,
):
    """
    ``Line`` longitudinal de referencia para laterales: última capa longitudinal hacia el
    interior (``−n`` con ``n`` exterior de la cara) + ``_MODELO_LINE_CARA_SUPERIOR_OFFSET_MM``.
    Válido para **cara superior o inferior** — misma convención que la armadura principal.
    """
    if line_tramo_capa0 is None or normal_cara_exterior is None:
        return None
    try:
        n_cap = max(1, int(n_capas_superiores or 1))
    except Exception:
        n_cap = 1
    try:
        step_mm = float(step_mm_entre_capas)
    except Exception:
        step_mm = 0.0
    seg_ultima_capa = line_tramo_capa0
    off_capas_mm = float(max(0, n_cap - 1)) * step_mm
    if off_capas_mm > 1e-9:
        seg_u = _linea_desplazada_mm_reverso_normal_cara(
            line_tramo_capa0,
            normal_cara_exterior,
            off_capas_mm,
        )
        if seg_u is not None:
            seg_ultima_capa = seg_u
    try:
        off_lat_mm = float(_MODELO_LINE_CARA_SUPERIOR_OFFSET_MM)
    except Exception:
        off_lat_mm = 0.0
    seg_prof = seg_ultima_capa
    if off_lat_mm > 1e-9:
        seg_l = _linea_desplazada_mm_reverso_normal_cara(
            seg_ultima_capa,
            normal_cara_exterior,
            off_lat_mm,
        )
        if seg_l is not None:
            seg_prof = seg_l
    return seg_prof


def _offset_mm_curva_laterales_vs_cara_superior(
    n_capas_superiores,
    step_mm_entre_capas,
    recubrimiento_extra_mm=0.0,
    rebar_long_bar_type=None,
):
    """
    Distancia mm desde la **cara activa** (sup./inf.) al eje de la curva guía de laterales
    (recub. + estribo + Ø longitudinal / 2, pasos entre capas y offset lateral), alineado con
    :func:`_aplicar_offsets_armadura_superior_desde_linea`. ``recubrimiento_extra_mm``: p. ej. Ø estribo.
    """
    try:
        n_cap = max(1, int(n_capas_superiores or 1))
    except Exception:
        n_cap = 1
    try:
        step_mm = float(step_mm_entre_capas)
    except Exception:
        step_mm = 0.0
    off_capas = float(max(0, n_cap - 1)) * step_mm
    try:
        rex = max(0.0, float(recubrimiento_extra_mm or 0.0))
    except Exception:
        rex = 0.0
    half_long = _media_diametro_nominal_rebar_mm(rebar_long_bar_type)
    return (
        float(_OFFSET_NORMAL_MM)
        + rex
        + half_long
        + off_capas
        + float(_MODELO_LINE_CARA_SUPERIOR_OFFSET_MM)
    )


def _crear_model_line_cara_superior_offset_y_marcador_normal(
    document,
    host_framing,
    line_tramo_capa0,
    normal_cara_superior,
    norm_createfromcurves,
    n_capas_superiores,
    step_mm_entre_capas,
    refinar_n_hint_con_cara_superior_framing=True,
):
    """
    Curva guía de **laterales** (+ marcador opcional): 1.ª capa → última capa → offset lateral.
    Crea ``ModelCurve`` y trazo en ``norm_createfromcurves`` desde el punto medio.

    Si ``refinar_n_hint_con_cara_superior_framing`` es ``True`` (cara superior), se ajusta el
    plano del dibujo con :func:`obtener_cara_superior_framing`. En **cara inferior**, pasar
    ``False`` para conservar la ``normal_cara_superior`` del trazo (en realidad ``n`` de la cara inf.).
    """
    if (
        document is None
        or host_framing is None
        or line_tramo_capa0 is None
    ):
        return 0
    n_out = 0
    seg_prof = _linea_guia_laterales_cara_superior(
        line_tramo_capa0,
        normal_cara_superior,
        n_capas_superiores,
        step_mm_entre_capas,
    )
    if seg_prof is None:
        return 0
    n_hint = normal_cara_superior
    if refinar_n_hint_con_cara_superior_framing:
        cara_sup = obtener_cara_superior_framing(host_framing)
        if cara_sup is not None:
            try:
                nh = cara_sup.FaceNormal
                if nh is not None and nh.GetLength() > 1e-12:
                    n_hint = nh.Normalize()
            except Exception:
                pass
    if _crear_model_curve(document, seg_prof, n_hint):
        n_out += 1
    if norm_createfromcurves is None:
        return n_out
    try:
        Lm = _mm_to_ft(float(_MARCADOR_NORMAL_CURVA_MM))
    except Exception:
        Lm = _mm_to_ft(250.0)
    if Lm < 1e-12:
        return n_out
    mid = _punto_medio_linea_bound(seg_prof)
    if mid is None:
        return n_out
    try:
        nb = norm_createfromcurves.Normalize()
        ln_m = Line.CreateBound(mid, mid + nb.Multiply(Lm))
        if _crear_model_curve(document, ln_m, n_hint):
            n_out += 1
    except Exception:
        pass
    return n_out


def vista_permite_detail_curve(view):
    if view is None:
        return False
    if isinstance(view, View3D):
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    return True


def _punto_insercion_etiqueta_rebar(document, view, rb):
    """Punto para ``IndependentTag`` sobre armadura longitudinal (vigas)."""
    if rb is None:
        return None
    if _rebar_centerline_midpoint_xyz is not None:
        try:
            p = _rebar_centerline_midpoint_xyz(rb)
            if p is not None:
                return p
        except Exception:
            pass
    try:
        bb = rb.get_BoundingBox(view)
        if bb is not None:
            return (bb.Min + bb.Max) * 0.5
    except Exception:
        pass
    try:
        bb0 = rb.get_BoundingBox(None)
        if bb0 is not None:
            return (bb0.Min + bb0.Max) * 0.5
    except Exception:
        pass
    return None


def _referencias_candidatas_etiqueta_rebar(document, view, rb):
    """Referencias a probar con ``IndependentTag.Create`` (compat. API / barras en conjunto)."""
    refs = []
    seen = set()

    def _add_ref(r):
        if r is None:
            return
        try:
            k = r.ConvertToStableRepresentation(document)
        except Exception:
            try:
                k = unicode(r)
            except Exception:
                k = id(r)
        if k in seen:
            return
        seen.add(k)
        refs.append(r)

    try:
        subs = rb.GetSubelements() if hasattr(rb, "GetSubelements") else None
    except Exception:
        subs = None
    if subs:
        for sub in subs:
            if sub is None:
                continue
            try:
                if hasattr(sub, "GetReference"):
                    _add_ref(sub.GetReference())
            except Exception:
                continue
    try:
        npos = int(getattr(rb, "NumberOfBarPositions", 0))
    except Exception:
        try:
            npos = (
                int(rb.GetNumberOfBarPositions())
                if hasattr(rb, "GetNumberOfBarPositions")
                else 0
            )
        except Exception:
            npos = 0
    if npos > 0:
        idxs = (0, max(0, npos - 1), int(npos / 2))
        for idx in idxs:
            try:
                if hasattr(rb, "GetReferenceToBarPosition"):
                    _add_ref(rb.GetReferenceToBarPosition(idx))
                elif hasattr(rb, "GetReferenceForBarPosition"):
                    _add_ref(rb.GetReferenceForBarPosition(idx))
            except Exception:
                continue
    try:
        _add_ref(Reference(rb))
    except Exception:
        pass

    def _collect_geom_refs(geom_elem):
        if geom_elem is None:
            return
        for go in geom_elem:
            if go is None:
                continue
            try:
                rgo = getattr(go, "Reference", None)
                if rgo is not None:
                    _add_ref(rgo)
            except Exception:
                pass
            try:
                gi = (
                    go.GetInstanceGeometry()
                    if hasattr(go, "GetInstanceGeometry")
                    else None
                )
                if gi is not None:
                    _collect_geom_refs(gi)
            except Exception:
                pass

    try:
        opts = Options()
        opts.ComputeReferences = True
        opts.IncludeNonVisibleObjects = False
        try:
            opts.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        try:
            opts.View = view
        except Exception:
            pass
        _collect_geom_refs(rb.get_Geometry(opts))
    except Exception:
        pass
    return refs


def _vec_dot(a, b):
    try:
        return float(a.X) * float(b.X) + float(a.Y) * float(b.Y) + float(a.Z) * float(b.Z)
    except Exception:
        return 0.0


def _vec_len_sq(v):
    return _vec_dot(v, v)


def _vec_normalize_xyz(v):
    L = math.sqrt(_vec_len_sq(v))
    if L < 1e-12:
        return None
    try:
        return XYZ(v.X / L, v.Y / L, v.Z / L)
    except Exception:
        return None


def _proyectar_vector_en_plano_perp_normal(v, plane_unit_normal):
    """Componente de ``v`` en el plano perpendicular a ``plane_unit_normal``."""
    try:
        n = plane_unit_normal
        d = _vec_dot(v, n)
        return XYZ(v.X - d * n.X, v.Y - d * n.Y, v.Z - d * n.Z)
    except Exception:
        return None


def _rebar_tangente_media_en_plano_vista(rebar, view_dir_unit):
    """Tangente unitaria a la curva de centro (param 0.5), proyectada en el plano de la vista."""
    if rebar is None or _rebar_centerline_dominant_curve is None:
        return None
    c = _rebar_centerline_dominant_curve(rebar)
    if c is None:
        return None
    tan = None
    try:
        tr = c.ComputeDerivatives(0.5, True)
        if tr is not None:
            tan = tr.BasisX
    except Exception:
        tan = None
    if tan is None:
        try:
            tan = c.ComputeDerivatives(0.0, True).BasisX
        except Exception:
            tan = None
    if tan is None:
        return None
    tan_p = _proyectar_vector_en_plano_perp_normal(tan, view_dir_unit)
    return _vec_normalize_xyz(tan_p)


def _promedio_vectores_unitarios_xyz(vectors):
    sx = sy = sz = 0.0
    n = 0
    for v in vectors or []:
        if v is None:
            continue
        try:
            sx += float(v.X)
            sy += float(v.Y)
            sz += float(v.Z)
            n += 1
        except Exception:
            continue
    if n == 0:
        return None
    return _vec_normalize_xyz(XYZ(sx / n, sy / n, sz / n))


def _tag_coincide_rebar_ids(tag, rebar_ids_int):
    invalid = ElementId.InvalidElementId
    try:
        for tid in tag.GetTaggedLocalElementIds():
            try:
                if int(tid.IntegerValue) in rebar_ids_int:
                    return True
            except Exception:
                continue
    except Exception:
        pass
    try:
        for leid in tag.GetTaggedElementIds():
            try:
                link_inst = leid.LinkInstanceId
                if (
                    link_inst is not None
                    and link_inst != invalid
                    and int(link_inst.IntegerValue) >= 0
                ):
                    continue
            except Exception:
                pass
            for attr in (u"LinkedElementId", u"HostElementId"):
                try:
                    eid = getattr(leid, attr, None)
                    if eid is None:
                        continue
                    if int(eid.IntegerValue) in rebar_ids_int:
                        return True
                except Exception:
                    continue
    except Exception:
        pass
    return False


def _collect_independent_tags_for_rebar_lote(document, view, rebar_element_ids):
    """
    ``IndependentTag`` en ``view`` que referencian algún ``Rebar`` de ``rebar_element_ids``.
    """
    if document is None or view is None or not rebar_element_ids:
        return []
    try:
        vid = view.Id
    except Exception:
        return []
    rebar_ids_int = set()
    for rid in rebar_element_ids:
        try:
            rebar_ids_int.add(int(rid.IntegerValue))
        except Exception:
            continue
    if not rebar_ids_int:
        return []
    tags = []
    try:
        coll = (
            FilteredElementCollector(document)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []
    for el in coll or []:
        if el is None or not isinstance(el, IndependentTag):
            continue
        try:
            if el.OwnerViewId != vid:
                continue
        except Exception:
            continue
        if not _tag_coincide_rebar_ids(el, rebar_ids_int):
            continue
        tags.append(el)
    return tags


def _tangente_media_lote_rebars_en_plano_vista(document, rebar_element_ids, vdir):
    if document is None or not rebar_element_ids or vdir is None:
        return None
    tangentes = []
    for rid in rebar_element_ids:
        try:
            rb = document.GetElement(rid)
        except Exception:
            rb = None
        if rb is None or not isinstance(rb, Rebar):
            continue
        t = _rebar_tangente_media_en_plano_vista(rb, vdir)
        if t is not None:
            tangentes.append(t)
    return _promedio_vectores_unitarios_xyz(tangentes)


def _tag_anchor_along_in_view(tag, view, t_unit):
    """Escalar para ordenar etiquetas a lo largo del eje longitudinal en la vista."""
    if tag is None or t_unit is None:
        return 0.0
    try:
        bb = tag.get_BoundingBox(view)
        if bb is not None:
            mx = 0.5 * (float(bb.Min.X) + float(bb.Max.X))
            my = 0.5 * (float(bb.Min.Y) + float(bb.Max.Y))
            mz = 0.5 * (float(bb.Min.Z) + float(bb.Max.Z))
            return _vec_dot(XYZ(mx, my, mz), t_unit)
    except Exception:
        pass
    try:
        h = tag.TagHeadPosition
        return _vec_dot(h, t_unit)
    except Exception:
        return 0.0


def _tag_pair_bbox_intersects_view(tag_a, tag_b, view):
    """Intersección AABB en ``view`` si no está disponible ``_independent_tag_intersects_obstacles``."""
    if tag_a is None or tag_b is None or view is None:
        return False
    try:
        ba = tag_a.get_BoundingBox(view)
        bb = tag_b.get_BoundingBox(view)
    except Exception:
        return False
    if ba is None or bb is None:
        return False
    try:
        return (
            float(ba.Min.X) < float(bb.Max.X)
            and float(ba.Max.X) > float(bb.Min.X)
            and float(ba.Min.Y) < float(bb.Max.Y)
            and float(ba.Max.Y) > float(bb.Min.Y)
            and float(ba.Min.Z) < float(bb.Max.Z)
            and float(ba.Max.Z) > float(bb.Min.Z)
        )
    except Exception:
        return False


def _bbox_pair_overlaps_with_margin(ba, bb, margin_ft):
    """Intersección de dos ``BoundingBoxXYZ`` con ``margin_ft`` inflado por lado (espacio mínimo)."""
    if ba is None or bb is None:
        return False
    try:
        m = 0.5 * max(0.0, float(margin_ft))
        return (
            float(ba.Min.X) - m < float(bb.Max.X) + m
            and float(ba.Max.X) + m > float(bb.Min.X) - m
            and float(ba.Min.Y) - m < float(bb.Max.Y) + m
            and float(ba.Max.Y) + m > float(bb.Min.Y) - m
            and float(ba.Min.Z) - m < float(bb.Max.Z) + m
            and float(ba.Max.Z) + m > float(bb.Min.Z) - m
        )
    except Exception:
        return False


def _tags_overlap_with_clearance(tag_a, tag_b, view, clearance_mm):
    """Solape ``get_BoundingBox(view)`` con holgura mínima ``clearance_mm`` entre cajas."""
    if tag_a is None or tag_b is None or view is None:
        return False
    margin_ft = _mm_to_ft(max(0.0, float(clearance_mm)))
    try:
        ba = tag_a.get_BoundingBox(view)
        bb = tag_b.get_BoundingBox(view)
    except Exception:
        return False
    if ba is None or bb is None:
        return False
    return _bbox_pair_overlaps_with_margin(ba, bb, margin_ft)


def _tags_overlap_in_view(tag_a, tag_b, view):
    if _independent_tag_intersects_obstacles is not None:
        try:
            return _independent_tag_intersects_obstacles(tag_a, view, [tag_b])
        except Exception:
            pass
    return _tag_pair_bbox_intersects_view(tag_a, tag_b, view)


def _framing_host_desde_lote_rebars(document, rebar_element_ids):
    """Primer muro/elemento ``Structural Framing`` host de algún ``Rebar`` del lote."""
    if document is None or not rebar_element_ids:
        return None
    for rid in rebar_element_ids:
        try:
            rb = document.GetElement(rid)
        except Exception:
            rb = None
        if rb is None or not isinstance(rb, Rebar):
            continue
        try:
            hid = rb.GetHostId()
        except Exception:
            hid = None
        if hid is None or hid == ElementId.InvalidElementId:
            continue
        try:
            host = document.GetElement(hid)
        except Exception:
            host = None
        if host is None:
            continue
        try:
            if host.Category is None:
                continue
            if int(host.Category.Id.IntegerValue) != _FRAMING_CAT:
                continue
        except Exception:
            continue
        return host
    return None


def _normal_cara_hormigon_viga_en_plano_vista(
    host_framing, vdir_unit, es_cara_inferior
):
    """
    Normal exterior de la cara superior o inferior del hormigón, proyectada en el plano
    de la vista (perpendicular a ``vdir_unit``).
    """
    if host_framing is None or vdir_unit is None:
        return None
    face = (
        obtener_cara_inferior_framing(host_framing)
        if es_cara_inferior
        else obtener_cara_superior_framing(host_framing)
    )
    if face is None:
        return None
    try:
        n = face.FaceNormal
    except Exception:
        n = None
    if n is None:
        return None
    n = _vec_normalize_xyz(n)
    if n is None:
        return None
    n_p = _proyectar_vector_en_plano_perp_normal(n, vdir_unit)
    if n_p is None or _vec_len_sq(n_p) < 1e-20:
        n_p = None
    else:
        n_p = _vec_normalize_xyz(n_p)
    if n_p is None:
        z_hint = XYZ(0, 0, -1.0) if es_cara_inferior else XYZ(0, 0, 1.0)
        n_p2 = _proyectar_vector_en_plano_perp_normal(z_hint, vdir_unit)
        n_p = _vec_normalize_xyz(n_p2)
    return n_p


def _perp_y_tangente_etiquetas_lote(
    document, rebar_element_ids, vdir, es_cara_inferior
):
    """
    Eje ``perp`` (misma dirección que el alineado hacia afuera del hormigón) y tangente
    media ``t_avg`` del lote. ``perp`` puede ser ``None`` si no es determinable.
    """
    if document is None or not rebar_element_ids or vdir is None:
        return None, None
    es_inf = bool(es_cara_inferior)
    host_fm = _framing_host_desde_lote_rebars(document, rebar_element_ids)
    n_face = _normal_cara_hormigon_viga_en_plano_vista(host_fm, vdir, es_inf)
    t_avg = _tangente_media_lote_rebars_en_plano_vista(
        document, rebar_element_ids, vdir
    )
    perp = None
    if n_face is not None:
        perp = n_face
    elif t_avg is not None:
        try:
            perp = vdir.CrossProduct(t_avg)
        except Exception:
            perp = None
        perp = _vec_normalize_xyz(perp)
    if perp is None:
        return None, t_avg
    n_ref = n_face
    if n_ref is None:
        z_hint = XYZ(0, 0, -1.0) if es_inf else XYZ(0, 0, 1.0)
        n_ref = _vec_normalize_xyz(
            _proyectar_vector_en_plano_perp_normal(z_hint, vdir)
        )
    if n_ref is not None and _vec_dot(perp, n_ref) < 0:
        try:
            perp = XYZ(-perp.X, -perp.Y, -perp.Z)
        except Exception:
            pass
    return perp, t_avg


def _tangente_separacion_conserva_alineacion(t_avg, perp_aline):
    """
    Componente de la tangente **ortogonal** a ``perp_aline`` (en el espacio modelo):
    desplazar la cabecera en esa dirección no altera ``dot(head, perp_aline)``.
    """
    if t_avg is None:
        return None
    if perp_aline is None:
        return t_avg
    d = _vec_dot(t_avg, perp_aline)
    if abs(d) < 1e-10:
        return t_avg
    v = XYZ(
        t_avg.X - d * perp_aline.X,
        t_avg.Y - d * perp_aline.Y,
        t_avg.Z - d * perp_aline.Z,
    )
    out = _vec_normalize_xyz(v)
    return out if out is not None else t_avg


def _clamp_head_along_tslide_vs_origin(head, origin, t_slide, max_abs_ft):
    """
    Limita ``|dot(head - origin, t_slide)|`` a ``max_abs_ft`` manteniendo el resto del
    vector ``head - origin`` (p. ej. componente tras snap en ``perp``).
    """
    if head is None or origin is None or t_slide is None:
        return head
    try:
        m = max(0.0, float(max_abs_ft))
    except Exception:
        m = 0.0
    if m < 1e-12:
        return head
    try:
        dx = float(head.X) - float(origin.X)
        dy = float(head.Y) - float(origin.Y)
        dz = float(head.Z) - float(origin.Z)
        tx = float(t_slide.X)
        ty = float(t_slide.Y)
        tz = float(t_slide.Z)
        s = dx * tx + dy * ty + dz * tz
        if abs(s) <= m + 1e-12:
            return head
        s_cl = max(-m, min(m, s))
        ox, oy, oz = float(origin.X), float(origin.Y), float(origin.Z)
        rx = ox + s_cl * tx + (dx - s * tx)
        ry = oy + s_cl * ty + (dy - s * ty)
        rz = oz + s_cl * tz + (dz - s * tz)
        return XYZ(rx, ry, rz)
    except Exception:
        return head


def _snap_tag_heads_to_max_perp_projection(tags, perp):
    """
    Fuerza la misma proyección ``dot(TagHeadPosition, perp) = max`` en todo el lote.
    Útil tras ``Regenerate`` o pasos iterativos que introducen drift perpendicular
    al eje de separación (escalonado visual en ficha).
    """
    if not tags or perp is None:
        return
    s_heads = []
    for tg in tags:
        if tg is None:
            continue
        try:
            h = tg.TagHeadPosition
        except Exception:
            continue
        try:
            s_heads.append((tg, h, _vec_dot(h, perp)))
        except Exception:
            continue
    if not s_heads:
        return
    try:
        ref_s = max(sh for _, _, sh in s_heads)
    except Exception:
        return
    for tg, head, s0 in s_heads:
        d = ref_s - s0
        if abs(d) < 1e-12:
            continue
        try:
            tg.TagHeadPosition = XYZ(
                head.X + d * perp.X,
                head.Y + d * perp.Y,
                head.Z + d * perp.Z,
            )
        except Exception:
            continue


def _alinear_etiquetas_rebar_mismo_lote(
    document, view, rebar_element_ids, es_cara_inferior=False
):
    """
    Alinea cabeceras del lote hacia el **exterior del hormigón** (fibra de la cara
    activa) y las iguala proyectando sobre la normal de esa cara en el plano de la vista.
    Respaldo: eje ⟂ a la tangente media si no hay host de vigas legible.
    """
    if document is None or view is None:
        return
    if not rebar_element_ids:
        return
    try:
        view_dir = view.ViewDirection
    except Exception:
        return
    vdir = _vec_normalize_xyz(view_dir)
    if vdir is None:
        return
    perp, _ = _perp_y_tangente_etiquetas_lote(
        document, rebar_element_ids, vdir, es_cara_inferior
    )
    if perp is None:
        return
    tags = _collect_independent_tags_for_rebar_lote(
        document, view, rebar_element_ids
    )
    if not tags:
        return
    projs = []
    for tag in tags:
        try:
            head = tag.TagHeadPosition
        except Exception:
            head = None
        if head is None:
            continue
        try:
            projs.append((tag, head, _vec_dot(head, perp)))
        except Exception:
            continue
    if not projs:
        return
    try:
        ref_s = (
            max(s for _, _, s in projs)
            + _TAG_ALIGN_PERP_EXTRA_FT
            + _TAG_OUTSIDE_CONCRETE_EXTRA_FT
        )
    except Exception:
        return
    for tag, head, s0 in projs:
        try:
            shift = ref_s - s0
            tag.TagHeadPosition = XYZ(
                head.X + shift * perp.X,
                head.Y + shift * perp.Y,
                head.Z + shift * perp.Z,
            )
        except Exception:
            continue
    try:
        document.Regenerate()
    except Exception:
        pass


def _etiquetar_un_rebar_categoria_vista(document, view, rb, avisos):
    """
    Crea una etiqueta por categoría para ``Rebar``. Si falla, añade aviso (acotado).
    Devuelve True si se creó al menos un ``IndependentTag``.
    """
    if document is None or view is None or rb is None:
        return False
    if not isinstance(rb, Rebar):
        return False
    p = _punto_insercion_etiqueta_rebar(document, view, rb)
    if p is None:
        nerr = sum(
            1 for a in avisos if u"Etiqueta rebar:" in (a or u"")
        )
        if nerr < _MAX_AVISOS_FALLA_ETIQUETA_REBAR:
            try:
                avisos.append(
                    u"Etiqueta rebar Id {0}: sin punto de inserción.".format(
                        rb.Id.IntegerValue
                    )
                )
            except Exception:
                avisos.append(u"Etiqueta rebar: sin punto de inserción.")
        return False
    refs = _referencias_candidatas_etiqueta_rebar(document, view, rb)
    if not refs:
        nerr = sum(1 for a in avisos if u"Etiqueta rebar:" in (a or u""))
        if nerr < _MAX_AVISOS_FALLA_ETIQUETA_REBAR:
            try:
                avisos.append(
                    u"Etiqueta rebar Id {0}: sin referencia válida.".format(
                        rb.Id.IntegerValue
                    )
                )
            except Exception:
                avisos.append(u"Etiqueta rebar: sin referencia válida.")
        return False
    primera = (
        (_ETIQUETAR_REBAR_ORIENT, _ETIQUETAR_REBAR_LEADER),
    )
    resto = []
    for o in (TagOrientation.Horizontal, TagOrientation.Vertical):
        for ldr in (True, False):
            if (o, ldr) not in primera:
                resto.append((o, ldr))
    for orient, leader in list(primera) + resto:
        for ref in refs:
            try:
                tag = IndependentTag.Create(
                    document,
                    view.Id,
                    ref,
                    leader,
                    TagMode.TM_ADDBY_CATEGORY,
                    orient,
                    p,
                )
                if tag is not None:
                    return True
            except Exception:
                continue
    nerr = sum(1 for a in avisos if u"Etiqueta rebar:" in (a or u""))
    if nerr < _MAX_AVISOS_FALLA_ETIQUETA_REBAR:
        try:
            avisos.append(
                u"Etiqueta rebar Id {0}: no se pudo crear (tipo por categoría).".format(
                    rb.Id.IntegerValue
                )
            )
        except Exception:
            avisos.append(
                u"Etiqueta rebar: no se pudo crear (compruebe familia de etiqueta Rebar)."
            )
    return False


def _mm_to_ft(mm):
    try:
        return float(mm) / 304.8
    except Exception:
        return 0.0


def _separar_etiquetas_rebar_solapadas_lote(
    document, view, rebar_element_ids, es_cara_inferior=False
):
    """
    Tras alinear cabeceras: (1) separación longitudinal con tope y ``snap`` en ``perp``;
    (2) solapes residuales: una o más filas extra en ``+perp`` (sin snap), con presupuesto
    por etiqueta. Holgura AABB configurable.
    """
    if document is None or view is None or not rebar_element_ids:
        return
    tags = _collect_independent_tags_for_rebar_lote(
        document, view, rebar_element_ids
    )
    if len(tags) < 2:
        return
    try:
        view_dir = view.ViewDirection
    except Exception:
        return
    vdir = _vec_normalize_xyz(view_dir)
    if vdir is None:
        return
    perp_align, t_avg = _perp_y_tangente_etiquetas_lote(
        document, rebar_element_ids, vdir, es_cara_inferior
    )
    if t_avg is None:
        try:
            rd = view.RightDirection
        except Exception:
            rd = None
        if rd is not None:
            t_avg = _vec_normalize_xyz(
                _proyectar_vector_en_plano_perp_normal(rd, vdir)
            )
    if t_avg is None:
        try:
            up = view.UpDirection
        except Exception:
            up = None
        if up is not None:
            t_avg = _vec_normalize_xyz(
                _proyectar_vector_en_plano_perp_normal(up, vdir)
            )
    if t_avg is None:
        return
    t_slide = _tangente_separacion_conserva_alineacion(t_avg, perp_align)
    if t_slide is None:
        return
    try:
        if _vec_len_sq(t_slide) < 1e-20:
            return
    except Exception:
        return
    try:
        document.Regenerate()
    except Exception:
        pass
    origin_heads = {}
    for tg in tags:
        if tg is None:
            continue
        try:
            h0 = tg.TagHeadPosition
            oid = int(tg.Id.IntegerValue)
            origin_heads[oid] = XYZ(float(h0.X), float(h0.Y), float(h0.Z))
        except Exception:
            try:
                origin_heads[id(tg)] = tg.TagHeadPosition
            except Exception:
                pass
    step_long_ft = max(_mm_to_ft(float(_TAG_SEPARATE_STEP_MM)), 1e-9)
    max_slide_ft = max(0.0, _mm_to_ft(float(_TAG_SEPARATE_MAX_OFFSET_ALONG_SLIDE_MM)))
    clr_mm = float(_TAG_SEPARATE_CLEARANCE_MM or 0.0)
    for _ in range(int(_TAG_SEPARATE_MAX_ITER)):
        try:
            document.Regenerate()
        except Exception:
            pass
        ordered = sorted(
            tags,
            key=lambda tg: _tag_anchor_along_in_view(tg, view, t_slide),
        )
        pairs = []
        for i in range(len(ordered)):
            for j in range(i + 1, len(ordered)):
                ta, tb = ordered[i], ordered[j]
                if _tags_overlap_with_clearance(ta, tb, view, clr_mm):
                    pairs.append((i, j))
        if not pairs:
            break
        moved_ids = set()
        for i, j in pairs:
            tg = ordered[j]
            try:
                tid = int(tg.Id.IntegerValue)
            except Exception:
                tid = id(tg)
            if tid in moved_ids:
                continue
            try:
                h = tg.TagHeadPosition
                okey = tid if tid in origin_heads else None
                if okey is None and id(tg) in origin_heads:
                    okey = id(tg)
                step_use = step_long_ft
                if okey is not None and okey in origin_heads and max_slide_ft > 1e-12:
                    orig = origin_heads[okey]
                    cur = _vec_dot(
                        XYZ(
                            float(h.X) - float(orig.X),
                            float(h.Y) - float(orig.Y),
                            float(h.Z) - float(orig.Z),
                        ),
                        t_slide,
                    )
                    nxt = cur + step_long_ft
                    if nxt > max_slide_ft:
                        step_use = max(0.0, max_slide_ft - cur)
                        if step_use < 1e-7:
                            continue
                new_h = XYZ(
                    h.X + t_slide.X * step_use,
                    h.Y + t_slide.Y * step_use,
                    h.Z + t_slide.Z * step_use,
                )
                if (
                    okey is not None
                    and okey in origin_heads
                    and max_slide_ft > 1e-12
                ):
                    new_h = _clamp_head_along_tslide_vs_origin(
                        new_h,
                        origin_heads[okey],
                        t_slide,
                        max_slide_ft,
                    )
                tg.TagHeadPosition = new_h
                moved_ids.add(tid)
            except Exception:
                continue
        if perp_align is not None:
            _snap_tag_heads_to_max_perp_projection(tags, perp_align)
            if max_slide_ft > 1e-12:
                for tg in tags:
                    if tg is None:
                        continue
                    try:
                        tid = int(tg.Id.IntegerValue)
                    except Exception:
                        tid = id(tg)
                    ok = tid if tid in origin_heads else id(tg)
                    if ok not in origin_heads:
                        continue
                    try:
                        tg.TagHeadPosition = _clamp_head_along_tslide_vs_origin(
                            tg.TagHeadPosition,
                            origin_heads[ok],
                            t_slide,
                            max_slide_ft,
                        )
                    except Exception:
                        pass
    try:
        document.Regenerate()
    except Exception:
        pass
    if perp_align is not None:
        _snap_tag_heads_to_max_perp_projection(tags, perp_align)
        try:
            document.Regenerate()
        except Exception:
            pass
    if max_slide_ft > 1e-12:
        for tg in tags:
            if tg is None:
                continue
            try:
                tid = int(tg.Id.IntegerValue)
            except Exception:
                tid = id(tg)
            ok = tid if tid in origin_heads else id(tg)
            if ok not in origin_heads:
                continue
            try:
                tg.TagHeadPosition = _clamp_head_along_tslide_vs_origin(
                    tg.TagHeadPosition,
                    origin_heads[ok],
                    t_slide,
                    max_slide_ft,
                )
            except Exception:
                pass

    # Fase 2: solapes que queden a pesar del tope longitudinal → segunda/tercera fila en +perp.
    if perp_align is not None:
        row_ft = max(0.0, _mm_to_ft(float(_TAG_SEPARATE_RESIDUAL_ROW_STEP_MM)))
        perp_cap_mm = float(_TAG_SEPARATE_RESIDUAL_MAX_PERP_EXTRA_MM)
        if es_cara_inferior:
            perp_cap_mm *= 1.2
        max_perp_extra = max(0.0, _mm_to_ft(perp_cap_mm))
        if row_ft > 1e-12 and max_perp_extra > 1e-12:
            perp_budget = {}
            for tg in tags:
                if tg is None:
                    continue
                try:
                    k0 = int(tg.Id.IntegerValue)
                except Exception:
                    k0 = id(tg)
                perp_budget[k0] = 0.0
            stagnant = 0
            for _rit in range(int(_TAG_SEPARATE_RESIDUAL_MAX_ITER)):
                try:
                    document.Regenerate()
                except Exception:
                    pass
                ordered = sorted(
                    tags,
                    key=lambda tg2: _tag_anchor_along_in_view(tg2, view, t_slide),
                )
                overlap_pairs = []
                for ii in range(len(ordered)):
                    for jj in range(ii + 1, len(ordered)):
                        ta, tb = ordered[ii], ordered[jj]
                        if _tags_overlap_with_clearance(ta, tb, view, clr_mm):
                            overlap_pairs.append((ii, jj))
                if not overlap_pairs:
                    break
                moved_any = False
                max_sub = max(1, int(_TAG_SEPARATE_RESIDUAL_SUBSTEPS_PER_PAIR))
                for i, j in overlap_pairs:
                    for _sub in range(max_sub):
                        ta_ij = ordered[i]
                        tb_ij = ordered[j]
                        if not _tags_overlap_with_clearance(
                            ta_ij, tb_ij, view, clr_mm
                        ):
                            break
                        moved_sub = False
                        for idx_mv in (j, i):
                            tg = ordered[idx_mv]
                            try:
                                tid = int(tg.Id.IntegerValue)
                            except Exception:
                                tid = id(tg)
                            bk = tid if tid in perp_budget else id(tg)
                            if bk not in perp_budget:
                                perp_budget[bk] = 0.0
                            used = float(perp_budget.get(bk, 0.0))
                            if used + row_ft > max_perp_extra + 1e-12:
                                continue
                            try:
                                h = tg.TagHeadPosition
                                tg.TagHeadPosition = XYZ(
                                    float(h.X) + float(perp_align.X) * row_ft,
                                    float(h.Y) + float(perp_align.Y) * row_ft,
                                    float(h.Z) + float(perp_align.Z) * row_ft,
                                )
                                perp_budget[bk] = used + row_ft
                                moved_sub = True
                                moved_any = True
                                okey = tid if tid in origin_heads else id(tg)
                                if okey in origin_heads and max_slide_ft > 1e-12:
                                    tg.TagHeadPosition = (
                                        _clamp_head_along_tslide_vs_origin(
                                            tg.TagHeadPosition,
                                            origin_heads[okey],
                                            t_slide,
                                            max_slide_ft,
                                        )
                                    )
                            except Exception:
                                continue
                            break
                        if not moved_sub:
                            break
                if not moved_any:
                    stagnant += 1
                    if stagnant >= 6:
                        break
                else:
                    stagnant = 0
            try:
                document.Regenerate()
            except Exception:
                pass


def _hook_extension_mm_por_extremo_desde_bar_type(document, rebar_bar_type):
    """
    mm de gancho por extremo para evaluación >12 m: tabla fija
    :mod:`bimtools_rebar_hook_lengths` según Ø nominal del ``rebar_bar_type``.
    ``document`` se mantiene por compatibilidad con llamadas existentes.
    """
    _ = document
    try:
        d_mm = _rebar_nominal_diameter_mm(rebar_bar_type)
    except Exception:
        d_mm = None
    return float(hook_length_mm_from_nominal_diameter_mm(d_mm))


def _largo_total_barra_mm_eje_mas_dos_ganchos(document, rebar_bar_type, L_eje_mm):
    """
    Estimación de largo físico con ganchos en ambos extremos: eje (curva recta del tramo)
    más dos veces la extensión de gancho del tipo de barra.
    """
    L = max(0.0, float(L_eje_mm or 0.0))
    h = _hook_extension_mm_por_extremo_desde_bar_type(document, rebar_bar_type)
    return L + 2.0 * h


def _read_width_depth_ft_local(document, elem, curve):
    if _read_width_depth_ft is not None:
        try:
            return _read_width_depth_ft(document, elem, curve)
        except Exception:
            pass
    w = d = 0.5
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            dx = abs(bb.Max.X - bb.Min.X)
            dy = abs(bb.Max.Y - bb.Min.Y)
            dz = abs(bb.Max.Z - bb.Min.Z)
            dims = sorted([dx, dy, dz], reverse=True)
            if len(dims) >= 3:
                w = float(sorted(dims[1:])[0])
                d = float(sorted(dims[1:])[1])
    except Exception:
        pass
    return w, d


def _layout_max_spacing_array_length_ft(
    document,
    elemento,
    rebar_bar_type=None,
    diametro_estribo_mm=0.0,
):
    """
    Longitud en **pies** del vector ``SetLayoutAsFixedNumber`` (distribución transversal):

    ``ancho nominal (mm) − :data:`_LAYOUT_ARRAY_SIDE_CLEARANCE_MM`
    − 2×Ø_estribo − Ø_longitudinal``,

    coherente con :func:`_aplicar_offsets_armadura_superior_desde_linea`
    (``25+Øe+Ø_long/2`` por cara → entre centros extremos van dos medios diámetros).
    ``diametro_estribo_mm``: mismo criterio que ``recubrimiento_extra_estribo_mm``.
    """
    if document is None or elemento is None:
        return None
    crv_loc = _curva_location_framing(elemento)
    if crv_loc is None:
        return None
    w_ft, _ = _read_width_depth_ft_local(document, elemento, crv_loc)
    try:
        w_mm = float(w_ft) * 304.8
        cle = float(_LAYOUT_ARRAY_SIDE_CLEARANCE_MM)
        d_est = max(0.0, float(diametro_estribo_mm or 0.0))
        d_long_mm = 2.0 * _media_diametro_nominal_rebar_mm(rebar_bar_type)
        span_mm = max(
            0.0,
            w_mm - cle - 2.0 * d_est - d_long_mm,
        )
        return _mm_to_ft(span_mm)
    except Exception:
        return None


def _cantidad_barras_desde_espaciado_transversal_mm(span_mm, esp_mm, max_barras=99):
    """
    Número de posiciones con separación aproximada ``esp_mm`` en un ancho útil ``span_mm``
    (misma referencia que ``_layout_max_spacing_array_length_ft`` en mm), tipo *fixed number*.
    """
    try:
        s = float(span_mm)
        e = float(esp_mm)
        if e <= 1e-9 or s <= 0.0:
            return None
        n = int(s / e) + 1
        return max(1, min(int(max_barras), n))
    except Exception:
        return None


def _desplazar_punto_mm_along_vec(pt, vec_unit, mm):
    if pt is None or vec_unit is None:
        return None
    try:
        return pt + vec_unit.Multiply(_mm_to_ft(mm))
    except Exception:
        return None


def _tangente_unitaria_linea(line):
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < 1e-12:
            return None
        return d.Normalize()
    except Exception:
        return None


def _eje_y_interno_curva_createfromcurves(line_tramo):
    """
    Eje **Y** del marco interno de la curva (``Curve.ComputeDerivatives`` en Revit):
    ``Transform.BasisY`` en el punto medio, unitario. Sirve como ``norm`` de
    ``CreateFromCurves*`` cuando debe coincidir con el triedro paramétr del trazo.
    Si el resultado es degenerado o no es ⟂ a la tangente, retorna ``None``.
    """
    if line_tramo is None:
        return None
    try:
        der = line_tramo.ComputeDerivatives(0.5, True)
        if der is None:
            return None
        bx = der.BasisX
        by = der.BasisY
        if by is None or by.GetLength() < 1e-12:
            return None
        yn = by.Normalize()
        if bx is not None and bx.GetLength() >= 1e-12:
            try:
                if abs(float(yn.DotProduct(bx.Normalize()))) > 1e-3:
                    return None
            except Exception:
                return None
        return yn
    except Exception:
        return None


def _presentacion_rebar_show_middle_en_vista(rebar, db_view):
    """
    En la vista dada, fija la presentación del conjunto en **Middle** (equivalente UI *Show Middle*),
    si la API lo admite para ese ``Rebar`` y ``View``.
    """
    if rebar is None or db_view is None:
        return
    try:
        if not rebar.CanApplyPresentationMode(db_view):
            return
        rebar.SetPresentationMode(db_view, RebarPresentationMode.Middle)
    except Exception:
        pass


def _norm_createfromcurves_reversa_normal_cara_superior(n_face, line_tramo):
    """
    ``norm`` para ``CreateFromCurves*``: **−n** con ``n`` = normal unitaria **exterior** de la
    cara (superior o inferior), si es casi ⟂ a la tangente del tramo; si no, ``None``.
    """
    if n_face is None or line_tramo is None:
        return None
    t = _tangente_unitaria_linea(line_tramo)
    try:
        neg_n = n_face.Normalize().Negate()
    except Exception:
        return None
    if neg_n is None or neg_n.GetLength() < 1e-12:
        return None
    try:
        neg_n = neg_n.Normalize()
    except Exception:
        return None
    if t is not None:
        try:
            if abs(float(neg_n.DotProduct(t))) > 1e-3:
                return None
        except Exception:
            return None
    return neg_n


def _norm_createfromcurves_desde_cara_y_tramo(n_face, line_tramo):
    """
    Normal ``CreateFromCurves`` = **eje Y** del triedro de la fibra en cara superior:
    ``n_face × tangente`` (misma base que ``v_dir`` en :func:`_curva_armadura_superior_en_fibra`),
    ⟂ al eje de la barra. Si ``n`` y ``t`` son casi paralelos, respaldo: ``−n_face`` y luego
    ``−Z`` proyectados ⟂ ``t``.
    """
    if n_face is None or line_tramo is None:
        return None
    t = _tangente_unitaria_linea(line_tramo)
    if t is None:
        return None

    def _proyectar_ort(vec):
        try:
            a = vec.Normalize()
            w = a - t.Multiply(a.DotProduct(t))
            if w.GetLength() < 1e-12:
                return None
            return w.Normalize()
        except Exception:
            return None

    try:
        n = n_face.Normalize()
        y = n.CrossProduct(t)
        if y.GetLength() >= 1e-12:
            return y.Normalize()
    except Exception:
        pass

    interior = None
    try:
        interior = n_face.Normalize().Negate()
    except Exception:
        interior = None
    w = _proyectar_ort(interior) if interior is not None else None
    if w is None:
        try:
            w = _proyectar_ort(XYZ.BasisZ.Negate())
        except Exception:
            w = None
    return w


def _aplicar_offsets_armadura_superior_desde_linea(
    document,
    elemento_host,
    base_line,
    n_face,
    cara,
    recubrimiento_extra_mm=0.0,
    rebar_bar_type=None,
):
    """
    A partir de una línea base (fibra proyectada o eje Location unificado), aplica
    desplazamiento en ``n_face`` y en ``n × tangente`` según recubrimiento + estribo
    + media barra longitudinal: ``25 + extra + Ø_long/2`` mm hacia interior desde cada
    cara de referencia (:data:`_OFFSET_NORMAL_MM` / :data:`_OFFSET_V_ALERO_MM`),
    con ``extra`` típ. Ø nominal del estribo (``recubrimiento_extra_mm``).
    """
    if (
        document is None
        or elemento_host is None
        or base_line is None
        or n_face is None
        or cara is None
    ):
        return None, None, None
    try:
        rex = max(0.0, float(recubrimiento_extra_mm or 0.0))
    except Exception:
        rex = 0.0
    half_long = _media_diametro_nominal_rebar_mm(rebar_bar_type)
    off_n = float(_OFFSET_NORMAL_MM) + rex + half_long
    off_v = float(_OFFSET_V_ALERO_MM) + rex + half_long
    t = _tangente_unitaria_linea(base_line)
    if t is None:
        return None, None, None
    try:
        v_dir = n_face.CrossProduct(t).Normalize()
    except Exception:
        return None, None, None
    if v_dir.GetLength() < 1e-12:
        return None, None, None
    crv_loc = _curva_location_framing(elemento_host)
    w_ft, _ = _read_width_depth_ft_local(document, elemento_host, crv_loc)
    q0 = _desplazar_punto_mm_along_vec(
        base_line.GetEndPoint(0), n_face, -off_n
    )
    q1 = _desplazar_punto_mm_along_vec(
        base_line.GetEndPoint(1), n_face, -off_n
    )
    if q0 is None or q1 is None:
        return None, None, None
    shift_ft = max(0.0, float(w_ft) * 0.5 - _mm_to_ft(off_v))
    try:
        off_v = v_dir.Multiply(shift_ft)
        q0 = q0 + off_v
        q1 = q1 + off_v
    except Exception:
        return None, None, None
    try:
        if q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
            return None, None, None
        ln = Line.CreateBound(q0, q1)
        return ln, n_face, cara
    except Exception:
        return None, None, None


def _aplicar_offsets_armadura_inferior_desde_linea(
    document,
    elemento_host,
    base_line,
    n_face,
    cara,
    recubrimiento_extra_mm=0.0,
    rebar_bar_type=None,
):
    """
    Recubrimiento hacia el **interior del host** desde la cara inferior: ``n_face`` es la
    normal exterior de Revit. Mismo criterio que la superior:
    ``25 + extra + Ø_long/2`` mm en normal y en ancho ``(b/2 − mismo)`` en ``n × tangente``.
    """
    if (
        document is None
        or elemento_host is None
        or base_line is None
        or n_face is None
        or cara is None
    ):
        return None, None, None
    try:
        rex = max(0.0, float(recubrimiento_extra_mm or 0.0))
    except Exception:
        rex = 0.0
    half_long = _media_diametro_nominal_rebar_mm(rebar_bar_type)
    off_n = float(_OFFSET_NORMAL_MM) + rex + half_long
    off_v = float(_OFFSET_V_ALERO_MM) + rex + half_long
    t = _tangente_unitaria_linea(base_line)
    if t is None:
        return None, None, None
    try:
        v_dir = n_face.CrossProduct(t).Normalize()
    except Exception:
        return None, None, None
    if v_dir.GetLength() < 1e-12:
        return None, None, None
    crv_loc = _curva_location_framing(elemento_host)
    w_ft, _ = _read_width_depth_ft_local(document, elemento_host, crv_loc)
    q0 = _desplazar_punto_mm_along_vec(
        base_line.GetEndPoint(0), n_face, -off_n
    )
    q1 = _desplazar_punto_mm_along_vec(
        base_line.GetEndPoint(1), n_face, -off_n
    )
    if q0 is None or q1 is None:
        return None, None, None
    shift_ft = max(0.0, float(w_ft) * 0.5 - _mm_to_ft(off_v))
    try:
        off_v = v_dir.Multiply(shift_ft)
        q0 = q0 + off_v
        q1 = q1 + off_v
    except Exception:
        return None, None, None
    try:
        if q0.DistanceTo(q1) < _MIN_LINE_LEN_FT:
            return None, None, None
        ln = Line.CreateBound(q0, q1)
        return ln, n_face, cara
    except Exception:
        return None, None, None


def _curva_armadura_superior_en_fibra(
    document,
    elemento,
    recubrimiento_extra_mm=0.0,
    rebar_bar_type=None,
):
    """
    Línea del trazo en la fibra superior: proyección al plano superior, desplazamiento
    según :func:`_aplicar_offsets_armadura_superior_desde_linea` (recub. + estribo + Ø_long/2).

    **Nota:** por defecto la herramienta usa este trazado (:data:`_TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA`
    ``False``). Con ``True`` se usa :func:`_curva_armadura_superior_desde_location_unificada`; ambos comparten
    offsets vía :func:`_aplicar_offsets_armadura_superior_desde_linea`.
    """
    if document is None or elemento is None:
        return None, None, None
    cara = obtener_cara_superior_framing(elemento)
    if cara is None:
        return None, None, None
    try:
        n_face = cara.FaceNormal.Normalize()
    except Exception:
        return None, None, None
    base = linea_largo_sobre_plano_cara_superior(elemento, cara)
    if base is None:
        return None, None, None
    return _aplicar_offsets_armadura_superior_desde_linea(
        document,
        elemento,
        base,
        n_face,
        cara,
        recubrimiento_extra_mm,
        rebar_bar_type,
    )


def _curva_armadura_inferior_en_fibra(
    document,
    elemento,
    recubrimiento_extra_mm=0.0,
    rebar_bar_type=None,
):
    """
    Línea del trazo en la fibra inferior: proyección al plano inferior; mismos offsets
    que la superior (recub. + estribo + Ø_long/2).
    """
    if document is None or elemento is None:
        return None, None, None
    cara = obtener_cara_inferior_framing(elemento)
    if cara is None:
        return None, None, None
    try:
        n_face = cara.FaceNormal.Normalize()
    except Exception:
        return None, None, None
    base = linea_largo_sobre_plano_cara_inferior(elemento, cara)
    if base is None:
        return None, None, None
    return _aplicar_offsets_armadura_inferior_desde_linea(
        document,
        elemento,
        base,
        n_face,
        cara,
        recubrimiento_extra_mm,
        rebar_bar_type,
    )


def _unificar_lineas_colineales(curvas, referencia):
    """Une proyección longitudinal de tramos paralelos sobre la recta de ``referencia``."""
    if not curvas:
        return None
    ref = referencia or curvas[0]
    try:
        p0 = ref.GetEndPoint(0)
        p1 = ref.GetEndPoint(1)
        ax = (p1 - p0).Normalize()
    except Exception:
        return None
    lo = hi = None
    for cv in curvas:
        for i in (0, 1):
            try:
                t = float((cv.GetEndPoint(i) - p0).DotProduct(ax))
            except Exception:
                continue
            lo = t if lo is None else min(lo, t)
            hi = t if hi is None else max(hi, t)
    if lo is None or hi is None or hi - lo < _TOL_SPLIT_FT:
        return None
    try:
        qa = p0 + ax.Multiply(lo)
        qb = p0 + ax.Multiply(hi)
        return Line.CreateBound(qa, qb)
    except Exception:
        return None


def _expand_merged_line_with_location_endpoints(merged_line, elementos_chain):
    """
    Alarga el tramo ya unificado (p. ej. fibra en cara superior) para que su proyección
    longitudinal cubra los extremos del ``LocationCurve`` de cada viga de la cadena.

    Evita que los puntos 0,25 / 0,75 del eje de referencia queden fuera del segmento
    usado para ``work``/``work2``, caso frecuente cuando la fibra proyectada es más
    corta que la línea analítica del elemento.
    """
    if merged_line is None or not elementos_chain:
        return merged_line
    try:
        qa = merged_line.GetEndPoint(0)
        qb = merged_line.GetEndPoint(1)
        dseg = qb - qa
        ax = dseg.Normalize()
        span = float(dseg.DotProduct(ax))
        if span < _MIN_LINE_LEN_FT:
            return merged_line
        lo = 0.0
        hi = span
    except Exception:
        return merged_line
    for el in elementos_chain:
        if el is None:
            continue
        crv = _curva_location_framing(el)
        if crv is None:
            continue
        for idx in (0, 1):
            try:
                pt = crv.GetEndPoint(idx)
                t = float((pt - qa).DotProduct(ax))
                if t < lo:
                    lo = t
                if t > hi:
                    hi = t
            except Exception:
                continue
    if hi - lo < _MIN_LINE_LEN_FT:
        return merged_line
    try:
        return Line.CreateBound(qa + ax.Multiply(lo), qa + ax.Multiply(hi))
    except Exception:
        return merged_line


def _curva_armadura_superior_desde_location_unificada(
    document,
    chain,
    recubrimiento_extra_mm=0.0,
    rebar_bar_type=None,
):
    """
    Trazo alternativo (si :data:`_TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA` es ``True``): por cada viga,
    ``LocationCurve`` → segmento; unificación colineal; extensión al haz de los extremos analíticos;
    mismos offsets (:func:`_aplicar_offsets_armadura_superior_desde_linea`) con cara superior del primer host.
    """
    if document is None or not chain:
        return None, None, None
    host0 = chain[0]
    if host0 is None:
        return None, None, None
    curvas_loc = []
    for e in chain:
        if e is None:
            continue
        lb = _line_bound_desde_location_curve(_curva_location_framing(e))
        if lb is not None:
            curvas_loc.append(lb)
    if not curvas_loc:
        return None, None, None
    uni = _unificar_lineas_colineales(curvas_loc, curvas_loc[0])
    if uni is None:
        return None, None, None
    merged_axis = _expand_merged_line_with_location_endpoints(uni, chain)
    cara = obtener_cara_superior_framing(host0)
    if cara is None:
        return None, None, None
    try:
        n_face = cara.FaceNormal.Normalize()
    except Exception:
        return None, None, None
    return _aplicar_offsets_armadura_superior_desde_linea(
        document,
        host0,
        merged_axis,
        n_face,
        cara,
        recubrimiento_extra_mm,
        rebar_bar_type,
    )


def _curva_armadura_inferior_desde_location_unificada(
    document,
    chain,
    recubrimiento_extra_mm=0.0,
    rebar_bar_type=None,
):
    """
    Trazo unificado en cara inferior (análogo a :func:`_curva_armadura_superior_desde_location_unificada`).
    """
    if document is None or not chain:
        return None, None, None
    host0 = chain[0]
    if host0 is None:
        return None, None, None
    curvas_loc = []
    for e in chain:
        if e is None:
            continue
        lb = _line_bound_desde_location_curve(_curva_location_framing(e))
        if lb is not None:
            curvas_loc.append(lb)
    if not curvas_loc:
        return None, None, None
    uni = _unificar_lineas_colineales(curvas_loc, curvas_loc[0])
    if uni is None:
        return None, None, None
    merged_axis = _expand_merged_line_with_location_endpoints(uni, chain)
    cara = obtener_cara_inferior_framing(host0)
    if cara is None:
        return None, None, None
    try:
        n_face = cara.FaceNormal.Normalize()
    except Exception:
        return None, None, None
    return _aplicar_offsets_armadura_inferior_desde_linea(
        document,
        host0,
        merged_axis,
        n_face,
        cara,
        recubrimiento_extra_mm,
        rebar_bar_type,
    )


def _merged_location_line_para_cadena(chain):
    """Eje unificado (solo Location, sin offsets) para estimaciones de longitud."""
    if not chain:
        return None
    curvas_loc = []
    for e in chain:
        if e is None:
            continue
        lb = _line_bound_desde_location_curve(_curva_location_framing(e))
        if lb is not None:
            curvas_loc.append(lb)
    if not curvas_loc:
        return None
    uni = _unificar_lineas_colineales(curvas_loc, curvas_loc[0])
    if uni is None:
        return None
    return _expand_merged_line_with_location_endpoints(uni, chain)


def _largo_eje_mm_max_fibra_y_location_cadena(chain, uni_fiber):
    """
    Respaldo mm de eje: máximo entre fibra unificada y eje ``LocationCurve`` unificado.
    """
    L = 0.0
    if uni_fiber is not None:
        try:
            L = max(L, float(uni_fiber.Length) * 304.8)
        except Exception:
            pass
    uni_loc = _merged_location_line_para_cadena(chain)
    if uni_loc is not None:
        try:
            L = max(L, float(uni_loc.Length) * 304.8)
        except Exception:
            pass
    return float(L)


def _merged_armadura_superior_cara_para_cadena(document, chain, rebar_bar_type=None):
    """
    Curva ``merged`` (offsets cara superior) por cadena, **misma** que en colocación:
    unificación + extensión a extremos de Location; sin estiramiento ±2 m ni recortes.
    """
    if document is None or not chain:
        return None
    if _TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA:
        merged, _n, _ = _curva_armadura_superior_desde_location_unificada(
            document, chain, 0.0, rebar_bar_type
        )
        return merged
    curvas_arm = []
    for e in chain:
        if e is None:
            continue
        ln, _n_f, _cara = _curva_armadura_superior_en_fibra(
            document, e, 0.0, rebar_bar_type
        )
        if ln is not None:
            curvas_arm.append(ln)
    if not curvas_arm:
        return None
    merged = _unificar_lineas_colineales(curvas_arm, curvas_arm[0])
    if merged is None:
        return None
    return _expand_merged_line_with_location_endpoints(merged, chain)


def _merged_armadura_inferior_cara_para_cadena(document, chain, rebar_bar_type=None):
    """
    Curva fusionada con offsets de cara inferior por cadena (sin ±2 m ni recortes de obstáculo).
    """
    if document is None or not chain:
        return None
    if _TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA:
        merged, _n, _ = _curva_armadura_inferior_desde_location_unificada(
            document, chain, 0.0, rebar_bar_type
        )
        return merged
    curvas_arm = []
    for e in chain:
        if e is None:
            continue
        ln, _n_f, _cara = _curva_armadura_inferior_en_fibra(
            document, e, 0.0, rebar_bar_type
        )
        if ln is not None:
            curvas_arm.append(ln)
    if not curvas_arm:
        return None
    merged = _unificar_lineas_colineales(curvas_arm, curvas_arm[0])
    if merged is None:
        return None
    return _expand_merged_line_with_location_endpoints(merged, chain)


def _largo_eje_mm_tras_unificar_extender_recortar(merged, obst_elems):
    """
    Longitud de eje (mm) tras el mismo pipeline previo al troceo por empalmes que en
    colocación: ``merged`` → :func:`_extender_linea_mm` (±:data:`_EXTENSION_ENDS_MM`) →
    recorte por caras de obstáculos → :func:`_recortar_extremos_linea_mm`
    (:data:`_TRIM_AXIS_ENDS_MM`) en un solo tramo continuo.
    """
    if merged is None:
        return None
    extended = _extender_linea_mm(merged, _EXTENSION_ENDS_MM)
    if extended is None:
        return None
    seg_obst = _trim_extremos_linea_por_interseccion_caras_obstaculo(
        extended, obst_elems or []
    )
    if seg_obst is None:
        return None
    seg_trim = _recortar_extremos_linea_mm(seg_obst, _TRIM_AXIS_ENDS_MM)
    if seg_trim is None:
        return None
    try:
        return float(seg_trim.Length) * 304.8
    except Exception:
        return None


def _build_chains_local(document, framing_list):
    if _build_collinear_chains_from_elements is not None:
        try:
            ch = _build_collinear_chains_from_elements(document, framing_list)
            if ch:
                return ch
        except Exception:
            pass
    return [[e] for e in framing_list]


def _extender_linea_mm(line, mm_each_side):
    if line is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        t = (p1 - p0).Normalize()
        d = _mm_to_ft(mm_each_side)
        e0 = p0 - t.Multiply(d)
        e1 = p1 + t.Multiply(d)
        return Line.CreateBound(e0, e1)
    except Exception:
        return None


def _recortar_extremos_linea_mm(line, mm_each_side):
    """Recorte simétrico en ambos extremos (un solo tramo o extremos libres de corrida)."""
    return _recortar_extremos_linea_mm_selectivo(
        line, mm_each_side, trim_start=True, trim_end=True
    )


def _recortar_extremos_linea_mm_selectivo(
    line, mm_each_side, trim_start=True, trim_end=True
):
    """
    Recorta ``mm_each_side`` solo en los extremos indicados.

    En corridas con traslape, los extremos hacia la junta **no** deben recortarse con el
    mismo margen de eje que los apoyos: si no, se acorta el solape medido (p. ej. 1140 → 1090 mm).
    """
    if line is None:
        return None
    try:
        L = float(line.Length)
        trim = _mm_to_ft(mm_each_side)
        ntrim = int(bool(trim_start)) + int(bool(trim_end))
        if ntrim == 0:
            return line
        if L <= float(ntrim) * trim + _MIN_LINE_LEN_FT:
            return None
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        t = (p1 - p0).Normalize()
        a = p0 + t.Multiply(trim) if trim_start else p0
        b = p1 - t.Multiply(trim) if trim_end else p1
        if (b - a).GetLength() < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(a, b)
    except Exception:
        return None


def _param_line_plane_intersection_distance(line, plane):
    """Distancia desde el origen de la línea al plano a lo largo de la tangente (0…Length)."""
    if line is None or plane is None:
        return None
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        d_raw = p1 - p0
        L = float(d_raw.GetLength())
        if L < _MIN_LINE_LEN_FT:
            return None
        du = d_raw.Normalize()
        n = plane.Normal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        o = plane.Origin
        denom = float(du.DotProduct(n))
        if abs(denom) < 1e-12:
            return None
        s = float((o - p0).DotProduct(n)) / denom
        if s < -_TOL_SPLIT_FT or s > L + _TOL_SPLIT_FT:
            return None
        return max(0.0, min(L, s))
    except Exception:
        return None


def _plano_division_desde_location_curve_fraccion_normalizada(elemento, u_norm):
    crv = _curva_location_framing(elemento)
    if crv is None:
        return None
    try:
        uu = float(u_norm)
        if uu < 0.0 or uu > 1.0:
            return None
        origin = crv.Evaluate(uu, True)
    except Exception:
        return None
    u = None
    try:
        der = crv.ComputeDerivatives(uu, True)
        if der is not None:
            tx = der.BasisX
            if tx is not None and tx.GetLength() >= 1e-12:
                u = tx.Normalize()
    except Exception:
        u = None
    if u is None:
        try:
            p0 = crv.GetEndPoint(0)
            p1 = crv.GetEndPoint(1)
            d = p1 - p0
            if d.GetLength() < _MIN_LINE_LEN_FT:
                return None
            u = d.Normalize()
        except Exception:
            return None
    try:
        return Plane.CreateByNormalAndOrigin(u, origin)
    except Exception:
        return None


def _plano_division_empalme_desde_location_curve(elemento):
    return _plano_division_desde_location_curve_fraccion_normalizada(elemento, 0.5)


def _parametros_corte_por_planos_empalme_location(curva, elementos_framing):
    out = []
    if curva is None or not elementos_framing:
        return out
    for el in elementos_framing:
        if el is None:
            continue
        pl = _plano_division_empalme_desde_location_curve(el)
        if pl is None:
            continue
        t = _param_line_plane_intersection_distance(curva, pl)
        if t is not None:
            out.append(t)
    return out


def _parametros_corte_por_planos_fracciones_location(curva, elementos_framing, fracciones):
    """
    Parámetros de corte (distancia desde el origen de ``curva``) por intersección con
    planos ``Plane.CreateByNormalAndOrigin(tangente, punto_en_curva)`` en cada fracción
    normalizada del LocationCurve de cada elemento (p. ej. 0,25 y 0,75).
    """
    out = []
    if curva is None or not elementos_framing or not fracciones:
        return out
    for el in elementos_framing:
        if el is None:
            continue
        for fn in fracciones:
            pl = _plano_division_desde_location_curve_fraccion_normalizada(el, fn)
            if pl is None:
                continue
            t = _param_line_plane_intersection_distance(curva, pl)
            if t is not None:
                out.append(t)
    return out


def _param_corte_suples_proyeccion_evaluate_sobre_linea(linea_trabajo, elemento, u_norm):
    """
    Distancia desde el origen de ``linea_trabajo`` hasta la proyección del punto de
    troceo **u_norm** sobre el segmento inicio→fin del eje con cotas estructurales
    en extremos (coherente con :func:`_plano_division_suple_desde_location_y_cotas_extremos`).
    """
    if linea_trabajo is None or elemento is None:
        return None
    try:
        uu = float(u_norm)
        if uu < 0.0 or uu > 1.0:
            return None
    except Exception:
        return None
    p_ext0, p_ext1 = _extremos_location_con_cotas_estructurales(elemento)
    pt = _punto_fraccion_extremos_con_cotas(p_ext0, p_ext1, uu)
    if pt is None:
        crv = _curva_location_framing(elemento)
        if crv is None:
            return None
        try:
            pt = crv.Evaluate(uu, True)
        except Exception:
            return None
    try:
        p0 = linea_trabajo.GetEndPoint(0)
        p1 = linea_trabajo.GetEndPoint(1)
        d_raw = p1 - p0
        L = float(d_raw.GetLength())
        if L < _MIN_LINE_LEN_FT:
            return None
        du = d_raw.Normalize()
        s = float((pt - p0).DotProduct(du))
        if s <= _TOL_SPLIT_FT or s >= L - _TOL_SPLIT_FT:
            return None
        return s
    except Exception:
        return None


def _parametros_corte_suples_evaluate_location_sobre_linea(
    linea_trabajo, elementos_chain, fracciones
):
    """Cortes suples: fracciones normalizadas del Location de cada viga de la cadena."""
    out = []
    if linea_trabajo is None or not elementos_chain or not fracciones:
        return out
    for el in elementos_chain:
        if el is None:
            continue
        for fn in fracciones:
            t = _param_corte_suples_proyeccion_evaluate_sobre_linea(
                linea_trabajo, el, fn
            )
            if t is not None:
                out.append(t)
    return out


def _linea_eje_recortada_tapas_viga_framing(elemento_framing):
    """
    Eje de la viga recortado entre **tapas** del sólido (misma lógica que la curva de estribos
    en ``geometria_estribos_viga``). Si no hay dos tapas, se usa la cuerda del ``LocationCurve``.
    """
    if elemento_framing is None:
        return None
    try:
        from geometria_estribos_viga import (
            _curva_location_framing,
            _linea_entre_tapas_extremas_viga,
            _solido_principal,
        )
    except Exception:
        return None
    crv = _curva_location_framing(elemento_framing)
    if crv is None:
        return None
    line_full = _line_bound_desde_location_curve(crv)
    if line_full is None:
        return None
    solid = _solido_principal(elemento_framing)
    if solid is None:
        return line_full
    line_work = _linea_entre_tapas_extremas_viga(line_full, solid)
    if line_work is None:
        return line_full
    return line_work


def _plano_division_suple_tapas_fraccion(elemento, u_norm):
    """
    Plano ⟂ eje en el punto ``u_norm`` ∈ [0,1] sobre el tramo **entre tapas** (estribos).
    """
    if elemento is None:
        return None
    try:
        uu = float(u_norm)
        if uu < 0.0 or uu > 1.0:
            return None
    except Exception:
        return None
    line_trim = _linea_eje_recortada_tapas_viga_framing(elemento)
    if line_trim is None:
        return None
    try:
        p0 = line_trim.GetEndPoint(0)
        p1 = line_trim.GetEndPoint(1)
        d = p1 - p0
        if d.GetLength() < _MIN_LINE_LEN_FT:
            return None
        u = d.Normalize()
        origin = p0 + d.Multiply(uu)
        return Plane.CreateByNormalAndOrigin(u, origin)
    except Exception:
        return None


def _parametros_corte_suples_tapas_sobre_linea(linea_trabajo, elementos_chain, fracciones):
    """
    Parámetros de corte en ``linea_trabajo``: planos en ``fracciones`` del eje **entre tapas**
    de cada viga (como estribos); intersección con la línea de suple. Respaldo: método con cotas
    estructurales si el plano tapas no corta.
    """
    out = []
    if linea_trabajo is None or not elementos_chain or not fracciones:
        return out
    for el in elementos_chain:
        if el is None:
            continue
        for fn in fracciones:
            pl = _plano_division_suple_tapas_fraccion(el, fn)
            if pl is None:
                tfb = _param_corte_suples_proyeccion_evaluate_sobre_linea(
                    linea_trabajo, el, fn
                )
                if tfb is not None:
                    out.append(tfb)
                continue
            t = _param_line_plane_intersection_distance(linea_trabajo, pl)
            if t is not None:
                out.append(t)
            else:
                tfb = _param_corte_suples_proyeccion_evaluate_sobre_linea(
                    linea_trabajo, el, fn
                )
                if tfb is not None:
                    out.append(tfb)
    return out


def _vector_en_plano_perpendicular_a_normal(plane_normal):
    """Vector unitario en el plano de corte (⟂ ``plane_normal``), para trazar el plano."""
    if plane_normal is None:
        return None
    try:
        u = plane_normal.Normalize()
    except Exception:
        return None
    if u is None or u.GetLength() < 1e-12:
        return None
    for w in (XYZ.BasisZ, XYZ.BasisX, XYZ.BasisY):
        try:
            v = w.CrossProduct(u)
            if v.GetLength() >= 1e-12:
                return v.Normalize()
        except Exception:
            continue
    return None


def _crear_marcadores_planos_division_suples(
    document, curva_suples, elementos_framing, fracciones, n_face_ref
):
    """
    ``ModelLine`` para ver dónde quedan los planos de división suples:

    - Segmento centrado en el **origen del plano** (fracción sobre inicio→fin con cotas
      estructurales en extremos; ver :func:`_plano_division_suple_desde_location_y_cotas_extremos`).
      Eje del segmento: dirección en el plano de corte (plano ⟂ eje viga).
    - Si hay intersección con ``curva_suples``, trazo corto en **Z** en ese punto
      (la curva de suples está desplazada frente al LocationCurve).

    Returns:
        Número de ``ModelCurve`` creadas.
    """
    if (
        not _DIBUJAR_MARCADORES_PLANOS_SUPLES
        or document is None
        or curva_suples is None
        or not elementos_framing
        or not fracciones
    ):
        return 0
    half = 0.5 * _mm_to_ft(_MARCADOR_PLANO_SUPLES_EN_PLANO_MM)
    h_tick = _mm_to_ft(_MARCADOR_PLANO_SUPLES_TICK_Z_MM)
    n = 0
    for el in elementos_framing:
        if el is None:
            continue
        for fn in fracciones:
            pl = _plano_division_suple_tapas_fraccion(el, fn)
            if pl is None:
                pl = _plano_division_suple_desde_location_y_cotas_extremos(el, fn)
            if pl is None:
                continue
            try:
                origin = pl.Origin
                npl = pl.Normal
            except Exception:
                continue
            inp = _vector_en_plano_perpendicular_a_normal(npl)
            if inp is None or half < 1e-12:
                continue
            try:
                seg_pl = Line.CreateBound(
                    origin - inp.Multiply(half),
                    origin + inp.Multiply(half),
                )
                if _crear_model_curve(document, seg_pl, n_face_ref):
                    n += 1
            except Exception:
                pass
            t_hit = _param_line_plane_intersection_distance(curva_suples, pl)
            if t_hit is None or h_tick < 1e-12:
                continue
            try:
                p0s = curva_suples.GetEndPoint(0)
                p1s = curva_suples.GetEndPoint(1)
                du = (p1s - p0s).Normalize()
                pt_hit = p0s + du.Multiply(float(t_hit))
                zseg = Line.CreateBound(
                    pt_hit,
                    pt_hit + XYZ.BasisZ.Multiply(h_tick),
                )
                if _crear_model_curve(document, zseg, du):
                    n += 1
            except Exception:
                pass
    return n


def _linea_desplazada_mm_reverso_normal_cara(line, normal_cara, mm):
    """
    Traslada ``line`` (``Line``) ``mm`` milímetros en la dirección opuesta a la normal de
    cara (interior típico si la normal apunta hacia afuera): ``−normalize(n)·mm``.
    """
    if line is None or normal_cara is None:
        return None
    try:
        n = normal_cara.Normalize()
        d = n.Multiply(-_mm_to_ft(float(mm)))
        p0 = line.GetEndPoint(0) + d
        p1 = line.GetEndPoint(1) + d
        if p0.DistanceTo(p1) < _MIN_LINE_LEN_FT:
            return None
        return Line.CreateBound(p0, p1)
    except Exception:
        return None


def _dedupe_sorted_cut_params(params, L):
    if L < _MIN_LINE_LEN_FT:
        return []
    out = []
    for x in sorted(params):
        if x <= _TOL_SPLIT_FT or x >= L - _TOL_SPLIT_FT:
            continue
        if not out or abs(x - out[-1]) > _TOL_SPLIT_FT:
            out.append(float(x))
    return out


def _split_line_by_distances(line, cuts):
    try:
        p0 = line.GetEndPoint(0)
        du = (line.GetEndPoint(1) - p0).Normalize()
        L = float(line.Length)
    except Exception:
        return []
    bounds = [0.0] + list(cuts) + [L]
    segs = []
    for i in range(len(bounds) - 1):
        a, b = bounds[i], bounds[i + 1]
        if b - a < _MIN_LINE_LEN_FT:
            continue
        try:
            qa = p0 + du.Multiply(a)
            qb = p0 + du.Multiply(b)
            segs.append(Line.CreateBound(qa, qb))
        except Exception:
            continue
    return segs


def _suple_superior_mantener_trozo(k_pc):
    """Suple superior: trozos impares en numeración 1-based (1, 3, 5…)."""
    return (k_pc + 1) % 2 != 0


def _suple_inferior_mantener_trozo(k_pc, n_pieces):
    """
    Suple inferior: trozos numerados **1…n** desde el inicio del ``work2``.
    Quedarse solo con los de numeración **par** (2, 4, 6…); eliminar los **impares**
    (1, 3, 5…). Índice ``k_pc`` 0-based → conservar si ``(k_pc + 1) % 2 == 0``.
    Un solo tramo (sin troceo útil) se conserva para no vaciar la corrida.
    """
    if n_pieces == 1:
        return True
    return (k_pc + 1) % 2 == 0


def _suple_fracciones_plan_troceo(es_inferior):
    """Fracciones del eje para planos ⟂ al Location que trocean el suple."""
    return (
        _SUPLES_INFERIOR_LOCATION_FRACCIONES
        if es_inferior
        else _SUPLES_LOCATION_FRACCIONES
    )


def _traslapo_longitudinal_mm_desde_bar_type(rebar_bar_type):
    """
    Largo **total** de traslape (mm) entre tramos consecutivos.

    Misma tabla e interpolación que ``lap_mm_para_bar_type`` en
    ``enfierrado_shaft_hashtag`` (Borde losa gancho / shaft: ``LAP_LENGTH_MM_...``).
    Si no importa ese módulo, respaldo ``40×Ø`` mm nominal.

    Returns:
        ``(lap_mm, texto_aviso)`` — ``texto_aviso`` opcional para el resumen.
    """
    if rebar_bar_type is None:
        return 0.0, None
    if lap_mm_para_bar_type is not None:
        try:
            lap, txt, _d = lap_mm_para_bar_type(rebar_bar_type)
            lap = float(lap)
            if lap > 0:
                return lap, (txt or None)
        except Exception:
            pass
    try:
        d = float(_rebar_nominal_diameter_mm(rebar_bar_type))
        if d <= 0:
            return 0.0, None
        lap = max(0.0, float(_TRASLAPO_LONG_MULT_DIAM_FALLBACK) * d)
        return lap, None
    except Exception:
        return 0.0, None


def _split_line_by_distances_con_traslapos_empalme(line, cuts, lap_mm):
    """
    Igual que :func:`_split_line_by_distances` pero en cada corte interior extiende
    los tramos ``± lap/2`` para generar solape entre barras consecutivas.

    Returns:
        ``(segmentos, idx_intervalo)``: ``idx_intervalo[k]`` es el índice del tramo
        en ``[0 .. len(cuts)]`` (mismo orden que ``bounds`` en el split), para enlazar
        el corte ``j`` con los rebars de los intervalos ``j`` y ``j+1``.
    """
    if not cuts:
        if line is None:
            return [], []
        return [line], [0]
    try:
        p0 = line.GetEndPoint(0)
        du = (line.GetEndPoint(1) - p0).Normalize()
        L = float(line.Length)
    except Exception:
        return [], []
    sc = sorted(float(x) for x in cuts)
    bounds = [0.0] + sc + [L]
    n_int = len(bounds) - 1
    lap_half_ft = 0.0
    if lap_mm and float(lap_mm) > 0:
        lap_half_ft = 0.5 * _mm_to_ft(float(lap_mm))
    segs = []
    idxs = []
    for i in range(n_int):
        a, b = bounds[i], bounds[i + 1]
        if lap_half_ft > 1e-12:
            if i > 0:
                a = max(0.0, a - lap_half_ft)
            if i < n_int - 1:
                b = min(L, b + lap_half_ft)
        if b - a < _MIN_LINE_LEN_FT:
            continue
        try:
            qa = p0 + du.Multiply(a)
            qb = p0 + du.Multiply(b)
            segs.append(Line.CreateBound(qa, qb))
            idxs.append(i)
        except Exception:
            continue
    return segs, idxs


def _puntos_segmento_traslape_sobre_work(work, cut_param_ft, lap_mm):
    """
    Extremos 3D del tramo de solape sobre la curva ``work`` (extendida), centrado en
    ``cut_param_ft`` y con longitud total ``lap_mm`` (mitad a cada lado del corte).
    """
    if work is None:
        return None, None
    try:
        p0 = work.GetEndPoint(0)
        du = (work.GetEndPoint(1) - p0).Normalize()
        L = float(work.Length)
    except Exception:
        return None, None
    half = 0.5 * _mm_to_ft(float(lap_mm)) if lap_mm and float(lap_mm) > 0 else 0.0
    if half < 1e-12:
        return None, None
    try:
        c = float(cut_param_ft)
    except Exception:
        return None, None
    a = max(0.0, c - half)
    b = min(L, c + half)
    if b - a < _mm_to_ft(2.0):
        return None, None
    try:
        return (p0 + du.Multiply(a), p0 + du.Multiply(b))
    except Exception:
        return None, None


def _colocar_detail_item_traslape_en_vista(document, view, family_symbol, p0_3d, p1_3d):
    """
    Coloca un detail component **line-based** en ``view``; ``p0_3d``/``p1_3d`` se proyectan
    al plano de corte de la vista (igual criterio que ``DetailCurve`` en este módulo).
    Returns:
        ``(ok, mensaje_error | None, instancia | None)``
    """
    if document is None or view is None or family_symbol is None:
        return False, u"Parámetros incompletos para detail de traslape.", None
    if p0_3d is None or p1_3d is None:
        return False, None, None
    pl = _obtener_plano_detalle_vista(view)
    if pl is None:
        return False, u"Sin plano de vista para proyectar el traslape.", None
    q0 = _proyectar_punto_al_plano(p0_3d, pl)
    q1 = _proyectar_punto_al_plano(p1_3d, pl)
    if q0 is None or q1 is None:
        return False, u"No se pudo proyectar el segmento de traslape.", None
    tol = max(_MIN_LINE_LEN_FT, _mm_to_ft(1.0))
    if q0.DistanceTo(q1) <= tol:
        return False, None, None
    try:
        ln = Line.CreateBound(q0, q1)
    except Exception:
        return False, u"Línea de detail inválida.", None
    try:
        if not bool(getattr(family_symbol, "IsActive", True)):
            family_symbol.Activate()
            try:
                document.Regenerate()
            except Exception:
                pass
    except Exception:
        pass
    try:
        inst = document.Create.NewFamilyInstance(ln, family_symbol, view)
        return True, None, inst
    except Exception as ex:
        try:
            return False, unicode(ex), None
        except Exception:
            return False, u"NewFamilyInstance (detail traslape)", None


def _iter_planar_faces_elemento(elemento):
    """``PlanarFace`` de la geometría del elemento (sólidos / instancias)."""
    if elemento is None:
        return
    opts = Options()
    opts.ComputeReferences = False
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        opts.IncludeNonVisibleObjects = True
    except Exception:
        pass
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
                faces = obj.Faces
            except Exception:
                continue
            for f in faces:
                if isinstance(f, PlanarFace):
                    yield f
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
                        faces = g2.Faces
                    except Exception:
                        continue
                    for f in faces:
                        if isinstance(f, PlanarFace):
                            yield f


def _plano_desde_face(face):
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        n = face.FaceNormal
        if n is None or n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        bb_uv = face.GetBoundingBox()
        if bb_uv is None:
            return None
        u_mid = 0.5 * (float(bb_uv.Min.U) + float(bb_uv.Max.U))
        v_mid = 0.5 * (float(bb_uv.Min.V) + float(bb_uv.Max.V))
        o = face.Evaluate(UV(u_mid, v_mid))
        if o is None:
            return None
        return Plane.CreateByNormalAndOrigin(n, o)
    except Exception:
        return None


def _punto_sobre_cara_planar(pt, face):
    if pt is None or face is None:
        return False
    try:
        r = face.Project(pt)
    except Exception:
        return False
    if r is None:
        return False
    try:
        d = float(r.Distance)
    except Exception:
        try:
            d = float(r.Distance)
        except Exception:
            return True
    return d <= _TOL_ON_FACE_FT


def _params_interseccion_linea_caras(line, faces):
    """Parámetros a lo largo de ``line`` (0…Length) donde corta una ``PlanarFace``."""
    out = []
    if line is None or not faces:
        return out
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        L = float(line.Length)
        if L < _MIN_LINE_LEN_FT:
            return out
    except Exception:
        return out
    for face in faces:
        if face is None or not isinstance(face, PlanarFace):
            continue
        try:
            if face.Area < _MIN_FACE_AREA_FT2:
                continue
        except Exception:
            pass
        pl = _plano_desde_face(face)
        if pl is None:
            continue
        t = _param_line_plane_intersection_distance(line, pl)
        if t is None:
            continue
        try:
            du = (p1 - p0).Normalize()
            q = p0 + du.Multiply(t)
        except Exception:
            continue
        if not _punto_sobre_cara_planar(q, face):
            continue
        out.append(float(t))
    return out


def _dedupe_sorted_params(params, L):
    """Une parámetros duplicados sin descartar cortes cerca de 0 o L (zonas de extendido)."""
    ps = sorted(params)
    out = []
    for x in ps:
        x = max(0.0, min(L, float(x)))
        if x <= 1e-10 or x >= L - 1e-10:
            continue
        if not out or abs(x - out[-1]) > _TOL_SPLIT_FT:
            out.append(x)
    return out


def _trim_extremos_linea_por_interseccion_caras_obstaculo_selectivo(
    line,
    elementos_obstaculos,
    aplicar_en_inicio=True,
    aplicar_en_fin=True,
    zona_extremo_mm=None,
    min_face_hits_en_extremo=None,
):
    """
    Igual que la versión completa, pero permite recortar solo el extremo inicial (cara
    de apoyo en el ``GetEndPoint(0)``), solo el final, o ambos. Útil para suples: el
    troceo en 0,25/0,75 no debe recortarse por pilares en los extremos **internos** del
    tramo.

    Args:
        zona_extremo_mm: Si no es ``None``, distancia (mm) desde cada extremo de la línea
            para clasificar intersecciones con caras de obstáculos (estribos / eje sin
            extender). Se limita a ``≤ 48%`` del largo del tramo para vigas cortas.
            Si es ``None``, se usa la zona extendida habitual (``_EXTENSION_ENDS_MM`` +
            ``_OBST_FACE_ZONE_EXTRA_MM``) del trazo longitudinal.
        min_face_hits_en_extremo: Umbral de cruces en la zona de cada extremo. Por defecto
            ``_MIN_FACE_HITS_EN_EXTREMO`` (2). **Suple inferior:** pasar ``1`` para recortar
            con la **primera colisión** en inicio y en fin.
    """
    if line is None:
        return None
    obs = [e for e in (elementos_obstaculos or []) if e is not None]
    if not obs:
        return line
    faces = []
    for el in obs:
        for f in _iter_planar_faces_elemento(el):
            faces.append(f)
    if not faces:
        return line
    params = _params_interseccion_linea_caras(line, faces)
    try:
        L = float(line.Length)
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        du = (p1 - p0).Normalize()
    except Exception:
        return None
    if L < _MIN_LINE_LEN_FT:
        return None
    ts = _dedupe_sorted_params(params, L)
    if zona_extremo_mm is not None:
        try:
            zona_ft = _mm_to_ft(float(zona_extremo_mm))
        except Exception:
            zona_ft = _mm_to_ft(_EXTENSION_ENDS_MM + float(_OBST_FACE_ZONE_EXTRA_MM))
        cap_ft = max(_MIN_LINE_LEN_FT * 3.0, float(L) * 0.48)
        zona_ft = min(zona_ft, cap_ft)
    else:
        zona_ft = _mm_to_ft(_EXTENSION_ENDS_MM + float(_OBST_FACE_ZONE_EXTRA_MM))

    hs = [t for t in ts if t <= zona_ft]
    he = [t for t in ts if t >= L - zona_ft]
    try:
        _nx = (
            int(min_face_hits_en_extremo)
            if min_face_hits_en_extremo is not None
            else _MIN_FACE_HITS_EN_EXTREMO
        )
        if _nx < 1:
            _nx = 1
    except Exception:
        _nx = _MIN_FACE_HITS_EN_EXTREMO

    t_lo = 0.0
    if aplicar_en_inicio and len(hs) >= _nx:
        t_lo = float(hs[0])

    t_hi = L
    if aplicar_en_fin and len(he) >= _nx:
        t_hi = float(he[-1])

    if t_hi <= t_lo + _MIN_LINE_LEN_FT:
        t_lo = 0.0
        t_hi = L

    if t_hi - t_lo < _MIN_LINE_LEN_FT:
        return None
    try:
        return Line.CreateBound(p0 + du.Multiply(t_lo), p0 + du.Multiply(t_hi))
    except Exception:
        return None


def _trim_extremos_linea_por_interseccion_caras_obstaculo(line, elementos_obstaculos):
    """
    Por cada extremo de la curva extendida: si en la zona cercana hay al menos
    ``_MIN_FACE_HITS_EN_EXTREMO`` intersecciones distintas con caras de pilares/muros,
    elimina el tramo exterior en ese extremo (punta más lejana). Evaluación independiente
    en inicio y fin sobre la misma curva.
    """
    return _trim_extremos_linea_por_interseccion_caras_obstaculo_selectivo(
        line, elementos_obstaculos, True, True
    )


def _sketch_plane_para_linea_en_modelo(line, n_hint):
    """Plano que contiene la línea (normal ⟂ tangente; ``n_hint`` desempata)."""
    try:
        p0 = line.GetEndPoint(0)
        p1 = line.GetEndPoint(1)
        t = (p1 - p0).Normalize()
    except Exception:
        return None
    bn = None
    if n_hint is not None:
        try:
            bn = t.CrossProduct(n_hint)
            if bn.GetLength() < 1e-12:
                bn = None
        except Exception:
            bn = None
    if bn is None or bn.GetLength() < 1e-12:
        try:
            bn = t.CrossProduct(XYZ.BasisZ)
            if bn.GetLength() < 1e-12:
                bn = t.CrossProduct(XYZ.BasisX)
        except Exception:
            return None
    try:
        bn = bn.Normalize()
        pl = Plane.CreateByNormalAndOrigin(bn, p0)
        return pl
    except Exception:
        return None


def _crear_model_curve(document, line, n_hint):
    pl = _sketch_plane_para_linea_en_modelo(line, n_hint)
    if pl is None:
        return False
    try:
        sp = SketchPlane.Create(document, pl)
        document.Create.NewModelCurve(line, sp)
        return True
    except Exception:
        return False


def _crear_marcadores_model_line_ejes_xz(document, seg, n_face_ref):
    """
    En el punto medio de la curva-guía de la barra: dos ``ModelLine`` cortos —
    uno según la **tangente** (eje X local / eje de la barra) y otro según **Z global**
    del proyecto (``XYZ.BasisZ``). Útil para verificar orientación en vistas y 3D.
    Si la tangente es casi paralela a Z, solo se dibuja el eje X local.
    Returns:
        Número de ``ModelCurve`` creadas (0, 1 ó 2).
    """
    if document is None or seg is None or not _DIBUJAR_MARCADORES_EJE_XZ:
        return 0
    t = _tangente_unitaria_linea(seg)
    if t is None:
        return 0
    try:
        mid = seg.Evaluate(0.5, True)
    except Exception:
        return 0
    Lm = _mm_to_ft(_MARCADOR_EJE_XZ_MM)
    if Lm < 1e-12:
        return 0
    n = 0
    try:
        ln_x = Line.CreateBound(mid, mid + t.Multiply(Lm))
        if _crear_model_curve(document, ln_x, n_face_ref):
            n += 1
    except Exception:
        pass
    zdir = XYZ.BasisZ
    try:
        if abs(float(t.DotProduct(zdir))) > float(_MARCADOR_EJE_XZ_PARALELO_TOL):
            return n
        ln_z = Line.CreateBound(mid, mid + zdir.Multiply(Lm))
        if _crear_model_curve(document, ln_z, t):
            n += 1
    except Exception:
        pass
    return n


def _rebar_bar_positions(rebar, fallback):
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return max(1, int(fallback or 1))


def estimar_largo_mm_trazo_superior(
    document, framing_list, obstaculos, rebar_bar_type=None
):
    """
    Estimación (mm) del largo **físico** máximo por cadena: curva ya unificada en cara
    superior, luego estiramiento ±:data:`_EXTENSION_ENDS_MM`, recortes por obstáculos y
    recorte de eje (:data:`_TRIM_AXIS_ENDS_MM`) — igual que un tramo único antes del
    troceo por empalmes — más **dos patas** (:mod:`bimtools_rebar_hook_lengths`).
    Sirve para el panel de empalmes (> 12 m **incl. ganchos**).
    """
    if document is None or not framing_list:
        return None
    obst_elems = obstaculos or []
    chains = _build_chains_local(document, framing_list)
    max_total_mm = 0.0
    for chain in chains:
        merged = _merged_armadura_superior_cara_para_cadena(
            document, chain, rebar_bar_type
        )
        if merged is None:
            continue
        L_eje_mm = _largo_eje_mm_tras_unificar_extender_recortar(merged, obst_elems)
        if L_eje_mm is None:
            L_eje_mm = _largo_eje_mm_max_fibra_y_location_cadena(chain, merged)
        L_total_mm = _largo_total_barra_mm_eje_mas_dos_ganchos(
            document, rebar_bar_type, L_eje_mm
        )
        max_total_mm = max(max_total_mm, float(L_total_mm))
    if max_total_mm <= 1e-9:
        return None
    return float(max_total_mm)


def estimar_largo_mm_trazo_inferior(
    document, framing_list, obstaculos, rebar_bar_type=None
):
    """
    Análogo a :func:`estimar_largo_mm_trazo_superior` para el trazo en cara inferior
    (cadena unificada + extensión, recortes y ganchos).
    """
    if document is None or not framing_list:
        return None
    obst_elems = obstaculos or []
    chains = _build_chains_local(document, framing_list)
    max_total_mm = 0.0
    for chain in chains:
        merged = _merged_armadura_inferior_cara_para_cadena(
            document, chain, rebar_bar_type
        )
        if merged is None:
            continue
        L_eje_mm = _largo_eje_mm_tras_unificar_extender_recortar(merged, obst_elems)
        if L_eje_mm is None:
            L_eje_mm = _largo_eje_mm_max_fibra_y_location_cadena(chain, merged)
        L_total_mm = _largo_total_barra_mm_eje_mas_dos_ganchos(
            document, rebar_bar_type, L_eje_mm
        )
        max_total_mm = max(max_total_mm, float(L_total_mm))
    if max_total_mm <= 1e-9:
        return None
    return float(max_total_mm)


def crear_detail_lines_largo_cara_superior_en_vista(
    document,
    view,
    elementos_framing,
    elementos_obstaculos=None,
    rebar_bar_type=None,
    rebar_cantidad=1,
    framing_empalme_element_ids=None,
    aplicar_troceo_empalmes_framing=True,
    n_capas_superiores=1,
    rebar_bar_types_capas=None,
    rebar_cantidades_capas=None,
    rebar_bar_type_suple=None,
    rebar_cantidad_suple=1,
    crear_armadura_suple=False,
    n_capas_suple=1,
    rebar_bar_types_suple=None,
    rebar_cantidades_suple=None,
    rebar_espaciado_mm=None,
    rebar_espaciado_suple_mm=None,
    es_cara_inferior=False,
    gestionar_transaccion=True,
    crear_laterales_cara_superior=False,
    laterales_rebar_bar_type=None,
    laterales_cantidad=1,
    crear_model_lines=True,
    recubrimiento_extra_estribo_mm=0.0,
    crear_model_line_eje_curva_suple_troceo=False,
):
    """
    Crea ``Rebar`` (y opcionalmente ``ModelCurve`` / ``DetailCurve`` / detail item de
    traslape según flags del módulo). El detail item usa la misma familia/tipo que
    ``barras_bordes_losa_gancho_empotramiento._find_fixed_lap_detail_symbol_id``.
    Ganchos: siempre el ``RebarHookType`` nombrado
    ``HOOK_GANCHO_90_STANDARD_NAME`` (p. ej. Standard - 90 deg.) en los extremos
    donde corresponda gancho.

    **Capas superiores:** ``n_capas_superiores`` en 1–3; la capa 1 usa el trazo habitual;
    cada capa adicional traslada la curva ``_OFFSET_SUPLES_SEGUNDA_CAPA_MM`` mm en ``−n``.
    Si se pasan ``rebar_bar_types_capas`` / ``rebar_cantidades_capas`` (listas de longitud
    ``n_capas_superiores``), cada capa usa su tipo y cantidad; si no, se repiten
    ``rebar_bar_type`` y ``rebar_cantidad``.

    **Suple:** mismo eje que la 1.ª capa (``work`` estirado ±2 m, etc.) desplazado
    ``(N_capas_main + k) × _OFFSET_SUPLES_SEGUNDA_CAPA_MM`` mm en ``−n`` para cada
    capa de suple ``k`` (0-based); ``n_capas_suple`` en 1–2 con tipos/cantidades en
    listas opcionales. **Suple superior:** curva recortada (obstáculos + eje); troceo **0,25/0,75**
    del eje entre **tapas** por viga (como estribos); trozos **pares** descartados.
    **Suple inferior:** la curva de suple se **recorta** antes del troceo: **primera colisión**
    con obstáculos en cada zona de extremo (un solo cruce por lado, no dos), luego
    ``_TRIM_AXIS_ENDS_MM``. Después los planos en **0,10** y **0,90** del eje entre tapas
    (estribos), proyectados a esa curva. Trozos **pares** (2, 4, …).

    Si ``rebar_espaciado_mm`` (o ``rebar_espaciado_suple_mm``) es > 0, la cantidad en ancho
    se deriva del ancho útil y el espaciado; si no aplica, se usan ``rebar_cantidad`` /
    ``rebar_cantidad_suple``.

    **Cara inferior:** mismo eje en plano inferior; con troceo por empalmes se colocan
    también **detail item de traslape** y **cota** (misma familia y lógica que en superior).

    Si ``gestionar_transaccion`` es ``False``, no se abre ni confirma ``Transaction`` aquí:
    debe existir una transacción ya iniciada en el documento (p. ej. cara superior + inferior
    en una sola operación).

    **Laterales:** si ``crear_laterales_cara_superior`` (UI: *Barras laterales*, según se coloque
    cara superior y/o inferior) y hay ``laterales_rebar_bar_type``, por tramo un ``Rebar`` en la
    curva guía (:func:`_linea_guia_laterales_cara_superior`; misma lógica para ambas caras),
    ``norm`` preferente :func:`_norm_createfromcurves_reversa_normal_cara_superior` (−n; respaldos
    si ``n ∥`` tramo), sin ganchos, ``SetLayoutAsFixedNumber``, canto − 2×
    :func:`_offset_mm_curva_laterales_vs_cara_superior` (incl. ``recubrimiento_extra_estribo_mm``).
    **Recubrimiento extra estribo:** ``recubrimiento_extra_estribo_mm`` (típ. Ø nominal del estribo)
    en offsets cara→eje y, en ancho, como ``2×Ø_estribo`` dentro de
    :func:`_layout_max_spacing_array_length_ft` (**ancho − 50 − 2×Ø_estribo − Ø_long** mm).

    ``crear_model_lines``: si ``False``, no se crean ``ModelCurve``/``ModelLine`` de guía
    (proyección cara paralela, tramo, marcadores XZ, planos suple); no afecta a ``Rebar`` ni
    a ``DetailCurve`` ni a detail items de traslape.

    ``crear_model_line_eje_curva_suple_troceo``: si ``True``, por cada capa de **suple** se crea
    una ``ModelLine`` sobre la curva desplazada ``work2`` (eje donde se proyectan los cortes
    0,25/0,75 o 0,10/0,90) para verificar planos de división; **independiente** de
    ``crear_model_lines``.

    Returns:
        ``tuple``: (n_model_curves, n_rebar_positions, avisos)
    """
    avisos = []
    _es_inf = bool(es_cara_inferior)
    _dibujar_ml = bool(crear_model_lines)
    _ml_eje_suple_troceo = bool(crear_model_line_eje_curva_suple_troceo)
    try:
        rex_mm = max(0.0, float(recubrimiento_extra_estribo_mm or 0.0))
    except Exception:
        rex_mm = 0.0
    if document is None:
        return 0, 0, [u"No hay documento."]
    framing_list = [e for e in (elementos_framing or []) if e is not None]
    if not framing_list:
        return 0, 0, [u"Sin vigas."]

    obst_elems = elementos_obstaculos or []

    emp_ids = list(framing_empalme_element_ids or [])
    emp_elems = []
    for eid in emp_ids:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None or not el.IsValidObject:
            continue
        try:
            if el.Category and int(el.Category.Id.IntegerValue) == _FRAMING_CAT:
                emp_elems.append(el)
        except Exception:
            pass

    plane_view = None
    if view is not None and vista_permite_detail_curve(view):
        plane_view = _obtener_plano_detalle_vista(view)

    n_curves = 0
    n_rebar = 0

    chains = _build_chains_local(document, framing_list)

    t = None
    if gestionar_transaccion:
        t = Transaction(
            document,
            (
                u"BIMTools — Cara inferior vigas (detalle)"
                if _es_inf
                else u"BIMTools — Cara superior vigas (detalle)"
            ),
        )
        t.Start()
    try:
        lap_detail_symbol = None
        n_lap_details = 0
        n_lap_cotas = 0
        aviso_refs_lap_familia = None
        rebar_ids_barrido_ganchos = []
        if (
            _DIBUJAR_DETAIL_ITEM_TRASLAPE
            and view is not None
            and vista_permite_detail_curve(view)
            and _find_fixed_lap_detail_symbol_id is not None
        ):
            sid, s_err = _find_fixed_lap_detail_symbol_id(document)
            if sid is not None:
                try:
                    sy = document.GetElement(sid)
                    if isinstance(sy, FamilySymbol):
                        lap_detail_symbol = sy
                except Exception:
                    pass
            if lap_detail_symbol is None and s_err:
                avisos.append(
                    s_err + u" No se colocarán detail items de traslape."
                )

        try:
            _nc_raw = n_capas_superiores
            try:
                _nc_raw = int(_nc_raw)
            except Exception:
                _nc_raw = 1
            _n_capas_sup_def = max(1, min(3, int(_nc_raw or 1)))
        except Exception:
            _n_capas_sup_def = 1

        try:
            _nsr_suple = int(n_capas_suple or 1)
        except Exception:
            _nsr_suple = 1
        _n_suple_lay = max(1, min(2, _nsr_suple))
        _tlist_suple = (
            list(rebar_bar_types_suple) if rebar_bar_types_suple else []
        )
        _clist_suple = (
            list(rebar_cantidades_suple) if rebar_cantidades_suple else []
        )
        _suple_resuelve_tipo = rebar_bar_type_suple is not None
        if _tlist_suple:
            for _si_chk in range(_n_suple_lay):
                _t_one = (
                    _tlist_suple[_si_chk]
                    if _si_chk < len(_tlist_suple)
                    else None
                )
                if _t_one is not None:
                    _suple_resuelve_tipo = True
                    break
        if crear_armadura_suple and not _suple_resuelve_tipo:
            avisos.append(
                u"Suples: activados pero sin RebarBarType (diámetro); no se crea el suple."
            )

        _lbl_cara = u"inferior" if _es_inf else u"superior"
        for chain in chains:
            host0 = chain[0] if chain else None
            if _es_inf:
                if _TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA:
                    merged, n_face_ref, _ = (
                        _curva_armadura_inferior_desde_location_unificada(
                            document,
                            chain,
                            rex_mm,
                            rebar_bar_type,
                        )
                    )
                else:
                    merged = n_face_ref = None
            else:
                if _TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA:
                    merged, n_face_ref, _ = (
                        _curva_armadura_superior_desde_location_unificada(
                            document,
                            chain,
                            rex_mm,
                            rebar_bar_type,
                        )
                    )
                else:
                    merged = n_face_ref = None
            if _es_inf:
                if _TRAZO_INFERIOR_USAR_LOCATION_UNIFICADA:
                    if merged is None or n_face_ref is None:
                        try:
                            ids_txt = u", ".join(
                                unicode(getattr(el.Id, "IntegerValue", u""))
                                for el in chain
                                if el is not None
                            )
                        except Exception:
                            ids_txt = u""
                        avisos.append(
                            u"Cadena [{0}]: sin eje Location/unificación u offsets (cara {1}).".format(
                                ids_txt,
                                _lbl_cara,
                            )
                        )
                        continue
                else:
                    curvas_arm = []
                    n_face_ref = None
                    for e in chain:
                        ln, n_f, _cara = _curva_armadura_inferior_en_fibra(
                            document, e, rex_mm, rebar_bar_type
                        )
                        if ln is None:
                            avisos.append(
                                u"Viga Id {0}: sin geometría de fibra {1}.".format(
                                    getattr(e.Id, "IntegerValue", u""),
                                    _lbl_cara,
                                )
                            )
                            continue
                        curvas_arm.append(ln)
                        if n_face_ref is None:
                            n_face_ref = n_f
                    if not curvas_arm or n_face_ref is None:
                        continue
                    merged = _unificar_lineas_colineales(
                        curvas_arm, curvas_arm[0]
                    )
                    if merged is None:
                        continue
                    merged = _expand_merged_line_with_location_endpoints(
                        merged, chain
                    )
            else:
                if _TRAZO_SUPERIOR_USAR_LOCATION_UNIFICADA:
                    if merged is None or n_face_ref is None:
                        try:
                            ids_txt = u", ".join(
                                unicode(getattr(el.Id, "IntegerValue", u""))
                                for el in chain
                                if el is not None
                            )
                        except Exception:
                            ids_txt = u""
                        avisos.append(
                            u"Cadena [{0}]: sin eje Location/unificación u offsets (cara {1}).".format(
                                ids_txt,
                                _lbl_cara,
                            )
                        )
                        continue
                else:
                    curvas_arm = []
                    n_face_ref = None
                    for e in chain:
                        ln, n_f, _cara = _curva_armadura_superior_en_fibra(
                            document, e, rex_mm, rebar_bar_type
                        )
                        if ln is None:
                            avisos.append(
                                u"Viga Id {0}: sin geometría de fibra {1}.".format(
                                    getattr(e.Id, "IntegerValue", u""),
                                    _lbl_cara,
                                )
                            )
                            continue
                        curvas_arm.append(ln)
                        if n_face_ref is None:
                            n_face_ref = n_f
                    if not curvas_arm or n_face_ref is None:
                        continue
                    merged = _unificar_lineas_colineales(
                        curvas_arm, curvas_arm[0]
                    )
                    if merged is None:
                        continue
                    merged = _expand_merged_line_with_location_endpoints(
                        merged, chain
                    )

            layout_span_ft = _layout_max_spacing_array_length_ft(
                document,
                host0,
                rebar_bar_type,
                diametro_estribo_mm=rex_mm,
            )

            L_eje_mm = _largo_eje_mm_tras_unificar_extender_recortar(
                merged, obst_elems
            )
            if L_eje_mm is None:
                L_eje_mm = _largo_eje_mm_max_fibra_y_location_cadena(chain, merged)
            L_total_mm = _largo_total_barra_mm_eje_mas_dos_ganchos(
                document, rebar_bar_type, L_eje_mm
            )
            _um_adj = float(_UMBRAL_LARGO_EMPALMES_MM) - float(
                _UMBRAL_EMPALMES_COMP_EPS_MM
            )
            need_warn_12 = L_total_mm > _um_adj and not emp_elems
            if need_warn_12:
                avisos.append(
                    u"Trazo eje ≈ {0:.0f} mm + ganchos → ≈ {1:.0f} mm > 12 m: "
                    u"defina vigas de empalme para trocear.".format(L_eje_mm, L_total_mm)
                )

            extended = _extender_linea_mm(merged, _EXTENSION_ENDS_MM)
            if extended is None:
                continue
            work = extended
            cuts = []
            if (
                aplicar_troceo_empalmes_framing
                and emp_elems
                and L_total_mm > _um_adj
            ):
                cuts.extend(
                    _parametros_corte_por_planos_empalme_location(work, emp_elems)
                )
            cuts = _dedupe_sorted_cut_params(cuts, float(work.Length))
            lap_mm = 0.0
            lap_txt = None
            if cuts:
                lap_mm, lap_txt = _traslapo_longitudinal_mm_desde_bar_type(rebar_bar_type)
                if lap_mm <= 0 and rebar_bar_type is not None:
                    avisos.append(
                        u"Empalme: no se pudo obtener Ø nominal; traslape no aplicado."
                    )
                elif lap_mm > 0:
                    if lap_txt:
                        avisos.append(lap_txt)
                    else:
                        try:
                            avisos.append(
                                u"Traslape en juntas de troceo: ≈ {0:.0f} mm (respaldo {1:.0f}×Ø).".format(
                                    lap_mm,
                                    float(_TRASLAPO_LONG_MULT_DIAM_FALLBACK),
                                )
                            )
                        except Exception:
                            pass
            if cuts:
                pieces, piece_interval_idx = _split_line_by_distances_con_traslapos_empalme(
                    work, cuts, lap_mm
                )
            else:
                pieces = [work]
                piece_interval_idx = [0]

            final_segs = []
            final_interval_indices = []
            n_pc = len(pieces)
            for i_pc, pc in enumerate(pieces):
                interval_i = piece_interval_idx[i_pc]
                seg_obst = _trim_extremos_linea_por_interseccion_caras_obstaculo(
                    pc, obst_elems
                )
                if seg_obst is None:
                    continue
                if n_pc <= 1:
                    seg_trim = _recortar_extremos_linea_mm(
                        seg_obst, _TRIM_AXIS_ENDS_MM
                    )
                else:
                    seg_trim = _recortar_extremos_linea_mm_selectivo(
                        seg_obst,
                        _TRIM_AXIS_ENDS_MM,
                        trim_start=(i_pc == 0),
                        trim_end=(i_pc == n_pc - 1),
                    )
                if seg_trim is not None:
                    final_segs.append(seg_trim)
                    final_interval_indices.append(interval_i)

            if not final_segs:
                avisos.append(u"Cadena sin tramos válidos tras troceo/recortes.")
                continue

            _nc = int(_n_capas_sup_def)
            if rebar_bar_types_capas is None:
                _types_cap = [
                    rebar_bar_type for _ in range(_nc)
                ]
            else:
                try:
                    _rbtc_seq = list(rebar_bar_types_capas)
                except Exception:
                    _rbtc_seq = []
                _types_cap = []
                for _ci in range(_nc):
                    if _ci < len(_rbtc_seq):
                        _types_cap.append(_rbtc_seq[_ci])
                    else:
                        _types_cap.append(rebar_bar_type)
                for _ti in range(len(_types_cap)):
                    if _types_cap[_ti] is None:
                        if rebar_bar_type is not None and _ti > 0:
                            try:
                                avisos.append(
                                    u"Capa superior {0}: no se resolvió el diámetro en UI; "
                                    u"se usa el tipo de la 1.ª capa.".format(_ti + 1)
                                )
                            except Exception:
                                pass
                        _types_cap[_ti] = rebar_bar_type
            if rebar_cantidades_capas is None:
                _cants_cap = [
                    max(1, int(rebar_cantidad or 1)) for _ in range(_nc)
                ]
            else:
                try:
                    _rcc_seq = list(rebar_cantidades_capas)
                except Exception:
                    _rcc_seq = []
                _cants_cap = []
                for _ci in range(_nc):
                    if _ci < len(_rcc_seq):
                        _cants_cap.append(max(1, int(_rcc_seq[_ci] or 1)))
                    else:
                        _cants_cap.append(max(1, int(rebar_cantidad or 1)))
            _fallback_cant_capa = max(1, int(rebar_cantidad or 1))
            while len(_types_cap) < _nc:
                _types_cap.append(rebar_bar_type)
            while len(_cants_cap) < _nc:
                _cants_cap.append(
                    _cants_cap[-1] if _cants_cap else _fallback_cant_capa
                )
            del _types_cap[_nc:]
            del _cants_cap[_nc:]

            # Misma normal CreateFromCurves que si el trazo fuera un solo tramo (curva unificada
            # previa al troceo); no recalcular por cada ``seg`` tras traslapos.
            norm_prio_rebar = _norm_createfromcurves_desde_cara_y_tramo(n_face_ref, merged)
            if norm_prio_rebar is None:
                try:
                    norm_prio_rebar = _norm_createfromcurves_desde_cara_y_tramo(
                        n_face_ref, work
                    )
                except Exception:
                    norm_prio_rebar = None

            n_final = len(final_segs)
            host_candidatos_emp = _candidatos_host_empalme(chain, emp_elems)
            step_capas_mm = float(_OFFSET_SUPLES_SEGUNDA_CAPA_MM)
            if _dibujar_ml and _DIBUJAR_MODEL_LINE_PROYECCION_CARA_PARALELA:
                for _idx_seg_ml, seg_ml in enumerate(final_segs):
                    h_ml = host0
                    if n_final > 1:
                        pm_ml = _punto_medio_linea_bound(seg_ml)
                        if pm_ml is not None:
                            h_ml = _host_framing_para_segmento_rebar(
                                pm_ml,
                                host_candidatos_emp,
                                host0,
                            )
                    norm_ml = norm_prio_rebar
                    if norm_ml is None:
                        norm_ml = _norm_createfromcurves_desde_cara_y_tramo(
                            n_face_ref, seg_ml
                        )
                    n_curves += (
                        _crear_model_line_cara_superior_offset_y_marcador_normal(
                            document,
                            h_ml,
                            seg_ml,
                            n_face_ref,
                            norm_ml,
                            _n_capas_sup_def,
                            step_capas_mm,
                            refinar_n_hint_con_cara_superior_framing=(
                                not _es_inf
                            ),
                        )
                    )
            rebar_by_interval = {}
            for capa_idx in range(_n_capas_sup_def):
                rta_c = _types_cap[capa_idx]
                cantidad_c = _cants_cap[capa_idx]
                layout_span_ft_c = _layout_max_spacing_array_length_ft(
                    document,
                    host0,
                    rta_c,
                    diametro_estribo_mm=rex_mm,
                )
                span_mm_layout_c = None
                try:
                    if (
                        layout_span_ft_c is not None
                        and float(layout_span_ft_c) > 1e-12
                    ):
                        span_mm_layout_c = float(layout_span_ft_c) * 304.8
                except Exception:
                    span_mm_layout_c = None
                if rebar_espaciado_mm is not None:
                    try:
                        esp_u = float(rebar_espaciado_mm)
                    except Exception:
                        esp_u = 0.0
                    if esp_u > 0 and span_mm_layout_c is not None:
                        c_esp = _cantidad_barras_desde_espaciado_transversal_mm(
                            span_mm_layout_c, esp_u
                        )
                        if c_esp is not None:
                            cantidad_c = c_esp

                off_mm = float(capa_idx) * step_capas_mm
                for idx_seg, seg in enumerate(final_segs):
                    seg_use = seg
                    if off_mm > 1e-9:
                        seg_use = _linea_desplazada_mm_reverso_normal_cara(
                            seg, n_face_ref, off_mm
                        )
                    if seg_use is None:
                        avisos.append(
                            u"Armadura {0} capa {1}: tramo {2} sin geometría tras "
                            u"desplazar {3:.0f} mm en −n.".format(
                                _lbl_cara,
                                capa_idx + 1,
                                idx_seg + 1,
                                off_mm,
                            )
                        )
                        continue
                    if _dibujar_ml and _DIBUJAR_MODEL_LINE_TRAMO:
                        if _crear_model_curve(document, seg_use, n_face_ref):
                            n_curves += 1
                    if _dibujar_ml and _DIBUJAR_MARCADORES_EJE_XZ:
                        n_curves += _crear_marcadores_model_line_ejes_xz(
                            document, seg_use, n_face_ref
                        )
                    if _DIBUJAR_DETAIL_CURVE_VISTA and plane_view is not None:
                        try:
                            dv = _proyectar_linea_al_plano_vista(
                                seg_use, plane_view
                            )
                            if dv is not None:
                                document.Create.NewDetailCurve(view, dv)
                        except Exception:
                            pass

                    interval_i = final_interval_indices[idx_seg]
                    if rta_c is not None:
                        if crear_rebar_desde_curva_linea_con_ganchos is None:
                            avisos.append(
                                u"Módulo rebar_fundacion_cara_inferior no disponible."
                            )
                            break
                        host = host0
                        if n_final > 1:
                            pm = _punto_medio_linea_bound(seg_use)
                            host = _host_framing_para_segmento_rebar(
                                pm, host_candidatos_emp, host0
                            )
                        norm_prio = norm_prio_rebar
                        if off_mm > 1e-9:
                            n_seg = _norm_createfromcurves_desde_cara_y_tramo(
                                n_face_ref, seg_use
                            )
                            if n_seg is not None:
                                norm_prio = n_seg
                        if norm_prio is None:
                            norm_prio = _norm_createfromcurves_desde_cara_y_tramo(
                                n_face_ref, seg_use
                            )
                        try:
                            z_hook_ref = (
                                n_face_ref.Normalize().Negate()
                                if n_face_ref is not None
                                else XYZ.BasisZ.Negate()
                            )
                        except Exception:
                            z_hook_ref = XYZ.BasisZ.Negate()
                        # Con varios tramos (traslapes/empalmes): ganchos solo en extremos libres
                        # de la corrida (inicio del primer tramo, fin del último).
                        _un_tramo = n_final <= 1
                        _gi = _un_tramo or (idx_seg == 0)
                        _gf = _un_tramo or (idx_seg == n_final - 1)
                        norm_bar_create = None
                        if norm_prio is not None:
                            try:
                                norm_bar_create = norm_prio.Negate()
                            except Exception:
                                norm_bar_create = norm_prio
                        rb, err, _nv = crear_rebar_desde_curva_linea_con_ganchos(
                            document,
                            host,
                            rta_c,
                            seg_use,
                            hook_type_name=HOOK_GANCHO_90_STANDARD_NAME,
                            marco_cara_uvn=None,
                            cara_paralela=None,
                            eje_referencia_z_ganchos=z_hook_ref,
                            normales_prioridad=(
                                [norm_bar_create]
                                if norm_bar_create is not None
                                else None
                            ),
                            gancho_en_inicio=_gi,
                            gancho_en_fin=_gf,
                        )
                        if rb is None:
                            if err:
                                avisos.append(err)
                        else:
                            rebar_by_interval.setdefault(capa_idx, {})[
                                interval_i
                            ] = rb
                            try:
                                rebar_ids_barrido_ganchos.append(rb.Id)
                            except Exception:
                                pass
                            try:
                                if aplicar_layout_fixed_number_rebar is not None:
                                    span_ft = layout_span_ft_c
                                    if n_final > 1 and host is not None:
                                        try:
                                            ls_h = (
                                                _layout_max_spacing_array_length_ft(
                                                    document,
                                                    host,
                                                    rta_c,
                                                    diametro_estribo_mm=rex_mm,
                                                )
                                            )
                                            if (
                                                ls_h is not None
                                                and float(ls_h) > 1e-12
                                            ):
                                                span_ft = ls_h
                                        except Exception:
                                            pass
                                    if (
                                        cantidad_c > 1
                                        and (
                                            span_ft is None
                                            or float(span_ft) < 1e-12
                                        )
                                    ):
                                        diam = float(
                                            _rebar_nominal_diameter_mm(rta_c)
                                            or 16.0
                                        )
                                        sep_mm = max(25.0, diam + 25.0)
                                        span_ft = _mm_to_ft(
                                            sep_mm * float(cantidad_c - 1)
                                        )
                                    ok_lay, err_lay = (
                                        aplicar_layout_fixed_number_rebar(
                                            rb,
                                            document,
                                            cantidad_c,
                                            float(span_ft or 0.0),
                                        )
                                    )
                                    if not ok_lay and err_lay:
                                        avisos.append(err_lay)
                            except Exception:
                                pass
                            if (
                                (_gi or _gf)
                                and _enforce_rebar_hook_types_by_name is not None
                            ):
                                try:
                                    _enforce_rebar_hook_types_by_name(
                                        rb,
                                        document,
                                        HOOK_GANCHO_90_STANDARD_NAME,
                                        _gi,
                                        _gf,
                                        avisos=avisos,
                                    )
                                except Exception:
                                    pass
                            n_rebar += _rebar_bar_positions(rb, cantidad_c)

            if (
                crear_laterales_cara_superior
                and laterales_rebar_bar_type is not None
                and crear_rebar_desde_curva_linea_con_ganchos is not None
                and aplicar_layout_fixed_number_rebar is not None
            ):
                _lat_tag = u"Laterales (inferior)" if _es_inf else u"Laterales"
                n_lat_eff = max(1, int(laterales_cantidad or 1))
                for seg_lat_src in final_segs:
                    seg_lat = _linea_guia_laterales_cara_superior(
                        seg_lat_src,
                        n_face_ref,
                        _n_capas_sup_def,
                        step_capas_mm,
                    )
                    if seg_lat is None:
                        continue
                    host_l = host0
                    if n_final > 1:
                        pm_l = _punto_medio_linea_bound(seg_lat)
                        if pm_l is not None:
                            host_l = _host_framing_para_segmento_rebar(
                                pm_l,
                                host_candidatos_emp,
                                host0,
                            )
                    crv_h = _curva_location_framing(host_l)
                    if crv_h is None:
                        crv_h = seg_lat_src
                    try:
                        _wf_unused, d_ft = _read_width_depth_ft_local(
                            document, host_l, crv_h
                        )
                        h_mm = float(d_ft) * 304.8
                    except Exception:
                        h_mm = 0.0
                    off_vs_cara = _offset_mm_curva_laterales_vs_cara_superior(
                        _n_capas_sup_def,
                        step_capas_mm,
                        rex_mm,
                        rebar_bar_type,
                    )
                    span_mm = float(h_mm) - 2.0 * float(off_vs_cara)
                    if span_mm < 5.0:
                        avisos.append(
                            u"{0}: canto útil insuficiente (altura {1:.0f} mm, "
                            u"2×offset {2:.0f} mm); tramo omitido.".format(
                                _lat_tag,
                                h_mm,
                                2.0 * off_vs_cara,
                            )
                        )
                        continue
                    span_ft = _mm_to_ft(span_mm)
                    # ``norm`` prioritario: **−n** (normal exterior de la cara activa).
                    norm_ml = _norm_createfromcurves_reversa_normal_cara_superior(
                        n_face_ref,
                        seg_lat,
                    )
                    if norm_ml is None:
                        norm_ml = _eje_y_interno_curva_createfromcurves(seg_lat)
                    if norm_ml is None:
                        norm_ml = _norm_createfromcurves_desde_cara_y_tramo(
                            n_face_ref,
                            seg_lat,
                        )
                    try:
                        z_hook_ref_l = (
                            n_face_ref.Normalize().Negate()
                            if n_face_ref is not None
                            else XYZ.BasisZ.Negate()
                        )
                    except Exception:
                        z_hook_ref_l = XYZ.BasisZ.Negate()
                    rb_lat, err_lat, _nv_lat = (
                        crear_rebar_desde_curva_linea_con_ganchos(
                            document,
                            host_l,
                            laterales_rebar_bar_type,
                            seg_lat,
                            hook_type_name=HOOK_GANCHO_90_STANDARD_NAME,
                            marco_cara_uvn=None,
                            cara_paralela=None,
                            eje_referencia_z_ganchos=z_hook_ref_l,
                            normales_prioridad=(
                                [norm_ml] if norm_ml is not None else None
                            ),
                            gancho_en_inicio=False,
                            gancho_en_fin=False,
                        )
                    )
                    if rb_lat is None:
                        if err_lat:
                            avisos.append(
                                u"{0}: {1}".format(_lat_tag, err_lat)
                            )
                        continue
                    try:
                        rebar_ids_barrido_ganchos.append(rb_lat.Id)
                    except Exception:
                        pass
                    ok_ll, err_ll = aplicar_layout_fixed_number_rebar(
                        rb_lat,
                        document,
                        n_lat_eff,
                        float(span_ft),
                    )
                    if not ok_ll and err_ll:
                        avisos.append(
                            u"{0} layout: {1}".format(_lat_tag, err_ll)
                        )
                    _presentacion_rebar_show_middle_en_vista(rb_lat, view)
                    n_rebar += _rebar_bar_positions(rb_lat, n_lat_eff)

            if (
                crear_armadura_suple
                and crear_rebar_desde_curva_linea_con_ganchos is not None
            ):
                for si in range(_n_suple_lay):
                    bar_suple_i = rebar_bar_type_suple
                    if _tlist_suple and si < len(_tlist_suple):
                        if _tlist_suple[si] is not None:
                            bar_suple_i = _tlist_suple[si]
                    if bar_suple_i is None:
                        continue
                    cantidad_inf = max(1, int(rebar_cantidad_suple or 1))
                    if _clist_suple and si < len(_clist_suple):
                        try:
                            cantidad_inf = max(
                                1, int(_clist_suple[si] or 1)
                            )
                        except Exception:
                            pass
                    suple_off_mm = float(_n_capas_sup_def + si) * float(
                        _OFFSET_SUPLES_SEGUNDA_CAPA_MM
                    )
                    work2 = _linea_desplazada_mm_reverso_normal_cara(
                        work,
                        n_face_ref,
                        suple_off_mm,
                    )
                    if work2 is None:
                        avisos.append(
                            u"Suples: no se pudo desplazar el trazo (−n, "
                            u"{0:.0f} mm; capa suple {1}/{2}); revise geometría.".format(
                                float(suple_off_mm),
                                int(si + 1),
                                int(_n_suple_lay),
                            )
                        )
                        continue
                    work2_obst = work2
                    if obst_elems:
                        if _es_inf:
                            wo = _trim_extremos_linea_por_interseccion_caras_obstaculo_selectivo(
                                work2,
                                obst_elems,
                                True,
                                True,
                                min_face_hits_en_extremo=1,
                            )
                        else:
                            wo = _trim_extremos_linea_por_interseccion_caras_obstaculo(
                                work2, obst_elems
                            )
                        if wo is not None:
                            work2_obst = wo
                    work2_trim = _recortar_extremos_linea_mm(
                        work2_obst, _TRIM_AXIS_ENDS_MM
                    )
                    if work2_trim is None:
                        work2_trim = work2_obst
                    if _ml_eje_suple_troceo:
                        try:
                            if _crear_model_curve(
                                document, work2_trim, n_face_ref
                            ):
                                n_curves += 1
                        except Exception:
                            pass
                    _frac_suple = _suple_fracciones_plan_troceo(_es_inf)
                    if _dibujar_ml:
                        n_curves += _crear_marcadores_planos_division_suples(
                            document,
                            work2_trim,
                            chain,
                            _frac_suple,
                            n_face_ref,
                        )
                    # Planos en fracciones del eje entre tapas (estribos); curva = work2 ya recortada.
                    cuts2 = _parametros_corte_suples_tapas_sobre_linea(
                        work2_trim,
                        chain,
                        _frac_suple,
                    )
                    cuts2 = _dedupe_sorted_cut_params(
                        cuts2, float(work2_trim.Length)
                    )
                    pieces2 = _split_line_by_distances(work2_trim, cuts2)
                    if not pieces2:
                        pieces2 = [work2_trim]
                    n_pc2_all = len(pieces2)
                    final_segs2 = []
                    orig_k2 = []
                    for k_pc, pc2 in enumerate(pieces2):
                        if _es_inf:
                            if not _suple_inferior_mantener_trozo(k_pc, n_pc2_all):
                                continue
                        elif not _suple_superior_mantener_trozo(k_pc):
                            continue
                        # Suples: recorte obst. solo extremos libres de la corrida completa.
                        if not obst_elems:
                            seg_obst = pc2
                        elif n_pc2_all <= 1:
                            seg_obst = _trim_extremos_linea_por_interseccion_caras_obstaculo(
                                pc2, obst_elems
                            )
                        elif k_pc == 0:
                            seg_obst = (
                                _trim_extremos_linea_por_interseccion_caras_obstaculo_selectivo(
                                    pc2, obst_elems, True, False
                                )
                            )
                        elif k_pc == n_pc2_all - 1:
                            seg_obst = (
                                _trim_extremos_linea_por_interseccion_caras_obstaculo_selectivo(
                                    pc2, obst_elems, False, True
                                )
                            )
                        else:
                            seg_obst = pc2
                        if seg_obst is None:
                            continue
                        if n_pc2_all <= 1:
                            seg_trim = _recortar_extremos_linea_mm(
                                seg_obst, _TRIM_AXIS_ENDS_MM
                            )
                        else:
                            seg_trim = _recortar_extremos_linea_mm_selectivo(
                                seg_obst,
                                _TRIM_AXIS_ENDS_MM,
                                trim_start=(k_pc == 0),
                                trim_end=(k_pc == n_pc2_all - 1),
                            )
                        if seg_trim is not None:
                            final_segs2.append(seg_trim)
                            orig_k2.append(k_pc)
                    if not final_segs2:
                        avisos.append(
                            u"Suples: sin tramos válidos tras recortes."
                        )
                    else:
                        layout_span_inf = _layout_max_spacing_array_length_ft(
                            document,
                            host0,
                            bar_suple_i,
                            diametro_estribo_mm=rex_mm,
                        )
                        span_mm_suple = None
                        try:
                            if (
                                layout_span_inf is not None
                                and float(layout_span_inf) > 1e-12
                            ):
                                span_mm_suple = float(layout_span_inf) * 304.8
                        except Exception:
                            span_mm_suple = None
                        if rebar_espaciado_suple_mm is not None:
                            try:
                                esp_su = float(rebar_espaciado_suple_mm)
                            except Exception:
                                esp_su = 0.0
                            if esp_su > 0 and span_mm_suple is not None:
                                c_su = (
                                    _cantidad_barras_desde_espaciado_transversal_mm(
                                        span_mm_suple, esp_su
                                    )
                                )
                                if c_su is not None:
                                    cantidad_inf = c_su
                        host_candidatos_sup = _candidatos_host_empalme(
                            chain, emp_elems
                        )
                        n_kept2 = len(final_segs2)
                        for idx2, seg2 in enumerate(final_segs2):
                            k_orig = orig_k2[idx2]
                            if _dibujar_ml and _DIBUJAR_MODEL_LINE_TRAMO:
                                if _crear_model_curve(
                                    document, seg2, n_face_ref
                                ):
                                    n_curves += 1
                            if _dibujar_ml and _DIBUJAR_MARCADORES_EJE_XZ:
                                n_curves += _crear_marcadores_model_line_ejes_xz(
                                    document, seg2, n_face_ref
                                )
                            if (
                                _DIBUJAR_DETAIL_CURVE_VISTA
                                and plane_view is not None
                            ):
                                try:
                                    dv = _proyectar_linea_al_plano_vista(
                                        seg2, plane_view
                                    )
                                    if dv is not None:
                                        document.Create.NewDetailCurve(
                                            view, dv
                                        )
                                except Exception:
                                    pass
                            norm_prio2 = norm_prio_rebar
                            if norm_prio2 is None:
                                norm_prio2 = (
                                    _norm_createfromcurves_desde_cara_y_tramo(
                                        n_face_ref, seg2
                                    )
                                )
                            try:
                                z_hook_ref2 = (
                                    n_face_ref.Normalize().Negate()
                                    if n_face_ref is not None
                                    else XYZ.BasisZ.Negate()
                                )
                            except Exception:
                                z_hook_ref2 = XYZ.BasisZ.Negate()
                            host2 = host0
                            if n_pc2_all > 1:
                                pm2 = _punto_medio_linea_bound(seg2)
                                host2 = _host_framing_para_segmento_rebar(
                                    pm2, host_candidatos_sup, host0
                                )
                            norm_bar_create2 = None
                            if norm_prio2 is not None:
                                try:
                                    norm_bar_create2 = norm_prio2.Negate()
                                except Exception:
                                    norm_bar_create2 = norm_prio2
                            if _es_inf:
                                # Suple inferior: sin ganchos en extremos (barra recta entre cortes).
                                _gi2 = False
                                _gf2 = False
                            else:
                                _gi2 = k_orig == 0
                                _gf2 = k_orig == n_pc2_all - 1
                            rb2, err2, _nv2 = (
                                crear_rebar_desde_curva_linea_con_ganchos(
                                    document,
                                    host2,
                                    bar_suple_i,
                                    seg2,
                                    hook_type_name=HOOK_GANCHO_90_STANDARD_NAME,
                                    marco_cara_uvn=None,
                                    cara_paralela=None,
                                    eje_referencia_z_ganchos=z_hook_ref2,
                                    normales_prioridad=(
                                        [norm_bar_create2]
                                        if norm_bar_create2 is not None
                                        else None
                                    ),
                                    gancho_en_inicio=_gi2,
                                    gancho_en_fin=_gf2,
                                )
                            )
                            if rb2 is None:
                                if err2:
                                    avisos.append(
                                        u"Suples: {0}".format(err2)
                                    )
                            else:
                                try:
                                    rebar_ids_barrido_ganchos.append(rb2.Id)
                                except Exception:
                                    pass
                                try:
                                    if aplicar_layout_fixed_number_rebar is not None:
                                        span2 = layout_span_inf
                                        if n_kept2 > 1 and host2 is not None:
                                            try:
                                                ls_h2 = (
                                                    _layout_max_spacing_array_length_ft(
                                                        document,
                                                        host2,
                                                        bar_suple_i,
                                                        diametro_estribo_mm=rex_mm,
                                                    )
                                                )
                                                if (
                                                    ls_h2 is not None
                                                    and float(ls_h2) > 1e-12
                                                ):
                                                    span2 = ls_h2
                                            except Exception:
                                                pass
                                        if (
                                            cantidad_inf > 1
                                            and (
                                                span2 is None
                                                or float(span2) < 1e-12
                                            )
                                        ):
                                            diam2 = float(
                                                _rebar_nominal_diameter_mm(
                                                    bar_suple_i
                                                )
                                                or 16.0
                                            )
                                            sep2 = max(25.0, diam2 + 25.0)
                                            span2 = _mm_to_ft(
                                                sep2 * float(cantidad_inf - 1)
                                            )
                                        ok_l2, err_l2 = (
                                            aplicar_layout_fixed_number_rebar(
                                                rb2,
                                                document,
                                                cantidad_inf,
                                                float(span2 or 0.0),
                                            )
                                        )
                                        if not ok_l2 and err_l2:
                                            avisos.append(err_l2)
                                except Exception:
                                    pass
                                if (
                                    (_gi2 or _gf2)
                                    and _enforce_rebar_hook_types_by_name
                                    is not None
                                ):
                                    try:
                                        _enforce_rebar_hook_types_by_name(
                                            rb2,
                                            document,
                                            HOOK_GANCHO_90_STANDARD_NAME,
                                            _gi2,
                                            _gf2,
                                            avisos=avisos,
                                        )
                                    except Exception:
                                        pass
                                n_rebar += _rebar_bar_positions(
                                    rb2, cantidad_inf
                                )

            if (
                _DIBUJAR_DETAIL_ITEM_TRASLAPE
                and lap_detail_symbol is not None
                and view is not None
                and vista_permite_detail_curve(view)
                and cuts
                and lap_mm > 0
            ):
                for capa_lap in range(_n_capas_sup_def):
                    cap_map = rebar_by_interval.get(capa_lap) or {}
                    try:
                        rta_lap = _types_cap[capa_lap]
                    except Exception:
                        rta_lap = rebar_bar_type
                    lap_mm_c = lap_mm
                    if rta_lap is not None:
                        try:
                            _lmc, _ = _traslapo_longitudinal_mm_desde_bar_type(
                                rta_lap
                            )
                            if _lmc and float(_lmc) > 0:
                                lap_mm_c = float(_lmc)
                        except Exception:
                            pass
                    off_lap_mm = float(capa_lap) * float(step_capas_mm)
                    for j in range(len(cuts)):
                        c = float(cuts[j])
                        pa, pb = _puntos_segmento_traslape_sobre_work(
                            work, c, lap_mm_c
                        )
                        if pa is None or pb is None:
                            continue
                        if off_lap_mm > 1e-9 and n_face_ref is not None:
                            try:
                                nrm = n_face_ref.Normalize()
                                d_cap = nrm.Multiply(-_mm_to_ft(off_lap_mm))
                                pa = pa + d_cap
                                pb = pb + d_cap
                            except Exception:
                                pass
                        ra_el = cap_map.get(j)
                        rb_el = cap_map.get(j + 1)
                        ok_d, err_d, lap_inst = (
                            _colocar_detail_item_traslape_en_vista(
                                document,
                                view,
                                lap_detail_symbol,
                                pa,
                                pb,
                            )
                        )
                        if ok_d and lap_inst is not None:
                            dim_eid = None
                            if (
                                _get_named_left_right_refs_from_detail_instance
                                is not None
                                and _create_overlap_dimension_from_detail_refs
                                is not None
                            ):
                                ref_l, ref_r, ref_err = (
                                    _get_named_left_right_refs_from_detail_instance(
                                        lap_inst
                                    )
                                )
                                if ref_l is not None and ref_r is not None:
                                    axis_u = None
                                    try:
                                        dv = pb - pa
                                        if dv.GetLength() > 1e-9:
                                            axis_u = dv.Normalize()
                                    except Exception:
                                        axis_u = None
                                    inward_xy = None
                                    inward_3d = None
                                    if n_face_ref is not None:
                                        try:
                                            inv = n_face_ref.Negate()
                                            if inv.GetLength() > 1e-12:
                                                inward_3d = inv.Normalize()
                                            inward_xy = XYZ(
                                                float(inv.X), float(inv.Y), 0.0
                                            )
                                            if inward_xy.GetLength() > 1e-9:
                                                inward_xy = inward_xy.Normalize()
                                            else:
                                                inward_xy = None
                                        except Exception:
                                            inward_xy = None
                                            inward_3d = None
                                    ok_dim, msg_dim, dim_data = (
                                        _create_overlap_dimension_from_detail_refs(
                                            document,
                                            view,
                                            ref_l,
                                            ref_r,
                                            pa,
                                            pb,
                                            axis_u,
                                            lateral_hint=None,
                                            line_offset_mm=450.0,
                                            inward_dir_xy=inward_xy,
                                            inward_dir_3d=inward_3d,
                                            use_view_plane_dim_line=True,
                                            flip_dimension_side=False,
                                        )
                                    )
                                    if ok_dim and dim_data and dim_data.get(
                                        "dim_id"
                                    ):
                                        try:
                                            dim_eid = ElementId(
                                                int(dim_data["dim_id"])
                                            )
                                        except Exception:
                                            dim_eid = None
                                        if dim_eid is not None:
                                            n_lap_cotas += 1
                                    elif msg_dim:
                                        avisos.append(
                                            u"Cota traslape (viga): {0}".format(
                                                msg_dim
                                            )
                                        )
                                elif ref_err and aviso_refs_lap_familia is None:
                                    aviso_refs_lap_familia = ref_err
                            if (
                                ra_el is not None
                                and rb_el is not None
                                and set_lap_detail_vigas_rebar_link is not None
                            ):
                                try:
                                    set_lap_detail_vigas_rebar_link(
                                        lap_inst,
                                        ra_el.Id,
                                        rb_el.Id,
                                        dim_eid,
                                    )
                                except Exception:
                                    pass
                            n_lap_details += 1
                        elif err_d:
                            avisos.append(err_d)

        if aviso_refs_lap_familia:
            avisos.append(aviso_refs_lap_familia)
        if n_lap_details > 0:
            avisos.append(
                u"Detail Items de traslape colocados: {0}.".format(
                    int(n_lap_details)
                )
            )
        if n_lap_cotas > 0:
            avisos.append(
                u"Cotas de longitud de traslape (vigas): {0}.".format(int(n_lap_cotas))
            )
        # Tras crear barras y layout/detail: barrido de ganchos en la misma Transaction.
        # Errores inesperados solo generan aviso; la colocación ya hecha sigue confirmándose al Commit.
        if (
            rebar_ids_barrido_ganchos
            and _sweep_rebar_hook_types_to_name is not None
        ):
            try:
                _sweep_rebar_hook_types_to_name(
                    document,
                    rebar_ids_barrido_ganchos,
                    HOOK_GANCHO_90_STANDARD_NAME,
                    avisos=avisos,
                )
            except Exception as ex:
                try:
                    avisos.append(
                        u"Barrido de ganchos: {0}".format(unicode(ex))
                    )
                except Exception:
                    avisos.append(
                        u"Barrido de ganchos: error inesperado al ajustar tipos."
                    )
        if (
            rebar_ids_barrido_ganchos
            and _apply_armadura_largo_total_to_rebars is not None
        ):
            try:
                _apply_armadura_largo_total_to_rebars(
                    document, rebar_ids_barrido_ganchos, avisos
                )
            except Exception:
                pass
        n_etiquetas_rebar_vista = 0
        if (
            _ETIQUETAR_REBAR_EN_VISTA_ACTIVA
            and view is not None
            and vista_permite_detail_curve(view)
            and rebar_ids_barrido_ganchos
        ):
            try:
                document.Regenerate()
            except Exception:
                pass
            ids_etiqueta = []
            vistos_tag = set()
            for rid in rebar_ids_barrido_ganchos:
                try:
                    iid = int(rid.IntegerValue)
                except Exception:
                    continue
                if iid in vistos_tag:
                    continue
                vistos_tag.add(iid)
                el = document.GetElement(rid)
                if el is None or not isinstance(el, Rebar):
                    continue
                ids_etiqueta.append(rid)
            if ids_etiqueta:
                if etiquetar_rebars_creados_en_vista is not None:
                    n_tag, avis_tag, err_tag = etiquetar_rebars_creados_en_vista(
                        document,
                        view,
                        ids_etiqueta,
                        family_name=_ETIQUETAR_REBAR_FAMILIA_NOMBRE,
                        fixed_type_name=None,
                        use_transaction=False,
                    )
                    n_etiquetas_rebar_vista = int(n_tag or 0)
                    if avis_tag:
                        avisos.extend(list(avis_tag))
                    if err_tag:
                        avisos.append(err_tag)
                else:
                    for rid in ids_etiqueta:
                        el = document.GetElement(rid)
                        if el is None or not isinstance(el, Rebar):
                            continue
                        if _etiquetar_un_rebar_categoria_vista(
                            document, view, el, avisos
                        ):
                            n_etiquetas_rebar_vista += 1
                    avisos.append(
                        u"Etiquetas: no se cargó etiquetar_rebars_creados_en_vista; "
                        u"se usó tipo por categoría (no familia EST)."
                    )
            if ids_etiqueta and n_etiquetas_rebar_vista > 0:
                try:
                    _alinear_etiquetas_rebar_mismo_lote(
                        document,
                        view,
                        ids_etiqueta,
                        es_cara_inferior=_es_inf,
                    )
                except Exception:
                    pass
                try:
                    _separar_etiquetas_rebar_solapadas_lote(
                        document,
                        view,
                        ids_etiqueta,
                        es_cara_inferior=_es_inf,
                    )
                except Exception:
                    pass
        if n_etiquetas_rebar_vista > 0:
            avisos.append(
                u"Etiquetas de armadura en vista activa: {0}.".format(
                    int(n_etiquetas_rebar_vista)
                )
            )
        if gestionar_transaccion:
            try:
                t.Commit()
            except Exception as ex:
                try:
                    t.RollBack()
                except Exception:
                    pass
                try:
                    avisos.append(u"Transacción: {0}".format(unicode(ex)))
                except Exception:
                    avisos.append(u"Error al confirmar la transacción.")
    except Exception as ex:
        if gestionar_transaccion and t is not None:
            try:
                t.RollBack()
            except Exception:
                pass
        try:
            avisos.append(u"{0}".format(unicode(ex)))
        except Exception:
            avisos.append(u"Error durante la creación de geometría.")

    return n_curves, n_rebar, avisos
