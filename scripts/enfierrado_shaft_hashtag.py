# -*- coding: utf-8 -*-
"""
Enfierrado alrededor de hueco (shaft) en losa: barras por cara vertical seleccionada.

Geometría base por cara:
  - Se toma el perímetro exterior con GetEdgesAsCurveLoops().
  - Se extrae la(s) curva(s) inferior(es) horizontal(es) del loop exterior.
  - No se combinan perímetros entre caras; cada cara se procesa en forma independiente.
  - Se recortan extremos por recubrimiento + 1/2 diámetro nominal y se valida punto medio en hormigón.

Creación: Rebar.CreateFromCurves (Standard/StirrupTie, varias normales); sin ganchos si la API lo admite;
fallback CreateFromCurvesAndShape con RebarShape recto + RebarHookType del proyecto.

Uso programático: crear_detail_curves_tramos_shaft_hashtag(doc, view, host, refs, ...)
para depuración; crear_enfierrado_shaft_hashtag(doc, host, refs, ...) para barras.

Desde herramientas de borde de losa (pyRevit / RPS): usar los módulos que llaman a crear_enfierrado_shaft_hashtag / helpers de este archivo.

Extensión de tramos: según diámetro nominal del RebarBarType (tabla fija + interpolación), no por área del hueco.

Revit 2024–2026 | pyRevit (IronPython 2.7). ElementId: .Value (2026+) o .IntegerValue.
"""

import math
import sys
import clr
import os
from collections import defaultdict

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import System
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Curve,
    ElementId,
    ElementType,
    Family,
    Floor,
    FilteredElementCollector,
    FamilySymbol,
    GeometryInstance,
    IndependentTag,
    Line,
    LocationCurve,
    Options,
    Plane,
    PlanarFace,
    Reference,
    Solid,
    SolidCurveIntersectionMode,
    SolidCurveIntersectionOptions,
    SpecTypeId,
    StorageType,
    SubTransaction,
    TagMode,
    TagOrientation,
    Transaction,
    UnitUtils,
    UnitTypeId,
    ViewDetailLevel,
    View3D,
    ViewPlan,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import (
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
)
from Autodesk.Revit.UI import TaskDialog


def _eid_int(element_id):
    """ElementId → int (2026+: Value; 2024–25: IntegerValue). Inline: evita fallos de sys.path."""
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


# Parámetro compartido: suma de los largos de tramo de forma (A, B, C… en mm), como en la paleta.
ARMADURA_LARGO_TOTAL_PARAM_NAMES = (u"Armadura_Largo Total",)
# Rotación de gancho (instancia Rebar): BuiltIn + nombres EN/ES para LookupParameter.
REBAR_HOOK_ROTATION_START_BIP_NAMES = ("REBAR_HOOK_ROTATION_AT_START",)
REBAR_HOOK_ROTATION_END_BIP_NAMES = ("REBAR_HOOK_ROTATION_AT_END",)
REBAR_HOOK_ROTATION_START_LOOKUP_NAMES = (
    u"Hook Rotation At Start",
    u"Rotación de gancho al inicio",
    u"Rotación gancho al inicio",
)
REBAR_HOOK_ROTATION_END_LOOKUP_NAMES = (
    u"Hook Rotation At End",
    u"Rotación de gancho al final",
    u"Rotación gancho al final",
)

# Recubrimiento desde la cara del hueco hacia el hormigón (mm) — offset lateral y eje.
COVER_MM_DEFAULT = 25.0
# Recubrimiento en extremos del tramo inferior (recorte longitudinal en la arista del vano;
# cara superficial → eje de barra en cada extremo). Independiente del nominal lateral.
SHAFT_END_COVER_MM_DEFAULT = 35.0
# Separación fija entre capas hacia interior del host (mm).
LAYER_SPACING_MM_DEFAULT = 100.0
# Extra lateral XY para la primera capa (k==0): aleja del borde del host para evitar recortes automáticos (mm).
SHAFT_FIRST_LAYER_LATERAL_EXTRA_MM_DEFAULT = 10.0
# 0 = desactivado. Si > 0: una DetailLine por cara en tag_view, copia de la primera arista inferior
# del loop exterior desplazada en planta (mm) perpendicular al tramo, para inspeccionar orden/origen.
SHAFT_DEBUG_FIRST_BOTTOM_EDGE_OFFSET_MM = 0.0
# Sonda axial (mm) para detectar en qué extremo el anclaje entra en hormigón (SCI).
# Misma sonda para DetailLine debug y para estirón real en empotramiento adaptivo.
SHAFT_EMBED_HOST_PROBE_MM = 100.0
# Alias histórico (DetailLine): igual que SHAFT_EMBED_HOST_PROBE_MM.
SHAFT_DEBUG_FIRST_BOTTOM_EDGE_EXTEND_EACH_END_MM = SHAFT_EMBED_HOST_PROBE_MM
# Reservado por si se reintroduce layout / segundo eje (actualmente no usado).
DUPLEX_SPACING_MM_DEFAULT = 50.0

# Empuje de IndependentTag de rebar en XY: pasos crecientes (mm) hasta que la AABB de la etiqueta
# no intersecte las AABB de la barra etiquetada, otras barras del lote ni etiquetas ya colocadas.
REBAR_TAG_NUDGE_STEPS_MM = (
    50,
    75,
    100,
    125,
    150,
    200,
    250,
    300,
    400,
    500,
    650,
    800,
    1000,
    1250,
    1500,
    2000,
)

# Línea de llamada (leader) con quiebre: fracción recorrida barra→cabecera y desvío lateral en el plano de la vista.
REBAR_TAG_LEADER_JOG_ALONG = 0.42
REBAR_TAG_LEADER_JOG_OFFSET_MM = 80.0
# Nombres de tipo de etiqueta (``_norm_text`` del ``FamilySymbol``) con leader en L ortogonal.
# Demás tipos: ``_apply_rebar_tag_leader_elbow_jog``.
REBAR_TAG_ORTHOGONAL_LEADER_TYPE_KEYS = frozenset((u"01",))
# L (tipo ortogonal): tramo vertical mínimo desde el ancla y estante horizontal mínimo hacia ``Right``.
# Si la cabecera está colineal con el ancla en ``Up``, sin estante horizontal Revit no dibuja codo.
REBAR_TAG_ORTHOGONAL_DROP_MIN_MM = 100.0
REBAR_TAG_ORTHOGONAL_SHELF_MIN_MM = 120.0
# Tras el codo, desplazar la cabecera más allá del estante (anclaje del símbolo ≠ centro del texto).
REBAR_TAG_ORTHOGONAL_HEAD_EXTRA_R_MM = 160.0
REBAR_TAG_ORTHOGONAL_HEAD_EXTRA_U_MM = 35.0

# Si True: en hueco rectangular (4 caras, 1 tramo/cara) iguala longitudes H/V y alinea centros entre pares.
# Si False: solo geometría por cara + estiro por Ø (comportamiento anterior a esa corrección).
SHAFT_SYMMETRIZE_ALIGN_RECT4 = False

# Extensión por lado (mm) según diámetro nominal de barra (mm). Orden creciente por Ø.
EXTENSION_MM_BY_BAR_DIAMETER_MM = (
    (8, 570),
    (10, 710),
    (12, 860),
    (16, 1140),
    (18, 1290),
    (22, 1960),
    (25, 2230),
    (28, 2500),
    (32, 2850),
    (36, 3210),
)

# Longitud **total** de traslapo (mm) entre tramos colindantes (p. ej. vigas cara superior tras troceo).
# Por defecto mismas filas que desarrollo/empotramiento; sustituir si el reglamento de traslapos difiere.
LAP_LENGTH_MM_BY_BAR_DIAMETER_MM = tuple(EXTENSION_MM_BY_BAR_DIAMETER_MM)

# Largo de pata de gancho (mm) según Ø nominal — usado al partir (recto + pata = chunk).
SHAFT_HOOK_LEG_MM_BY_BAR_DIAMETER_MM = (
    (8, 160),
    (10, 200),
    (12, 240),
    (16, 320),
    (18, 360),
    (22, 440),
    (25, 500),
    (28, 570),
    (32, 650),
    (36, 720),
)

# Regla de división para barras largas en pasada/shaft.
MAX_SINGLE_BAR_LENGTH_MM = 12000.0
LONG_BAR_SPLIT_CHUNK_MM = 6000.0
# Tras offset a eje en el host: longitud en mm → siguiente múltiplo de este paso (simétrico respecto al centro).
SHAFT_BAR_LENGTH_ROUND_STEP_MM = 10.0
# Herramienta enfierrado shaft / borde losa: si True, ``Left`` en ambos extremos (ignora geometría).
# False: orientación desde ``_compute_hook_orientations_pair_from_inverse_normal_xyz`` (gancho hacia interior).
SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT = False
# Tras crear barras: rotación de gancho 180° — Revit 2026 ``SetTerminationRotationAngle``; antes vía parámetro/API antigua.
SHAFT_REBAR_HOOK_ROTATION_PARAM_DEGREES = 180.0


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _ft_to_mm(ft):
    return float(ft) * 304.8

_FIXED_DIMSTYLE_NAME = u"Linear - 2.5mm Arial"
_FIXED_DIMSTYLE_CACHE = {}


def _get_fixed_dimension_type_id(doc):
    """
    Busca el DimensionType por nombre exacto para cotas lineales.
    Si no existe, retorna None (fallback silencioso al estilo por defecto).
    """
    if doc is None:
        return None
    try:
        cache_key = id(doc)
    except Exception:
        cache_key = None
    if cache_key is not None and cache_key in _FIXED_DIMSTYLE_CACHE:
        return _FIXED_DIMSTYLE_CACHE.get(cache_key)
    found_id = None
    target_name = unicode(_FIXED_DIMSTYLE_NAME).strip().lower()
    dim_types = []
    # Compat 2024+: preferimos DimensionType cuando está disponible.
    try:
        from Autodesk.Revit import DB as _RDB

        dt_cls = getattr(_RDB, "DimensionType", None)
    except Exception:
        dt_cls = None
    if dt_cls is not None:
        try:
            dim_types = list(FilteredElementCollector(doc).OfClass(dt_cls))
        except Exception:
            dim_types = []
    if not dim_types:
        # Fallback compatible: recorrer ElementType de categoría cotas.
        try:
            for et in FilteredElementCollector(doc).OfClass(ElementType):
                try:
                    cat = et.Category
                    if cat is None:
                        continue
                    if _eid_int(cat.Id) != int(BuiltInCategory.OST_Dimensions):
                        continue
                    dim_types.append(et)
                except Exception:
                    continue
        except Exception:
            dim_types = []
    try:
        for dt in dim_types:
            try:
                nm = unicode(dt.Name or u"").strip().lower()
            except Exception:
                nm = u""
            if nm == target_name:
                found_id = dt.Id
                break
    except Exception:
        found_id = None
    if found_id is None:
        # Fallback robusto: algunos templates exponen nombre por parámetro.
        try:
            for dt in dim_types:
                try:
                    p = dt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    nm = unicode(p.AsString() if p is not None else u"").strip().lower()
                except Exception:
                    nm = u""
                if nm == target_name:
                    found_id = dt.Id
                    break
        except Exception:
            found_id = None
    if cache_key is not None:
        _FIXED_DIMSTYLE_CACHE[cache_key] = found_id
    return found_id


def _try_apply_fixed_dimension_type(doc, dim):
    """Aplica el estilo fijo de cota si está disponible; si no, no hace nada."""
    if doc is None or dim is None:
        return
    dim_type_id = _get_fixed_dimension_type_id(doc)
    if dim_type_id is None:
        return
    try:
        dim.ChangeTypeId(dim_type_id)
    except Exception:
        return
    # Verificación post-aplicación (algunos contextos de Revit ignoran el cambio).
    try:
        cur_tid = dim.GetTypeId()
        if _eid_int(cur_tid) != _eid_int(dim_type_id):
            return
    except Exception:
        pass


def _norm_text(txt):
    if txt is None:
        return u""
    # Forzar conversión robusta desde objetos .NET/IronPython a texto.
    try:
        s = System.Convert.ToString(txt)
        if s is not None:
            return unicode(s).strip().lower()
    except Exception:
        pass
    try:
        return unicode(txt).strip().lower()
    except Exception:
        return str(txt).strip().lower()


def _norm_family_name(txt):
    """
    Normaliza nombre de familia para comparación robusta.
    """
    s = _norm_text(txt)
    if not s:
        return s
    # Normalización suave: separadores comunes a espacio.
    for ch in ("_", "-", ".", ","):
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    return s


def _norm_alnum_key(txt):
    """
    Clave alfanumérica robusta para comparar nombres ignorando separadores/símbolos.
    """
    s = _norm_family_name(txt)
    if not s:
        return u""
    return u"".join(ch for ch in s if ch.isalnum())


def extension_mm_por_diametro_nominal_mm(d_mm):
    """
    Extensión a cada lado según Ø nominal (mm): tabla + interpolación lineal entre filas;
    fuera del rango de la tabla se usa la fila extrema.
    Returns:
        (extend_each_side_mm:float, descripcion:unicode)
    """
    table = EXTENSION_MM_BY_BAR_DIAMETER_MM
    try:
        d = float(d_mm)
    except Exception:
        d = float(table[0][0])
    d_min, e_min = table[0]
    d_max, e_max = table[-1]
    if d <= d_min:
        return float(e_min), u"Ø≤{0} mm → {1} mm/lado (tabla)".format(int(d_min), int(e_min))
    if d >= d_max:
        return float(e_max), u"Ø≥{0} mm → {1} mm/lado (tabla)".format(int(d_max), int(e_max))
    for i in range(len(table) - 1):
        d_lo, e_lo = table[i]
        d_hi, e_hi = table[i + 1]
        if d_lo <= d <= d_hi:
            if abs(d - d_lo) < 1e-9:
                return float(e_lo), u"Ø{0} mm → {1} mm/lado".format(int(d_lo), int(e_lo))
            if abs(d - d_hi) < 1e-9:
                return float(e_hi), u"Ø{0} mm → {1} mm/lado".format(int(d_hi), int(e_hi))
            span = float(d_hi) - float(d_lo)
            if span <= 1e-12:
                return float(e_lo), u"Ø{0} mm → {1} mm/lado".format(int(d_lo), int(e_lo))
            t = (d - float(d_lo)) / span
            ext = float(e_lo) + t * (float(e_hi) - float(e_lo))
            return ext, (
                u"Interpol. Ø{0:.1f} mm → {1:.0f} mm/lado (entre Ø{2} y Ø{3})"
                .format(d, ext, int(d_lo), int(d_hi))
            )
    return float(e_min), u"(tabla)"


def _nearest_tabulated_bar_diameter_mm(d_mm, table=None):
    """
    Proyecta un Ø nominal (mm) al valor de fila de la tabla (``EXTENSION_MM_...`` por defecto).
    Si dos filas empatan en distancia, se elige el Ø mayor (tabla más exigente).
    """
    if table is None:
        table = EXTENSION_MM_BY_BAR_DIAMETER_MM
    try:
        d = float(d_mm)
    except Exception:
        return float(table[0][0])
    best = None
    for k, _ in table:
        kk = float(k)
        key = (abs(d - kk), -kk)
        if best is None or key < best[0]:
            best = (key, kk)
    return float(best[1]) if best is not None else float(table[0][0])


def extension_mm_para_bar_type(bar_type):
    """
    Usa el diámetro nominal del RebarBarType, proyectado al Ø de tabla más cercano
    antes de aplicar EXTENSION_MM_BY_BAR_DIAMETER_MM (coherente con la tabla por Ø nominal).
    Returns:
        (extend_each_side_mm:float, descripcion:unicode, d_nominal_mm:float|None)
    """
    if bar_type is None:
        return 0.0, u"", None
    try:
        d_mm = _ft_to_mm(float(bar_type.BarNominalDiameter))
    except Exception:
        return 0.0, u"Sin diámetro nominal en el tipo de barra.", None
    d_tab = _nearest_tabulated_bar_diameter_mm(d_mm)
    ext, txt = extension_mm_por_diametro_nominal_mm(d_tab)
    if abs(d_tab - d_mm) > 0.05:
        try:
            txt = u"{0} [Ø tipo API {1:.1f} mm → Ø tabla {2:.0f} mm]".format(
                txt, d_mm, d_tab
            )
        except Exception:
            pass
    return ext, txt, d_mm


def lap_mm_por_diametro_nominal_mm(d_mm):
    """
    Longitud **total** de traslapo (mm) según Ø nominal: tabla ``LAP_LENGTH_MM_...`` + interpolación.
    Returns:
        (lap_mm:float, descripcion:unicode)
    """
    table = LAP_LENGTH_MM_BY_BAR_DIAMETER_MM
    try:
        d = float(d_mm)
    except Exception:
        d = float(table[0][0])
    d_min, e_min = table[0]
    d_max, e_max = table[-1]
    if d <= d_min:
        return float(e_min), u"Traslapos: Ø≤{0} mm → {1} mm".format(int(d_min), int(e_min))
    if d >= d_max:
        return float(e_max), u"Traslapos: Ø≥{0} mm → {1} mm".format(int(d_max), int(e_max))
    for i in range(len(table) - 1):
        d_lo, e_lo = table[i]
        d_hi, e_hi = table[i + 1]
        if d_lo <= d <= d_hi:
            if abs(d - d_lo) < 1e-9:
                return float(e_lo), u"Traslapos: Ø{0} mm → {1} mm".format(int(d_lo), int(e_lo))
            if abs(d - d_hi) < 1e-9:
                return float(e_hi), u"Traslapos: Ø{0} mm → {1} mm".format(int(d_hi), int(e_hi))
            span = float(d_hi) - float(d_lo)
            if span <= 1e-12:
                return float(e_lo), u"Traslapos: Ø{0} mm → {1} mm".format(int(d_lo), int(e_lo))
            t = (d - float(d_lo)) / span
            lap = float(e_lo) + t * (float(e_hi) - float(e_lo))
            return lap, (
                u"Traslapos interpol. Ø{0:.1f} mm → {1:.0f} mm (entre Ø{2} y Ø{3})"
                .format(d, lap, int(d_lo), int(d_hi))
            )
    return float(e_min), u"(tabla traslapos)"


def lap_mm_para_bar_type(bar_type):
    """
    Traslapo total (mm) según ``RebarBarType`` (Ø a fila más cercana de ``LAP_LENGTH_MM_...``).
    Returns:
        (lap_mm:float, descripcion:unicode, d_nominal_mm:float|None)
    """
    if bar_type is None:
        return 0.0, u"", None
    try:
        d_mm = _ft_to_mm(float(bar_type.BarNominalDiameter))
    except Exception:
        return 0.0, u"Sin diámetro nominal en el tipo de barra.", None
    d_tab = _nearest_tabulated_bar_diameter_mm(d_mm, LAP_LENGTH_MM_BY_BAR_DIAMETER_MM)
    lap, txt = lap_mm_por_diametro_nominal_mm(d_tab)
    if abs(d_tab - d_mm) > 0.05:
        try:
            txt = u"{0} [Ø API {1:.1f} mm → Ø tabla {2:.0f} mm]".format(txt, d_mm, d_tab)
        except Exception:
            pass
    return lap, txt, d_mm


def hook_leg_mm_por_diametro_nominal_mm(d_mm):
    """
    Largo de pata de gancho (mm) según Ø nominal: tabla + interpolación lineal entre filas;
    fuera del rango se usa la fila extrema.
    """
    table = SHAFT_HOOK_LEG_MM_BY_BAR_DIAMETER_MM
    try:
        d = float(d_mm)
    except Exception:
        d = float(table[0][0])
    d_min, h_min = table[0]
    d_max, h_max = table[-1]
    if d <= d_min:
        return float(h_min), u"Ø≤{0} mm → pata {1} mm".format(int(d_min), int(h_min))
    if d >= d_max:
        return float(h_max), u"Ø≥{0} mm → pata {1} mm".format(int(d_max), int(h_max))
    for i in range(len(table) - 1):
        d_lo, h_lo = table[i]
        d_hi, h_hi = table[i + 1]
        if d_lo <= d <= d_hi:
            if abs(d - d_lo) < 1e-9:
                return float(h_lo), u"Ø{0} mm → pata {1} mm".format(int(d_lo), int(h_lo))
            if abs(d - d_hi) < 1e-9:
                return float(h_hi), u"Ø{0} mm → pata {1} mm".format(int(d_hi), int(h_hi))
            span = float(d_hi) - float(d_lo)
            if span <= 1e-12:
                return float(h_lo), u"Ø{0} mm → pata {1} mm".format(int(d_lo), int(h_lo))
            t = (d - float(d_lo)) / span
            leg = float(h_lo) + t * (float(h_hi) - float(h_lo))
            return leg, (
                u"Pata gancho interpol. Ø{0:.1f} mm → {1:.0f} mm (entre Ø{2} y Ø{3})"
                .format(d, leg, int(d_lo), int(d_hi))
            )
    return float(h_min), u"(tabla pata)"


def hook_leg_mm_para_bar_type(bar_type):
    """Largo de pata (mm) según diámetro nominal del RebarBarType."""
    if bar_type is None:
        return 0.0
    try:
        d_mm = _ft_to_mm(float(bar_type.BarNominalDiameter))
    except Exception:
        return 0.0
    leg, _t = hook_leg_mm_por_diametro_nominal_mm(d_mm)
    return float(leg)


def _xyz_add(a, b):
    return XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)


def _xyz_sub(a, b):
    return XYZ(a.X - b.X, a.Y - b.Y, a.Z - b.Z)


def _dot_xyz(a, b):
    if a is None or b is None:
        return 0.0
    return (
        float(a.X) * float(b.X)
        + float(a.Y) * float(b.Y)
        + float(a.Z) * float(b.Z)
    )


def _xyz_scale(v, s):
    return XYZ(v.X * s, v.Y * s, v.Z * s)


def _unit_xy(v):
    ln = math.sqrt(v.X * v.X + v.Y * v.Y)
    if ln < 1e-12:
        return None
    return XYZ(v.X / ln, v.Y / ln, 0.0)


def _unit_3d(v):
    ln = math.sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z)
    if ln < 1e-12:
        return None
    return XYZ(v.X / ln, v.Y / ln, v.Z / ln)


def _is_vertical_planar_face(face):
    if not isinstance(face, PlanarFace):
        return False
    try:
        n = face.FaceNormal
        if n is None:
            return False
        return abs(float(n.Z)) < 0.2
    except Exception:
        return False


def _top_horizontal_edge_segment(face):
    """
    En una cara vertical del hueco, toma la arista horizontal (misma Z) más alta en Z.
    Devuelve (p0, p1) en 3D o None.
    """
    best = None
    best_z = None
    try:
        loops = face.GetEdgesAsCurveLoops()
    except Exception:
        return None
    if loops is None:
        return None
    for cl in loops:
        if cl is None:
            continue
        for c in cl:
            if c is None or not c.IsBound:
                continue
            try:
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
            except Exception:
                continue
            if abs(p0.Z - p1.Z) > 1e-4:
                continue
            zm = (p0.Z + p1.Z) * 0.5
            if best is None or zm > best_z:
                best = (p0, p1)
                best_z = zm
    return best


def _curve_loop_total_length(curve_loop):
    total = 0.0
    if curve_loop is None:
        return total
    for c in curve_loop:
        if c is None:
            continue
        try:
            total += float(c.Length)
        except Exception:
            continue
    return total


def _outer_curve_loop(face):
    """Devuelve el loop exterior de la cara (aprox.: mayor perímetro)."""
    try:
        loops = face.GetEdgesAsCurveLoops()
    except Exception:
        return None
    if loops is None:
        return None
    best = None
    best_len = -1.0
    for cl in loops:
        if cl is None:
            continue
        ln = _curve_loop_total_length(cl)
        if ln > best_len:
            best_len = ln
            best = cl
    return best


def _is_curve_horizontal(curve):
    if curve is None:
        return False
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
    except Exception:
        return False
    return abs(float(p0.Z) - float(p1.Z)) <= 1e-4


def _is_line_curve(curve):
    if curve is None:
        return False
    try:
        # En IronPython/Revit, el chequeo por nombre evita falsos negativos de isinstance.
        return curve.GetType().Name == "Line"
    except Exception:
        return isinstance(curve, Line)


def _bottom_curves_from_outer_loop(face):
    """
    Desde el perímetro exterior de la cara, obtiene curvas del borde inferior.
    Prioriza tramos horizontales en la cota Z mínima del loop exterior.
    """
    outer = _outer_curve_loop(face)
    if outer is None:
        return None, u"No se pudo obtener el perímetro exterior de la cara."

    curves = []
    z_min = None
    for c in outer:
        if c is None:
            continue
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
        except Exception:
            continue
        zc = min(float(p0.Z), float(p1.Z))
        if z_min is None or zc < z_min:
            z_min = zc
        curves.append((c, p0, p1, zc))

    if not curves or z_min is None:
        return None, u"El perímetro exterior no contiene curvas válidas."

    z_tol = max(1e-4, _mm_to_ft(1.0))
    out = []
    for c, p0, p1, zc in curves:
        if zc > z_min + z_tol:
            continue
        if not _is_curve_horizontal(c):
            continue
        if not _is_line_curve(c):
            continue
        out.append(c)

    if not out:
        return None, u"No se encontró curva inferior horizontal en el perímetro exterior."
    return out, None


def _detail_line_first_bottom_curve_debug(
    doc,
    view,
    face,
    q0,
    q1,
    offset_mm,
    solids,
    extend_ft_for_embed,
    avisos=None,
    face_1based=None,
):
    """
    Debug exacto pedido:
      1) Extender axialmente 100 mm desde cada extremo (solo para "detectar colisión").
      2) Evaluar colisión con el host (SCI) y guardar qué extremo colisiona.
      3) Restaurar a la curva original (sin esos 100 mm).
      4) Aplicar estiramiento por Ø (extend_ft_for_embed) SOLO en el/los extremo(s) que colisionó.
      5) Dibujar DetailCurve desplazada en planta (offset_mm).
    """
    if doc is None or view is None or face is None or q0 is None or q1 is None:
        return
    try:
        off_mm = float(offset_mm)
    except Exception:
        off_mm = 0.0
    if off_mm <= 1e-6:
        return
    try:
        probe_mm = float(SHAFT_DEBUG_FIRST_BOTTOM_EDGE_EXTEND_EACH_END_MM)
    except Exception:
        probe_mm = 0.0
    probe_mm = max(0.0, probe_mm)
    try:
        emb_ft = max(0.0, float(extend_ft_for_embed))
    except Exception:
        emb_ft = 0.0

    d_xy = _xyz_sub(q1, q0)
    u = _unit_xy(XYZ(float(d_xy.X), float(d_xy.Y), 0.0))
    if u is None:
        return
    perp = _rotate90_xy(u)
    if perp is None:
        return

    p0 = XYZ(float(q0.X), float(q0.Y), float(q0.Z))
    p1 = XYZ(float(q1.X), float(q1.Y), float(q1.Z))

    # 1) Colisión con sonda axial a 100 mm
    host_hit_start = False
    host_hit_end = False
    if probe_mm > 1e-9 and solids:
        probe_ft = _mm_to_ft(probe_mm)
        s_tip0 = _xyz_add(p0, _xyz_scale(u, -probe_ft))
        s_tip1 = _xyz_add(p1, _xyz_scale(u, probe_ft))
        host_hit_start = bool(_endpoint_inside_host_concrete(s_tip0, u, solids))
        host_hit_end = bool(_endpoint_inside_host_concrete(s_tip1, u, solids))

    # 3) Restaurar y 4) estirar SOLO donde colisionó
    p0r = p0
    p1r = p1
    if emb_ft > 1e-12:
        if host_hit_start:
            p0r = _xyz_add(p0r, _xyz_scale(u, -emb_ft))
        if host_hit_end:
            p1r = _xyz_add(p1r, _xyz_scale(u, emb_ft))

    # 5) DetailCurve desplazada en planta
    off_ft = _mm_to_ft(off_mm)
    sh = _xyz_scale(perp, off_ft)
    z = _z_plano_para_detail_curves(view)
    lbl = int(face_1based) if face_1based is not None else 0
    try:
        if avisos is not None:
            if not solids:
                avisos.append(
                    u"Debug cara [{0}]: sin sólidos de host; no se evaluó colisión 100 mm."
                    .format(lbl)
                )
            elif probe_mm <= 1e-9:
                avisos.append(
                    u"Debug cara [{0}]: sonda 0 mm; no se evaluó colisión."
                    .format(lbl)
                )
            else:
                st = u"SI" if host_hit_start else u"NO"
                en = u"SI" if host_hit_end else u"NO"
                if emb_ft > 1e-12:
                    dest = u"inicio" if host_hit_start else u""
                    if host_hit_end and host_hit_start:
                        dest = u"inicio+fin"
                    elif host_hit_end and not host_hit_start:
                        dest = u"fin"
                    elif (not host_hit_end) and (not host_hit_start):
                        dest = u"ninguno"
                    avisos.append(
                        u"Debug cara [{0}]: colisión 100 mm → inicio={1}, fin={2}. "
                        u"Estirón Ø (tab)={3:.0f} mm aplicado en: {4}."
                        .format(lbl, st, en, _ft_to_mm(emb_ft), dest)
                    )
                else:
                    avisos.append(
                        u"Debug cara [{0}]: colisión 100 mm → inicio={1}, fin={2}. "
                        u"Estirón Ø tab=0 (extend_ft_for_embed=0)."
                        .format(lbl, st, en)
                    )
    except Exception:
        pass

    try:
        a = XYZ(float(p0r.X) + float(sh.X), float(p0r.Y) + float(sh.Y), float(z))
        b = XYZ(float(p1r.X) + float(sh.X), float(p1r.Y) + float(sh.Y), float(z))
        ln = Line.CreateBound(a, b)
        doc.Create.NewDetailCurve(view, ln)
    except Exception:
        pass


def _line_segment_from_curve_with_margins(curve, margin_start_ft, margin_end_ft):
    """
    Convierte una curva inferior a segmento lineal utilizable para Rebar.
    Recorta el extremo GetEndPoint(0) en ``margin_start_ft`` y el (1) en ``margin_end_ft``
    a lo largo del eje q0→q1 (mismo convenio que el estirón SCI en `_extend_segment_both_ends`).
    Suele ser cover en arista + Ø/2 para no pegar el eje al borde; en empotramiento adaptivo
    puede anularse el recorte en el extremo que queda anclado en hormigón adyacente.
    """
    if curve is None:
        return None, u"Curva nula."
    if not _is_line_curve(curve):
        return None, u"La curva inferior no es lineal (Line)."
    try:
        m0 = max(0.0, float(margin_start_ft))
    except Exception:
        m0 = 0.0
    try:
        m1 = max(0.0, float(margin_end_ft))
    except Exception:
        m1 = 0.0
    try:
        q0 = curve.GetEndPoint(0)
        q1 = curve.GetEndPoint(1)
    except Exception:
        return None, u"No se pudo obtener el tramo inferior."
    try:
        raw = _xyz_sub(q1, q0)
        ln = math.sqrt(raw.X * raw.X + raw.Y * raw.Y + raw.Z * raw.Z)
    except Exception:
        return None, u"Curva inferior inválida."
    if ln <= 1e-9:
        return None, u"Curva inferior degenerada."
    if m0 > 1e-12 or m1 > 1e-12:
        if ln <= m0 + m1 + _mm_to_ft(0.1):
            return None, u"Tramo inferior muy corto para el recubrimiento+Ø/2."
        u = _unit_3d(raw)
        if u is None:
            return None, u"Dirección de tramo inválida."
        q0 = _xyz_add(q0, _xyz_scale(u, m0))
        q1 = _xyz_sub(q1, _xyz_scale(u, m1))
    try:
        seg = Line.CreateBound(q0, q1)
        if float(seg.Length) <= 1e-6:
            return None, u"Segmento inferior degenerado tras recorte."
    except Exception:
        return None, u"Segmento inferior inválido tras recorte."
    return (q0, q1), None


def _line_segment_from_curve_with_margin(curve, margin_ft):
    """Wrapper: mismo recorte longitudinal en ambos extremos."""
    try:
        m = max(0.0, float(margin_ft))
    except Exception:
        m = 0.0
    return _line_segment_from_curve_with_margins(curve, m, m)


def _host_bbox_center(host):
    if host is None:
        return None
    try:
        bb = host.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return None
    try:
        mn = bb.Min
        mx = bb.Max
        return XYZ(
            0.5 * (float(mn.X) + float(mx.X)),
            0.5 * (float(mn.Y) + float(mx.Y)),
            0.5 * (float(mn.Z) + float(mx.Z)),
        )
    except Exception:
        return None


def _lateral_offset_from_face_normal_xy(n_xy, distance_ft, mid, solids):
    """
    Vector XY (mismo plano que n_xy) para trasladar el tramo hacia el hormigón.
    Usa la projected FaceNormal de la cara del tramo (no un centro global del hueco).
    Con sólidos del host, el signo se resuelve con _infer_inward_xy_sign; si no hay
    datos, se asume convención Revit (normal hacia fuera del sólido → interior = -n_xy).
    """
    d = max(0.0, float(distance_ft))
    if n_xy is None or d <= 1e-12:
        return XYZ(0.0, 0.0, 0.0)
    # Cuando el offset lateral es "solo recubrimiento" (capa 1), la sonda
    # con una distancia mínima fija (25mm) puede quedar demasiado cerca del
    # borde del vano y dar signo ambiguo al seleccionar 2-4 caras.
    # Al elevar el mínimo de sonda (~50mm) estabilizamos la inferencia del signo.
    probe_min_ft = _mm_to_ft(0.5 * float(LAYER_SPACING_MM_DEFAULT))
    probe_ft = max(d, probe_min_ft)
    sgn = None
    if solids:
        sgn = _infer_inward_xy_sign(mid, n_xy, solids, probe_ft)
    if sgn is None:
        sgn = -1.0
    return _xyz_scale(n_xy, float(sgn) * d)


def _offset_segment_into_concrete(q0, q1, face, host, cover_ft, solids=None):
    """
    Aplica offset lateral por recubrimiento desde la cara del shaft hacia el interior del hormigón.
    El sentido en planta se deriva de la misma PlanarFace que aporta el tramo (v. _lateral_offset_from_face_normal_xy).
    """
    if q0 is None or q1 is None or face is None:
        return q0, q1
    try:
        n = face.FaceNormal
    except Exception:
        n = None
    n_xy = _unit_xy(XYZ(n.X, n.Y, 0.0)) if n is not None else None
    if n_xy is None:
        return q0, q1

    mid = XYZ(
        0.5 * (float(q0.X) + float(q1.X)),
        0.5 * (float(q0.Y) + float(q1.Y)),
        0.5 * (float(q0.Z) + float(q1.Z)),
    )
    off = _lateral_offset_from_face_normal_xy(n_xy, float(cover_ft), mid, solids or [])
    return _xyz_add(q0, off), _xyz_add(q1, off)


def _offset_segment_layer_inside_host(q0, q1, face, host, extra_offset_ft, solids=None):
    """
    Desplaza un tramo horizontal más hacia interior del host; mismo criterio de signo
    que el recubrimiento (normal de la misma cara + sonda en sólidos).
    """
    if q0 is None or q1 is None or face is None:
        return q0, q1
    extra = max(0.0, float(extra_offset_ft))
    if extra <= 1e-9:
        return q0, q1
    try:
        n = face.FaceNormal
    except Exception:
        n = None
    n_xy = _unit_xy(XYZ(n.X, n.Y, 0.0)) if n is not None else None
    if n_xy is None:
        return q0, q1

    mid = XYZ(
        0.5 * (float(q0.X) + float(q1.X)),
        0.5 * (float(q0.Y) + float(q1.Y)),
        0.5 * (float(q0.Z) + float(q1.Z)),
    )
    off = _lateral_offset_from_face_normal_xy(n_xy, extra, mid, solids or [])
    return _xyz_add(q0, off), _xyz_add(q1, off)


def _apply_vertical_cover_inside_host(q0, q1, host, cover_ft):
    """
    Asegura que la barra quede dentro del host en Z considerando recubrimiento.
    Si el tramo está más cerca de la cara inferior/superior del host, lo desplaza
    hacia el interior al menos cover_ft.
    """
    if q0 is None or q1 is None or host is None:
        return q0, q1
    try:
        bb = host.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return q0, q1

    try:
        zmin = float(bb.Min.Z)
        zmax = float(bb.Max.Z)
        z0 = float(q0.Z)
        z1 = float(q1.Z)
        zm = 0.5 * (z0 + z1)
        cover = max(0.0, float(cover_ft))
    except Exception:
        return q0, q1

    if zmax <= zmin + 1e-9:
        return q0, q1

    target_lo = zmin + cover
    target_hi = zmax - cover
    if target_hi <= target_lo + 1e-9:
        return q0, q1

    # Si está más cerca del fondo o bajo recubrimiento inferior, subir.
    # Si está más cerca de la cara superior o sobre recubrimiento superior, bajar.
    dz = 0.0
    if zm < target_lo:
        dz = target_lo - zm
    elif zm > target_hi:
        dz = target_hi - zm
    else:
        dist_lo = abs(zm - zmin)
        dist_hi = abs(zmax - zm)
        if dist_lo < cover:
            dz = cover - dist_lo
        elif dist_hi < cover:
            dz = -(cover - dist_hi)

    if abs(dz) <= 1e-9:
        return q0, q1
    shift = XYZ(0.0, 0.0, dz)
    return _xyz_add(q0, shift), _xyz_add(q1, shift)


def _extend_segment_both_ends(q0, q1, extend_ft):
    """Extiende un tramo lineal en ambas direcciones por la misma distancia."""
    if q0 is None or q1 is None:
        return q0, q1
    ext = max(0.0, float(extend_ft))
    if ext <= 1e-9:
        return q0, q1
    d = _xyz_sub(q1, q0)
    u = _unit_3d(d)
    if u is None:
        return q0, q1
    return _xyz_add(q0, _xyz_scale(u, -ext)), _xyz_add(q1, _xyz_scale(u, ext))


def _extend_segment_one_end(q0, q1, extend_ft, end_index=1):
    """
    Extiende un tramo lineal solo desde un extremo.

    end_index:
        0 -> mueve q0 hacia atrás (opuesto a la dirección q0->q1)
        1 -> mueve q1 hacia adelante (dirección q0->q1)
    """
    if q0 is None or q1 is None:
        return q0, q1
    ext = max(0.0, float(extend_ft))
    if ext <= 1e-9:
        return q0, q1
    d = _xyz_sub(q1, q0)
    u = _unit_3d(d)
    if u is None:
        return q0, q1
    if int(end_index) == 0:
        return _xyz_add(q0, _xyz_scale(u, -ext)), q1
    return q0, _xyz_add(q1, _xyz_scale(u, ext))


def _round_up_length_mm_to_step(length_mm, step_mm):
    """Redondea hacia arriba a múltiplo de step_mm (en mm)."""
    try:
        ln = float(length_mm)
        st = float(step_mm)
    except Exception:
        return None
    if st <= 1e-9:
        return None
    # Evita errores por coma flotante cuando ya es múltiplo exacto.
    n = ln / st
    return st * float(int(math.ceil(n - 1e-12)))


def _round_down_length_mm_to_step(length_mm, step_mm):
    """Redondea hacia abajo a múltiplo de step_mm (en mm)."""
    try:
        ln = float(length_mm)
        st = float(step_mm)
    except Exception:
        return None
    if st <= 1e-9:
        return None
    n = ln / st
    return st * float(int(math.floor(n + 1e-12)))


def _round_up_segment_length_by_extending_one_end(q0, q1, step_mm=None, end_index=1):
    """
    Si la longitud del segmento (mm) no es múltiplo de step_mm, extiende SOLO un extremo
    para alcanzar el siguiente múltiplo hacia arriba.

    Devuelve:
        (new_q0, new_q1, rounded_applied:bool, delta_mm:float)
    """
    if step_mm is None:
        step_mm = SHAFT_BAR_LENGTH_ROUND_STEP_MM
    if q0 is None or q1 is None:
        return q0, q1, False, 0.0
    ln_ft = _segment_length_ft(q0, q1)
    if ln_ft <= 1e-9:
        return q0, q1, False, 0.0
    ln_mm = _ft_to_mm(ln_ft)
    target_mm = _round_up_length_mm_to_step(ln_mm, step_mm)
    if target_mm is None:
        return q0, q1, False, 0.0
    delta_mm = float(target_mm) - float(ln_mm)
    if delta_mm <= 1e-6:
        return q0, q1, False, 0.0
    dq0, dq1 = _extend_segment_one_end(q0, q1, _mm_to_ft(delta_mm), end_index=end_index)
    return dq0, dq1, True, float(delta_mm)


def _round_up_segment_length_symmetric_about_mid(q0, q1, step_mm=None):
    """
    Sube la longitud al siguiente múltiplo de ``step_mm`` repartiendo el incremento por
    ambos extremos de forma equidistante (mismo punto medio y dirección que el tramo).
    """
    if step_mm is None:
        step_mm = SHAFT_BAR_LENGTH_ROUND_STEP_MM
    if q0 is None or q1 is None:
        return q0, q1, False, 0.0
    ln_ft = _segment_length_ft(q0, q1)
    if ln_ft <= 1e-9:
        return q0, q1, False, 0.0
    ln_mm = _ft_to_mm(ln_ft)
    target_mm = _round_up_length_mm_to_step(ln_mm, step_mm)
    if target_mm is None:
        return q0, q1, False, 0.0
    delta_mm = float(target_mm) - float(ln_mm)
    if delta_mm <= 1e-6:
        return q0, q1, False, 0.0
    new_len_ft = _mm_to_ft(float(target_mm))
    nq0, nq1 = _resize_segment_keep_mid_dir(q0, q1, new_len_ft)
    return nq0, nq1, True, float(delta_mm)


def _round_up_segment_length_adaptive_embed_end(
    q0,
    q1,
    stretch_kept0,
    stretch_kept1,
    step_mm=None,
):
    """
    Sube la longitud al siguiente múltiplo de step (solo tramo recto). El incremento
    se aplica en el extremo **libre** (sin estirón de anclaje por Ø), para no mover el
    extremo empotrado y respetar la cota tabulada. ``stretch_kept*`` True = ese extremo
    es el de anclaje (sonda SCI / empotramiento).

    Si **ambos** extremos retienen estirón, el redondeo no se reparte en ambos cantos:
    el incremento va solo al extremo **final** de la curva (``q1``, dirección q0→q1).
    Si ninguno retiene estirón, se mantiene el reparto simétrico.
    """
    if step_mm is None:
        step_mm = SHAFT_BAR_LENGTH_ROUND_STEP_MM
    if q0 is None or q1 is None:
        return q0, q1, False, 0.0
    k0 = bool(stretch_kept0)
    k1 = bool(stretch_kept1)
    if k0 and (not k1):
        # Anclaje en q0 → redondeo solo hacia q1 (end_index=1).
        return _round_up_segment_length_by_extending_one_end(
            q0, q1, step_mm, end_index=1
        )
    if k1 and (not k0):
        # Anclaje en q1 → redondeo solo hacia q0 (end_index=0).
        return _round_up_segment_length_by_extending_one_end(
            q0, q1, step_mm, end_index=0
        )
    if k0 and k1:
        # Dos anclajes tabulados: no repartir el delta entre extremos; solo en q1.
        return _round_up_segment_length_by_extending_one_end(
            q0, q1, step_mm, end_index=1
        )
    return _round_up_segment_length_symmetric_about_mid(q0, q1, step_mm)


def _split_long_bar_with_overlap_segments(
    q0,
    q1,
    max_len_mm,
    chunk_mm,
    overlap_mm,
    hook_dev_first_mm=0.0,
    hook_dev_last_mm=0.0,
):
    """
    Devuelve:
      - subtramos: lista de (p0, p1) para una barra base
      - joints: lista de dict por unión con:
          {lap_start, lap_end, axis_u, has_lap}

    - Si largo <= max_len_mm, retorna el tramo original.
    - Si largo > max_len_mm, divide en tramos consecutivos y aplica traslapo nominal.

    Con gancho en el **primer** subtramo y más tramos detrás, ``chunk_mm`` es el cupo
    **total** habitual del tramo (recto + pata en desarrollo; p. ej. 6000 mm): el ``Line``
    del primer Rebar se acorta en ``hook_dev_first_mm`` para dejar la pata dentro del cupo.
    Tramos siguientes con ``chunk_mm`` recto pleno salvo regla del último tramo con
    ``hook_dev_last_mm``. El avance respeta el traslapo nominal ``overlap_mm``.
    """
    if q0 is None or q1 is None:
        return [], []
    ln_ft = _segment_length_ft(q0, q1)
    if ln_ft <= 1e-9:
        return [], []
    try:
        max_len_ft = _mm_to_ft(float(max_len_mm))
    except Exception:
        max_len_ft = _mm_to_ft(MAX_SINGLE_BAR_LENGTH_MM)
    if ln_ft <= max_len_ft + 1e-9:
        return [(q0, q1)], []

    try:
        step_ft = _mm_to_ft(float(chunk_mm))
    except Exception:
        step_ft = _mm_to_ft(LONG_BAR_SPLIT_CHUNK_MM)
    step_ft = max(step_ft, _mm_to_ft(100.0))

    try:
        ov_ft = max(0.0, _mm_to_ft(float(overlap_mm)))
    except Exception:
        ov_ft = 0.0

    d = _xyz_sub(q1, q0)
    u = _unit_3d(d)
    if u is None:
        return [(q0, q1)], []

    # Para que los tramos no "crezcan" por sumar traslapo en su largo:
    # - largo objetivo por tramo = step_ft (6 m por defecto)
    # - avance entre inicios = step_ft - ov_ft
    # Así se mantiene el traslapo entre tramos sin inflar longitudes.
    stride_ft = step_ft - ov_ft
    if stride_ft <= _mm_to_ft(10.0):
        stride_ft = step_ft
        ov_ft = 0.0

    try:
        hook_f_ft = max(0.0, float(_mm_to_ft(float(hook_dev_first_mm))))
    except Exception:
        hook_f_ft = 0.0
    try:
        hook_l_ft = max(0.0, float(_mm_to_ft(float(hook_dev_last_mm))))
    except Exception:
        hook_l_ft = 0.0

    out = []
    joints = []
    covered_ft = 0.0
    total_ft = float(ln_ft)
    for _ in range(1000):
        if covered_ft >= total_ft - 1e-9:
            break
        remaining_ft = total_ft - covered_ft
        # Tramos intermedios a 6m; último tramo toma el remanente real.
        if remaining_ft > step_ft + 1e-9:
            seg_len_ft = step_ft
            has_next = True
        else:
            seg_len_ft = remaining_ft
            has_next = False
        is_first = len(out) == 0
        if has_next and hook_f_ft > 1e-9 and is_first:
            cand = min(remaining_ft, step_ft - hook_f_ft)
            if cand >= stride_ft - 1e-9 and cand > 1e-6:
                cand_mm = _ft_to_mm(cand)
                cand_mm_dn = _round_down_length_mm_to_step(
                    cand_mm, SHAFT_BAR_LENGTH_ROUND_STEP_MM
                )
                if cand_mm_dn is not None:
                    cand_dn_ft = _mm_to_ft(cand_mm_dn)
                    if cand_dn_ft >= stride_ft - 1e-9 and cand_dn_ft > 1e-6:
                        cand = cand_dn_ft
                seg_len_ft = cand
        if (not has_next) and hook_l_ft > 1e-9 and len(out) > 0:
            seg_len_ft = max(0.0, remaining_ft - hook_l_ft)
            # Si la pata supera el remanente, no restar (evita largo 0).
            if seg_len_ft < _mm_to_ft(10.0):
                seg_len_ft = remaining_ft
        # Avance 0 → bucle infinito; Line de longitud 0 → API Revit inestable.
        if seg_len_ft < 1e-9:
            seg_len_ft = remaining_ft
        if seg_len_ft < 1e-9:
            break
        advance_ft = seg_len_ft if not has_next else stride_ft
        if has_next and is_first and hook_f_ft > 1e-9 and seg_len_ft < step_ft - 1e-6:
            advance_ft = max(_mm_to_ft(5.0), seg_len_ft - ov_ft)
        p0 = _xyz_add(q0, _xyz_scale(u, covered_ft))
        p1 = _xyz_add(p0, _xyz_scale(u, seg_len_ft))
        out.append((p0, p1))
        if has_next and ov_ft > 1e-9:
            # Zona de traslapo = [inicio siguiente, fin tramo actual]
            next_start = _xyz_add(p0, _xyz_scale(u, advance_ft))
            joints.append(
                {
                    "lap_start": next_start,
                    "lap_end": p1,
                    "axis_u": u,
                    "has_lap": True,
                }
            )
        covered_ft += advance_ft
    if not out:
        return [(q0, q1)], []
    return out, joints


def _resolve_shaft_bar_type(doc, forced_bar_type_id=None):
    bar_type = None
    exact_bt = False
    delta_bt = None
    if forced_bar_type_id is not None and forced_bar_type_id != ElementId.InvalidElementId:
        try:
            bt_forced = doc.GetElement(forced_bar_type_id)
            if isinstance(bt_forced, RebarBarType):
                bar_type = bt_forced
        except Exception:
            bar_type = None
    if bar_type is None:
        bar_type, exact_bt, delta_bt = resolver_bar_type_por_diametro_mm(doc, 12.0)
    return bar_type, exact_bt, delta_bt


def evaluar_si_excede_12m_en_shaft(
    doc,
    host,
    refs,
    cover_mm=None,
    forced_bar_type_id=None,
    ignore_empotramientos=True,
    empotramiento_adaptivo_extremos=False,
):
    """
    Evalúa si alguna barra potencial superaría 12m antes de crear Rebar.
    Alineado con `crear_enfierrado_shaft_hashtag`: sin estirón por empotramiento en
    geometría cuando ignore_empotramientos=True.
    Con empotramiento_adaptivo_extremos=True e ignore_empotramientos=False, usa el
    mismo recorte de estirón por extremo que la creación de barras.

    Returns: (excede:bool, max_len_mm:float, err:unicode|None)
    """
    if doc is None or host is None or not refs:
        return False, 0.0, u"No hay documento, host o caras."
    if not isinstance(host, Floor):
        return False, 0.0, u"El host no es una losa (Floor)."
    bar_type, _exact_bt, _delta_bt = _resolve_shaft_bar_type(doc, forced_bar_type_id)
    if bar_type is None:
        return False, 0.0, u"No hay RebarBarType compatible."
    cov = float(cover_mm if cover_mm is not None else COVER_MM_DEFAULT)
    cover_ft = _mm_to_ft(cov)
    bar_diam_ft = float(_bar_nominal_diameter_ft(bar_type) or 0.0)
    cover_center_ft = float(cover_ft) + 0.5 * max(bar_diam_ft, 1e-6)
    solids = _collect_solids_from_host_element(host)
    extend_each_side_mm, _txt, _d = extension_mm_para_bar_type(bar_type)
    extend_ft = _mm_to_ft(extend_each_side_mm)
    extend_ft_geom = 0.0 if bool(ignore_empotramientos) else float(extend_ft)
    max_layer_extra_sci = _mm_to_ft(SHAFT_FIRST_LAYER_LATERAL_EXTRA_MM_DEFAULT)
    adaptive_clip = bool(empotramiento_adaptivo_extremos) and (not bool(ignore_empotramientos))
    max_len_ft = 0.0
    for rf in refs:
        try:
            face = host.GetGeometryObjectFromReference(rf)
        except Exception:
            face = None
        if not isinstance(face, PlanarFace):
            continue
        (_fk, segments), err = _horizontal_offset_segments_for_face(
            face,
            cover_ft,
            None,
            host,
            bar_type,
            solids,
            stretch_ft_for_sci=extend_ft_geom,
            layer_extra_ft_for_sci=max_layer_extra_sci,
            adaptive_embed_end_clip=adaptive_clip,
        )
        if err or (not segments):
            continue
        for q0, q1 in segments:
            if adaptive_clip and extend_ft_geom > 1e-12:
                q0e, q1e, sk0, sk1 = _shaft_probe_hit_asymmetric_extend_then_offset_into_host(
                    q0,
                    q1,
                    face,
                    host,
                    extend_ft_geom,
                    cover_center_ft,
                    _mm_to_ft(SHAFT_FIRST_LAYER_LATERAL_EXTRA_MM_DEFAULT),
                    solids,
                )
                q0e, q1e, _rounded, _delta = _round_up_segment_length_adaptive_embed_end(
                    q0e, q1e, sk0, sk1
                )
            else:
                q0e, q1e = _shaft_extend_then_offset_into_host(
                    q0,
                    q1,
                    face,
                    host,
                    extend_ft_geom,
                    cover_center_ft,
                    _mm_to_ft(SHAFT_FIRST_LAYER_LATERAL_EXTRA_MM_DEFAULT),
                    solids,
                )
                q0e, q1e, _rounded, _delta = _round_up_segment_length_symmetric_about_mid(
                    q0e, q1e
                )
            ln = _segment_length_ft(q0e, q1e)
            if ln > max_len_ft:
                max_len_ft = ln
    max_len_mm = _ft_to_mm(max_len_ft)
    return (max_len_mm > 12000.0 + 1e-6), float(max_len_mm), None


def _create_overlap_dimension_between_markers(
    doc,
    view,
    lap_start,
    lap_end,
    axis_u,
    lateral_hint=None,
    marker_len_mm=5.0,
    line_offset_mm=450.0,
):
    """
    Crea cota de traslapo entre dos marcadores (inicio/fin de traslapo).
    Retorna (ok:bool, msg:unicode|None, data:dict|None).
    """
    if doc is None or view is None or lap_start is None or lap_end is None:
        return False, u"Parámetros incompletos para cota de traslapo.", None
    dxy = _unit_xy(axis_u if axis_u is not None else _xyz_sub(lap_end, lap_start))
    if dxy is None:
        return False, u"No se pudo obtener dirección de traslapo.", None
    tdir = _rotate90_xy(dxy)
    if tdir is None:
        if lateral_hint is not None:
            tdir = _unit_xy(lateral_hint)
        if tdir is None:
            tdir = XYZ(1.0, 0.0, 0.0)
    out_vec = _xyz_scale(tdir, _mm_to_ft(float(line_offset_mm)))
    # Los marcadores (líneas auxiliares) deben quedar en la posición real de la barra.
    m0 = lap_start
    m1 = lap_end
    # Línea auxiliar perpendicular a la barra:
    # _create_marker_detailcurve traza la tangente de la "normal" recibida.
    # Si enviamos dxy como normal, la tangente resultante queda a 90° del eje.
    marker_face_nxy = dxy
    if marker_face_nxy is None:
        marker_face_nxy = XYZ(1.0, 0.0, 0.0)
    dc0, ref0 = _create_marker_detailcurve(
        doc, view, m0, marker_face_nxy, length_mm=float(marker_len_mm)
    )
    dc1, ref1 = _create_marker_detailcurve(
        doc, view, m1, marker_face_nxy, length_mm=float(marker_len_mm)
    )
    if ref0 is None or ref1 is None:
        try:
            if dc0 is not None:
                doc.Delete(dc0.Id)
        except Exception:
            pass
        try:
            if dc1 is not None:
                doc.Delete(dc1.Id)
        except Exception:
            pass
        return False, u"No fue posible crear referencias de marcadores para traslapo.", None

    try:
        from Autodesk.Revit.DB import ReferenceArray

        z = _z_plano_para_detail_curves(view)
        # La línea de cota sí puede ir desplazada para legibilidad, pero midiendo
        # entre referencias ubicadas sobre la barra.
        a = XYZ(float(m0.X) + float(out_vec.X), float(m0.Y) + float(out_vec.Y), float(z))
        b = XYZ(float(m1.X) + float(out_vec.X), float(m1.Y) + float(out_vec.Y), float(z))
        dim_line = Line.CreateBound(a, b)
        ra = ReferenceArray()
        ra.Append(ref0)
        ra.Append(ref1)
        dim = doc.Create.NewDimension(view, dim_line, ra)
    except Exception:
        dim = None
    if dim is None:
        try:
            if dc0 is not None:
                doc.Delete(dc0.Id)
        except Exception:
            pass
        try:
            if dc1 is not None:
                doc.Delete(dc1.Id)
        except Exception:
            pass
        return False, u"Revit no permitió crear la cota de traslapo.", None
    nrefs = _dimension_reference_count(dim)
    if (nrefs is None) or (nrefs != 2):
        try:
            doc.Delete(dim.Id)
        except Exception:
            pass
        try:
            if dc0 is not None:
                doc.Delete(dc0.Id)
        except Exception:
            pass
        try:
            if dc1 is not None:
                doc.Delete(dc1.Id)
        except Exception:
            pass
        return False, u"Cota de traslapo inválida (referencias extra).", None
    _try_apply_fixed_dimension_type(doc, dim)
    data = {
        "dim_id": _eid_int(dim.Id),
        "marker_a_id": _eid_int(dc0.Id) if dc0 is not None else None,
        "marker_b_id": _eid_int(dc1.Id) if dc1 is not None else None,
    }
    return True, None, data


def _endpoint_inside_host_concrete(point_xyz, axis_unit, solids):
    """
    True si un tramo corto centrado en `point_xyz` a lo largo de `axis_unit` queda
    dentro del hormigón (SolidCurveIntersection curve inside), coherente con
    _rebar_mid_stub_inside_concrete.
    """
    if point_xyz is None or axis_unit is None:
        return False
    if not solids:
        return True
    t = axis_unit
    tlen = math.sqrt(t.X * t.X + t.Y * t.Y + t.Z * t.Z)
    if tlen < 1e-9:
        return False
    stub = max(1e-3, min(0.02, 0.05))  # ~15 mm; estable en pies
    ux, uy, uz = t.X / tlen, t.Y / tlen, t.Z / tlen
    px = float(point_xyz.X)
    py = float(point_xyz.Y)
    pz = float(point_xyz.Z)
    p_a = XYZ(px - ux * stub, py - uy * stub, pz - uz * stub)
    p_b = XYZ(px + ux * stub, py + uy * stub, pz + uz * stub)
    try:
        ln = Line.CreateBound(p_a, p_b)
        clen = float(ln.Length)
    except Exception:
        return False
    if clen < 1e-9:
        return False
    mid_param = 0.5 * clen
    for solid in solids:
        try:
            if solid.Volume < 1e-12:
                continue
            intervals = _solid_line_inside_param_intervals(ln, solid)
            merged = _merge_axis_intervals(intervals)
            for a, b in merged:
                lo = max(0.0, float(a))
                hi = min(clen, float(b))
                if lo <= mid_param <= hi:
                    return True
        except Exception:
            continue
    return False


def _shaft_adaptive_extend_then_offset_into_host(
    q0,
    q1,
    face,
    host,
    extend_ft,
    lateral_cover_ft,
    layer_extra_ft,
    solids,
    vertical_cover_ft=None,
    embed_clip_avisos=None,
    clip_label=None,
):
    """
    Estira por tabla en ambos extremos; si un extremo estirado no queda dentro del
    host (SCI), revierte ese estirón. Luego offset lateral, capa extra y vertical
    como _shaft_extend_then_offset_into_host.
    """
    ext = max(0.0, float(extend_ft))
    if ext <= 1e-9:
        qea, qeb = _shaft_extend_then_offset_into_host(
            q0,
            q1,
            face,
            host,
            extend_ft,
            lateral_cover_ft,
            layer_extra_ft,
            solids,
            vertical_cover_ft=vertical_cover_ft,
        )
        return qea, qeb, True, True
    s0, s1 = _extend_segment_both_ends(q0, q1, ext)
    d = _xyz_sub(q1, q0)
    u = _unit_3d(d)
    if u is None:
        qea, qeb = _shaft_extend_then_offset_into_host(
            q0,
            q1,
            face,
            host,
            extend_ft,
            lateral_cover_ft,
            layer_extra_ft,
            solids,
            vertical_cover_ft=vertical_cover_ft,
        )
        return qea, qeb, True, True
    kept0 = bool(_endpoint_inside_host_concrete(s0, u, solids))
    kept1 = bool(_endpoint_inside_host_concrete(s1, u, solids))
    p0 = s0 if kept0 else q0
    p1 = s1 if kept1 else q1
    try:
        if embed_clip_avisos is not None:
            if (not kept0) and _segment_length_ft(p0, s0) > 1e-5:
                lbl = clip_label or u"Tramo"
                embed_clip_avisos.append(
                    u"{0}: empotramiento recortado en extremo inicio (estirón fuera del host)."
                    .format(lbl)
                )
            if (not kept1) and _segment_length_ft(p1, s1) > 1e-5:
                lbl = clip_label or u"Tramo"
                embed_clip_avisos.append(
                    u"{0}: empotramiento recortado en extremo fin (estirón fuera del host)."
                    .format(lbl)
                )
    except Exception:
        pass
    qe0, qe1 = _offset_segment_into_concrete(
        p0, p1, face, host, lateral_cover_ft, solids=solids
    )
    le = float(layer_extra_ft) if layer_extra_ft is not None else 0.0
    if le > 1e-9:
        qe0, qe1 = _offset_segment_layer_inside_host(qe0, qe1, face, host, le, solids=solids)
    v_cov = (
        vertical_cover_ft
        if vertical_cover_ft is not None
        else lateral_cover_ft
    )
    qe0, qe1 = _apply_vertical_cover_inside_host(qe0, qe1, host, v_cov)
    return qe0, qe1, kept0, kept1


def _shaft_probe_hit_asymmetric_extend_then_offset_into_host(
    q0,
    q1,
    face,
    host,
    extend_ft,
    lateral_cover_ft,
    layer_extra_ft,
    solids,
    vertical_cover_ft=None,
    embed_clip_avisos=None,
    clip_label=None,
    probe_mm=None,
):
    """
    Estirón por tabla solo en extremos donde la sonda ``probe_mm`` (desde cada vértice
    hacia el exterior del tramo) queda dentro del host (SCI). Misma lógica que el debug
    de DetailLine. Sin sólidos o sin dirección válida: delega en
    ``_shaft_adaptive_extend_then_offset_into_host``.
    """
    ext = max(0.0, float(extend_ft))
    if ext <= 1e-9 or (not solids):
        return _shaft_adaptive_extend_then_offset_into_host(
            q0,
            q1,
            face,
            host,
            extend_ft,
            lateral_cover_ft,
            layer_extra_ft,
            solids,
            vertical_cover_ft=vertical_cover_ft,
            embed_clip_avisos=embed_clip_avisos,
            clip_label=clip_label,
        )
    d = _xyz_sub(q1, q0)
    u = _unit_3d(d)
    if u is None:
        return _shaft_adaptive_extend_then_offset_into_host(
            q0,
            q1,
            face,
            host,
            extend_ft,
            lateral_cover_ft,
            layer_extra_ft,
            solids,
            vertical_cover_ft=vertical_cover_ft,
            embed_clip_avisos=embed_clip_avisos,
            clip_label=clip_label,
        )
    try:
        pm = float(probe_mm) if probe_mm is not None else float(SHAFT_EMBED_HOST_PROBE_MM)
    except Exception:
        pm = float(SHAFT_EMBED_HOST_PROBE_MM)
    pm = max(0.0, pm)

    hit0 = False
    hit1 = False
    if pm > 1e-9:
        probe_ft = _mm_to_ft(pm)
        s_tip0 = _xyz_add(
            XYZ(float(q0.X), float(q0.Y), float(q0.Z)), _xyz_scale(u, -probe_ft)
        )
        s_tip1 = _xyz_add(
            XYZ(float(q1.X), float(q1.Y), float(q1.Z)), _xyz_scale(u, probe_ft)
        )
        hit0 = bool(_endpoint_inside_host_concrete(s_tip0, u, solids))
        hit1 = bool(_endpoint_inside_host_concrete(s_tip1, u, solids))

    p0 = XYZ(float(q0.X), float(q0.Y), float(q0.Z))
    p1 = XYZ(float(q1.X), float(q1.Y), float(q1.Z))
    if hit0:
        p0 = _xyz_add(p0, _xyz_scale(u, -ext))
    if hit1:
        p1 = _xyz_add(p1, _xyz_scale(u, ext))

    qe0, qe1 = _offset_segment_into_concrete(
        p0, p1, face, host, lateral_cover_ft, solids=solids
    )
    le = float(layer_extra_ft) if layer_extra_ft is not None else 0.0
    if le > 1e-9:
        qe0, qe1 = _offset_segment_layer_inside_host(
            qe0, qe1, face, host, le, solids=solids
        )
    v_cov = (
        vertical_cover_ft
        if vertical_cover_ft is not None
        else lateral_cover_ft
    )
    qe0, qe1 = _apply_vertical_cover_inside_host(qe0, qe1, host, v_cov)
    return qe0, qe1, hit0, hit1


def _shaft_extend_then_offset_into_host(
    q0,
    q1,
    face,
    host,
    extend_ft,
    lateral_cover_ft,
    layer_extra_ft,
    solids,
    vertical_cover_ft=None,
):
    """
    Orden para armadura: estirar por diámetro sobre el tramo en la cara, luego
    desplazar al interior del host (recubrimiento lateral + capa extra) y ajuste Z.

    `lateral_cover_ft`: cara del vano → eje de barra en planta (cover + Ø/2), sin sumar
    longitud de gancho (esa reserva solo debe influir en Z / espesor, no en el offset XY).

    `vertical_cover_ft`: si se informa, se usa en el clamp vertical del host; si no, el mismo
    valor que el lateral.
    """
    qe0, qe1 = _extend_segment_both_ends(q0, q1, extend_ft)
    qe0, qe1 = _offset_segment_into_concrete(
        qe0, qe1, face, host, lateral_cover_ft, solids=solids
    )
    le = float(layer_extra_ft) if layer_extra_ft is not None else 0.0
    if le > 1e-9:
        qe0, qe1 = _offset_segment_layer_inside_host(qe0, qe1, face, host, le, solids=solids)
    v_cov = (
        vertical_cover_ft
        if vertical_cover_ft is not None
        else lateral_cover_ft
    )
    qe0, qe1 = _apply_vertical_cover_inside_host(qe0, qe1, host, v_cov)
    return qe0, qe1


def _segment_length_ft(q0, q1):
    if q0 is None or q1 is None:
        return 0.0
    d = _xyz_sub(q1, q0)
    return math.sqrt(d.X * d.X + d.Y * d.Y + d.Z * d.Z)


def _stretch_kept_at_ln_api_end(
    ln_line,
    api_end_idx,
    q_corner0,
    q_corner1,
    stretch_kept0,
    stretch_kept1,
    corner_match_tol_ft=None,
):
    """
    True si este extremo de la curva de barra se asocia al canto donde se **retuvo** el
    estirón de empotramiento (sonda SCI / prolongación dentro del host).

    Regla de producto: el **gancho** va en el extremo que **no** retiene ese estirón (canto
    libre, sin colisión SCI en la prolongación: ``stretch_kept`` False en ese vértice).
    Cuando esta función devuelve True para un extremo API, el llamador **quita** el gancho
    ahí para dejarlo solo en el extremo sin retención.

    Primero se intenta coincidencia por tolerancia con ``q_corner0`` o ``q_corner1`` (mismo
    vértice que el tramo global antes del troceo); si ninguno coincide, se usa el canto más
    cercano (evita asignar el flag del canto equivocado en uniones de traslapo).
    """
    if ln_line is None or q_corner0 is None or q_corner1 is None:
        return True
    try:
        tol = float(corner_match_tol_ft) if corner_match_tol_ft is not None else _mm_to_ft(2.0)
    except Exception:
        tol = _mm_to_ft(2.0)
    try:
        tol = max(tol, 1e-6)
    except Exception:
        tol = _mm_to_ft(2.0)
    try:
        p = ln_line.GetEndPoint(int(api_end_idx))
    except Exception:
        return True
    try:
        d0 = float(_segment_length_ft(p, q_corner0))
        d1 = float(_segment_length_ft(p, q_corner1))
    except Exception:
        return bool(stretch_kept0)
    if d0 <= tol:
        return bool(stretch_kept0)
    if d1 <= tol:
        return bool(stretch_kept1)
    if d0 <= d1 + 1e-9:
        return bool(stretch_kept0)
    return bool(stretch_kept1)


def _resize_segment_keep_mid_dir(q0, q1, new_len_ft):
    """Mantiene punto medio y dirección; ajusta longitud a new_len_ft."""
    if q0 is None or q1 is None:
        return q0, q1
    new_len = float(new_len_ft)
    if new_len <= 1e-9:
        return q0, q1
    mid = XYZ(
        0.5 * (float(q0.X) + float(q1.X)),
        0.5 * (float(q0.Y) + float(q1.Y)),
        0.5 * (float(q0.Z) + float(q1.Z)),
    )
    d = _xyz_sub(q1, q0)
    u = _unit_3d(d)
    if u is None:
        return q0, q1
    half = 0.5 * new_len
    return _xyz_sub(mid, _xyz_scale(u, half)), _xyz_add(mid, _xyz_scale(u, half))


def _segment_rebuild_at_mid_dir_len(mid_xyz, u, len_ft):
    """Reconstruye tramo con punto medio mid_xyz, dirección unitaria u y longitud len_ft."""
    if u is None or mid_xyz is None:
        return None, None
    ln = max(0.0, float(len_ft))
    if ln <= 1e-9:
        return None, None
    half = 0.5 * ln
    return _xyz_sub(mid_xyz, _xyz_scale(u, half)), _xyz_add(mid_xyz, _xyz_scale(u, half))


def _pending_segments_index(row):
    """
    Índice de segmentos para las variantes conocidas de `pending`:
      - Rebars / DetailCurves alineados: (face_1based, rf, face, fk, segments) -> 4
      - Formato legacy: (..., segments) con 4 ítems -> 3
      - Tupla corta (face_1based, fk, segments) -> 2
    """
    if row is None:
        return None
    try:
        n = len(row)
    except Exception:
        return None
    if n >= 5:
        return 4
    if n == 4:
        return 3
    if n >= 3:
        return 2
    return None


def _pending_get_segments(row):
    idx = _pending_segments_index(row)
    if idx is None:
        return None
    try:
        return row[idx]
    except Exception:
        return None


def _pending_set_segment(pending, pi, si, q0, q1):
    try:
        row = pending[pi]
    except Exception:
        return False
    idx = _pending_segments_index(row)
    if idx is None:
        return False
    segs = _pending_get_segments(row)
    if not segs or si < 0 or si >= len(segs):
        return False
    segs[si] = (q0, q1)
    return True


def _dot_xy_raw(a, b):
    if a is None or b is None:
        return 0.0
    return float(a.X) * float(b.X) + float(a.Y) * float(b.Y)


def _pending_refs_hv_rect4(pending):
    """
    4 caras, un tramo por cara: devuelve (refs_h, refs_v) con refs_* = list[(pi, si)].
    None si no aplica.
    """
    if pending is None or len(pending) != 4:
        return None
    refs_h = []
    refs_v = []
    for pi, _row in enumerate(pending):
        segs = _pending_get_segments(_row)
        if not segs or len(segs) != 1:
            return None
        q0, q1 = segs[0]
        d = _xyz_sub(q1, q0)
        ln = _segment_length_ft(q0, q1)
        if ln <= 1e-9:
            return None
        if abs(d.X) >= abs(d.Y):
            refs_h.append((pi, 0))
        else:
            refs_v.append((pi, 0))
    if len(refs_h) != 2 or len(refs_v) != 2:
        return None
    return refs_h, refs_v


def _align_pending_hv_midpoints_rect4(pending, avisos):
    """
    Tras igualar longitudes: alinea el punto medio en X de los dos tramos ~horizontales
    y en Y de los dos ~verticales (promedio del par). Así los empotramientos cotados
    desde las esquinas del hueco coinciden en rectángulos simétricos.
    """
    cl = _pending_refs_hv_rect4(pending)
    if cl is None:
        return
    refs_h, refs_v = cl

    mids_h = []
    for pi, si in refs_h:
        segs = _pending_get_segments(pending[pi])
        if not segs:
            continue
        q0, q1 = segs[si]
        mids_h.append(
            XYZ(
                0.5 * (float(q0.X) + float(q1.X)),
                0.5 * (float(q0.Y) + float(q1.Y)),
                0.5 * (float(q0.Z) + float(q1.Z)),
            )
        )
    if len(mids_h) != 2:
        return
    mx = 0.5 * (mids_h[0].X + mids_h[1].X)
    for idx, (pi, si) in enumerate(refs_h):
        segs = _pending_get_segments(pending[pi])
        if not segs:
            continue
        q0, q1 = segs[si]
        ln = _segment_length_ft(q0, q1)
        u = _unit_3d(_xyz_sub(q1, q0))
        if u is None:
            continue
        mid = mids_h[idx]
        new_mid = XYZ(float(mx), float(mid.Y), float(mid.Z))
        nq0, nq1 = _segment_rebuild_at_mid_dir_len(new_mid, u, ln)
        if nq0 is not None and nq1 is not None:
            _pending_set_segment(pending, pi, si, nq0, nq1)

    mids_v = []
    for pi, si in refs_v:
        segs = _pending_get_segments(pending[pi])
        if not segs:
            continue
        q0, q1 = segs[si]
        mids_v.append(
            XYZ(
                0.5 * (float(q0.X) + float(q1.X)),
                0.5 * (float(q0.Y) + float(q1.Y)),
                0.5 * (float(q0.Z) + float(q1.Z)),
            )
        )
    if len(mids_v) != 2:
        return
    my = 0.5 * (mids_v[0].Y + mids_v[1].Y)
    for idx, (pi, si) in enumerate(refs_v):
        segs = _pending_get_segments(pending[pi])
        if not segs:
            continue
        q0, q1 = segs[si]
        ln = _segment_length_ft(q0, q1)
        u = _unit_3d(_xyz_sub(q1, q0))
        if u is None:
            continue
        mid = mids_v[idx]
        new_mid = XYZ(float(mid.X), float(my), float(mid.Z))
        nq0, nq1 = _segment_rebuild_at_mid_dir_len(new_mid, u, ln)
        if nq0 is not None and nq1 is not None:
            _pending_set_segment(pending, pi, si, nq0, nq1)

    try:
        avisos.append(
            u"Alineación pares: mismo centro X entre horizontales, mismo centro Y entre verticales."
        )
    except Exception:
        pass


def _symmetrize_pending_shaft_segments_if_rect4(pending, avisos):
    """
    Hueco shaft rectangular alineado a X/Y con 4 caras y un tramo por cara:
    iguala longitud de los dos tramos con dirección dominante en X y la de los dos en Y
    (promedio), para evitar diferencias de pocos mm por geometría/offsets por cara.
    """
    # Solo si los dos tramos paralelos ya son casi iguales (ruido de coma flotante /
    # offsets). Si difieren mucho (trapecio u otra geometría), no forzar.
    _symm_max_delta_mm = 25.0
    _symm_max_delta_ft = float(_symm_max_delta_mm) / 304.8

    if pending is None or len(pending) != 4:
        return
    lens_h = []
    lens_v = []
    refs_h = []  # (pi, si)
    refs_v = []
    for pi, _row in enumerate(pending):
        segs = _pending_get_segments(_row)
        if not segs or len(segs) != 1:
            return
        q0, q1 = segs[0]
        d = _xyz_sub(q1, q0)
        ln = _segment_length_ft(q0, q1)
        if ln <= 1e-9:
            return
        if abs(d.X) >= abs(d.Y):
            lens_h.append(ln)
            refs_h.append((pi, 0))
        else:
            lens_v.append(ln)
            refs_v.append((pi, 0))
    if len(lens_h) != 2 or len(lens_v) != 2:
        return
    if abs(lens_h[0] - lens_h[1]) > _symm_max_delta_ft:
        return
    if abs(lens_v[0] - lens_v[1]) > _symm_max_delta_ft:
        return
    target_h = 0.5 * (lens_h[0] + lens_h[1])
    target_v = 0.5 * (lens_v[0] + lens_v[1])
    for pi, si in refs_h:
        segs = _pending_get_segments(pending[pi])
        if not segs:
            continue
        q0, q1 = segs[si]
        nq0, nq1 = _resize_segment_keep_mid_dir(q0, q1, target_h)
        _pending_set_segment(pending, pi, si, nq0, nq1)
    for pi, si in refs_v:
        segs = _pending_get_segments(pending[pi])
        if not segs:
            continue
        q0, q1 = segs[si]
        nq0, nq1 = _resize_segment_keep_mid_dir(q0, q1, target_v)
        _pending_set_segment(pending, pi, si, nq0, nq1)
    try:
        avisos.append(
            u"Simetría hueco (4 caras): tramos horizontales {0:.1f} mm, verticales {1:.1f} mm (promedio)."
            .format(float(target_h) * 304.8, float(target_v) * 304.8)
        )
    except Exception:
        pass


def _unique_xy_points(points, tol_ft):
    out = []
    seen = set()
    t = max(float(tol_ft), 1e-9)
    for p in points:
        if p is None:
            continue
        try:
            k = (int(round(float(p.X) / t)), int(round(float(p.Y) / t)))
        except Exception:
            continue
        if k in seen:
            continue
        seen.add(k)
        out.append((float(p.X), float(p.Y)))
    return out


def _convex_hull_xy(points_xy):
    """Convex hull 2D (monotonic chain). Devuelve vértices en orden antihorario."""
    pts = sorted(set(points_xy))
    if len(pts) <= 1:
        return pts

    def _cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower = []
    for p in pts:
        while len(lower) >= 2 and _cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)

    upper = []
    for p in reversed(pts):
        while len(upper) >= 2 and _cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)

    return lower[:-1] + upper[:-1]


def _polygon_area_xy(points_xy):
    """Área de polígono simple en XY por fórmula del shoelace."""
    n = len(points_xy)
    if n < 3:
        return 0.0
    s = 0.0
    for i in range(n):
        x1, y1 = points_xy[i]
        x2, y2 = points_xy[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return abs(0.5 * s)


def calcular_area_hueco_planta_desde_caras_m2(host, refs):
    """
    Reconstruye contorno de hueco en planta desde caras seleccionadas:
    - toma extremos de curvas inferiores por cara,
    - arma un contorno en XY (convex hull),
    - devuelve área en m2.
    """
    if host is None or not refs:
        return None, u"No hay host o caras para calcular área."

    pts = []
    for rf in refs:
        try:
            face = host.GetGeometryObjectFromReference(rf)
        except Exception:
            face = None
        if not isinstance(face, PlanarFace) or not _is_vertical_planar_face(face):
            continue
        curves, _err = _bottom_curves_from_outer_loop(face)
        if not curves:
            continue
        for c in curves:
            try:
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                pts.append(XYZ(float(p0.X), float(p0.Y), 0.0))
                pts.append(XYZ(float(p1.X), float(p1.Y), 0.0))
            except Exception:
                continue

    pts_xy = _unique_xy_points(pts, _mm_to_ft(1.0))
    if len(pts_xy) < 3:
        return None, u"No hay suficientes puntos para reconstruir el área del hueco."

    hull = _convex_hull_xy(pts_xy)
    if len(hull) < 3:
        return None, u"No se pudo construir un contorno cerrado en planta."

    area_ft2 = _polygon_area_xy(hull)
    if area_ft2 <= 1e-9:
        return None, u"Área en planta inválida o nula."
    area_m2 = area_ft2 * 0.09290304
    return area_m2, None


def _canonical_dir_xy(d):
    if d is None:
        return None
    if d.X < -1e-9 or (abs(d.X) <= 1e-9 and d.Y < -1e-9):
        return XYZ(-d.X, -d.Y, 0.0)
    return XYZ(d.X, d.Y, 0.0)


def _franja_key(mid_xy, dxy):
    """
    Identifica una franja de refuerzo en planta: misma dirección de barra puede repetirse
    en caras opuestas del hueco (dos ejes distintos); se distingue por la coordenada
    constante perpendicular al eje.
    """
    c = _canonical_dir_xy(dxy)
    if c is None:
        return None
    if abs(c.Y) >= abs(c.X):
        # Eje principal ~Y: franja fija en X
        return (u"y_axis", round(mid_xy.X, 4), round(c.X, 5), round(c.Y, 5))
    # Eje principal ~X: franja fija en Y
    return (u"x_axis", round(mid_xy.Y, 4), round(c.X, 5), round(c.Y, 5))


def _segment_dir_xy(q0, q1):
    if q0 is None or q1 is None:
        return None
    return _unit_xy(_xyz_sub(q1, q0))


def _parallel_family_key_xy(dxy):
    """
    Clave de grupo para tramos paralelos (ignorando el sentido).
    """
    c = _canonical_dir_xy(dxy)
    if c is None:
        return None
    return (round(float(c.X), 3), round(float(c.Y), 3))


def _normalize_pending_segments_direction(pending, avisos, parallel_dot_tol=0.996):
    """
    Agrupa segmentos por dirección paralela y cuenta evaluación (no invierte tramos:
    no se altera el sentido q0->q1 de las curvas).
    """
    _ = parallel_dot_tol  # parámetro conservado por compatibilidad; sin inversión de tramos.
    if not pending:
        return 0, 0, 0

    groups = {}  # family_key -> list[(pi, si, dxy, length_ft, franja)]
    evaluados = 0
    for pi, row in enumerate(pending):
        segs = _pending_get_segments(row)
        if not segs:
            continue
        for si, seg in enumerate(segs):
            try:
                q0, q1 = seg
            except Exception:
                continue
            ln = _segment_length_ft(q0, q1)
            if ln <= 1e-9:
                continue
            dxy = _segment_dir_xy(q0, q1)
            if dxy is None:
                continue
            mid = XYZ(
                0.5 * (float(q0.X) + float(q1.X)),
                0.5 * (float(q0.Y) + float(q1.Y)),
                0.0,
            )
            family_key = _parallel_family_key_xy(dxy)
            if family_key is None:
                continue
            fr_key = _franja_key(mid, dxy)
            groups.setdefault(family_key, []).append((pi, si, dxy, ln, fr_key))
            evaluados += 1

    flipped = 0
    grouped = 0
    for _family_key, entries in groups.items():
        if not entries:
            continue
        grouped += 1

    try:
        if evaluados > 0:
            avisos.append(
                u"Normalización dirección: {0} segmento(s) evaluado(s), {1} grupo(s) paralelo(s); "
                u"no se invierten tramos (política sin reversión de curva)."
                .format(int(evaluados), int(grouped))
            )
    except Exception:
        pass

    return evaluados, grouped, flipped


def _first_bar_type(document):
    for bt in FilteredElementCollector(document).OfClass(RebarBarType):
        return bt
    return None


def resolver_bar_type_por_diametro_mm(document, target_mm):
    """
    Busca RebarBarType con diámetro nominal más cercano a target_mm.
    Returns:
      (bar_type, exact_match:bool, delta_mm:float)
    """
    if document is None:
        return None, False, None
    best = None
    best_delta = None
    target = float(target_mm)
    for bt in FilteredElementCollector(document).OfClass(RebarBarType):
        try:
            d_mm = _ft_to_mm(float(bt.BarNominalDiameter))
        except Exception:
            continue
        delta = abs(d_mm - target)
        if best is None or delta < best_delta:
            best = bt
            best_delta = delta
    if best is None:
        return None, False, None
    return best, (best_delta <= 0.25), best_delta


def _bar_nominal_diameter_ft(bar_type):
    if bar_type is None:
        return 0.0
    try:
        return float(bar_type.BarNominalDiameter)
    except Exception:
        return 0.0


def _bbox_corners_xyz(bb):
    if bb is None:
        return []
    mn, mx = bb.Min, bb.Max
    out = []
    for x in (float(mn.X), float(mx.X)):
        for y in (float(mn.Y), float(mx.Y)):
            for z in (float(mn.Z), float(mx.Z)):
                out.append(XYZ(x, y, z))
    return out


def _collect_solids_from_host_element(elem):
    if elem is None:
        return []
    opts = Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    try:
        opts.DetailLevel = ViewDetailLevel.Fine
    except Exception:
        pass
    try:
        ge = elem.get_Geometry(opts)
    except Exception:
        return []
    if ge is None:
        return []
    out = []
    for obj in ge:
        if obj is None:
            continue
        if isinstance(obj, Solid) and obj.Volume > 1e-12:
            out.append(obj)
        elif isinstance(obj, GeometryInstance):
            try:
                sub = obj.GetInstanceGeometry()
                if sub is not None:
                    for g in sub:
                        if isinstance(g, Solid) and g.Volume > 1e-12:
                            out.append(g)
            except Exception:
                pass
    return out


def _merge_axis_intervals(intervals):
    if not intervals:
        return []
    iv = sorted(intervals, key=lambda x: x[0])
    merged = []
    for a, b in iv:
        if b < a:
            a, b = b, a
        if not merged or a > merged[-1][1] + 1e-9:
            merged.append([a, b])
        else:
            merged[-1][1] = max(merged[-1][1], b)
    return merged


def _solid_line_inside_param_intervals(line, solid):
    out = []
    try:
        scio = SolidCurveIntersectionOptions()
        try:
            scio.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
        except Exception:
            pass
        sci = solid.IntersectWithCurve(line, scio)
    except Exception:
        return out
    if sci is None:
        return out
    try:
        n = int(sci.SegmentCount)
    except Exception:
        return out
    if n < 1:
        return out
    for i in range(n):
        try:
            ext = sci.GetCurveSegmentExtents(i)
            s0 = float(ext.StartParameter)
            s1 = float(ext.EndParameter)
            if s1 < s0:
                s0, s1 = s1, s0
            out.append((s0, s1))
        except Exception:
            continue
    return out


def _best_merged_span_for_midpoint(merged, curve_len, mid_param):
    if not merged:
        return None
    for a, b in merged:
        if a <= mid_param <= b:
            return (float(a), float(b))
    best = None
    best_len = -1.0
    for a, b in merged:
        lo = max(0.0, float(a))
        hi = min(float(curve_len), float(b))
        span = hi - lo
        if span > best_len:
            best_len = span
            best = (float(a), float(b))
    return best if best_len > 1e-9 else None


def _rebar_mid_stub_inside_concrete(qa, qb, solids):
    """
    True si un tramo corto sobre el eje de la barra en el punto medio queda
    dentro del hormigón (SCI). Filtra tramos que solo cumplen bbox (hueco dentro del AABB).
    """
    if not solids:
        return True
    mid_x = 0.5 * (float(qa.X) + float(qb.X))
    mid_y = 0.5 * (float(qa.Y) + float(qb.Y))
    mid_z = 0.5 * (float(qa.Z) + float(qb.Z))
    t = _xyz_sub(qb, qa)
    tlen = math.sqrt(t.X * t.X + t.Y * t.Y + t.Z * t.Z)
    if tlen < 1e-9:
        return False
    stub = max(1e-3, min(0.02, 0.01 * tlen))
    ux, uy, uz = t.X / tlen, t.Y / tlen, t.Z / tlen
    p_a = XYZ(mid_x - ux * stub, mid_y - uy * stub, mid_z - uz * stub)
    p_b = XYZ(mid_x + ux * stub, mid_y + uy * stub, mid_z + uz * stub)
    try:
        ln = Line.CreateBound(p_a, p_b)
        clen = float(ln.Length)
    except Exception:
        return False
    if clen < 1e-9:
        return False
    mid_param = 0.5 * clen
    for solid in solids:
        try:
            if solid.Volume < 1e-12:
                continue
            intervals = _solid_line_inside_param_intervals(ln, solid)
            merged = _merge_axis_intervals(intervals)
            for a, b in merged:
                lo = max(0.0, float(a))
                hi = min(clen, float(b))
                if lo <= mid_param <= hi:
                    return True
        except Exception:
            continue
    return False


def _rebar_mid_stub_inside_concrete_score(qa, qb, solids):
    """
    Similar a _rebar_mid_stub_inside_concrete, pero en vez de booleano devuelve
    un "score" cuantitativo (longitud del intervalo SCI que contiene el punto
    medio) para desempatar el signo cuando ambos ± dan SCI.
    """
    if not solids:
        return 0.0
    mid_x = 0.5 * (float(qa.X) + float(qb.X))
    mid_y = 0.5 * (float(qa.Y) + float(qb.Y))
    mid_z = 0.5 * (float(qa.Z) + float(qb.Z))
    t = _xyz_sub(qb, qa)
    tlen = math.sqrt(t.X * t.X + t.Y * t.Y + t.Z * t.Z)
    if tlen < 1e-9:
        return 0.0
    stub = max(1e-3, min(0.02, 0.01 * tlen))
    ux, uy, uz = t.X / tlen, t.Y / tlen, t.Z / tlen
    p_a = XYZ(mid_x - ux * stub, mid_y - uy * stub, mid_z - uz * stub)
    p_b = XYZ(mid_x + ux * stub, mid_y + uy * stub, mid_z + uz * stub)
    try:
        ln = Line.CreateBound(p_a, p_b)
        clen = float(ln.Length)
    except Exception:
        return 0.0
    if clen < 1e-9:
        return 0.0
    mid_param = 0.5 * clen
    best = 0.0
    for solid in solids:
        try:
            if solid.Volume < 1e-12:
                continue
            intervals = _solid_line_inside_param_intervals(ln, solid)
            merged = _merge_axis_intervals(intervals)
            for a, b in merged:
                lo = max(0.0, float(a))
                hi = min(clen, float(b))
                if lo <= mid_param <= hi:
                    score = max(0.0, hi - lo)
                    if score > best:
                        best = score
        except Exception:
            continue
    return float(best)


def _probe_offset_inside_concrete(mid_edge, direction, sign, probe_ft, solids):
    """Desplaza mid_edge en ±direction (normalizada) y comprueba hormigón con sonda vertical corta."""
    u = _unit_3d(direction) if direction is not None else None
    if u is None:
        return False
    pt = _xyz_add(mid_edge, _xyz_scale(u, float(sign) * float(probe_ft)))
    qa = _xyz_add(pt, XYZ(0.0, 0.0, -probe_ft))
    qb = _xyz_add(pt, XYZ(0.0, 0.0, probe_ft))
    return _rebar_mid_stub_inside_concrete(qa, qb, solids)


def _probe_offset_inside_concrete_score(mid_edge, direction, sign, probe_ft, solids):
    """Igual que _probe_offset_inside_concrete, pero devuelve un score SCI (>=0)."""
    u = _unit_3d(direction) if direction is not None else None
    if u is None:
        return 0.0
    pt = _xyz_add(mid_edge, _xyz_scale(u, float(sign) * float(probe_ft)))
    qa = _xyz_add(pt, XYZ(0.0, 0.0, -probe_ft))
    qb = _xyz_add(pt, XYZ(0.0, 0.0, probe_ft))
    return _rebar_mid_stub_inside_concrete_score(qa, qb, solids)


def _infer_inward_xy_sign(mid_edge, n_xy, solids, probe_ft):
    """
    Para la cara del hueco, elige el signo de n_xy que apunta al hormigón
    (solo un lado suele dar SCI válido frente al vacío).
    """
    if not solids or n_xy is None:
        return None
    score_neg = _probe_offset_inside_concrete_score(mid_edge, n_xy, -1.0, probe_ft, solids)
    score_pos = _probe_offset_inside_concrete_score(mid_edge, n_xy, 1.0, probe_ft, solids)
    ins_neg = score_neg > 1e-9
    ins_pos = score_pos > 1e-9
    if ins_pos and not ins_neg:
        return 1.0
    if ins_neg and not ins_pos:
        return -1.0
    # Si ambos son válidos, elegimos el que penetra "más" (mayor score).
    if ins_pos and ins_neg:
        if score_pos > score_neg + 1e-9:
            return 1.0
        if score_neg > score_pos + 1e-9:
            return -1.0
    return None


def _axis_solid_span_params(p0, p1, lateral_offset, solids):
    if not solids:
        return None
    try:
        off = lateral_offset if lateral_offset is not None else XYZ(0, 0, 0)
        a = _xyz_add(p0, off)
        b = _xyz_add(p1, off)
        ln_seg = Line.CreateBound(a, b)
        clen = float(ln_seg.Length)
    except Exception:
        return None
    if clen < 1e-9:
        return None
    acc = []
    for solid in solids:
        acc.extend(_solid_line_inside_param_intervals(ln_seg, solid))
    merged = _merge_axis_intervals(acc)
    mid = 0.5 * clen
    span = _best_merged_span_for_midpoint(merged, clen, mid)
    if span is None:
        return None
    s0, s1 = span
    if s1 <= s0 + 1e-9:
        return None
    return (max(0.0, s0), min(clen, s1))


def _axis_bbox_t_span_clamped(p0, p1, host, margin):
    raw = _xyz_sub(p1, p0)
    ln = math.sqrt(raw.X * raw.X + raw.Y * raw.Y + raw.Z * raw.Z)
    if ln < 1e-9:
        return None
    ax = _unit_3d(raw)
    if ax is None:
        return None
    m = float(margin)
    bb = host.get_BoundingBox(None) if host else None
    if bb is None:
        return (m, ln - m)
    hc = _bbox_corners_xyz(bb)
    if not hc:
        return (m, ln - m)
    ts = []
    for c in hc:
        v = _xyz_sub(c, p0)
        ts.append(float(v.X * ax.X + v.Y * ax.Y + v.Z * ax.Z))
    t_min = min(ts)
    t_max = max(ts)
    t_lo = t_min + m
    t_hi = t_max - m
    t_u0 = max(t_lo, 0.0, m)
    t_u1 = min(t_hi, ln, ln - m)
    if t_u1 <= t_u0 + 1e-6:
        return None
    return (t_u0, t_u1)


def _face_info_entry_from_face(face, solids):
    """
    Entrada mínima tipo ``face_infos`` para ``_pick_end_for_anchorage`` / cotas,
    coherente con la construcción en ``crear_enfierrado_shaft_hashtag``.
    """
    if face is None:
        return {}
    face_ref = None
    try:
        fr = getattr(face, "Reference", None)
        if fr is not None:
            face_ref = fr
    except Exception:
        face_ref = None
    nxy = None
    try:
        nxy = _face_normal_xy(face)
    except Exception:
        nxy = None
    inward_dir = None
    if nxy is not None:
        try:
            probe_ft = _mm_to_ft(50.0)
            sgn_in = _infer_inward_xy_sign(
                _face_origin(face), nxy, solids or [], probe_ft
            )
        except Exception:
            sgn_in = None
        if sgn_in is None:
            sgn_in = -1.0
        try:
            inward_dir = _xyz_scale(nxy, float(sgn_in))
        except Exception:
            inward_dir = None
    n_bar_create = _face_normal_3d(face)
    return {
        "ref": face_ref,
        "face": face,
        "nxy": nxy,
        "inward_dir": inward_dir,
        "n_bar_create": n_bar_create,
    }


def _horizontal_offset_segments_for_face(
    face,
    cover_ft,
    used_strip_keys,
    host,
    bar_type,
    solids,
    stretch_ft_for_sci=None,
    layer_extra_ft_for_sci=0.0,
    adaptive_embed_end_clip=False,
    embed_clip_avisos=None,
):
    """
    Tramo en la arista inferior de la cara (recorte longitudinal por cover+Ø/2 en los extremos).
    Con ``adaptive_embed_end_clip`` y estirón SCI: (1) ensayo SCI como antes; si **ambos**
    extremos retienen estirón dentro del host, no se aplica ese margen en arista (evita acortar
    el empotramiento respecto al valor tabulado). (2) si hay ambigüedad, ``_pick_end_for_anchorage``
    sobre geometría ya estirada/desplazada; (3) si aún falla, el vértice más cercano al plano
    de la cara seleccionada se trata como lado empotrado (el muro fuera del sólido de la losa
    suele dar SCI ambiguo).
    Los (q0,q1) devueltos siguen en el plano de la cara; el estiramiento por Ø y el
    offset al hormigón se aplican al crear la barra.

    Si stretch_ft_for_sci no es None, valida SCI con la misma secuencia que al crear
    (estirar → offset lateral → vertical). layer_extra_ft_for_sci comprueba la capa más
    profunda cuando hay varias capas.

    Returns:
        ((fk, list of (q0, q1)), None) o (None, err).
    """
    if not _is_vertical_planar_face(face):
        return (None, None), u"No es una cara planar vertical (normal Z baja)."

    bottom_curves, err_bottom = _bottom_curves_from_outer_loop(face)
    if err_bottom:
        return (None, None), err_bottom

    # `cover_ft`: recubrimiento lateral desde la cara seleccionada (cara → eje en planta).
    # Extremos del tramo: SHAFT_END_COVER_MM_DEFAULT (p. ej. 35 mm) + Ø/2, no obliga a
    # igualar el nominal lateral (p. ej. 25 mm).
    bar_diam_ft = float(_bar_nominal_diameter_ft(bar_type) or 0.0)
    half_d = 0.5 * max(bar_diam_ft, 1e-6)
    cover_center_ft = float(cover_ft) + half_d
    try:
        end_cov_mm = float(SHAFT_END_COVER_MM_DEFAULT)
    except Exception:
        end_cov_mm = float(COVER_MM_DEFAULT)
    margin = float(_mm_to_ft(max(0.0, end_cov_mm))) + half_d
    lex_sci = float(layer_extra_ft_for_sci or 0.0)
    segments = []
    errs = []

    for c in bottom_curves:
        margin_start = float(margin)
        margin_end = float(margin)
        if (
            bool(adaptive_embed_end_clip)
            and stretch_ft_for_sci is not None
            and float(stretch_ft_for_sci) > 1e-12
        ):
            try:
                p0r = c.GetEndPoint(0)
                p1r = c.GetEndPoint(1)
                qa_f, qb_f, sk0, sk1 = _shaft_probe_hit_asymmetric_extend_then_offset_into_host(
                    p0r,
                    p1r,
                    face,
                    host,
                    float(stretch_ft_for_sci),
                    cover_center_ft,
                    max(0.0, lex_sci),
                    solids,
                )
                resolved = False
                if sk0 and (not sk1):
                    margin_start, margin_end = 0.0, float(margin)
                    resolved = True
                elif sk1 and (not sk0):
                    margin_start, margin_end = float(margin), 0.0
                    resolved = True
                elif sk0 and sk1:
                    # Ambos cantos retienen estirón dentro del host: no aplicar recorte
                    # longitudinal por SHAFT_END_COVER+Ø/2 en la arista inferior; ese
                    # recorte sumaba al estirón tabulado y acortaba el empotramiento
                    # efectivo (p. ej. ~860 mm esperados → ~820 mm medidos).
                    margin_start, margin_end = 0.0, 0.0
                    resolved = True
                if not resolved:
                    fi_one = _face_info_entry_from_face(face, solids)
                    try:
                        dxy_p = _unit_xy(_xyz_sub(qb_f, qa_f))
                    except Exception:
                        dxy_p = None
                    exp_mm = _ft_to_mm(float(stretch_ft_for_sci))
                    end_idx = None
                    if dxy_p is not None and fi_one.get("nxy") is not None:
                        _ei, fp = _pick_end_for_anchorage(
                            qa_f,
                            qb_f,
                            dxy_p,
                            [fi_one],
                            expected_mm=exp_mm,
                            tol_mm=80.0,
                        )
                        if fp is not None:
                            end_idx = int(_ei)
                    if end_idx is not None:
                        if end_idx == 0:
                            margin_start, margin_end = 0.0, float(margin)
                        else:
                            margin_start, margin_end = float(margin), 0.0
                        resolved = True
                    if not resolved:
                        d0 = _point_plane_distance_ft_xy(p0r, face)
                        d1 = _point_plane_distance_ft_xy(p1r, face)
                        eps_ft = _mm_to_ft(25.0)
                        if d0 is not None and d1 is not None:
                            if float(d0) + eps_ft < float(d1):
                                margin_start, margin_end = 0.0, float(margin)
                                resolved = True
                            elif float(d1) + eps_ft < float(d0):
                                margin_start, margin_end = float(margin), 0.0
                                resolved = True
                    if not resolved:
                        o = _face_origin(face)
                        inward_dir = fi_one.get("inward_dir")
                        if o is not None and inward_dir is not None:
                            try:
                                v0 = XYZ(
                                    float(p0r.X) - float(o.X),
                                    float(p0r.Y) - float(o.Y),
                                    0.0,
                                )
                                v1 = XYZ(
                                    float(p1r.X) - float(o.X),
                                    float(p1r.Y) - float(o.Y),
                                    0.0,
                                )
                                di0 = float(_dot_xy(v0, inward_dir))
                                di1 = float(_dot_xy(v1, inward_dir))
                                delta_in = _mm_to_ft(20.0)
                                # Menor avance hacia el interior del hueco (en planta) suele ser la
                                # esquina de encuentro con muro / cara de anclaje.
                                if di0 + delta_in < di1:
                                    margin_start, margin_end = 0.0, float(margin)
                                elif di1 + delta_in < di0:
                                    margin_start, margin_end = float(margin), 0.0
                            except Exception:
                                pass
            except Exception:
                margin_start = float(margin)
                margin_end = float(margin)

        seg, err_seg = _line_segment_from_curve_with_margins(
            c, margin_start, margin_end
        )
        if err_seg:
            errs.append(err_seg)
            continue
        q0, q1 = seg
        min_len_seg = max(float(margin_start) + float(margin_end), 1e-4)
        try:
            ln_face = Line.CreateBound(q0, q1)
            if float(ln_face.Length) < min_len_seg:
                errs.append(u"Tramo inferior muy corto.")
                continue
        except Exception:
            errs.append(u"Tramo inferior inválido.")
            continue
        if stretch_ft_for_sci is not None:
            if bool(adaptive_embed_end_clip) and float(stretch_ft_for_sci) > 1e-12:
                qa, qb, _sk0, _sk1 = _shaft_probe_hit_asymmetric_extend_then_offset_into_host(
                    q0,
                    q1,
                    face,
                    host,
                    float(stretch_ft_for_sci),
                    cover_center_ft,
                    max(0.0, lex_sci),
                    solids,
                    embed_clip_avisos=embed_clip_avisos,
                )
            else:
                qa, qb = _shaft_extend_then_offset_into_host(
                    q0,
                    q1,
                    face,
                    host,
                    float(stretch_ft_for_sci),
                    cover_center_ft,
                    max(0.0, lex_sci),
                    solids,
                )
        else:
            qa, qb = _offset_segment_into_concrete(q0, q1, face, host, cover_center_ft, solids=solids)
            qa, qb = _apply_vertical_cover_inside_host(qa, qb, host, cover_center_ft)
            if lex_sci > 1e-9:
                qa, qb = _offset_segment_layer_inside_host(
                    qa, qb, face, host, lex_sci, solids=solids
                )
                qa, qb = _apply_vertical_cover_inside_host(qa, qb, host, cover_center_ft)
        if solids and not _rebar_mid_stub_inside_concrete(qa, qb, solids):
            errs.append(u"Tramo inferior fuera del hormigón.")
            continue
        segments.append((q0, q1))

    if not segments:
        if errs:
            return (None, None), errs[0]
        return (None, None), u"No hay tramo inferior válido para la cara."

    return (None, segments), None


def _hook_orientation_pairs_shaft():
    """Pares (inicio, fin) para CreateFromCurves (patrón armadura_vigas_capas)."""
    return (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )


def _inverse_face_normal_xy(nxy):
    """
    Dirección en planta = -n_xy (inverso de la proyección XY de FaceNormal de Revit).
    Es la «normal invertida» de la cara para orientar ganchos hacia el interior del host,
    sin depender de la sonda SCI (que a veces alinea inward_dir con +n_xy).
    """
    if nxy is None:
        return None
    try:
        return _xyz_scale(nxy, -1.0)
    except Exception:
        return None


def _face_y_vector_3d(face):
    """
    Vector 3D = eje **Y** del sistema interno de la ``PlanarFace`` (``YVector``), **sin normalizar**.

    No se usa como ``n_bar_create`` (creación de barra usa :func:`_face_normal_3d`); útil como helper.
    """
    if face is None or not isinstance(face, PlanarFace):
        return None
    try:
        y = getattr(face, "YVector", None)
    except Exception:
        y = None
    if y is None:
        return None
    try:
        return XYZ(float(y.X), float(y.Y), float(y.Z))
    except Exception:
        return None


def _face_normal_3d(face):
    """
    Vector 3D = **normal inversa de cara** (``-FaceNormal``) **sin normalizar**.
    Referencia ``n_bar_create`` para ``Rebar.CreateFromCurves`` y ganchos (proyectada ⟂ tangente).
    """
    if face is None:
        return None
    try:
        n = face.FaceNormal
    except Exception:
        n = None
    if n is None:
        return None
    try:
        return XYZ(-float(n.X), -float(n.Y), -float(n.Z))
    except Exception:
        return None


def _negative_unit_face_normal_3d(face):
    """Legado: ``+FaceNormal`` sin normalizar (opuesto a :func:`_face_normal_3d`)."""
    u = _face_normal_3d(face)
    if u is None:
        return None
    try:
        return XYZ(-float(u.X), -float(u.Y), -float(u.Z))
    except Exception:
        return None


def _rebar_plane_normal_from_face_normal_and_tangent(face_normal, tangent):
    """
    Vector (no necesariamente unitario) ⟂ ``tangent`` en el plano de la barra: proyección del
    vector de referencia de cara (p. ej. ``-FaceNormal``) sobre el plano perpendicular a ``tangent``:

    ``nb - (nb·ut)/(ut·ut) * ut``

    Sin normalizar el resultado. ``Rebar.CreateFromCurves`` y ganchos comparten esta referencia.
    """
    if face_normal is None or tangent is None:
        return None
    try:
        nb = face_normal
        ut = tangent
        utu = float(_dot_3d_xyz(ut, ut))
        if utu < 1e-18:
            return None
        nbu = float(_dot_3d_xyz(nb, ut))
        s = nbu / utu
        vx = float(nb.X) - s * float(ut.X)
        vy = float(nb.Y) - s * float(ut.Y)
        vz = float(nb.Z) - s * float(ut.Z)
        v = XYZ(vx, vy, vz)
        if (vx * vx + vy * vy + vz * vz) > 1e-18:
            return v
        try:
            cp = ut.CrossProduct(XYZ.BasisZ)
        except Exception:
            cp = None
        if cp is not None:
            c2 = float(cp.X) ** 2 + float(cp.Y) ** 2 + float(cp.Z) ** 2
            if c2 > 1e-18:
                return cp
        try:
            cp = ut.CrossProduct(XYZ.BasisX)
        except Exception:
            cp = None
        if cp is not None:
            c2 = float(cp.X) ** 2 + float(cp.Y) ** 2 + float(cp.Z) ** 2
            if c2 > 1e-18:
                return cp
        return None
    except Exception:
        return None


def _dot_3d_xyz(a, b):
    try:
        return (
            float(a.X) * float(b.X)
            + float(a.Y) * float(b.Y)
            + float(a.Z) * float(b.Z)
        )
    except Exception:
        return 0.0


def _compute_hook_orientations_pair_from_inverse_normal_xyz(
    curve, inward_xy_unit, nvec, hook_end0, hook_end1
):
    """
    Calcula **ambos** ``RebarHookOrientation`` usando **el mismo** ``nvec`` (normal de creación
    de la barra: ``-FaceNormal`` proyectada ⟂ tangente, sin exigir unitarios) en los dos extremos: laterales
    ``nvec×t`` y ``nvec×(-t)``. La referencia «hacia el interior» en planta es ``inward_xy_unit`` (-n_xy).
    """
    if curve is None or inward_xy_unit is None or nvec is None:
        return None
    d = _unit_xy(inward_xy_unit)
    if d is None:
        return None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        t = XYZ(
            float(p1.X) - float(p0.X),
            float(p1.Y) - float(p0.Y),
            float(p1.Z) - float(p0.Z),
        )
    except Exception:
        return None
    o0 = RebarHookOrientation.Left
    o1 = RebarHookOrientation.Left
    try:
        lat0 = nvec.CrossProduct(t)
        lat0_xy = _unit_xy(XYZ(lat0.X, lat0.Y, 0.0))
        tneg = XYZ(-float(t.X), -float(t.Y), -float(t.Z))
        lat1 = nvec.CrossProduct(tneg)
        lat1_xy = _unit_xy(XYZ(lat1.X, lat1.Y, 0.0))
    except Exception:
        return None
    # Criterio invertido respecto al signo del dot: alinea la pierna del gancho con el interior en planta.
    if bool(hook_end0) and lat0_xy is not None:
        o0 = (
            RebarHookOrientation.Right
            if float(_dot_xy(lat0_xy, d)) < 0.0
            else RebarHookOrientation.Left
        )
    if bool(hook_end1) and lat1_xy is not None:
        o1 = (
            RebarHookOrientation.Right
            if float(_dot_xy(lat1_xy, d)) < 0.0
            else RebarHookOrientation.Left
        )
    return o0, o1


def _order_z_norms_for_face_inward(curve, inward_xy):
    """
    Orden de prueba de ±BasisZ para CreateFromCurves: alinear la regla de la mano derecha
    (normal × tangente) con la dirección hacia interior del hormigón en planta (equivalente
    a usar la normal de cara invertida respecto al exterior de la losa).
    """
    z_pos = XYZ.BasisZ
    try:
        z_neg = XYZ.BasisZ.Negate()
    except Exception:
        z_neg = XYZ(0.0, 0.0, -1.0)
    if curve is None or inward_xy is None:
        return [z_pos, z_neg]
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        t = _unit_xy(_xyz_sub(p1, p0))
        iw = _unit_xy(inward_xy)
    except Exception:
        return [z_pos, z_neg]
    if t is None or iw is None:
        return [z_pos, z_neg]
    try:
        c2 = float(t.X) * float(iw.Y) - float(t.Y) * float(iw.X)
    except Exception:
        return [z_pos, z_neg]
    if c2 >= 0.0:
        return [z_pos, z_neg]
    return [z_neg, z_pos]


def _curve_array_for_create_from_curves(one_line):
    """CreateFromCurves en IronPython: System.Array de Curve (base Line)."""
    ct = clr.GetClrType(Line).BaseType
    arr = System.Array.CreateInstance(ct, 1)
    arr[0] = one_line
    return arr


def _first_rebar_hook_type(document):
    if document is None:
        return None
    try:
        for ht in FilteredElementCollector(document).OfClass(RebarHookType):
            return ht
    except Exception:
        pass
    return None


def _rebar_hook_type_display_name(ht):
    """
    Nombre legible del RebarHookType (pyRevit/IronPython a veces deja .Name vacío).
    Replica el patrón usado en tags/shapes: SYMBOL_NAME_PARAM / ALL_MODEL_TYPE_NAME.
    """
    if ht is None:
        return u""
    try:
        n = unicode(getattr(ht, "Name", None) or u"").strip()
        if n:
            return n
    except Exception:
        pass
    for bip_name in (u"SYMBOL_NAME_PARAM", u"ALL_MODEL_TYPE_NAME"):
        try:
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            p = ht.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = unicode(p.AsString() or u"").strip()
                if s:
                    return s
        except Exception:
            continue
    return u""


def _hook_compare_string(value):
    """
    Convierte el valor pedido por Revit/IronPython/.NET a unicode para comparar
    siempre string con string (mismo tipo lógico). Solo trim y NBSP→espacio;
    sin lower ni tokenización (igualdad exacta al nombre mostrado).
    """
    if value is None:
        return u""
    t = None
    try:
        if isinstance(value, unicode):
            t = value
        else:
            t = unicode(value)
    except Exception:
        t = None
    if t is None:
        try:
            t = System.Convert.ToString(value)
        except Exception:
            t = None
    if t is None:
        return u""
    try:
        t = unicode(t)
    except Exception:
        return u""
    try:
        t = t.replace(u"\u00A0", u" ").strip()
    except Exception:
        t = u""
    return t


def _rebar_hook_type_by_name(document, hook_name):
    """
    Resuelve RebarHookType por igualdad exacta del nombre mostrado (unicode).
    Recorre todos los tipos del documento; compara `hook_name` y cada
    `_rebar_hook_type_display_name(ht)` vía `_hook_compare_string`; sin fuzzy
    ni fallback por ángulo.

    Tras encontrar coincidencia, se devuelve `document.GetElement(id)` para no
    depender del objeto devuelto por la lista materializada (referencia estable)
    y ante nombres duplicados se conserva el primero según orden de ElementId.
    """
    if document is None:
        return None
    target = _hook_compare_string(hook_name)
    if not target:
        return None
    try:
        hook_types = list(FilteredElementCollector(document).OfClass(RebarHookType))
    except Exception:
        return None
    found_id = None
    for ht in hook_types:
        disp = _hook_compare_string(_rebar_hook_type_display_name(ht))
        if disp != target:
            continue
        try:
            eid = ht.Id
        except Exception:
            continue
        if found_id is None:
            found_id = eid
        # Más de un tipo con el mismo nombre visible: se mantiene el primero.
        else:
            break
    if found_id is None:
        return None
    try:
        fresh = document.GetElement(found_id)
        if fresh is not None:
            return fresh
    except Exception:
        pass
    return None


def _rebar_shapes_shaft_fallback_ordered(document):
    """Formas rectas primero (menos conflictos), luego el resto — como laterales en vigas."""
    if document is None:
        return []
    try:
        all_s = list(FilteredElementCollector(document).OfClass(RebarShape))
    except Exception:
        return []
    simple = [s for s in all_s if getattr(s, "SimpleLine", False)]
    rest = [s for s in all_s if s not in simple]
    return simple + rest


def _normals_for_shaft_horizontal_line(ln):
    """
    Planos candidatos: XY (BasisZ) y el perpendicular horizontal al eje de la barra
    (algunos hosts/losas aceptan mejor una u otra).
    """
    out = []
    for v in (XYZ.BasisZ,):
        try:
            out.append(v)
            out.append(v.Negate())
        except Exception:
            out.append(XYZ.BasisZ)
            out.append(XYZ(0.0, 0.0, -1.0))
    try:
        d = _xyz_sub(ln.GetEndPoint(1), ln.GetEndPoint(0))
        u = _unit_3d(d)
        if u is not None:
            try:
                c = u.CrossProduct(XYZ.BasisZ)
            except Exception:
                c = None
            cp = _unit_3d(c) if c is not None else None
            if cp is not None:
                out.append(cp)
                try:
                    out.append(cp.Negate())
                except Exception:
                    out.append(_xyz_scale(cp, -1.0))
    except Exception:
        pass
    return [v for v in out if v is not None]


def _set_rebar_hook_type_id_at_end(rebar, end_idx, type_id, document, max_attempts=2):
    """
    Asigna HookTypeId en un extremo y comprueba lectura; si no coincide, Regenerate y reintenta.
    Así se reduce el caso en que Revit deja el gancho por defecto del RebarBarType
    (p. ej. 'Rebar Hook - 90 - 110 mm') tras CreateFromCurves.
    """
    if rebar is None:
        return False
    e = int(end_idx)
    for attempt in range(int(max_attempts)):
        try:
            rebar.SetHookTypeId(e, type_id)
        except Exception:
            return False
        try:
            if _eid_int(rebar.GetHookTypeId(e)) == _eid_int(type_id):
                return True
        except Exception:
            pass
        if document is not None and attempt + 1 < int(max_attempts):
            try:
                document.Regenerate()
            except Exception:
                pass
    return False


def _append_hook_set_failed_aviso(avisos, end_label, hook_display_hint):
    if avisos is None:
        return
    try:
        avisos.append(
            u"No se pudo fijar el gancho en {0} como '{1}'. Compruebe que el tipo de barra "
            u"admite ese RebarHookType (Tipo de barra → ganchos) o que el nombre coincide exactamente."
            .format(end_label, hook_display_hint or u"(pedido)")
        )
    except Exception:
        pass


def _create_single_rebar_from_line_no_hooks(document, host, bar_type, ln):
    """
    Creador sin ganchos:
    - Solo CreateFromCurves (sin CreateFromCurvesAndShape).
    - Prueba dos firmas sin gancho: (None, None) y (InvalidElementId, InvalidElementId),
      porque según versión/API activa de Revit puede resolverse una u otra.
    """
    if document is None or host is None or bar_type is None or ln is None:
        return None
    hid = ElementId.InvalidElementId
    ct = clr.GetClrType(Line).BaseType
    norms = [XYZ.BasisZ]
    try:
        norms.append(XYZ.BasisZ.Negate())
    except Exception:
        norms.append(XYZ(0.0, 0.0, -1.0))
    orient_pairs = _hook_orientation_pairs_shaft()
    hook_pairs = ((None, None), (hid, hid))

    def _try_create_from_curves(curve):
        arr = System.Array.CreateInstance(ct, 1)
        arr[0] = curve
        for h0, h1 in hook_pairs:
            for use_existing in (True, False):
                for create_new in (False, True):
                    for nvec in norms:
                        if nvec is None:
                            continue
                        for so, eo in orient_pairs:
                            try:
                                r = Rebar.CreateFromCurves(
                                    document,
                                    RebarStyle.Standard,
                                    bar_type,
                                    h0,
                                    h1,
                                    host,
                                    nvec,
                                    arr,
                                    so,
                                    eo,
                                    use_existing,
                                    create_new,
                                )
                                if r:
                                    # Seguridad extra: limpiar hooks si el entorno asignó alguno.
                                    _set_rebar_hook_type_id_at_end(r, 0, hid, document)
                                    _set_rebar_hook_type_id_at_end(r, 1, hid, document)
                                    return r
                            except Exception:
                                continue
        return None

    r = _try_create_from_curves(ln)
    if r:
        return r
    return None


def _round_line_axis_length_mm_ceil(ln, hook_end0, hook_end1, step_mm=None, d_half_mm=None):
    """
    Eje recto de ``ln``: largo en mm → siguiente múltiplo de ``step_mm`` hacia arriba (ceil).

    Con ``d_half_mm`` (mitad del Ø nominal), igual que en patas de curve loop: el tag en la
    juntura usa el valor «mostrado» ``eje_geom + d/2``; se redondea ese valor y se
    **descuenta** ``d/2`` para obtener la longitud geométrica del eje a crear.

    Estiramiento en un extremo del eje: solo gancho en p1 → mueve p0; si no → mueve p1.

    Debe invocarse sobre **la misma** ``Line`` que se pasará a ``CreateFromCurves`` (curve
    loop o barra recta de fallback).
    """
    if ln is None:
        return None
    if step_mm is None:
        step_mm = float(SHAFT_BAR_LENGTH_ROUND_STEP_MM)
    if step_mm <= 1e-9:
        step_mm = 10.0
    try:
        import math as _math

        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
        _bx = float(p1.X - p0.X)
        _by = float(p1.Y - p0.Y)
        _bz = float(p1.Z - p0.Z)
        _bl_ft = _math.sqrt(_bx * _bx + _by * _by + _bz * _bz)
        if _bl_ft <= 1e-9:
            return ln
        _bl_actual_mm = _bl_ft * 304.8
        _dh = float(d_half_mm) if d_half_mm is not None else 0.0
        if _dh > 1e-9:
            _bl_displayed_mm = float(_bl_actual_mm) + _dh
            _disp_rounded_mm = _math.ceil(_bl_displayed_mm / step_mm) * step_mm
            _bl_axis_target_mm = _disp_rounded_mm - _dh
        else:
            _bl_axis_target_mm = _math.ceil(_bl_actual_mm / step_mm) * step_mm
        _delta_mm = float(_bl_axis_target_mm) - float(_bl_actual_mm)
        if _delta_mm <= 1e-6:
            return ln
        _delta_ft = _delta_mm / 304.8
        _ux = _bx / _bl_ft
        _uy = _by / _bl_ft
        _uz = _bz / _bl_ft
        if hook_end1 and not hook_end0:
            p0 = XYZ(
                float(p0.X) - _ux * _delta_ft,
                float(p0.Y) - _uy * _delta_ft,
                float(p0.Z) - _uz * _delta_ft,
            )
        else:
            p1 = XYZ(
                float(p1.X) + _ux * _delta_ft,
                float(p1.Y) + _uy * _delta_ft,
                float(p1.Z) + _uz * _delta_ft,
            )
        return Line.CreateBound(p0, p1)
    except Exception:
        return ln


def _create_rebar_borde_losa_curve_loop(
    document,
    host,
    bar_type,
    ln,
    face_inward_xy,
    hook_end0=True,
    hook_end1=True,
    bar_plane_normal_3d=None,
):
    """
    Crea barra de borde losa modelando las patas de gancho como tramos de curva,
    sin usar RebarHookType en la API (alineado con el enfoque de Fundación Aislada).

    Construye la polilínea con List[Curve]:
      - Ambos ganchos:  [Line(q0→p0), eje, Line(p1→q1)]
      - Solo inicio:    [Line(q0→p0), eje]
      - Solo fin:       [eje, Line(p1→q1)]
    donde q0/q1 son extremos de patas desplazados en face_inward_xy.

    Prueba la polilínea en orden original e invertido.
    Normal: leg_dir × bar_dir (≈ ±Z para barras horizontales).
    Rebar.CreateFromCurves con InvalidElementId en ambos extremos.
    """
    if document is None or host is None or bar_type is None or ln is None:
        return None
    if face_inward_xy is None:
        return None
    if not hook_end0 and not hook_end1:
        return _create_single_rebar_from_line_no_hooks(document, host, bar_type, ln)

    # Dirección de la pata: inward XY horizontal
    try:
        d = XYZ(float(face_inward_xy.X), float(face_inward_xy.Y), 0.0)
        dl = float(d.GetLength())
        if dl < 1e-12:
            return None
        leg_dir = XYZ(d.X / dl, d.Y / dl, 0.0)
    except Exception:
        return None

    # Diámetro nominal (se reutiliza en ajuste del gancho y del segmento principal).
    try:
        d_nom_mm = _ft_to_mm(float(bar_type.BarNominalDiameter))
    except Exception:
        d_nom_mm = 0.0
    _d_half = d_nom_mm / 2.0

    # Longitud de la pata (tabla por diámetro).
    # El tag muestra (largo_polilinea + d/2) en la juntura del doblez; se resta d/2
    # al segmento para que el valor etiquetado coincida con el nominal de la tabla.
    try:
        leg_mm = float(hook_leg_mm_para_bar_type(bar_type))
        if leg_mm < 10.0:
            return None
        leg_mm_adj = leg_mm - _d_half if (leg_mm - _d_half) >= 5.0 else leg_mm
        leg_ft = leg_mm_adj / 304.8
    except Exception:
        return None

    # Puntos del eje y vértices de patas. ``ln`` debe ser el eje ya redondeado que el caller
    # pasa a CreateFromCurves (misma Line que en fallback recto).
    try:
        p0 = ln.GetEndPoint(0)
        p1 = ln.GetEndPoint(1)
        dx = float(leg_dir.X) * leg_ft
        dy = float(leg_dir.Y) * leg_ft
        q0 = XYZ(float(p0.X) + dx, float(p0.Y) + dy, float(p0.Z))
        q1 = XYZ(float(p1.X) + dx, float(p1.Y) + dy, float(p1.Z))
    except Exception:
        return None

    # Polilínea principal y variante invertida (como FA)
    # Siempre se crean Line.CreateBound nuevos — nunca se reutiliza el objeto ln original.
    def _build_ilist(reversed_order):
        """Construye List[Curve] en orden normal o invertido."""
        try:
            il = List[Curve]()
            bar_line = Line.CreateBound(p0, p1)
            bar_line_rev = Line.CreateBound(p1, p0)
            if hook_end0 and hook_end1:
                if not reversed_order:
                    il.Add(Line.CreateBound(q0, p0))
                    il.Add(bar_line)
                    il.Add(Line.CreateBound(p1, q1))
                else:
                    il.Add(Line.CreateBound(q1, p1))
                    il.Add(bar_line_rev)
                    il.Add(Line.CreateBound(p0, q0))
            elif hook_end0:
                if not reversed_order:
                    il.Add(Line.CreateBound(q0, p0))
                    il.Add(bar_line)
                else:
                    il.Add(bar_line_rev)
                    il.Add(Line.CreateBound(p0, q0))
            else:  # hook_end1
                if not reversed_order:
                    il.Add(bar_line)
                    il.Add(Line.CreateBound(p1, q1))
                else:
                    il.Add(Line.CreateBound(q1, p1))
                    il.Add(bar_line_rev)
            return il if il.Count > 0 else None
        except Exception:
            return None

    ilist_normal = _build_ilist(False)
    ilist_rev = _build_ilist(True)
    curve_variants = [v for v in (ilist_normal, ilist_rev) if v is not None]
    if not curve_variants:
        return None

    # Normales: solo del plano real de la barra (como FA — no se usan ±Z arbitrarios).
    normals = []
    seen_nv_build = set()
    def _add_normal(v):
        if v is None:
            return
        try:
            vl = float(v.GetLength())
            if vl < 1e-12:
                return
            u = XYZ(float(v.X) / vl, float(v.Y) / vl, float(v.Z) / vl)
            key = (round(float(u.X), 6), round(float(u.Y), 6), round(float(u.Z), 6))
            if key in seen_nv_build:
                return
            seen_nv_build.add(key)
            normals.append(u)
        except Exception:
            pass
    # Plano real: t_barra × t_pata (como FA usa t0×t1)
    try:
        t_bar = XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), float(p1.Z - p0.Z))
        t_leg = leg_dir
        n_geom = t_bar.CrossProduct(t_leg)
        _add_normal(n_geom)
        _add_normal(XYZ(-float(n_geom.X), -float(n_geom.Y), -float(n_geom.Z)))
    except Exception:
        pass
    # bar_plane_normal_3d si fue provisto (face normal 3D)
    _add_normal(bar_plane_normal_3d)
    if bar_plane_normal_3d is not None:
        try:
            _add_normal(XYZ(-float(bar_plane_normal_3d.X), -float(bar_plane_normal_3d.Y), -float(bar_plane_normal_3d.Z)))
        except Exception:
            pass

    inv = ElementId.InvalidElementId
    orient_pairs = (
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
        (RebarHookOrientation.Left, RebarHookOrientation.Right),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    )
    # (use_existing, create_new): igual que FA — (True,True) primero
    bool_pairs = ((True, True), (True, False))
    # hook IDs: igual que FA — prueba (None,None) y (inv,inv)
    hook_id_pairs = ((None, None), (inv, inv))

    seen_nv = set()
    _attempt_count = 0
    _last_ex = None
    for curves_ilist in curve_variants:
        for nvec in normals:
            if nvec is None:
                continue
            try:
                nv_key = (round(float(nvec.X), 6), round(float(nvec.Y), 6), round(float(nvec.Z), 6))
            except Exception:
                continue
            if nv_key in seen_nv:
                continue
            seen_nv.add(nv_key)
            for so, eo in orient_pairs:
                for use_ex, create_new in bool_pairs:
                    for h0, h1 in hook_id_pairs:
                        try:
                            _attempt_count += 1
                            r = Rebar.CreateFromCurves(
                                document,
                                RebarStyle.Standard,
                                bar_type,
                                h0,
                                h1,
                                host,
                                nvec,
                                curves_ilist,
                                so,
                                eo,
                                use_ex,
                                create_new,
                            )
                            if r:
                                _set_rebar_hook_type_id_at_end(r, 0, inv, document)
                                _set_rebar_hook_type_id_at_end(r, 1, inv, document)
                                return r
                        except Exception as _ex:
                            _last_ex = str(_ex)
                            continue
        # Reiniciar seen_nv para variante invertida (mismas normales, diferente curva)
        seen_nv = set()
    return None


def _try_rebar_create_from_curves_hook_pair(
    document,
    host,
    bar_type,
    curve,
    h0,
    h1,
    face_inward_xy=None,
    bordes_losa_hook_hacia_interior=False,
    hook_end0=True,
    hook_end1=True,
    bar_plane_normal_3d=None,
):
    """
    Intenta Rebar.CreateFromCurves con una sola pareja (h0, h1) de HookTypeId;
    misma rejilla que _create_single_rebar_from_line_with_hooks (norms, orient_pairs,
    use_existing / create_new). No invierte la curva: solo la geometría dada.
    Tras éxito, refuerza extremos con _set_rebar_hook_type_id_at_end.

    Si ``bar_plane_normal_3d`` no es None (**-FaceNormal** 3D, normal inversa de la cara, **sin normalizar**),
    se proyecta ⟂ tangente (también sin forzar unitarios); se prueba ``nvec`` y ``-nvec``;
    en **ambos** extremos los ``RebarHookOrientation`` se derivan con el mismo ``nvec`` y
    ``face_inward_xy`` (-n_xy). Si ``SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT`` es True, se fuerza
    ``Left`` en ambos extremos (sin usar esa geometría).

    Si solo ``face_inward_xy`` está definido, se mantiene el intento con ±Z ordenado por tangente.
    """
    if document is None or host is None or bar_type is None or curve is None:
        return None
    ct = clr.GetClrType(Line).BaseType
    norms = [XYZ.BasisZ]
    try:
        norms.append(XYZ.BasisZ.Negate())
    except Exception:
        norms.append(XYZ(0.0, 0.0, -1.0))
    orient_pairs = (
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
    ) if SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT else _hook_orientation_pairs_shaft()
    hook_pairs = ((h0, h1),)

    def _try_one(crv, norms_list, orient_pairs_list):
        arr = System.Array.CreateInstance(ct, 1)
        arr[0] = crv
        for hs0, hs1 in hook_pairs:
            for use_existing in (True, False):
                for create_new in (False, True):
                    for nvec in norms_list:
                        if nvec is None:
                            continue
                        for so, eo in orient_pairs_list:
                            try:
                                r = Rebar.CreateFromCurves(
                                    document,
                                    RebarStyle.Standard,
                                    bar_type,
                                    hs0,
                                    hs1,
                                    host,
                                    nvec,
                                    arr,
                                    so,
                                    eo,
                                    use_existing,
                                    create_new,
                                )
                                if r:
                                    # Limpiar primero: con (InvalidElementId, hid) Revit a menudo deja
                                    # el gancho por defecto del RebarBarType en el extremo 0.
                                    inv = ElementId.InvalidElementId
                                    _set_rebar_hook_type_id_at_end(r, 0, inv, document)
                                    _set_rebar_hook_type_id_at_end(r, 1, inv, document)
                                    _set_rebar_hook_type_id_at_end(r, 0, h0, document)
                                    _set_rebar_hook_type_id_at_end(r, 1, h1, document)
                                    return r
                            except Exception:
                                continue
        return None

    if bar_plane_normal_3d is not None:
        try:
            nb = bar_plane_normal_3d
            if nb is not None:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                tdir = XYZ(
                    float(p1.X) - float(p0.X),
                    float(p1.Y) - float(p0.Y),
                    float(p1.Z) - float(p0.Z),
                )
                ut = tdir
                if _segment_length_ft(p0, p1) > 1e-9:
                    n_plane = _rebar_plane_normal_from_face_normal_and_tangent(nb, ut)
                    if n_plane is not None:
                        nneg = XYZ(
                            -float(n_plane.X),
                            -float(n_plane.Y),
                            -float(n_plane.Z),
                        )
                        iw = _unit_xy(face_inward_xy) if face_inward_xy is not None else None
                        if iw is not None:
                            for nvec in (n_plane, nneg):
                                if SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT:
                                    o0 = RebarHookOrientation.Left
                                    o1 = RebarHookOrientation.Left
                                else:
                                    pr = _compute_hook_orientations_pair_from_inverse_normal_xyz(
                                        curve,
                                        iw,
                                        nvec,
                                        hook_end0,
                                        hook_end1,
                                    )
                                    if pr is None:
                                        continue
                                    o0, o1 = pr
                                r = _try_one(curve, [nvec], [(o0, o1)])
                                if r:
                                    return r
                        r = _try_one(curve, [n_plane, nneg], orient_pairs)
                        if r:
                            return r
        except Exception:
            pass

    if face_inward_xy is not None:
        try:
            iw = _unit_xy(face_inward_xy)
            if iw is not None:
                norms_face = _order_z_norms_for_face_inward(curve, face_inward_xy)
                for norms_try in (norms_face, list(reversed(norms_face))):
                    for nvec in norms_try:
                        if nvec is None:
                            continue
                        if SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT:
                            o0 = RebarHookOrientation.Left
                            o1 = RebarHookOrientation.Left
                        else:
                            pr = _compute_hook_orientations_pair_from_inverse_normal_xyz(
                                curve,
                                iw,
                                nvec,
                                hook_end0,
                                hook_end1,
                            )
                            if pr is None:
                                continue
                            o0, o1 = pr
                        r = _try_one(curve, [nvec], [(o0, o1)])
                        if r:
                            return r
        except Exception:
            pass

    r = _try_one(curve, norms, orient_pairs)
    if r:
        return r
    return None


def _create_single_rebar_from_line_with_hooks(
    document,
    host,
    bar_type,
    ln,
    hook_type,
    face_inward_xy=None,
    bordes_losa_hook_hacia_interior=False,
    bar_plane_normal_3d=None,
):
    """Crea una barra con ganchos en ambos extremos usando el hook indicado."""
    if document is None or host is None or bar_type is None or ln is None or hook_type is None:
        return None
    try:
        hid = hook_type.Id
    except Exception:
        hid = hook_type
    r = _try_rebar_create_from_curves_hook_pair(
        document,
        host,
        bar_type,
        ln,
        hid,
        hid,
        face_inward_xy=face_inward_xy,
        bordes_losa_hook_hacia_interior=bordes_losa_hook_hacia_interior,
        hook_end0=True,
        hook_end1=True,
        bar_plane_normal_3d=bar_plane_normal_3d,
    )
    if r:
        return r
    # Fallback: crear sin ganchos y luego asignar HookTypeId.
    try:
        r2 = _create_single_rebar_from_line_no_hooks(document, host, bar_type, ln)
        if r2 is None:
            return None
        _set_rebar_hook_type_id_at_end(r2, 0, hid, document)
        _set_rebar_hook_type_id_at_end(r2, 1, hid, document)
        return r2
    except Exception:
        return None


def _create_single_rebar_from_line_with_partial_hooks(
    document,
    host,
    bar_type,
    ln,
    hook_type,
    hook_end0,
    hook_end1,
    avisos=None,
    face_inward_xy=None,
    bordes_losa_hook_hacia_interior=False,
    bar_plane_normal_3d=None,
):
    """
    Crea una barra con gancho solo en un extremo.
    1) Intenta Rebar.CreateFromCurves con una sola pareja (hid, InvalidElementId)
       o (InvalidElementId, hid), misma rejilla que la barra con dos ganchos
       (alinea tipo Standard - 90 deg. con segmentos enteros).
    2) Si falla, fallback: no_hooks + limpieza + SetHookTypeId selectivo.

    `hook_end0` corresponde al extremo 0 (inicio de la curva `ln`)
    `hook_end1` corresponde al extremo 1 (fin de la curva `ln`)
    """
    if (
        document is None
        or host is None
        or bar_type is None
        or ln is None
        or hook_type is None
    ):
        return None

    try:
        hid = hook_type.Id
    except Exception:
        hid = hook_type
    inv = ElementId.InvalidElementId

    if bool(hook_end0) and bool(hook_end1):
        return _create_single_rebar_from_line_with_hooks(
            document,
            host,
            bar_type,
            ln,
            hook_type,
            face_inward_xy=face_inward_xy,
            bordes_losa_hook_hacia_interior=bordes_losa_hook_hacia_interior,
            bar_plane_normal_3d=bar_plane_normal_3d,
        )
    if not bool(hook_end0) and not bool(hook_end1):
        return _create_single_rebar_from_line_no_hooks(document, host, bar_type, ln)

    if bool(hook_end0):
        r_try = _try_rebar_create_from_curves_hook_pair(
            document,
            host,
            bar_type,
            ln,
            hid,
            inv,
            face_inward_xy=face_inward_xy,
            bordes_losa_hook_hacia_interior=bordes_losa_hook_hacia_interior,
            hook_end0=True,
            hook_end1=False,
            bar_plane_normal_3d=bar_plane_normal_3d,
        )
    else:
        r_try = _try_rebar_create_from_curves_hook_pair(
            document,
            host,
            bar_type,
            ln,
            inv,
            hid,
            face_inward_xy=face_inward_xy,
            bordes_losa_hook_hacia_interior=bordes_losa_hook_hacia_interior,
            hook_end0=False,
            hook_end1=True,
            bar_plane_normal_3d=bar_plane_normal_3d,
        )
    if r_try is not None:
        return r_try

    r2 = _create_single_rebar_from_line_no_hooks(document, host, bar_type, ln)
    if r2 is None:
        return None
    hook_lbl = _rebar_hook_type_display_name(hook_type)
    try:
        fn_cu = getattr(r2, "CanUseHookType", None)
        if (
            avisos is not None
            and fn_cu is not None
            and int(_eid_int(hid)) >= 0
            and not bool(fn_cu(hid))
        ):
            avisos.append(
                u"Revit: el tipo de barra seleccionado no admite el gancho '{0}' como combinación "
                u"válida. Edite el RebarBarType (familia de barra) o el RebarHookType en el proyecto."
                .format(hook_lbl)
            )
    except Exception:
        pass
    # Limpieza explícita en ambos extremos antes de asignar el gancho pedido (el tipo de barra
    # puede haber dejado un gancho por defecto en un solo extremo).
    _set_rebar_hook_type_id_at_end(r2, 0, inv, document)
    _set_rebar_hook_type_id_at_end(r2, 1, inv, document)
    if bool(hook_end0):
        if not _set_rebar_hook_type_id_at_end(r2, 0, hid, document):
            _append_hook_set_failed_aviso(avisos, u"inicio", hook_lbl)
        _set_rebar_hook_type_id_at_end(r2, 1, inv, document)
    else:
        _set_rebar_hook_type_id_at_end(r2, 0, inv, document)
        if not _set_rebar_hook_type_id_at_end(r2, 1, hid, document):
            _append_hook_set_failed_aviso(avisos, u"fin", hook_lbl)

    return r2


def _reapply_partial_hooks_after_fixed_number_layout(
    rebar, hook_type, hook_end0, hook_end1, avisos=None
):
    """
    Tras SetLayoutAsFixedNumber / FlipRebarSet, Revit puede alterar HookTypeId.
    Refuerza el gancho solo en los extremos pedidos y limpia el otro con InvalidElementId.
    """
    if rebar is None or hook_type is None:
        return
    if bool(hook_end0) == bool(hook_end1):
        # Ambos o ninguno: no es el caso de gancho parcial por división.
        return
    try:
        hid = hook_type.Id
    except Exception:
        hid = hook_type
    inv = ElementId.InvalidElementId
    doc = None
    try:
        doc = rebar.Document
    except Exception:
        doc = None
    hook_lbl = _rebar_hook_type_display_name(hook_type)
    _set_rebar_hook_type_id_at_end(rebar, 0, inv, doc)
    _set_rebar_hook_type_id_at_end(rebar, 1, inv, doc)
    if bool(hook_end0):
        if not _set_rebar_hook_type_id_at_end(rebar, 0, hid, doc):
            _append_hook_set_failed_aviso(avisos, u"inicio (post-layout)", hook_lbl)
        _set_rebar_hook_type_id_at_end(rebar, 1, inv, doc)
    else:
        _set_rebar_hook_type_id_at_end(rebar, 0, inv, doc)
        if not _set_rebar_hook_type_id_at_end(rebar, 1, hid, doc):
            _append_hook_set_failed_aviso(avisos, u"fin (post-layout)", hook_lbl)


def _reapply_both_hooks_after_fixed_number_layout(rebar, hook_type, avisos=None):
    """Tras Fixed Number, fuerza el mismo RebarHookType en ambos extremos (p. ej. Standard - 90 deg.)."""
    if rebar is None or hook_type is None:
        return
    try:
        hid = hook_type.Id
    except Exception:
        hid = hook_type
    doc = None
    try:
        doc = rebar.Document
    except Exception:
        doc = None
    hook_lbl = _rebar_hook_type_display_name(hook_type)
    if not _set_rebar_hook_type_id_at_end(rebar, 0, hid, doc):
        _append_hook_set_failed_aviso(avisos, u"inicio (post-layout)", hook_lbl)
    if not _set_rebar_hook_type_id_at_end(rebar, 1, hid, doc):
        _append_hook_set_failed_aviso(avisos, u"fin (post-layout)", hook_lbl)


def _enforce_rebar_hook_types_by_name(
    rebar, document, hook_type_name, hook_end0, hook_end1, avisos=None
):
    """
    Tras crear la barra y los retoques de layout/orientación/rotación, vuelve a resolver
    el RebarHookType por nombre y asigna HookTypeId en cada extremo. Así se corrige el
    caso en que Revit deja el gancho por defecto del RebarBarType pese a la creación
    solicitada (p. ej. «Standard - 90 deg.»).
    """
    if rebar is None or document is None:
        return
    try:
        nm = unicode(hook_type_name or u"").strip()
    except Exception:
        nm = u""
    if not nm:
        return
    if not (bool(hook_end0) or bool(hook_end1)):
        return
    ht = _rebar_hook_type_by_name(document, nm)
    if ht is None:
        try:
            avisos.append(
                u"No se encontró RebarHookType '{0}' para fijar el gancho tras crear la barra."
                .format(nm)
            )
        except Exception:
            pass
        return
    try:
        hid = ht.Id
    except Exception:
        hid = ht
    inv = ElementId.InvalidElementId
    hook_lbl = _rebar_hook_type_display_name(ht)
    _set_rebar_hook_type_id_at_end(rebar, 0, inv, document)
    _set_rebar_hook_type_id_at_end(rebar, 1, inv, document)
    if bool(hook_end0) and bool(hook_end1):
        if not _set_rebar_hook_type_id_at_end(rebar, 0, hid, document):
            _append_hook_set_failed_aviso(avisos, u"inicio (tipo por nombre)", hook_lbl)
        if not _set_rebar_hook_type_id_at_end(rebar, 1, hid, document):
            _append_hook_set_failed_aviso(avisos, u"fin (tipo por nombre)", hook_lbl)
    elif bool(hook_end0):
        if not _set_rebar_hook_type_id_at_end(rebar, 0, hid, document):
            _append_hook_set_failed_aviso(avisos, u"inicio (tipo por nombre)", hook_lbl)
        _set_rebar_hook_type_id_at_end(rebar, 1, inv, document)
    else:
        _set_rebar_hook_type_id_at_end(rebar, 0, inv, document)
        if not _set_rebar_hook_type_id_at_end(rebar, 1, hid, document):
            _append_hook_set_failed_aviso(avisos, u"fin (tipo por nombre)", hook_lbl)


def _sweep_rebar_hook_types_to_name(doc, rebar_ids, hook_type_name, avisos=None):
    """
    Barrido final (tras crear barras y layout): solo extremos que **tienen** gancho
    (HookTypeId válido). Se lee el nombre del ``RebarHookType`` actual; si no coincide
    exactamente con ``hook_type_name`` (p. ej. «Standard - 90 deg.»), se asigna ese tipo.
    Cubre ganchos por defecto del ``RebarBarType`` u otros distintos al solicitado.
    """
    if doc is None or not rebar_ids:
        return 0
    try:
        nm = unicode(hook_type_name or u"").strip()
    except Exception:
        nm = u""
    if not nm:
        return 0
    target_name_cmp = _hook_compare_string(nm)
    if not target_name_cmp:
        return 0
    ht_target = _rebar_hook_type_by_name(doc, nm)
    if ht_target is None:
        try:
            avisos.append(
                u"Barrido ganchos: no se encontró RebarHookType '{0}' en el proyecto."
                .format(nm)
            )
        except Exception:
            pass
        return 0
    try:
        hid_target = ht_target.Id
    except Exception:
        hid_target = ht_target
    inv = ElementId.InvalidElementId
    fixed = 0
    seen = set()
    for rid in rebar_ids:
        try:
            eid_int = _eid_int(rid)
        except Exception:
            continue
        if eid_int in seen:
            continue
        seen.add(eid_int)
        try:
            el = doc.GetElement(rid)
        except Exception:
            continue
        if el is None or not isinstance(el, Rebar):
            continue
        for end_idx in (0, 1):
            try:
                cur_id = el.GetHookTypeId(int(end_idx))
            except Exception:
                continue
            if cur_id is None:
                continue
            if _eid_int(cur_id) == _eid_int(inv):
                continue
            cur_name_cmp = u""
            try:
                cur_ht = doc.GetElement(cur_id)
                if isinstance(cur_ht, RebarHookType):
                    cur_name_cmp = _hook_compare_string(
                        _rebar_hook_type_display_name(cur_ht)
                    )
            except Exception:
                cur_name_cmp = u""
            if cur_name_cmp and cur_name_cmp == target_name_cmp:
                continue
            if _set_rebar_hook_type_id_at_end(el, int(end_idx), hid_target, doc):
                fixed += 1
            else:
                try:
                    avisos.append(
                        u"Barrido ganchos: no se pudo asignar '{0}' en barra Id {1}, extremo {2}."
                        .format(nm, eid_int, int(end_idx))
                    )
                except Exception:
                    pass
    if fixed > 0:
        try:
            avisos.append(
                u"Barrido ganchos: {0} extremo(s) ajustados a '{1}'."
                .format(int(fixed), nm)
            )
        except Exception:
            pass
    return int(fixed)


def _compute_bordes_hook_orientations(
    hook_end0,
    hook_end1,
    bordes_losa_hook_hacia_interior=False,
    hook_end0_inward=None,
    hook_end1_inward=None,
):
    """
    Par (Left/Right) coherente con la lógica de bordes de losa, sin instancia Rebar.
    Usado al crear con CreateFromCurves (normal de plano ±Z ordenada por cara) y tras layout.
    """
    if SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT:
        return RebarHookOrientation.Left, RebarHookOrientation.Left
    try:
        inward_default = bool(bordes_losa_hook_hacia_interior)
    except Exception:
        inward_default = False

    def _inward_for_api_end(eidx):
        v = hook_end0_inward if int(eidx) == 0 else hook_end1_inward
        if v is not None:
            return bool(v)
        return inward_default

    i0 = _inward_for_api_end(0) if bool(hook_end0) else False
    i1 = _inward_for_api_end(1) if bool(hook_end1) else False
    o0 = RebarHookOrientation.Right
    o1 = RebarHookOrientation.Left
    if bool(hook_end0) and bool(hook_end1):
        if i0 and i1:
            o0 = RebarHookOrientation.Left
            o1 = RebarHookOrientation.Right
        elif (not i0) and (not i1):
            pass
        elif i0 and (not i1):
            o0 = RebarHookOrientation.Right
            o1 = RebarHookOrientation.Left
        else:
            o0 = RebarHookOrientation.Right
            o1 = RebarHookOrientation.Left
    elif bool(hook_end0) and (not bool(hook_end1)):
        if i0:
            o0 = RebarHookOrientation.Right
        else:
            o0 = RebarHookOrientation.Right
    elif bool(hook_end1) and (not bool(hook_end0)):
        if i1:
            o1 = RebarHookOrientation.Left
        else:
            o1 = RebarHookOrientation.Left
    elif inward_default:
        o0 = RebarHookOrientation.Left
        o1 = RebarHookOrientation.Right
    return o0, o1


def _reapply_hook_orientations_after_layout(
    rebar,
    hook_end0,
    hook_end1,
    bordes_losa_hook_hacia_interior=False,
    hook_end0_inward=None,
    hook_end1_inward=None,
    inverse_normal_xy_for_hooks=None,
    hook_curve_for_hooks=None,
    bar_plane_normal_3d_for_hooks=None,
):
    """
    Tras FlipRebarSet / SetLayoutAsFixedNumber el tipo de gancho puede ser el correcto
    (H10 sin mismatch) pero la orientación Left/Right queda incoherente con la creación
    (primer par usado en CreateFromCurves: Right en extremo 0, Left en extremo 1).

    En bordes de losa (`ignore_empotramientos`) la pierna del 90° debe quedar hacia el
    interior del host.

    hook_end0_inward / hook_end1_inward: si no es None, fija si ese extremo usa orientación
    «hacia interior» (gancho habitual borde) independientemente del flag global.

    inverse_normal_xy_for_hooks / hook_curve_for_hooks: si ambos no son None (-n_xy unitario
    y la misma curva usada al crear), se recalculan **los dos** ``RebarHookOrientation``.

    bar_plane_normal_3d_for_hooks: si no es None, se usa **-FaceNormal** proyectada ⟂ tangente
    (misma lógica que en ``CreateFromCurves``) para ``nvec`` en ambos extremos;
    si es None, se usa el primer ±Z de ``_order_z_norms_for_face_inward``.
    """
    if rebar is None:
        return
    o0, o1 = None, None
    if SHAFT_REBAR_HOOK_ORIENT_ALWAYS_LEFT:
        o0 = RebarHookOrientation.Left
        o1 = RebarHookOrientation.Left
    elif inverse_normal_xy_for_hooks is not None and hook_curve_for_hooks is not None:
        try:
            iw = _unit_xy(inverse_normal_xy_for_hooks)
            if iw is not None:
                nvec = None
                if bar_plane_normal_3d_for_hooks is not None:
                    nb = bar_plane_normal_3d_for_hooks
                    if nb is not None:
                        p0 = hook_curve_for_hooks.GetEndPoint(0)
                        p1 = hook_curve_for_hooks.GetEndPoint(1)
                        tdir = XYZ(
                            float(p1.X) - float(p0.X),
                            float(p1.Y) - float(p0.Y),
                            float(p1.Z) - float(p0.Z),
                        )
                        ut = tdir
                        if _segment_length_ft(p0, p1) > 1e-9:
                            nvec = _rebar_plane_normal_from_face_normal_and_tangent(nb, ut)
                if nvec is None:
                    norms_ord = _order_z_norms_for_face_inward(
                        hook_curve_for_hooks, inverse_normal_xy_for_hooks
                    )
                    nvec = norms_ord[0] if norms_ord else None
                if nvec is not None:
                    pr = _compute_hook_orientations_pair_from_inverse_normal_xyz(
                        hook_curve_for_hooks,
                        iw,
                        nvec,
                        hook_end0,
                        hook_end1,
                    )
                    if pr is not None:
                        o0, o1 = pr
        except Exception:
            pass
    if o0 is None or o1 is None:
        o0, o1 = _compute_bordes_hook_orientations(
            hook_end0,
            hook_end1,
            bordes_losa_hook_hacia_interior=bordes_losa_hook_hacia_interior,
            hook_end0_inward=hook_end0_inward,
            hook_end1_inward=hook_end1_inward,
        )
    try:
        if bool(hook_end0):
            try:
                e0 = rebar.GetHookTypeId(0)
            except Exception:
                e0 = None
            if e0 is not None and int(_eid_int(e0)) >= 0:
                rebar.SetHookOrientation(0, o0)
    except Exception:
        pass
    try:
        if bool(hook_end1):
            try:
                e1 = rebar.GetHookTypeId(1)
            except Exception:
                e1 = None
            if e1 is not None and int(_eid_int(e1)) >= 0:
                rebar.SetHookOrientation(1, o1)
    except Exception:
        pass


def _rebar_parameter_by_built_in_or_lookup(rebar, bip_attr_names, lookup_display_names):
    """Resuelve parámetro por BuiltInParameter (primer nombre válido) o LookupParameter."""
    if rebar is None:
        return None
    for nm in bip_attr_names or ():
        try:
            bip = getattr(BuiltInParameter, nm)
            p = rebar.get_Parameter(bip)
            if p is not None:
                return p
        except Exception:
            continue
    for nm in lookup_display_names or ():
        try:
            p = rebar.LookupParameter(nm)
            if p is not None:
                return p
        except Exception:
            continue
    return None


def _set_rebar_double_parameter_if_writable(param, value):
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType == StorageType.Double:
            param.Set(float(value))
            return True
    except Exception:
        pass
    return False


def _set_rebar_hook_rotation_from_display_degrees(param, deg):
    """
    Escribe rotación de gancho (paleta) a partir de grados de **visualización**.
    Revit almacena ángulos en unidades internas; en 2024–2026 conviene derivar la conversión
    de ``GetUnitTypeId()`` del parámetro (no asumir solo ``UnitTypeId.Degrees`` global).
    """
    if param is None or param.IsReadOnly:
        return False
    if param.StorageType != StorageType.Double:
        return False
    try:
        d = float(deg)
    except Exception:
        return False
    internal_val = None
    try:
        uid = param.GetUnitTypeId()
        if uid is not None:
            internal_val = UnitUtils.ConvertToInternalUnits(d, uid)
    except Exception:
        internal_val = None
    if internal_val is None:
        try:
            internal_val = UnitUtils.ConvertToInternalUnits(d, UnitTypeId.Degrees)
        except Exception:
            try:
                internal_val = d * (math.pi / 180.0)
            except Exception:
                return False
    try:
        param.Set(float(internal_val))
        return True
    except Exception:
        pass
    # Revit a veces acepta solo SetValueString según localización / estado del parámetro.
    for s in (
        u"{0:.0f}°".format(d),
        u"{0}°".format(int(round(d))),
        u"{0:.0f} deg".format(d),
        unicode(int(round(d))),
    ):
        try:
            param.SetValueString(s)
            return True
        except Exception:
            continue
    return False


def _set_rebar_string_parameter_if_writable(param, value_u):
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType == StorageType.String:
            param.Set(unicode(value_u))
            return True
    except Exception:
        pass
    return False


def _norm_param_def_name(name):
    if name is None:
        return u""
    try:
        t = unicode(name).replace(u"\u00A0", u" ").strip()
    except Exception:
        return u""
    return t


def _find_instance_parameter_by_names(element, names):
    """
    LookupParameter y, si falla, barrido por Parameters (algunos SP solo aparecen al iterar).
    """
    if element is None or not names:
        return None
    targets = []
    for nm in names:
        try:
            targets.append(_norm_param_def_name(nm).lower())
        except Exception:
            continue
    targets = [t for t in targets if t]
    if not targets:
        return None
    for nm in names:
        try:
            p = element.LookupParameter(nm)
            if p is not None:
                return p
        except Exception:
            pass
    try:
        for p in _iter_rebar_instance_parameters(element):
            if p is None:
                continue
            try:
                dn = _norm_param_def_name(p.Definition.Name).lower()
            except Exception:
                continue
            if dn in targets:
                return p
    except Exception:
        pass
    return None


def _param_definition_is_length(param):
    """Mismo criterio que leer_dimensiones_rebar_rps: dimensiones de forma A, B, C…"""
    if param is None:
        return False
    try:
        dt = param.Definition.GetDataType()
        if dt is not None and SpecTypeId.Length is not None:
            return dt == SpecTypeId.Length
    except Exception:
        pass
    return False


def _is_shape_segment_letter_name(def_name):
    """Nombres de tramo estándar en formas de armadura: una letra A–Z."""
    if def_name is None:
        return False
    try:
        s = unicode(def_name).strip()
    except Exception:
        return False
    if len(s) != 1:
        return False
    c = s.upper()
    return u"A" <= c <= u"Z"


def _iter_rebar_instance_parameters(rebar):
    if rebar is None:
        return
    try:
        if hasattr(rebar, "GetOrderedParameters"):
            coll = rebar.GetOrderedParameters()
            if coll is not None:
                for p in coll:
                    yield p
                return
    except Exception:
        pass
    try:
        for p in rebar.Parameters:
            yield p
    except Exception:
        pass


def _rebar_total_mm_from_shape_segment_parameters(rebar):
    """
    Suma los valores de instancia de los tramos de forma (A, B, C, …): Double + Length.
    Coincide con los largos que muestra Revit en propiedades (no la geometría de eje).
    """
    if rebar is None:
        return None

    def _accumulate(require_length_datatype):
        total_mm = 0.0
        found = False
        for param in _iter_rebar_instance_parameters(rebar):
            if param is None or param.StorageType != StorageType.Double:
                continue
            if require_length_datatype and not _param_definition_is_length(param):
                continue
            try:
                dn = param.Definition.Name
            except Exception:
                dn = None
            if not _is_shape_segment_letter_name(dn):
                continue
            try:
                if not param.HasValue:
                    continue
            except Exception:
                continue
            try:
                total_mm += float(_ft_to_mm(float(param.AsDouble())))
                found = True
            except Exception:
                continue
        return (total_mm, found) if found else (None, False)

    total_mm, ok = _accumulate(True)
    if ok:
        return total_mm
    # Respaldo: nombre A–Z y Double (si GetDataType no devuelve Length en algún entorno).
    total_mm, ok = _accumulate(False)
    return total_mm if ok else None


def _read_armadura_largo_total_mm_from_param(p):
    """Lee el valor actual de ``Armadura_Largo Total`` en mm entero, o None."""
    if p is None:
        return None
    try:
        if not p.HasValue:
            return None
    except Exception:
        return None
    try:
        st = p.StorageType
        if st == StorageType.String:
            raw = p.AsString()
            if raw is None:
                return None
            s = (
                unicode(raw)
                .strip()
                .replace(u",", u".")
                .replace(u"mm", u"")
                .strip()
            )
            if not s:
                return None
            return int(round(float(s)))
        if st == StorageType.Double:
            return int(round(_ft_to_mm(float(p.AsDouble()))))
        if st == StorageType.Integer:
            return int(p.AsInteger())
    except Exception:
        return None
    return None


def _apply_armadura_largo_total_to_rebars(doc, rebar_ids, avisos):
    """
    Rellena Armadura_Largo Total = suma (A + B + C + …) en mm según parámetros de forma.
    Valor siempre entero en mm (texto sin decimales; longitud/entero coherentes).
    """
    if doc is None or not rebar_ids:
        return
    try:
        doc.Regenerate()
    except Exception:
        pass
    warned_missing = False
    warned_no_segments = False
    n_ok = 0
    for rid in rebar_ids:
        try:
            rb = doc.GetElement(rid)
        except Exception:
            rb = None
        if not isinstance(rb, Rebar):
            continue
        total_mm = _rebar_total_mm_from_shape_segment_parameters(rb)
        if total_mm is None:
            if not warned_no_segments:
                warned_no_segments = True
                try:
                    avisos.append(
                        u"Armadura_Largo Total: no se leyeron tramos A, B, … (Length) en alguna barra; "
                        u"revise la forma."
                    )
                except Exception:
                    pass
            continue
        try:
            total_mm_int = int(round(float(total_mm)))
        except Exception:
            continue
        p = _find_instance_parameter_by_names(rb, ARMADURA_LARGO_TOTAL_PARAM_NAMES)
        if p is None:
            if not warned_missing:
                warned_missing = True
                try:
                    avisos.append(
                        u"Armadura_Largo Total: no está el parámetro en la categoría "
                        u"Rebar (cárguelo y asígnelo al proyecto)."
                    )
                except Exception:
                    pass
            continue
        try:
            cur_mm = _read_armadura_largo_total_mm_from_param(p)
        except Exception:
            cur_mm = None
        if cur_mm is not None and cur_mm == total_mm_int:
            continue
        ok = False
        try:
            st = p.StorageType
        except Exception:
            st = None
        try:
            if st == StorageType.String:
                ok = _set_rebar_string_parameter_if_writable(
                    p, unicode(total_mm_int)
                )
            elif st == StorageType.Double:
                ok = _set_rebar_double_parameter_if_writable(
                    p, _mm_to_ft(float(total_mm_int))
                )
            elif st == StorageType.Integer:
                if not p.IsReadOnly:
                    p.Set(int(total_mm_int))
                    ok = True
        except Exception:
            ok = False
        if ok:
            n_ok += 1


def _set_rebar_termination_hook_rotation_radians(rebar, end_idx, radians):
    """
    Rotación fuera de plano del gancho/cruceta en el extremo de la barra (unidades internas: radianes).

    Revit 2026: ``Rebar.SetTerminationRotationAngle(end, angle)``. Versiones anteriores:
    ``SetHookRotationAngle(angle, end)`` (orden de argumentos distinto).
    Escribir solo los parámetros de instancia suele **no** actualizar la geometría en 2026.
    """
    if rebar is None:
        return False
    try:
        ri = int(end_idx)
    except Exception:
        return False
    try:
        r = float(radians)
    except Exception:
        return False
    if not (r == r):  # NaN
        return False
    try:
        fn = getattr(rebar, "SetTerminationRotationAngle", None)
        if fn is not None:
            fn(ri, r)
            return True
    except Exception:
        pass
    try:
        fn = getattr(rebar, "SetHookRotationAngle", None)
        if fn is not None:
            fn(r, ri)
            return True
    except Exception:
        pass
    return False


def _apply_rebar_hook_rotation_parameters_degrees(doc, rebar_ids, degrees, avisos=None):
    """
    Recorre Rebar creados: en cada extremo con ``RebarHookType`` válido, aplica la rotación
    de terminación (gancho). En Revit 2026 preferir ``SetTerminationRotationAngle``; si no,
    ``SetHookRotationAngle`` o parámetros de paleta como respaldo.
    """
    if doc is None or not rebar_ids:
        return
    try:
        deg = float(degrees)
    except Exception:
        return
    if abs(deg) < 1e-9:
        return
    try:
        rad = UnitUtils.ConvertToInternalUnits(deg, UnitTypeId.Degrees)
    except Exception:
        try:
            rad = deg * (math.pi / 180.0)
        except Exception:
            return
    try:
        doc.Regenerate()
    except Exception:
        pass
    warned_missing = False
    for rid in rebar_ids:
        try:
            rb = doc.GetElement(rid)
        except Exception:
            rb = None
        if not isinstance(rb, Rebar):
            continue
        for end_idx, bip_names, lookup_names in (
            (0, REBAR_HOOK_ROTATION_START_BIP_NAMES, REBAR_HOOK_ROTATION_START_LOOKUP_NAMES),
            (1, REBAR_HOOK_ROTATION_END_BIP_NAMES, REBAR_HOOK_ROTATION_END_LOOKUP_NAMES),
        ):
            try:
                hid = rb.GetHookTypeId(end_idx)
            except Exception:
                hid = None
            if hid is None or int(_eid_int(hid)) < 0:
                continue
            if _set_rebar_termination_hook_rotation_radians(rb, end_idx, rad):
                continue
            p = _rebar_parameter_by_built_in_or_lookup(rb, bip_names, lookup_names)
            if p is None:
                if not warned_missing and avisos is not None:
                    warned_missing = True
                    try:
                        avisos.append(
                            u"Rotación de gancho: no se pudo aplicar vía API ni encontrar el parámetro "
                            u"(Hook Rotation At Start/End / rotación gancho inicio-fin)."
                        )
                    except Exception:
                        pass
                continue
            _set_rebar_hook_rotation_from_display_degrees(p, deg)
    try:
        doc.Regenerate()
    except Exception:
        pass


def _rebar_quantity(rebar):
    try:
        return int(rebar.Quantity)
    except Exception:
        try:
            return int(rebar.NumberOfBarPositions)
        except Exception:
            return 1


def _apply_fixed_number_layout_n_bars(rebar, n_bars, array_length_ft):
    """
    Configura Layout Rule = Fixed Number con N barras.
    Prueba combinaciones de lado/include para maximizar compatibilidad.
    """
    if rebar is None:
        return False
    try:
        n = int(n_bars)
    except Exception:
        n = 2
    if n <= 1:
        # No es set: barra simple.
        return True
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return False

    arr_len = max(float(array_length_ft), _mm_to_ft(5.0))
    # Requisito: mantener set etiquetable; usar siempre extremos incluidos.
    combos = (
        (True, True, True),
        (False, True, True),
    )

    for bars_on_normal_side, include_first, include_last in combos:
        try:
            acc.SetLayoutAsFixedNumber(
                n,
                arr_len,
                bars_on_normal_side,
                include_first,
                include_last,
            )
            if _rebar_quantity(rebar) == n:
                return True
        except Exception:
            continue

    # Refrescar el set antes de reintento con Flip (referencias más estables para tag).
    try:
        rebar.Document.Regenerate()
    except Exception:
        pass

    try:
        acc.FlipRebarSet()
    except Exception:
        pass
    for bars_on_normal_side, include_first, include_last in combos:
        try:
            acc.SetLayoutAsFixedNumber(
                n,
                arr_len,
                bars_on_normal_side,
                include_first,
                include_last,
            )
            if _rebar_quantity(rebar) == n:
                return True
        except Exception:
            continue
    return False


def _apply_fixed_number_layout_two_bars(rebar, array_length_ft):
    """Compat: wrapper histórico (N=2)."""
    return _apply_fixed_number_layout_n_bars(rebar, 2, array_length_ft)


def _layout_length_from_host_thickness(host, cover_ft, fallback_ft):
    """
    Largo de propagación del set basado en espesor del host:
    espesor_Z - recubrimiento superior - recubrimiento inferior.
    """
    min_len = _mm_to_ft(5.0)
    if host is None:
        return max(float(fallback_ft), min_len)
    try:
        bb = host.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None:
        return max(float(fallback_ft), min_len)
    try:
        thickness = float(bb.Max.Z) - float(bb.Min.Z)
        eff = thickness - 2.0 * max(0.0, float(cover_ft))
        if eff > min_len:
            return eff
    except Exception:
        pass
    return max(float(fallback_ft), min_len)


def _view_ok_for_rebar_tags(view):
    if view is None:
        return False, u"No hay vista activa para etiquetar."
    try:
        if view.IsTemplate:
            return False, u"La vista activa es plantilla; no se etiquetan barras."
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False, u"Vista 3D: se omite etiquetado para evitar fallos."
    except Exception:
        pass
    try:
        if str(view.ViewType) == "Perspective":
            return False, u"Vista perspectiva: se omite etiquetado."
    except Exception:
        pass
    return True, None


def _bbox_center_xyz(element, view):
    if element is None:
        return None
    try:
        bb = element.get_BoundingBox(view)
    except Exception:
        bb = None
    if bb is None:
        try:
            bb = element.get_BoundingBox(None)
        except Exception:
            bb = None
    if bb is None or bb.Min is None or bb.Max is None:
        return None
    mn, mx = bb.Min, bb.Max
    return XYZ((mn.X + mx.X) * 0.5, (mn.Y + mx.Y) * 0.5, (mn.Z + mx.Z) * 0.5)


def _rebar_centerline_dominant_curve(rebar):
    """
    Curva de centro dominante (la más larga) o Location.Curve; misma lógica que el punto medio de tag.
    """
    if rebar is None:
        return None

    center_curves = []
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        mpo = getattr(MultiplanarOption, "IncludeOnlyPlanarCurves", None)
        if mpo is None:
            mpo = getattr(MultiplanarOption, "IncludeAllMultiplanarCurves", None)
        if mpo is not None:
            center_curves = list(rebar.GetCenterlineCurves(False, False, False, mpo, 0))
    except Exception:
        center_curves = []

    if not center_curves:
        try:
            center_curves = list(rebar.GetCenterlineCurves(False, False, False))
        except Exception:
            center_curves = []

    if center_curves:
        best = None
        best_len = -1.0
        for c in center_curves:
            if c is None:
                continue
            try:
                ln = float(c.Length)
            except Exception:
                ln = 0.0
            if ln > best_len:
                best = c
                best_len = ln
        if best is not None:
            return best

    try:
        loc = getattr(rebar, "Location", None)
        lc = getattr(loc, "Curve", None) if loc is not None else None
        if lc is not None:
            return lc
    except Exception:
        pass

    return None


def _rebar_centerline_curves_best_effort(rebar):
    """
    Curvas de centro del Rebar probando MultiplanarOption y flags de ganchos.
    Evita quedarse solo con IncludeOnlyPlanarCurves si el tramo útil no aparece.
    """
    if rebar is None:
        return []
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        mpos = []
        for name in ("IncludeAllMultiplanarCurves", "IncludeOnlyPlanarCurves"):
            mpo = getattr(MultiplanarOption, name, None)
            if mpo is not None:
                mpos.append(mpo)
        for mpo in mpos:
            for inc_hooks in (False, True):
                for hook_ext_adj in (False, True):
                    try:
                        lst = list(
                            rebar.GetCenterlineCurves(
                                False, inc_hooks, hook_ext_adj, mpo, 0
                            )
                        )
                        if lst:
                            return lst
                    except Exception:
                        pass
        try:
            lst = list(rebar.GetCenterlineCurves(False, False, False))
            if lst:
                return lst
        except Exception:
            pass
    except Exception:
        pass
    return []


def _rebar_segment_along_face_in_plan(p0, p1, face_nxy, max_dot=0.45):
    """
    True si el tramo en XY es paralelo a la arista acotada (dirección ⟂ a la normal de cara en planta).
    Filtra patas de gancho que suelen ser casi paralelas a la normal hacia el hormigón.
    ``max_dot``: cota superior de |dot(dir, nxy)| en planta (más alto = más permisivo).
    """
    if face_nxy is None:
        return True
    try:
        du = _unit_xy(_xyz_sub(p1, p0))
        if du is None:
            return False
        if abs(_dot_xy(du, face_nxy)) > float(max_dot):
            return False
    except Exception:
        return False
    return True


def _rebar_plan_endpoints_for_embed_anchorage(rebar, view, face_nxy=None):
    """
    Extremos del tramo en planta (Z del plano de la vista) para cotas de empotramiento.

    - Con ``face_nxy`` (normal de la cara en planta): se elige el segmento recto de centro
      **más largo en XY** que vaya **a lo largo del canto** (⟂ a la normal), no la pata de
      gancho ni un tramo casi paralelo a la normal.
    - ``Location.Curve`` a veces **no** refleja el tramo editado (shape / varios tramos) o
      coincide con un tramo corto: se usa solo como respaldo si pasa el mismo criterio.
    """
    if rebar is None or view is None:
        return None, None
    z = _z_plano_para_detail_curves(view)

    def _q(p):
        return XYZ(float(p.X), float(p.Y), float(z))

    def _best_line_from_curves(curves, max_dot_parallel):
        """max_dot_parallel None = sin filtro de cara (el más largo en XY)."""
        best = None
        best_xy = -1.0
        for c in curves:
            if c is None or not getattr(c, "IsBound", False):
                continue
            if not isinstance(c, Line):
                continue
            try:
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                if max_dot_parallel is not None and not _rebar_segment_along_face_in_plan(
                    p0, p1, face_nxy, max_dot=max_dot_parallel
                ):
                    continue
                dx = float(p1.X) - float(p0.X)
                dy = float(p1.Y) - float(p0.Y)
                lxy = math.sqrt(dx * dx + dy * dy)
                if lxy > best_xy + 1e-12:
                    best_xy = lxy
                    best = (p0, p1)
            except Exception:
                continue
        return best

    curves = _rebar_centerline_curves_best_effort(rebar)

    # 1) Tramo recto más largo en XY alineado con el canto (⟂ normal de cara en planta)
    if curves and face_nxy is not None:
        for md in (0.45, 0.65, 0.9):
            best = _best_line_from_curves(curves, max_dot_parallel=md)
            if best is not None:
                p0, p1 = best
                return _q(p0), _q(p1)

    # 2) Sin normal o sin tramo que pase filtro: el segmento recto más largo en XY
    if curves:
        best = _best_line_from_curves(curves, max_dot_parallel=None)
        if best is not None:
            p0, p1 = best
            return _q(p0), _q(p1)

    # 3) Location.Curve (solo si es coherente con la cara cuando hay normal)
    try:
        loc = getattr(rebar, "Location", None)
        if isinstance(loc, LocationCurve):
            c = loc.Curve
            if c is not None and getattr(c, "IsBound", False) and isinstance(c, Line):
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                if face_nxy is None or _rebar_segment_along_face_in_plan(
                    p0, p1, face_nxy, max_dot=0.9
                ):
                    return _q(p0), _q(p1)
    except Exception:
        pass

    c0 = _rebar_centerline_dominant_curve(rebar)
    if c0 is None or not getattr(c0, "IsBound", False):
        return None, None
    try:
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        return _q(p0), _q(p1)
    except Exception:
        return None, None


def _rebar_centerline_midpoint_xyz(rebar):
    """
    Punto medio de la curva que construye el rebar para inserción de tag.
    Prioriza la curva de centro más larga; fallback None si no es legible.
    """
    c = _rebar_centerline_dominant_curve(rebar)
    if c is None:
        return None
    try:
        return c.Evaluate(0.5, True)
    except Exception:
        return None


def _aabb_intersects_3d(bb_a, bb_b):
    """Intersección de dos BoundingBoxXYZ (ejes alineados al modelo)."""
    if bb_a is None or bb_b is None:
        return False
    try:
        amin, amax = bb_a.Min, bb_a.Max
        bmin, bmax = bb_b.Min, bb_b.Max
    except Exception:
        return False
    if amin is None or amax is None or bmin is None or bmax is None:
        return False
    try:
        return (
            amin.X < bmax.X
            and amax.X > bmin.X
            and amin.Y < bmax.Y
            and amax.Y > bmin.Y
            and amin.Z < bmax.Z
            and amax.Z > bmin.Z
        )
    except Exception:
        return False


def _rebar_tangent_unit_xy(rebar):
    """Vector unitario en XY tangente a la curva de centro al param 0.5 (fallback extremos)."""
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
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            tan = p1 - p0
        except Exception:
            return None
    if tan is None:
        return None
    try:
        ln = math.sqrt(float(tan.X) * float(tan.X) + float(tan.Y) * float(tan.Y))
    except Exception:
        return None
    if ln < 1e-9:
        return None
    return XYZ(float(tan.X) / ln, float(tan.Y) / ln, 0.0)


def _perpendicular_unit_xy(t_xy):
    """Una perpendicular en el plano XY a un vector XY dado (unitario)."""
    if t_xy is None:
        return None
    try:
        tx = float(t_xy.X)
        ty = float(t_xy.Y)
    except Exception:
        return None
    ln = math.sqrt(tx * tx + ty * ty)
    if ln < 1e-9:
        return None
    return XYZ(-ty / ln, tx / ln, 0.0)


def _independent_tag_intersects_obstacles(tag, view, obstacles):
    """
    True si la AABB de la etiqueta intersecta la AABB de algún obstáculo (misma vista si aplica).
    """
    if tag is None or not obstacles:
        return False
    try:
        tbb = tag.get_BoundingBox(view)
    except Exception:
        tbb = None
    if tbb is None:
        try:
            tbb = tag.get_BoundingBox(None)
        except Exception:
            tbb = None
    if tbb is None:
        return False
    for ob in obstacles:
        if ob is None:
            continue
        try:
            obb = ob.get_BoundingBox(view)
        except Exception:
            obb = None
        if obb is None:
            try:
                obb = ob.get_BoundingBox(None)
            except Exception:
                obb = None
        if obb is None:
            continue
        if _aabb_intersects_3d(tbb, obb):
            return True
    return False


def _nudge_rebar_independent_tag_clear(doc, view, tag, host_rebar, extra_obstacle_ids=None):
    """
    Desplaza TagHeadPosition en XY (Z fija) hasta que la AABB de la etiqueta no corte
    las AABB de la barra anfitriona, otras barras o etiquetas ya colocadas en el lote.
    """
    if doc is None or view is None or tag is None or host_rebar is None:
        return
    extra_obstacle_ids = extra_obstacle_ids or []
    obstacles = [host_rebar]
    seen = set()
    try:
        seen.add(_eid_int(host_rebar.Id))
    except Exception:
        pass
    for oid in extra_obstacle_ids:
        if oid is None:
            continue
        try:
            k = _eid_int(oid)
        except Exception:
            continue
        if k in seen:
            continue
        try:
            el = doc.GetElement(oid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            if not el.IsValidObject:
                continue
        except Exception:
            continue
        seen.add(k)
        obstacles.append(el)

    try:
        doc.Regenerate()
    except Exception:
        pass
    try:
        base = tag.TagHeadPosition
    except Exception:
        base = None
    if base is None:
        return

    try:
        if not _independent_tag_intersects_obstacles(tag, view, obstacles):
            return
    except Exception:
        pass

    t_xy = _rebar_tangent_unit_xy(host_rebar)
    perp = _perpendicular_unit_xy(t_xy)
    dirs = []
    if perp is not None:
        dirs.append(perp)
        dirs.append(XYZ(-float(perp.X), -float(perp.Y), 0.0))
    dirs.extend(
        (
            XYZ(1.0, 0.0, 0.0),
            XYZ(-1.0, 0.0, 0.0),
            XYZ(0.0, 1.0, 0.0),
            XYZ(0.0, -1.0, 0.0),
        )
    )

    steps_ft = [_mm_to_ft(float(mm)) for mm in REBAR_TAG_NUDGE_STEPS_MM]

    for d in dirs:
        try:
            dx = float(d.X)
            dy = float(d.Y)
            ln = math.sqrt(dx * dx + dy * dy)
            if ln < 1e-12:
                continue
            ux = dx / ln
            uy = dy / ln
        except Exception:
            continue
        for step_ft in steps_ft:
            if step_ft <= 1e-12:
                continue
            try:
                off = XYZ(
                    float(base.X) + ux * step_ft,
                    float(base.Y) + uy * step_ft,
                    float(base.Z),
                )
                tag.TagHeadPosition = off
            except Exception:
                continue
            try:
                doc.Regenerate()
            except Exception:
                pass
            try:
                if not _independent_tag_intersects_obstacles(tag, view, obstacles):
                    return
            except Exception:
                pass

    try:
        tag.TagHeadPosition = base
    except Exception:
        pass
    try:
        doc.Regenerate()
    except Exception:
        pass


def _family_symbol_type_name_norm(document, type_id):
    """Nombre de tipo de ``FamilySymbol`` normalizado con ``_norm_text`` (p. ej. ``01``)."""
    if document is None or type_id is None or type_id == ElementId.InvalidElementId:
        return u""
    try:
        sym = document.GetElement(type_id)
    except Exception:
        sym = None
    if sym is None:
        return u""
    tn = u""
    try:
        tn = _norm_text(sym.Name)
    except Exception:
        tn = u""
    if not tn:
        try:
            p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p is not None:
                tn = _norm_text(p.AsString())
        except Exception:
            pass
    return tn or u""


def _apply_rebar_tag_leader_elbow_orthogonal_l(doc, view, tag, ref_tagged, host_rebar):
    """
    Leader en L en el plano de la vista (tres puntos):
    1) Anclaje en la barra (``GetLeaderEnd`` / respaldo en eje o bbox).
    2) Codo: ``end + V·Up`` con |V| acotado: tramo vertical desde el ancla, misma línea “vertical de pantalla”.
    3) Cabecera: ``elbow + s·Right`` (estante hacia la derecha de vista si ``s_raw≈0``).
    ``SetLeaderElbow`` primero; luego ``TagHeadPosition`` al final del estante + offsets extra
    (Revit a veces deja la cabecera junto al ancla si solo se regenere una vez).
    """
    if doc is None or view is None or tag is None or ref_tagged is None or host_rebar is None:
        return
    try:
        if not bool(tag.HasLeader):
            return
    except Exception:
        return
    try:
        doc.Regenerate()
    except Exception:
        pass
    head = None
    try:
        head = tag.TagHeadPosition
    except Exception:
        pass
    if head is None:
        return
    end = None
    try:
        end = tag.GetLeaderEnd(ref_tagged)
    except Exception:
        end = None
    if end is None:
        end = _rebar_centerline_midpoint_xyz(host_rebar)
        if end is None:
            end = _bbox_center_xyz(host_rebar, view)
    if end is None:
        return
    n = _unit_3d(view.ViewDirection)
    r = _unit_3d(view.RightDirection)
    u = _unit_3d(view.UpDirection)
    if n is None or r is None or u is None:
        return
    d = _xyz_sub(head, end)
    dn = _dot_xyz(d, n)
    d_in = XYZ(
        float(d.X) - float(n.X) * float(dn),
        float(d.Y) - float(n.Y) * float(dn),
        float(d.Z) - float(n.Z) * float(dn),
    )
    t_raw = float(_dot_xyz(d_in, u))
    s_raw = float(_dot_xyz(d_in, r))
    if abs(t_raw) < 1e-9 and abs(s_raw) < 1e-9:
        return
    drop_ft = _mm_to_ft(float(REBAR_TAG_ORTHOGONAL_DROP_MIN_MM))
    shelf_ft = _mm_to_ft(float(REBAR_TAG_ORTHOGONAL_SHELF_MIN_MM))
    # Tramo vertical V: parte del vector ancla→cabecera en Up, nunca todo el tramo si no hay Right.
    abs_t = abs(t_raw)
    if abs_t < 1e-9:
        V = -drop_ft
    elif abs_t < drop_ft:
        V = math.copysign(drop_ft, t_raw)
    else:
        V_abs = min(abs_t * 0.55, max(drop_ft, abs_t - shelf_ft * 0.25))
        V_abs = max(drop_ft, V_abs)
        V = math.copysign(V_abs, t_raw)
    s_use = s_raw
    if abs(s_use) < shelf_ft:
        if s_use >= 0.0 or abs(s_raw) < 1e-9:
            s_use = shelf_ft
        else:
            s_use = -shelf_ft
    elbow = _xyz_add(end, _xyz_scale(u, float(V)))
    extra_r_ft = _mm_to_ft(float(REBAR_TAG_ORTHOGONAL_HEAD_EXTRA_R_MM))
    extra_u_ft = _mm_to_ft(float(REBAR_TAG_ORTHOGONAL_HEAD_EXTRA_U_MM))
    s_head = float(s_use) + math.copysign(extra_r_ft, float(s_use))
    new_head = _xyz_add(
        _xyz_add(elbow, _xyz_scale(r, float(s_head))),
        _xyz_scale(u, float(extra_u_ft)),
    )
    try:
        if elbow.DistanceTo(end) < 1e-4 or elbow.DistanceTo(new_head) < 1e-4:
            return
    except Exception:
        pass
    try:
        tag.SetLeaderElbow(ref_tagged, elbow)
    except Exception:
        pass
    try:
        doc.Regenerate()
    except Exception:
        pass
    for _ in (0, 1):
        try:
            tag.TagHeadPosition = new_head
        except Exception:
            break
        try:
            doc.Regenerate()
        except Exception:
            pass


def _apply_rebar_tag_leader_elbow_for_type(doc, view, tag, ref_tagged, host_rebar, tag_type_id):
    """
    Aplica codo de leader según el nombre del tipo de etiqueta: lista
    ``REBAR_TAG_ORTHOGONAL_LEADER_TYPE_KEYS`` → L ortogonal; resto → quiebre lateral (jog).
    """
    try:
        nm = _family_symbol_type_name_norm(doc, tag_type_id)
    except Exception:
        nm = u""
    if nm in REBAR_TAG_ORTHOGONAL_LEADER_TYPE_KEYS:
        _apply_rebar_tag_leader_elbow_orthogonal_l(doc, view, tag, ref_tagged, host_rebar)
    else:
        _apply_rebar_tag_leader_elbow_jog(doc, view, tag, ref_tagged, host_rebar)


def _apply_rebar_tag_leader_elbow_jog(doc, view, tag, ref_tagged, host_rebar):
    """
    Tras colocar la etiqueta con leader, define un quiebre con ``SetLeaderElbow`` entre el anclaje
    y la cabecera (desvío corto en el plano ⟂ ViewDirection).
    """
    if doc is None or view is None or tag is None or ref_tagged is None or host_rebar is None:
        return
    try:
        if not bool(tag.HasLeader):
            return
    except Exception:
        return
    try:
        doc.Regenerate()
    except Exception:
        pass
    head = None
    try:
        head = tag.TagHeadPosition
    except Exception:
        pass
    if head is None:
        return
    end = None
    try:
        end = tag.GetLeaderEnd(ref_tagged)
    except Exception:
        end = None
    if end is None:
        end = _rebar_centerline_midpoint_xyz(host_rebar)
        if end is None:
            end = _bbox_center_xyz(host_rebar, view)
    if end is None:
        return
    try:
        vx = float(head.X) - float(end.X)
        vy = float(head.Y) - float(end.Y)
        vz = float(head.Z) - float(end.Z)
        chord_len = math.sqrt(vx * vx + vy * vy + vz * vz)
    except Exception:
        return
    if chord_len < 1e-7:
        return
    ux = vx / chord_len
    uy = vy / chord_len
    uz = vz / chord_len
    n = None
    try:
        vn = view.ViewDirection
        if vn is not None and vn.GetLength() > 1e-12:
            n = vn.Normalize()
    except Exception:
        n = None
    px = py = pz = None
    if n is not None:
        try:
            cross = n.CrossProduct(XYZ(ux, uy, uz))
            cl = float(cross.GetLength())
            if cl > 1e-12:
                cross = cross.Normalize()
                px, py, pz = float(cross.X), float(cross.Y), float(cross.Z)
        except Exception:
            pass
    if px is None:
        for alt in (XYZ(-uy, ux, 0.0), XYZ(0.0, -uz, uy), XYZ(-uz, 0.0, ux)):
            try:
                al = float(alt.GetLength())
                if al > 1e-12:
                    alt = alt.Normalize()
                    px, py, pz = float(alt.X), float(alt.Y), float(alt.Z)
                    break
            except Exception:
                continue
        if px is None:
            return
    try:
        d_along = float(REBAR_TAG_LEADER_JOG_ALONG)
        d_along = max(0.12, min(0.88, d_along))
        jog_ft = _mm_to_ft(float(REBAR_TAG_LEADER_JOG_OFFSET_MM))
        ex = float(end.X) + ux * d_along * chord_len + px * jog_ft
        ey = float(end.Y) + uy * d_along * chord_len + py * jog_ft
        ez = float(end.Z) + uz * d_along * chord_len + pz * jog_ft
        elbow = XYZ(ex, ey, ez)
    except Exception:
        return
    try:
        tag.SetLeaderElbow(ref_tagged, elbow)
    except Exception:
        pass
    try:
        doc.Regenerate()
    except Exception:
        pass


def _rebar_shape_name(document, rebar):
    lab = _primary_rebar_shape_tag_key(document, rebar)
    if lab:
        return lab
    names = _rebar_shape_name_candidates(document, rebar)
    return names[0] if names else None


def _get_rebar_shape_element(document, rebar):
    """RebarShape asignado al Rebar (mismo criterio que la UI)."""
    if document is None or rebar is None:
        return None
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        return None
    try:
        return document.GetElement(sid)
    except Exception:
        return None


def _primary_rebar_shape_tag_key(document, rebar):
    """
    Clave normalizada del RebarShape para acoplarse a tipos de etiqueta (p. ej. EST_A):
    prioriza SYMBOL_NAME_PARAM y ALL_MODEL_TYPE_NAME del tipo de forma; suele coincidir
    con el código de forma que ve el usuario (p. ej. 01 / 02), evitando tomar antes
    otro candidato corto que exista en el mapa de tags.
    """
    shp = _get_rebar_shape_element(document, rebar)
    if shp is None:
        return None
    for bip_name in ("SYMBOL_NAME_PARAM", "ALL_MODEL_TYPE_NAME"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        try:
            p = shp.get_Parameter(bip)
            if p is None or not p.HasValue:
                continue
            if p.StorageType == StorageType.String:
                s = _norm_text(p.AsString())
                if s:
                    return s
        except Exception:
            continue
    try:
        s = _norm_text(getattr(shp, "Name", None))
        return s if s else None
    except Exception:
        return None


def _rebar_shape_name_candidates(document, rebar):
    """
    Obtiene candidatos de nombre de shape del rebar por múltiples rutas API.
    Orden: primero datos del RebarShape (símbolo/tipo antes que .Name); luego parámetros del Rebar.
    """
    out = []
    seen = set()
    if document is None or rebar is None:
        return out

    def push(raw):
        n = _norm_text(raw)
        if n and n not in seen:
            seen.add(n)
            out.append(n)

    shp = _get_rebar_shape_element(document, rebar)
    if shp is not None:
        for bip_name in ("SYMBOL_NAME_PARAM", "ALL_MODEL_TYPE_NAME"):
            bip = getattr(BuiltInParameter, bip_name, None)
            if bip is None:
                continue
            try:
                p = shp.get_Parameter(bip)
                if p is not None and p.HasValue:
                    push(p.AsString())
            except Exception:
                continue
        try:
            push(getattr(shp, "Name", None))
        except Exception:
            pass

    for bip_name in ("REBAR_SHAPE", "REBAR_SHAPE_ID"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is None:
            continue
        try:
            p = rebar.get_Parameter(bip)
        except Exception:
            p = None
        if p is None or not p.HasValue:
            continue
        try:
            if p.StorageType == StorageType.String:
                push(p.AsString())
            elif p.StorageType == StorageType.ElementId:
                eid = p.AsElementId()
                if eid is not None and eid != ElementId.InvalidElementId:
                    el = document.GetElement(eid)
                    if el is not None:
                        push(getattr(el, "Name", None))
        except Exception:
            continue

    return out


def _rebar_tag_family_exists_strict(document, family_name):
    """True si existe una Family con el mismo nombre normalizado/clave que ``family_name`` (sin fallback flexible)."""
    if document is None or not family_name:
        return False
    fam_norm = _norm_family_name(family_name)
    fam_key = _norm_alnum_key(family_name)
    if not fam_norm:
        return False
    for fam in FilteredElementCollector(document).OfClass(Family):
        if fam is None:
            continue
        try:
            fn = fam.Name
        except Exception:
            continue
        c = _norm_family_name(fn)
        ck = _norm_alnum_key(fn)
        if c == fam_norm:
            return True
        if fam_key and ck and ck == fam_key:
            return True
    return False


def _collect_rebar_tag_symbol_map(document, family_name):
    """
    Mapea nombre de tipo normalizado -> ElementId de símbolo de tag, para la familia dada.
    """
    out = {}
    if document is None:
        return out

    fam_norm = _norm_family_name(family_name)
    fam_key = _norm_alnum_key(family_name)

    def _family_name_matches(candidate_name, allow_flexible=False):
        c = _norm_family_name(candidate_name)
        ck = _norm_alnum_key(candidate_name)
        if not c or not fam_norm:
            return False
        # Match estricto por defecto: evita mezclar "...TAG" con "...TAG_DOUBLE QUANTITY".
        if c == fam_norm:
            return True
        if fam_key and ck and ck == fam_key:
            return True
        if not allow_flexible:
            return False
        # Fallback flexible (solo cuando no exista match estricto).
        if fam_norm in c or c in fam_norm:
            return True
        if fam_key and ck and (fam_key in ck or ck in fam_key):
            return True
        return False

    # 1) Resolver la familia por nombre (sin asumir categoría).
    fam_ids = []
    fam_elems = []
    for fam in FilteredElementCollector(document).OfClass(Family):
        if fam is None:
            continue
        try:
            fn = fam.Name
        except Exception:
            continue
        if _family_name_matches(fn, allow_flexible=False):
            fam_ids.append(fam.Id)
            fam_elems.append(fam)

    strict_family_found = bool(fam_elems)

    # Si no hay match estricto, permitir fallback flexible — una sola familia: la de nombre más largo.
    # Si no, "… REBAR TAG" queda contenido en "… TAG_TRIPLE QUANTITY" y se mezclan símbolos con la
    # familia base; el tag muestra solo la cantidad del set (p. ej. 2ø12) en lugar de multi-capa.
    if not fam_elems:
        flex_cands = []
        for fam in FilteredElementCollector(document).OfClass(Family):
            if fam is None:
                continue
            try:
                fn = fam.Name
            except Exception:
                continue
            if _family_name_matches(fn, allow_flexible=True):
                flex_cands.append(fam)
        if flex_cands:
            flex_cands.sort(
                key=lambda f: len(_norm_family_name(getattr(f, "Name", None) or u"")),
                reverse=True,
            )
            best = flex_cands[0]
            fam_elems = [best]
            try:
                fam_ids = [best.Id]
            except Exception:
                fam_ids = []
            # Los pasos 3–4 usan ``_family_name_matches``; fijar nombre resuelto para no mezclar
            # símbolos de otras familias al primer "01" coincidente.
            try:
                bn = best.Name
            except Exception:
                bn = None
            if bn:
                fam_norm = _norm_family_name(bn)
                fam_key = _norm_alnum_key(bn)
                strict_family_found = True

    # 1b) Camino principal: usar Family.GetFamilySymbolIds() y filtrar OST_RebarTags.
    rebar_tag_cat = int(BuiltInCategory.OST_RebarTags)
    for fam in fam_elems:
        try:
            sids = list(fam.GetFamilySymbolIds())
        except Exception:
            sids = []
        for sid in sids:
            if sid is None or sid == ElementId.InvalidElementId:
                continue
            sym = document.GetElement(sid)
            if sym is None:
                continue
            try:
                cat = sym.Category
                if cat is None or _eid_int(cat.Id) != rebar_tag_cat:
                    continue
            except Exception:
                continue
            tn = u""
            try:
                tn = _norm_text(sym.Name)
            except Exception:
                tn = u""
            if not tn:
                try:
                    p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p is not None:
                        tn = _norm_text(p.AsString())
                except Exception:
                    tn = u""
            if tn and tn not in out:
                out[tn] = sid
    if out:
        return out

    # 2) Extraer símbolos de esas familias vía FamilyId (mismo filtro de categoría).
    if fam_ids:
        fam_id_keys = set(_eid_int(fid) for fid in fam_ids if fid is not None)
        for sym in FilteredElementCollector(document).OfClass(FamilySymbol):
            if sym is None:
                continue
            try:
                sfam = sym.Family
                sfid = sfam.Id if sfam is not None else None
                if sfid is None or _eid_int(sfid) not in fam_id_keys:
                    continue
                cat = sym.Category
                if cat is None or _eid_int(cat.Id) != rebar_tag_cat:
                    continue
            except Exception:
                continue
            try:
                tn = _norm_text(sym.Name)
            except Exception:
                tn = u""
            if tn and tn not in out:
                out[tn] = sym.Id
        if out:
            return out

    # 3) Fallback legacy: FamilyName en símbolos de tags de rebar.
    col = (
        FilteredElementCollector(document)
        .OfClass(FamilySymbol)
        .OfCategory(BuiltInCategory.OST_RebarTags)
    )
    for sym in col:
        if sym is None:
            continue
        try:
            fn = sym.FamilyName
        except Exception:
            continue
        if not _family_name_matches(fn, allow_flexible=(not strict_family_found)):
            continue
        try:
            tn = _norm_text(sym.Name)
        except Exception:
            tn = u""
        if tn and tn not in out:
            out[tn] = sym.Id

    # 4) Último fallback: escanear todos los FamilySymbol por FamilyName flexible.
    if not out:
        for sym in FilteredElementCollector(document).OfClass(FamilySymbol):
            if sym is None:
                continue
            try:
                cat = sym.Category
                if cat is None or _eid_int(cat.Id) != rebar_tag_cat:
                    continue
            except Exception:
                continue
            fn = None
            try:
                fn = sym.FamilyName
            except Exception:
                try:
                    sfam = sym.Family
                    fn = sfam.Name if sfam is not None else None
                except Exception:
                    fn = None
            if not _family_name_matches(fn, allow_flexible=(not strict_family_found)):
                continue
            try:
                tn = _norm_text(sym.Name)
            except Exception:
                tn = u""
            if tn and tn not in out:
                out[tn] = sym.Id

    return out


def _resolve_fixed_tag_type_id(tag_map, type_name):
    if not tag_map:
        return None
    return tag_map.get(_norm_text(type_name))


def etiquetar_rebars_creados_en_vista(
    doc,
    view,
    rebar_ids,
    family_name=u"EST_A_STRUCTURAL REBAR TAG",
    fixed_type_name=None,
    use_transaction=True,
):
    """
    Etiqueta solo las barras indicadas usando tipo de tag cuyo nombre coincide
    con el nombre del shape de la barra, dentro de la familia dada.

    use_transaction: si False, el llamador debe estar dentro de un Transaction abierto
    (p. ej. misma transacción que la creación de rebars).
    """
    avisos = []
    if doc is None:
        return 0, [], u"No hay documento para etiquetar."
    if not rebar_ids:
        return 0, [], None
    ok_view, msg_view = _view_ok_for_rebar_tags(view)
    if not ok_view:
        return 0, [msg_view], None

    tag_map = _collect_rebar_tag_symbol_map(doc, family_name)
    if not tag_map:
        return 0, [], u"No se encontraron tipos de la familia de etiquetas '{}'.".format(family_name)

    fixed_type_id = None
    if fixed_type_name is not None:
        fixed_type_id = _resolve_fixed_tag_type_id(tag_map, fixed_type_name)
        if fixed_type_id is None:
            return 0, [], u"No se encontró tipo '{}' en familia '{}'.".format(
                fixed_type_name, family_name
            )

    placed_tag_ids = []

    def _tag_loop():
        creadas = 0
        loc = []
        for rid in rebar_ids:
            rb = doc.GetElement(rid)
            if rb is None:
                continue
            extra_obstacle_ids = []
            for r_other in rebar_ids:
                if r_other is None:
                    continue
                try:
                    if _eid_int(r_other) == _eid_int(rid):
                        continue
                except Exception:
                    if r_other == rid:
                        continue
                extra_obstacle_ids.append(r_other)
            extra_obstacle_ids.extend(placed_tag_ids)
            ok, err, tag_eid = _tag_single_rebar_subtx(
                doc,
                view,
                rb,
                tag_map,
                family_name,
                forced_tag_type_id=fixed_type_id,
                forced_type_name=fixed_type_name,
                extra_obstacle_ids=extra_obstacle_ids,
            )
            if ok:
                creadas += 1
                if tag_eid is not None:
                    placed_tag_ids.append(tag_eid)
            elif err:
                loc.append(err)
        return creadas, loc

    if use_transaction:
        t = Transaction(doc, u"BIMTools: etiquetado barras shaft")
        t.Start()
        try:
            creadas, loop_avis = _tag_loop()
            avisos.extend(loop_avis)
            t.Commit()
        except Exception as ex:
            t.RollBack()
            return 0, avisos, u"Error en transacción de etiquetado:\n{0}".format(ex)
        return creadas, avisos, None

    creadas, loop_avis = _tag_loop()
    avisos.extend(loop_avis)
    return creadas, avisos, None


def _rebar_reference_candidates_for_tag(doc, view, rb):
    """Referencias candidatas para IndependentTag.Create / AddReferences (orden de prioridad)."""
    if rb is None:
        return []
    refs = []
    seen_refs = set()

    def _add_ref(r):
        if r is None:
            return
        key = None
        try:
            key = r.ConvertToStableRepresentation(doc)
        except Exception:
            try:
                key = u"{}|{}".format(
                    _eid_int(r.ElementId) if r.ElementId is not None else -1,
                    _eid_int(getattr(r, "LinkedElementId", None))
                    if getattr(r, "LinkedElementId", None) is not None
                    else -1,
                )
            except Exception:
                key = id(r)
        if key in seen_refs:
            return
        seen_refs.add(key)
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
                    sref = sub.GetReference()
                    if sref is not None:
                        _add_ref(sref)
            except Exception:
                continue

    try:
        npos = int(getattr(rb, "NumberOfBarPositions", 0))
    except Exception:
        try:
            npos = int(rb.GetNumberOfBarPositions()) if hasattr(rb, "GetNumberOfBarPositions") else 0
        except Exception:
            npos = 0
    if npos > 0:
        idxs = [0, max(0, npos - 1)]
        if npos > 2:
            idxs.append(int(npos / 2))
        seen_idx = set()
        for idx in idxs:
            if idx in seen_idx:
                continue
            seen_idx.add(idx)
            try:
                if hasattr(rb, "GetReferenceToBarPosition"):
                    rpos = rb.GetReferenceToBarPosition(idx)
                elif hasattr(rb, "GetReferenceForBarPosition"):
                    rpos = rb.GetReferenceForBarPosition(idx)
                else:
                    rpos = None
                if rpos is not None:
                    _add_ref(rpos)
            except Exception:
                continue

    try:
        rself = Reference(rb)
        if rself is not None:
            _add_ref(rself)
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
                gi = go.GetInstanceGeometry() if hasattr(go, "GetInstanceGeometry") else None
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


def _tag_single_rebar_subtx(
    doc,
    view,
    rb,
    tag_map,
    family_name,
    forced_tag_type_id=None,
    forced_type_name=None,
    extra_obstacle_ids=None,
):
    if rb is None:
        return False, u"Rebar nulo para etiquetar.", None
    rid = rb.Id if rb is not None else ElementId.InvalidElementId
    try:
        doc.Regenerate()
    except Exception:
        pass
    if forced_tag_type_id is not None:
        tag_type_id = forced_tag_type_id
    else:
        shp_keys = _rebar_shape_name_candidates(doc, rb)
        if not shp_keys:
            return False, u"Rebar {}: sin shape para resolver etiqueta.".format(_eid_int(rid)), None
        tag_type_id = None
        primary = _primary_rebar_shape_tag_key(doc, rb)
        if primary and primary in tag_map:
            tag_type_id = tag_map.get(primary)
        if tag_type_id is None:
            for sk in shp_keys:
                if sk in tag_map:
                    tag_type_id = tag_map.get(sk)
                    break
        if tag_type_id is None:
            return False, u"Rebar {}: no existe tipo de tag para shapes {} en familia '{}'.".format(
                _eid_int(rid), ", ".join(shp_keys[:4]), family_name
            ), None
    p = _rebar_centerline_midpoint_xyz(rb)
    if p is None:
        p = _bbox_center_xyz(rb, view)
    if p is None:
        return False, u"Rebar {}: sin punto para insertar etiqueta.".format(_eid_int(rid)), None

    refs = _rebar_reference_candidates_for_tag(doc, view, rb)
    if not refs:
        return False, u"Rebar {}: no se pudo obtener referencia para etiquetar.".format(_eid_int(rid)), None

    st = SubTransaction(doc)
    st.Start()
    try:
        # Asegurar que el tipo de etiqueta esté activo.
        try:
            tag_sym = doc.GetElement(tag_type_id)
            if tag_sym is not None and hasattr(tag_sym, "IsActive") and not bool(tag_sym.IsActive):
                tag_sym.Activate()
                doc.Regenerate()
        except Exception:
            pass

        created = None
        last_ex = None
        orientations = (TagOrientation.Horizontal, TagOrientation.Vertical)
        leaders = (True, False)

        def _created_with_expected_type(tag):
            if tag is None:
                return False
            try:
                tid = tag.GetTypeId()
                return (tid is not None) and (tid == tag_type_id)
            except Exception:
                return False

        for ref in refs:
            if created is not None:
                break
            for orient in orientations:
                if created is not None:
                    break
                for add_leader in leaders:
                    if created is not None:
                        break
                    # Intento A: sobrecarga con typeId explícito.
                    try:
                        created = IndependentTag.Create(
                            doc,
                            tag_type_id,
                            view.Id,
                            ref,
                            add_leader,
                            orient,
                            p,
                        )
                        if created is not None and not _created_with_expected_type(created):
                            try:
                                doc.Delete(created.Id)
                            except Exception:
                                pass
                            created = None
                    except Exception as exa:
                        last_ex = exa
                        created = None
                    # Intento B: por categoría + SetTypeId.
                    if created is None:
                        try:
                            created = IndependentTag.Create(
                                doc,
                                view.Id,
                                ref,
                                add_leader,
                                TagMode.TM_ADDBY_CATEGORY,
                                orient,
                                p,
                            )
                            if created is not None:
                                try:
                                    created.SetTypeId(tag_type_id)
                                except Exception:
                                    pass
                                if not _created_with_expected_type(created):
                                    try:
                                        doc.Delete(created.Id)
                                    except Exception:
                                        pass
                                    created = None
                        except Exception as exb:
                            last_ex = exb
                            created = None
        if created is None:
            if last_ex is None:
                raise Exception(u"No se pudo crear IndependentTag con referencia de rebar.")
            raise last_ex
        try:
            _nudge_rebar_independent_tag_clear(
                doc, view, created, rb, extra_obstacle_ids=extra_obstacle_ids
            )
        except Exception:
            pass
        st.Commit()
        try:
            tid_out = created.Id
        except Exception:
            tid_out = None
        return True, None, tid_out
    except Exception as ex:
        try:
            st.RollBack()
        except Exception:
            pass
        if forced_type_name:
            return False, u"Rebar {}: fallo al crear tag tipo '{}' ({})".format(
                _eid_int(rid), forced_type_name, ex
            ), None
        return False, u"Rebar {}: fallo al crear tag ({})".format(_eid_int(rid), ex), None


def _tag_layer_stack_rebars_subtx(
    doc,
    view,
    group_rebar_ids_ordered,
    tag_map,
    family_name,
    forced_tag_type_id=None,
    forced_type_name=None,
    extra_obstacle_ids=None,
):
    """
    Misma sucesión de capas (misma cara/tramo/split): una etiqueta anclada a la última capa
    y referencias adicionales vía ``IndependentTag.AddReferences`` (Revit 2022+).
    Devuelve (ok, err_msg, listado_ElementId_etiquetas).
    """
    out_tag_ids = []
    if not group_rebar_ids_ordered:
        return False, u"Grupo de etiquetado vacío.", out_tag_ids
    if extra_obstacle_ids is None:
        extra_obstacle_ids = []
    try:
        primary_id = group_rebar_ids_ordered[-1]
    except Exception:
        return False, u"Grupo inválido.", out_tag_ids
    primary_rb = doc.GetElement(primary_id)
    if primary_rb is None:
        return False, u"Barra principal (última capa) no encontrada.", out_tag_ids

    ok, err, tag_eid = _tag_single_rebar_subtx(
        doc,
        view,
        primary_rb,
        tag_map,
        family_name,
        forced_tag_type_id=forced_tag_type_id,
        forced_type_name=forced_type_name,
        extra_obstacle_ids=extra_obstacle_ids,
    )
    if not ok or tag_eid is None:
        return False, err, out_tag_ids

    out_tag_ids.append(tag_eid)
    if len(group_rebar_ids_ordered) < 2:
        return True, None, out_tag_ids

    tag_el = doc.GetElement(tag_eid)
    add_fn = getattr(tag_el, "AddReferences", None) if tag_el is not None else None
    if add_fn is None:
        return True, None, out_tag_ids

    try:
        hb = getattr(tag_el, "HasTagBehavior", None)
        if hb is not None and callable(hb) and not bool(hb()):
            return True, None, out_tag_ids
    except Exception:
        pass

    refs_add = List[Reference]()
    for rid in group_rebar_ids_ordered[:-1]:
        rb_other = doc.GetElement(rid)
        if rb_other is None:
            continue
        cand = _rebar_reference_candidates_for_tag(doc, view, rb_other)
        if not cand:
            continue
        refs_add.Add(cand[0])

    if refs_add.Count < 1:
        return True, None, out_tag_ids

    st = SubTransaction(doc)
    try:
        st.Start()
    except Exception:
        return True, None, out_tag_ids
    try:
        add_fn(refs_add)
        st.Commit()
        return True, None, out_tag_ids
    except Exception as ex_mh:
        try:
            st.RollBack()
        except Exception:
            pass
        aviso_mh = u"Multihost capas: {0}".format(ex_mh)
        try:
            st_del = SubTransaction(doc)
            st_del.Start()
            try:
                doc.Delete(tag_eid)
            except Exception:
                pass
            st_del.Commit()
        except Exception:
            pass
        out_tag_ids = []
        sub_errs = []
        obs_fb = list(extra_obstacle_ids or [])
        for rid in group_rebar_ids_ordered:
            rb_i = doc.GetElement(rid)
            if rb_i is None:
                continue
            ok_i, err_i, tid_i = _tag_single_rebar_subtx(
                doc,
                view,
                rb_i,
                tag_map,
                family_name,
                forced_tag_type_id=forced_tag_type_id,
                forced_type_name=forced_type_name,
                extra_obstacle_ids=obs_fb,
            )
            if ok_i and tid_i is not None:
                out_tag_ids.append(tid_i)
                obs_fb.append(tid_i)
            elif err_i:
                sub_errs.append(err_i)
        if out_tag_ids:
            extra = u" " + u"; ".join(sub_errs[:2]) if sub_errs else u""
            return True, aviso_mh + u" (fallback: una etiqueta por capa)." + extra, out_tag_ids
        fail_msg = aviso_mh
        if sub_errs:
            fail_msg += u" " + u"; ".join(sub_errs[:2])
        return False, fail_msg, []


def etiquetar_grupos_rebar_multihost_capas_en_vista(
    doc,
    view,
    rebar_groups,
    all_rebar_ids,
    family_name=u"EST_A_STRUCTURAL REBAR TAG",
    fixed_type_name=None,
    use_transaction=True,
):
    """
    Etiqueta por grupos de la misma sucesión de capas (multihost cuando hay varias capas).
    ``rebar_groups``: listas de ElementId ordenadas por índice de capa creciente.
    """
    avisos = []
    if doc is None:
        return 0, [], u"No hay documento para etiquetar."
    if not rebar_groups:
        return 0, [], None
    ok_view, msg_view = _view_ok_for_rebar_tags(view)
    if not ok_view:
        return 0, [msg_view], None

    tag_map = _collect_rebar_tag_symbol_map(doc, family_name)
    if not tag_map:
        return 0, [], u"No se encontraron tipos de la familia de etiquetas '{}'.".format(family_name)

    fixed_type_id = None
    if fixed_type_name is not None:
        fixed_type_id = _resolve_fixed_tag_type_id(tag_map, fixed_type_name)
        if fixed_type_id is None:
            return 0, [], u"No se encontró tipo '{}' en familia '{}'.".format(
                fixed_type_name, family_name
            )

    placed_tag_ids = []

    def _group_id_set(group):
        s = set()
        for g in group or []:
            try:
                s.add(_eid_int(g))
            except Exception:
                pass
        return s

    def _tag_groups_loop():
        creadas = 0
        loc = []
        for group in rebar_groups:
            if not group:
                continue
            g_int = _group_id_set(group)
            extra_obs = []
            for rid in all_rebar_ids or []:
                if rid is None:
                    continue
                try:
                    if _eid_int(rid) in g_int:
                        continue
                except Exception:
                    continue
                extra_obs.append(rid)
            extra_obs.extend(placed_tag_ids)
            ok, err, tag_id_list = _tag_layer_stack_rebars_subtx(
                doc,
                view,
                group,
                tag_map,
                family_name,
                forced_tag_type_id=fixed_type_id,
                forced_type_name=fixed_type_name,
                extra_obstacle_ids=extra_obs,
            )
            if ok:
                creadas += 1
                for tid in tag_id_list or []:
                    if tid is not None:
                        placed_tag_ids.append(tid)
            if err:
                loc.append(err)
        return creadas, loc

    if use_transaction:
        t = Transaction(doc, u"BIMTools: etiquetado barras shaft (multihost capas)")
        t.Start()
        try:
            creadas, loop_avis = _tag_groups_loop()
            avisos.extend(loop_avis)
            t.Commit()
        except Exception as ex:
            t.RollBack()
            return 0, avisos, u"Error en transacción de etiquetado:\n{0}".format(ex)
        return creadas, avisos, None

    creadas, loop_avis = _tag_groups_loop()
    avisos.extend(loop_avis)
    return creadas, avisos, None


def _z_plano_para_detail_curves(view):
    """Z del plano de trabajo de la vista (p. ej. planta); fallback 0."""
    if view is None:
        return 0.0
    try:
        o = view.Origin
        if o is not None:
            return float(o.Z)
    except Exception:
        pass
    return 0.0


def _obtener_plano_trabajo_vista(document, view):
    """
    Plano geométrico de trabajo de la vista: preferir ``SketchPlane`` de la vista
    (alineado con ``NewFamilyInstance`` en sección/planta); si no, ``ViewDirection``
    + ``Origin``.
    """
    if document is None or view is None:
        return None
    try:
        sid = getattr(view, "SketchPlaneId", None)
        if sid is not None and sid != ElementId.InvalidElementId:
            sp = document.GetElement(sid)
            if sp is not None:
                try:
                    pl = sp.GetPlane()
                    if pl is not None:
                        return pl
                except Exception:
                    pass
    except Exception:
        pass
    try:
        o = view.Origin
        vd = getattr(view, "ViewDirection", None)
        if o is not None and vd is not None and vd.GetLength() > 1e-12:
            vd = vd.Normalize()
            return Plane.CreateByNormalAndOrigin(vd, o)
    except Exception:
        pass
    return None


def _proyectar_punto_al_plano_vista(document, view, pt):
    """
    Tras obtener el solape en **coordenadas de modelo 3D** (curva fuera del plano de la vista),
    proyecta el punto al **plano de trabajo actual** de la vista (válido en planta y sección).
    """
    if view is None or pt is None:
        return None
    plane = _obtener_plano_trabajo_vista(document, view)
    if plane is not None:
        try:
            n = plane.Normal
            o = plane.Origin
            if n is None or o is None:
                raise ValueError("plane")
            if n.GetLength() > 1e-12:
                n = n.Normalize()
            v = pt - o
            d = float(v.DotProduct(n))
            return pt - n.Multiply(d)
        except Exception:
            pass
    try:
        z = _z_plano_para_detail_curves(view)
        return XYZ(float(pt.X), float(pt.Y), float(z))
    except Exception:
        return None


def _proyectar_segmento_solape_al_plano_vista(document, view, p0, p1):
    """
    Curva de solape/traslapo ya definida en modelo (p0–p1) → extremos en el plano de la vista.
    """
    if p0 is None or p1 is None:
        return None, None
    q0 = _proyectar_punto_al_plano_vista(document, view, p0)
    q1 = _proyectar_punto_al_plano_vista(document, view, p1)
    return q0, q1


def _view_is_plan_only(view):
    """
    Restricción solicitada: crear cotas solo en vistas de planta (corte horizontal).

    Incluye planta arquitectónica, estructural (Engineering), techo reflejado y plano de área:
    en ellas ``NewDimension`` coincide con el uso habitual de cotas de empotramiento/empalme.
    """
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    try:
        if isinstance(view, ViewPlan):
            vt = getattr(view, "ViewType", None)
            _plan_types = (
                ViewType.FloorPlan,
                ViewType.EngineeringPlan,
                ViewType.CeilingPlan,
                ViewType.AreaPlan,
            )
            try:
                if vt in _plan_types:
                    return True
            except Exception:
                pass
            # IronPython / comparación de enums .NET: respaldo por valor entero
            try:
                iv = int(vt)
                if iv in tuple(int(x) for x in _plan_types):
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def _view_accepts_overlap_dimension(view):
    """
    Vista donde se puede colocar cota lineal de traslapo sobre refs del detail (planta, sección,
    alzado, detalle, etc.). Excluye 3D y plantillas — alineado con ``NewDimension`` en vista activa
    (p. ej. enfierrado vigas en la vista corriente).
    """
    if view is None:
        return False
    try:
        if view.IsTemplate:
            return False
    except Exception:
        pass
    try:
        if isinstance(view, View3D):
            return False
    except Exception:
        pass
    return True


def _dim_line_endpoints_traslapo_inward_to_offset_sign(inward_src, view_normal_n):
    """
    Dado un vector "hacia el interior del hormigón" (3D) y la normal de la vista,
    devuelve el vector en el plano de la vista que apunta hacia **afuera** del host
    (opuesto a la proyección del interior), normalizado, o None.
    """
    if inward_src is None or view_normal_n is None:
        return None
    try:
        n = view_normal_n
        if n.GetLength() < 1e-12:
            return None
        n = n.Normalize()
        iw3 = inward_src
        if iw3.GetLength() < 1e-12:
            return None
        iw3 = iw3.Normalize()
        di = float(iw3.DotProduct(n))
        iw_p = iw3 - n.Multiply(di)
        if iw_p.GetLength() < 1e-12:
            return None
        iw_p = iw_p.Normalize()
        return XYZ(-float(iw_p.X), -float(iw_p.Y), -float(iw_p.Z))
    except Exception:
        return None


def _dim_line_endpoints_traslapo_en_plano_vista(
    view,
    lap_start,
    lap_end,
    axis_u,
    lateral_hint,
    line_offset_mm,
    inward_dir_xy=None,
    inward_dir_3d=None,
    flip_dimension_side=False,
):
    """
    Extremos de la línea de cota paralelos al tramo de solape, **en el plano de la vista**
    (no solo XY + Z de origen). ``lap_start`` / ``lap_end`` deben ser puntos ya proyectados al plano.
    """
    if lap_start is None or lap_end is None or view is None:
        return None, None
    try:
        n = getattr(view, "ViewDirection", None)
        if n is None or n.GetLength() < 1e-12:
            return None, None
        n = n.Normalize()
    except Exception:
        return None, None
    u_raw = None
    try:
        if axis_u is not None and axis_u.GetLength() > 1e-12:
            u_raw = axis_u.Normalize()
    except Exception:
        u_raw = None
    if u_raw is None:
        try:
            u_raw = _unit_3d(lap_end - lap_start)
        except Exception:
            u_raw = None
    if u_raw is None and lateral_hint is not None:
        u_raw = _unit_3d(lateral_hint)
    if u_raw is None:
        return None, None
    try:
        du = float(u_raw.DotProduct(n))
        u_in = u_raw - n.Multiply(du)
        if u_in.GetLength() < 1e-12:
            return None, None
        u_in = u_in.Normalize()
    except Exception:
        return None, None
    try:
        tdir = n.CrossProduct(u_in)
        if tdir.GetLength() < 1e-12:
            tdir = u_in.CrossProduct(n)
        if tdir.GetLength() < 1e-12:
            return None, None
        tdir = tdir.Normalize()
    except Exception:
        return None, None
    try:
        off_ft = _mm_to_ft(float(line_offset_mm))
    except Exception:
        off_ft = _mm_to_ft(450.0)
    sign = 1.0
    ow = None
    if inward_dir_3d is not None:
        ow = _dim_line_endpoints_traslapo_inward_to_offset_sign(inward_dir_3d, n)
    if ow is None and inward_dir_xy is not None:
        try:
            iw = _unit_xy(inward_dir_xy)
            if iw is not None:
                iw_flat = XYZ(float(iw.X), float(iw.Y), 0.0)
                ow = _dim_line_endpoints_traslapo_inward_to_offset_sign(iw_flat, n)
        except Exception:
            ow = None
    if ow is not None:
        try:
            dp = float(tdir.DotProduct(ow))
            if dp < -1e-9:
                sign = -1.0
            elif dp > 1e-9:
                sign = 1.0
        except Exception:
            pass
    # Si ya hay interior/exterior vía normal de cara, ``flip`` lo invertiría y metería la cota en el host.
    if flip_dimension_side and ow is None:
        sign *= -1.0
    try:
        out_vec = tdir.Multiply(float(sign) * off_ft)
        a = lap_start + out_vec
        b = lap_end + out_vec
        return a, b
    except Exception:
        return None, None


def _unit_xy_from_xyz(v):
    if v is None:
        return None
    try:
        return _unit_xy(XYZ(float(v.X), float(v.Y), 0.0))
    except Exception:
        return None


def _dot_xy(a, b):
    try:
        return float(a.X) * float(b.X) + float(a.Y) * float(b.Y)
    except Exception:
        return 0.0


def _rotate90_xy(v):
    if v is None:
        return None
    try:
        return _unit_xy(XYZ(-float(v.Y), float(v.X), 0.0))
    except Exception:
        return None


def _face_origin(face):
    if face is None:
        return None
    try:
        return face.Origin
    except Exception:
        return None


def _face_normal_xy(face):
    if face is None:
        return None
    try:
        n = face.FaceNormal
    except Exception:
        n = None
    return _unit_xy_from_xyz(n) if n is not None else None


def _point_plane_distance_ft_xy(p, face):
    """
    Distancia escalar a plano vertical en planta (aprox): |(p-origin)·n_xy|.
    """
    if p is None or face is None:
        return None
    o = _face_origin(face)
    nxy = _face_normal_xy(face)
    if o is None or nxy is None:
        return None
    try:
        v = XYZ(float(p.X) - float(o.X), float(p.Y) - float(o.Y), 0.0)
        return abs(_dot_xy(v, nxy))
    except Exception:
        return None


def _signed_plane_offset_ft_xy(p, face_origin, nxy):
    """Devuelve (p-origin)·nxy en ft (solo XY), con signo."""
    if p is None or face_origin is None or nxy is None:
        return None
    try:
        v = XYZ(float(p.X) - float(face_origin.X), float(p.Y) - float(face_origin.Y), 0.0)
        return float(_dot_xy(v, nxy))
    except Exception:
        return None


def _project_point_to_face_plane_xy(p, face_origin, nxy):
    """
    Proyecta p al plano de la cara (en XY) usando la normal nxy.
    Devuelve el punto proyectado (Z intacta del input).
    """
    s = _signed_plane_offset_ft_xy(p, face_origin, nxy)
    if s is None:
        return p
    try:
        return XYZ(float(p.X) - float(nxy.X) * float(s), float(p.Y) - float(nxy.Y) * float(s), float(p.Z))
    except Exception:
        return p


def _marker_point_for_table_anchorage(end_pt, bar_dir_xy, end_index, face_origin, face_nxy, anchorage_mm):
    """
    Calcula el punto del marcador para que la cota muestre exactamente anchorage_mm:
      - Proyecta el endpoint al plano de la cara en XY
      - Avanza anchorage_mm sobre ±normal (mismo lado donde está el endpoint)
    """
    if end_pt is None or face_origin is None or face_nxy is None:
        return end_pt, end_pt
    nxy = face_nxy
    base_on_plane = _project_point_to_face_plane_xy(end_pt, face_origin, nxy)
    a_ft = _mm_to_ft(float(anchorage_mm))
    # Siempre hacia afuera: mantener el marcador en el mismo lado del plano donde está el endpoint.
    # (El filtro de cara/extremo ya asegura que ese lado corresponde a "hacia hormigón".)
    side = 1.0
    s = _signed_plane_offset_ft_xy(end_pt, face_origin, nxy)
    if s is not None and float(s) < 0.0:
        side = -1.0
    try:
        marker_pt = XYZ(
            float(base_on_plane.X) + float(nxy.X) * float(side) * float(a_ft),
            float(base_on_plane.Y) + float(nxy.Y) * float(side) * float(a_ft),
            float(base_on_plane.Z),
        )
    except Exception:
        marker_pt = base_on_plane
    return base_on_plane, marker_pt


def _choose_face_for_bar_end(end_pt, bar_dir_xy, face_infos, align_min_dot=0.85):
    """
    Escoge la cara (de las seleccionadas) cuya normal en planta sea casi colineal con el eje
    de la barra (perpendicular a la cara), y cuyo plano esté más cerca del punto extremo.

    face_infos: list of dict { 'ref':Reference, 'face':PlanarFace, 'nxy':XYZ }
    """
    if end_pt is None or bar_dir_xy is None or not face_infos:
        return None
    best = None
    best_dist = None
    for fi in face_infos:
        nxy = fi.get("nxy")
        if nxy is None:
            continue
        # La cara a acotar debe ser perpendicular a la barra: normal || barra_dir.
        if abs(_dot_xy(nxy, bar_dir_xy)) < float(align_min_dot):
            continue
        # Siempre hacia afuera: el punto debe estar del lado "hacia hormigón" del plano.
        inward_dir = fi.get("inward_dir")
        face = fi.get("face")
        o = _face_origin(face)
        if inward_dir is not None and o is not None:
            try:
                v = XYZ(float(end_pt.X) - float(o.X), float(end_pt.Y) - float(o.Y), 0.0)
                d_in = float(_dot_xy(v, inward_dir))
                if d_in <= 1e-6:
                    continue
            except Exception:
                pass
        d = _point_plane_distance_ft_xy(end_pt, fi.get("face"))
        if d is None:
            continue
        if best is None or d < best_dist:
            best = fi
            best_dist = d
    return best


def _pick_end_for_anchorage(q0e, q1e, bar_dir_xy, face_infos, expected_mm, tol_mm=5.0):
    """
    Según tu regla: escoger el primer extremo que cumpla empotramiento >= esperado (según Ø).
    Si ambos cumplen, se queda el primero.
    """
    ends = (q0e, q1e)
    exp_ft = _mm_to_ft(float(expected_mm))
    tol_ft = _mm_to_ft(float(tol_mm))
    best_fallback = None  # (idx, fi, d_in_ft)
    for idx, pt in enumerate(ends):
        fi = _choose_face_for_bar_end(pt, bar_dir_xy, face_infos)
        if fi is None:
            continue
        face = fi.get("face")
        d = _point_plane_distance_ft_xy(pt, face)
        if d is None:
            continue
        if d + tol_ft >= exp_ft:
            return idx, fi
        # Fallback: elegir el extremo más hacia afuera (mayor d_in sobre inward_dir).
        inward_dir = fi.get("inward_dir")
        o = _face_origin(face)
        d_in = None
        if inward_dir is not None and o is not None:
            try:
                v = XYZ(float(pt.X) - float(o.X), float(pt.Y) - float(o.Y), 0.0)
                d_in = float(_dot_xy(v, inward_dir))
            except Exception:
                d_in = None
        if d_in is None:
            d_in = float(d)
        if best_fallback is None or d_in > best_fallback[2] + 1e-9:
            best_fallback = (idx, fi, d_in)
    if best_fallback is not None:
        return best_fallback[0], best_fallback[1]
    return 0, None


def _create_marker_detailcurve(doc, view, at_pt, face_nxy, length_mm=200.0):
    """
    DetailCurve corto paralelo a la cara (dirección tangente), centrado en at_pt (en Z del plano).
    Devuelve (detailcurve, reference) o (None, None).
    """
    if doc is None or view is None or at_pt is None or face_nxy is None:
        return None, None
    tdir = _rotate90_xy(face_nxy)
    if tdir is None:
        return None, None
    z = _z_plano_para_detail_curves(view)
    half = 0.5 * _mm_to_ft(float(length_mm))
    c = XYZ(float(at_pt.X), float(at_pt.Y), float(z))
    p0 = _xyz_sub(c, _xyz_scale(tdir, half))
    p1 = _xyz_add(c, _xyz_scale(tdir, half))
    try:
        ln = Line.CreateBound(p0, p1)
        dc = doc.Create.NewDetailCurve(view, ln)
        if dc is None:
            return None, None
        try:
            return dc, dc.GeometryCurve.Reference
        except Exception:
            # Si no podemos obtener la Reference de la GeometryCurve, no intentamos
            # usar otras referencias (pueden no ser acotables o resolverse a otra cosa).
            return dc, None
    except Exception:
        return None, None


def _face_tangent_dir_xy(face, face_nxy):
    """
    Dirección tangente horizontal (en XY) tomada desde la propia PlanarFace.
    En caras verticales, uno de los vectores del plano suele ser horizontal y el otro vertical.
    """
    # 1) Preferir XVector/YVector de la cara.
    try:
        for cand in (getattr(face, "XVector", None), getattr(face, "YVector", None)):
            if cand is None:
                continue
            v = _unit_xy_from_xyz(cand)
            if v is None:
                continue
            return v
    except Exception:
        pass
    # 2) Fallback: rotación 90° de la normal en planta.
    return _rotate90_xy(face_nxy)


def _create_dimension_face_to_marker(
    doc,
    view,
    face_ref,
    marker_ref,
    dim_line_origin,
    face_nxy,
    face_obj=None,
    outside_offset_mm=250.0,
    line_len_mm=450.0,
    solids=None,
    dim_line_template=None,
):
    """
    Crea Dimension lineal entre Reference de cara y Reference del marcador.
    La línea de cota se ubica por fuera del hueco (opuesto a interior de hormigón).
    """
    if doc is None or view is None or face_ref is None or marker_ref is None or dim_line_origin is None or face_nxy is None:
        return None
    # Offset de la línea de cota: debe separar la cota de la barra (perpendicular a la barra).
    # Usamos la tangente de la cara en planta (perpendicular a la normal), y orientamos el signo
    # con la inferencia de "hacia hormigón" para ir hacia afuera del hueco.
    tdir = _face_tangent_dir_xy(face_obj, face_nxy) if face_obj is not None else _rotate90_xy(face_nxy)
    if tdir is None:
        return None
    sgn_in = None
    try:
        probe_ft = _mm_to_ft(50.0)
        sgn_in = _infer_inward_xy_sign(dim_line_origin, face_nxy, solids or [], probe_ft)
    except Exception:
        sgn_in = None
    if sgn_in is None:
        sgn_in = -1.0
    out_vec = _xyz_scale(tdir, float(-sgn_in) * _mm_to_ft(float(outside_offset_mm)))
    z = _z_plano_para_detail_curves(view)
    base = XYZ(float(dim_line_origin.X), float(dim_line_origin.Y), float(z))
    base = _xyz_add(base, out_vec)
    # Preferir usar la MISMA curva (línea) con que se crea la barra como dimLine,
    # proyectada al plano de la vista y desplazada hacia afuera.
    dim_line = None
    if dim_line_template is not None:
        try:
            p0 = dim_line_template.GetEndPoint(0)
            p1 = dim_line_template.GetEndPoint(1)
            a = XYZ(float(p0.X), float(p0.Y), float(z))
            b = XYZ(float(p1.X), float(p1.Y), float(z))
            a = _xyz_add(a, out_vec)
            b = _xyz_add(b, out_vec)
            dim_line = Line.CreateBound(a, b)
        except Exception:
            dim_line = None
    if dim_line is None:
        half = 0.5 * _mm_to_ft(float(line_len_mm))
        a = _xyz_sub(base, _xyz_scale(tdir, half))
        b = _xyz_add(base, _xyz_scale(tdir, half))
        try:
            dim_line = Line.CreateBound(a, b)
        except Exception:
            return None
    try:
        from Autodesk.Revit.DB import ReferenceArray
        ra = ReferenceArray()
        ra.Append(face_ref)
        ra.Append(marker_ref)
        dim = doc.Create.NewDimension(view, dim_line, ra)
        _try_apply_fixed_dimension_type(doc, dim)
        return dim
    except Exception:
        return None


def _dimension_reference_count(dim):
    """
    Intenta leer cuántas referencias reales tiene la Dimension.
    Si falla, devuelve None.
    """
    if dim is None:
        return None
    try:
        refs = dim.References
    except Exception:
        refs = None
    if refs is None:
        return None
    try:
        return int(refs.Size)
    except Exception:
        try:
            return int(refs.Count)
        except Exception:
            try:
                return int(len(list(refs)))
            except Exception:
                return None


def _place_line_based_detail_component(doc, view, family_symbol, p0, p1):
    """
    Coloca un Detail line-based: ``p0``/``p1`` son el solape en **modelo 3D**; primero se
    **proyectan** al plano de trabajo de la vista (SketchPlane / corte), luego ``Line`` + API.
    Retorna (ok:bool, err:unicode|None, inst_or_none).
    """
    if doc is None or view is None or family_symbol is None:
        return False, u"Parámetros incompletos para detalle de empalme.", None
    if p0 is None or p1 is None:
        return False, u"Segmento de empalme inválido.", None
    try:
        p0p, p1p = _proyectar_segmento_solape_al_plano_vista(doc, view, p0, p1)
        if p0p is None or p1p is None:
            return False, u"No se pudo proyectar el empalme al plano de la vista.", None
        if p0p.DistanceTo(p1p) <= _mm_to_ft(1.0):
            return (
                False,
                u"El solape proyectado en esta vista es nulo o demasiado corto (eje de barra paralelo al corte).",
                None,
            )
        p0, p1 = p0p, p1p
    except Exception:
        pass
    try:
        seg_len_ft = _segment_length_ft(p0, p1)
    except Exception:
        seg_len_ft = 0.0
    # Evita líneas degeneradas que rompen NewFamilyInstance(line, ...).
    if seg_len_ft <= _mm_to_ft(1.0):
        return False, u"Segmento de empalme demasiado corto.", None
    try:
        ln = Line.CreateBound(p0, p1)
    except Exception:
        return False, u"No se pudo construir la línea del empalme.", None
    try:
        if not bool(getattr(family_symbol, "IsActive", True)):
            family_symbol.Activate()
            try:
                doc.Regenerate()
            except Exception:
                pass
    except Exception:
        pass
    try:
        inst = doc.Create.NewFamilyInstance(ln, family_symbol, view)
        return True, None, inst
    except Exception as ex:
        # Último recurso: solo XY + Z de corte (vistas antiguas / API rara).
        try:
            z = _z_plano_para_detail_curves(view)
            p0f = XYZ(float(p0.X), float(p0.Y), float(z))
            p1f = XYZ(float(p1.X), float(p1.Y), float(z))
            if _segment_length_ft(p0f, p1f) <= _mm_to_ft(1.0):
                return False, u"{0}".format(ex), None
            ln_p = Line.CreateBound(p0f, p1f)
            inst = doc.Create.NewFamilyInstance(ln_p, family_symbol, view)
            return True, None, inst
        except Exception:
            return False, u"{0}".format(ex), None


def _get_named_left_right_refs_from_detail_instance(detail_inst):
    """
    Obtiene referencias izquierda/derecha desde la instancia del detail component
    usando nombres estables de referencia en familia.
    """
    if detail_inst is None:
        return None, None, u"No existe instancia de detail para extraer referencias."
    left_candidates = (u"Left", u"LEFT", u"Izquierda", u"IZQUIERDA")
    right_candidates = (u"Right", u"RIGHT", u"Derecha", u"DERECHA")
    ref_left = None
    ref_right = None
    for nm in left_candidates:
        try:
            r = detail_inst.GetReferenceByName(nm)
        except Exception:
            r = None
        if r is not None:
            ref_left = r
            break
    for nm in right_candidates:
        try:
            r = detail_inst.GetReferenceByName(nm)
        except Exception:
            r = None
        if r is not None:
            ref_right = r
            break
    if ref_left is None or ref_right is None:
        return None, None, u"La familia de empalme no expone referencias nombradas Left/Right."
    return ref_left, ref_right, None


def _create_overlap_dimension_from_detail_refs(
    doc,
    view,
    ref_left,
    ref_right,
    lap_start,
    lap_end,
    axis_u,
    lateral_hint=None,
    line_offset_mm=450.0,
    inward_dir_xy=None,
    inward_dir_3d=None,
    use_view_plane_dim_line=False,
    flip_dimension_side=False,
):
    """
    Crea cota de traslapo usando referencias izquierda/derecha del detail instance.
    La línea de cota se desplaza perpendicular al eje del empalme; si se pasa
    ``inward_dir_xy`` o ``inward_dir_3d`` (hacia el interior del hormigón desde la cara),
    el desplazamiento queda hacia **afuera** del host (lado opuesto al interior).
    Preferir ``inward_dir_3d`` cuando la normal de cara tenga componente fuerte en Z
    (cara inferior/superior en alzado): si solo se usa XY con Z=0, el vector puede degenerar.
    ``flip_dimension_side=True`` invierte el lado solo si **no** hay
    ``inward_dir_3d``/``inward_dir_xy`` útiles (respaldo del comportamiento antiguo en XY).
    Con dirección interior conocida, no debe usarse flip: anula la normal de cara.
    Con ``use_view_plane_dim_line=True``, los extremos de la línea de cota se calculan
    en el **plano de la vista** (vista actual: planta, sección, alzado); si es ``False``,
    se mantiene el comportamiento histórico (XY + Z de ``_z_plano_para_detail_curves``).
    Retorna (ok:bool, msg:unicode|None, data:dict|None).
    """
    if (
        doc is None
        or view is None
        or ref_left is None
        or ref_right is None
        or lap_start is None
        or lap_end is None
    ):
        return False, u"Parámetros incompletos para cota desde referencias del detail.", None
    dxy = _unit_xy(axis_u if axis_u is not None else _xyz_sub(lap_end, lap_start))
    if dxy is None:
        if not use_view_plane_dim_line:
            return False, u"No se pudo obtener dirección de traslapo.", None
        dxy = XYZ(1.0, 0.0, 0.0)
    tdir = _rotate90_xy(dxy)
    if tdir is None:
        if lateral_hint is not None:
            tdir = _unit_xy(lateral_hint)
        if tdir is None:
            tdir = XYZ(1.0, 0.0, 0.0)
    off_ft = _mm_to_ft(float(line_offset_mm))
    sign = 1.0
    if inward_dir_3d is not None:
        try:
            iw = inward_dir_3d
            if iw.GetLength() > 1e-12:
                iw = iw.Normalize()
                iw_xy = _unit_xy(XYZ(float(iw.X), float(iw.Y), 0.0))
                if iw_xy is not None:
                    ow = _unit_xy(XYZ(-float(iw_xy.X), -float(iw_xy.Y), 0.0))
                    if ow is not None:
                        dp = _dot_xy(tdir, ow)
                        if dp < -1e-9:
                            sign = -1.0
                        elif dp > 1e-9:
                            sign = 1.0
        except Exception:
            pass
    elif inward_dir_xy is not None:
        iw = _unit_xy(inward_dir_xy)
        if iw is not None:
            ow = _unit_xy(XYZ(-float(iw.X), -float(iw.Y), 0.0))
            if ow is not None:
                dp = _dot_xy(tdir, ow)
                if dp < -1e-9:
                    sign = -1.0
                elif dp > 1e-9:
                    sign = 1.0
    sign_fallback = float(sign) * (-1.0 if flip_dimension_side else 1.0)
    out_vec = _xyz_scale(tdir, float(sign_fallback) * off_ft)
    try:
        from Autodesk.Revit.DB import ReferenceArray

        a = b = None
        if use_view_plane_dim_line:
            a, b = _dim_line_endpoints_traslapo_en_plano_vista(
                view,
                lap_start,
                lap_end,
                axis_u,
                lateral_hint,
                line_offset_mm,
                inward_dir_xy=inward_dir_xy,
                inward_dir_3d=inward_dir_3d,
                flip_dimension_side=flip_dimension_side,
            )
        if a is None or b is None:
            z = _z_plano_para_detail_curves(view)
            a = XYZ(
                float(lap_start.X) + float(out_vec.X),
                float(lap_start.Y) + float(out_vec.Y),
                float(z),
            )
            b = XYZ(
                float(lap_end.X) + float(out_vec.X),
                float(lap_end.Y) + float(out_vec.Y),
                float(z),
            )
        dim_line = Line.CreateBound(a, b)
        ra = ReferenceArray()
        ra.Append(ref_left)
        ra.Append(ref_right)
        dim = doc.Create.NewDimension(view, dim_line, ra)
    except Exception:
        dim = None
    if dim is None:
        return False, u"Revit no permitió crear cota de traslapo desde referencias del detail.", None
    nrefs = _dimension_reference_count(dim)
    if (nrefs is None) or (nrefs != 2):
        try:
            doc.Delete(dim.Id)
        except Exception:
            pass
        return False, u"Cota de traslapo inválida (referencias extra).", None
    _try_apply_fixed_dimension_type(doc, dim)
    return True, None, {"dim_id": _eid_int(dim.Id)}

def crear_detail_curves_tramos_shaft_hashtag(
    doc,
    view,
    host,
    refs,
    cover_mm=None,
    ignore_empotramientos=True,
    forced_bar_type_id=None,
    hook_type_name=None,
    use_transaction=True,
):
    """
    Dibuja DetailCurve en la vista con la misma geometría que la armadura (capa 0) antes de
    partir por longitud máxima: tramo en cara → mismo pipeline que crear_enfierrado_shaft_hashtag
    (SCI, simetría si aplica) → offset al eje en el host → redondeo de longitud a múltiplo de
    SHAFT_BAR_LENGTH_ROUND_STEP_MM repartido por igual en ambos extremos (mismo centro).

    Returns:
        (creadas:int, avisos:list, err:unicode|None)
    """
    avisos = []
    if doc is None or host is None or not refs:
        return 0, [], u"No hay documento, host o caras."
    if view is None:
        return 0, [], u"No hay vista activa (necesaria para DetailCurve)."

    if not isinstance(host, Floor):
        return 0, [], u"El host no es una losa (Floor). Id: {}".format(_eid_int(host.Id))

    cov = float(cover_mm if cover_mm is not None else COVER_MM_DEFAULT)
    cover_ft = _mm_to_ft(cov)

    bar_type, exact_bt, delta_bt = _resolve_shaft_bar_type(doc, forced_bar_type_id)
    if bar_type is None:
        return 0, [], u"No hay ningún RebarBarType compatible en el proyecto."

    hook_type_resolved = None
    if hook_type_name:
        try:
            hook_type_resolved = _rebar_hook_type_by_name(doc, hook_type_name)
        except Exception:
            hook_type_resolved = None

    hook_extra_cover_ft = 0.0
    if bool(ignore_empotramientos) and hook_type_resolved is not None:
        try:
            hlen_ft = float(bar_type.GetHookLength(hook_type_resolved.Id) or 0.0)
        except Exception:
            hlen_ft = 0.0
        hook_extra_cover_ft = max(0.0, hlen_ft)
        try:
            bb = host.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is not None:
            try:
                host_thk_ft = float(bb.Max.Z - bb.Min.Z)
                max_extra = max(0.0, 0.5 * host_thk_ft - float(cover_ft))
                hook_extra_cover_ft = min(float(hook_extra_cover_ft), float(max_extra))
            except Exception:
                pass

    bar_diam_ft = float(_bar_nominal_diameter_ft(bar_type) or 0.0)
    cover_axis_ft = float(cover_ft) + 0.5 * max(bar_diam_ft, 1e-6)
    cover_vertical_ft = float(cover_axis_ft) + float(hook_extra_cover_ft)

    no_forcing = (
        forced_bar_type_id is None or forced_bar_type_id == ElementId.InvalidElementId
    )
    if bool(ignore_empotramientos):
        if no_forcing:
            extend_each_side_mm, extend_rule_text = extension_mm_por_diametro_nominal_mm(12.0)
        else:
            extend_each_side_mm, extend_rule_text, _d_nom = extension_mm_para_bar_type(bar_type)
    elif no_forcing:
        extend_each_side_mm, extend_rule_text = extension_mm_por_diametro_nominal_mm(12.0)
    else:
        extend_each_side_mm, extend_rule_text, _d_nom = extension_mm_para_bar_type(bar_type)
    extend_ft = _mm_to_ft(extend_each_side_mm)
    extend_ft_geom = 0.0 if bool(ignore_empotramientos) else float(extend_ft)
    if extend_rule_text and (not bool(ignore_empotramientos)):
        avisos.append(u"Extensión por lado: {0}".format(extend_rule_text))
    try:
        if (not bool(ignore_empotramientos)) and no_forcing and (not exact_bt) and delta_bt is not None:
            d_real_mm = _ft_to_mm(float(bar_type.BarNominalDiameter))
            avisos.append(
                u"Automático (objetivo Ø12 mm): tipo cercano Ø{0:.1f} mm (Δ {1:.1f} mm)."
                .format(d_real_mm, float(delta_bt))
            )
            avisos.append(
                u"Regla de empotramiento/traslapo aplicada con Ø12 objetivo."
            )
    except Exception:
        pass

    solids = _collect_solids_from_host_element(host)
    layer_spacing_ft = _mm_to_ft(LAYER_SPACING_MM_DEFAULT)
    first_layer_extra_ft = _mm_to_ft(SHAFT_FIRST_LAYER_LATERAL_EXTRA_MM_DEFAULT)
    max_layer_extra_sci = float(first_layer_extra_ft)
    sci_layer_extra_ft = max_layer_extra_sci
    z_plano = _z_plano_para_detail_curves(view)

    pending = []
    for i, rf in enumerate(refs):
        try:
            face = host.GetGeometryObjectFromReference(rf)
        except Exception:
            face = None
        if not isinstance(face, PlanarFace):
            avisos.append(u"Cara [{0}]: no PlanarFace.".format(i + 1))
            continue
        (fk, segments), err = _horizontal_offset_segments_for_face(
            face,
            cover_ft,
            None,
            host,
            bar_type,
            solids,
            stretch_ft_for_sci=extend_ft_geom,
            layer_extra_ft_for_sci=sci_layer_extra_ft,
        )
        if err:
            avisos.append(u"Cara [{0}]: {1}".format(i + 1, err))
            continue
        pending.append((i + 1, rf, face, fk, segments))

    if not pending:
        return 0, avisos, u"No hay tramos para dibujar (revisar caras o geometría)."

    if SHAFT_SYMMETRIZE_ALIGN_RECT4:
        _symmetrize_pending_shaft_segments_if_rect4(pending, avisos)
        _align_pending_hv_midpoints_rect4(pending, avisos)
    _normalize_pending_segments_direction(pending, avisos)

    creadas = 0
    fallos = 0
    t = None
    if bool(use_transaction):
        t = Transaction(doc, u"BIMTools: shaft — tramos (DetailCurve)")
        t.Start()
    try:
        for face_1based, rf_face, face, fk, segments in pending:
            for q0, q1 in segments:
                try:
                    q0e, q1e = _shaft_extend_then_offset_into_host(
                        q0,
                        q1,
                        face,
                        host,
                        extend_ft_geom,
                        cover_axis_ft,
                        first_layer_extra_ft,
                        solids,
                        vertical_cover_ft=cover_vertical_ft,
                    )
                    q0e, q1e, _, _ = _round_up_segment_length_symmetric_about_mid(q0e, q1e)
                except Exception:
                    fallos += 1
                    continue
                try:
                    ln = Line.CreateBound(q0e, q1e)
                except Exception:
                    fallos += 1
                    continue
                ok = False
                try:
                    doc.Create.NewDetailCurve(view, ln)
                    ok = True
                    creadas += 1
                except Exception:
                    pass
                if not ok:
                    try:
                        ln_p = Line.CreateBound(
                            XYZ(float(q0e.X), float(q0e.Y), z_plano),
                            XYZ(float(q1e.X), float(q1e.Y), z_plano),
                        )
                        doc.Create.NewDetailCurve(view, ln_p)
                        creadas += 1
                        ok = True
                    except Exception:
                        fallos += 1
        if t is not None:
            t.Commit()
    except Exception as ex:
        if t is not None:
            t.RollBack()
        return 0, avisos, u"Error al crear DetailCurves:\n{0}".format(ex)

    if fallos and not creadas:
        avisos.append(
            u"Ninguna DetailCurve en esta vista (prueba planta o sección; en 3D suele fallar)."
        )
    elif fallos:
        avisos.append(
            u"{0} tramo(s) no se dibujaron (vista o plano).".format(fallos)
        )

    return creadas, avisos, None


def crear_enfierrado_shaft_hashtag(
    doc,
    host,
    refs,
    cover_mm=None,
    duplex_spacing_mm=None,
    n_capas=1,
    n_barras_set=2,
    forced_bar_type_id=None,
    max_bar_length_mm=12000.0,
    lap_length_mm=None,
    tag_view=None,
    tag_family_name=u"EST_A_STRUCTURAL REBAR TAG",
    place_lap_details=False,
    lap_detail_view=None,
    lap_detail_symbol_id=None,
    hook_type_name=None,
    ignore_empotramientos=True,
    use_transaction=True,
    empotramiento_adaptivo_extremos=False,
    normalize_parallel_segment_direction=True,
    apply_armadura_largo_total=False,
    hook_orientation_from_face_normal=False,
):
    """
    Por cada cara válida: todos los tramos inferiores válidos (perímetro exterior de la cara)
    → Rebar.CreateFromCurves sin ganchos (un Rebar por tramo).
    Misma geometría base que crear_detail_curves_tramos_shaft_hashtag (capa 0, antes de partir por L máx.).

    Flujo resumido (cara → barras):
    1) Cara ``PlanarFace`` y tramos inferiores con offset (recubrimiento + Ø/2) vía
       ``_horizontal_offset_segments_for_face``.
    2) Estirón / empotramiento en extremos y evaluación de colisión (SCI, modo adaptivo).
    3) Decisión de ganchos por canto; troceo por longitud máx. y traslapes entre subtramos.
    4) Creación: con ``hook_orientation_from_face_normal``, la referencia es **-FaceNormal** (3D);
       se proyecta ⟂ tangente de la barra para el plano de ``CreateFromCurves``; en **ambos**
       extremos, ``RebarHookOrientation`` usa ese ``nvec`` y la referencia en planta **-n_xy**
       hacia el interior.

    Por defecto no se consideran empotramientos (extensiones por anclaje según Ø ni cotas de
    empotramiento). Para el comportamiento anterior, pase ignore_empotramientos=False.

    duplex_spacing_mm: separación del set para Fixed Number (N barras).
    n_barras_set: cantidad de barras del set (Fixed Number). Si <=1, deja barra simple.

    max_bar_length_mm: paso de troceo (subtramo) cuando la longitud del tramo supera
    MAX_SINGLE_BAR_LENGTH_MM (12 m). No define el umbral de partición: tramos ≤12 m quedan
    en un solo Rebar. Con ganchos, cada subtramo respeta el cupo (recto + patas); el traslapo
    entre subtramos conserva el nominal (overlap).

    lap_length_mm: si no es None, longitud de traslape/empalme (mm) entre subtramos al
    partir barras >12 m; sustituye el solape que antes se tomaba de ``extend_each_side_mm``
    (anclaje/empotramiento). Las reglas de empotramiento en extremos siguen usando
    ``extend_each_side_mm``.

    Returns:
        (creados:int, tags:int, rebar_ids:list[ElementId], avisos:list, err:unicode|None)
        — err no nulo si no se hizo Commit. ``rebar_ids`` solo incluye la última capa (retorno).
        Etiquetado: una ``IndependentTag`` por sucesión (cara/tramo/subtramo) con hosts adicionales
        vía ``AddReferences`` en capas inferiores cuando hay varias capas. Ganchos y Armadura_Largo
        Total recorren todas las capas.

    normalize_parallel_segment_direction: si True (defecto), ejecuta
    ``_normalize_pending_segments_direction`` (solo agrupa y cuenta tramos paralelos; no invierte
    el sentido de ningún segmento).

    apply_armadura_largo_total: si True, escribe ``Armadura_Largo Total`` como suma de los largos
    de tramo de forma (parámetros A, B, C… tipo longitud en la instancia), tras `Regenerate`,
    no como suma de curvas de eje.

    hook_orientation_from_face_normal: si True (p. ej. herramienta «Borde losa gancho y
    empotramiento»), la referencia de creación es **-FaceNormal** en 3D (``n_bar_create``, sin normalizar); la
    referencia en planta para las patas es **-n_xy**. Tras layout,
    ``_reapply_hook_orientations_after_layout`` reaplica el mismo criterio (mismo ``nvec`` en
    ambos extremos cuando la curva es perpendicular a esa normal).
    Tras crear, ``_apply_rebar_hook_rotation_parameters_degrees`` escribe
    ``SHAFT_REBAR_HOOK_ROTATION_PARAM_DEGREES`` en los parámetros de rotación de gancho por extremo
    que tenga hook (BuiltIn ``REBAR_HOOK_ROTATION_AT_START`` / ``_AT_END``).
    """
    avisos = []
    if doc is None or host is None or not refs:
        return 0, 0, [], [], u"No hay documento, host o caras."

    if not isinstance(host, Floor):
        return 0, 0, [], [], u"El host no es una losa (Floor). Id: {}".format(_eid_int(host.Id))

    cov = float(cover_mm if cover_mm is not None else COVER_MM_DEFAULT)
    cover_ft = _mm_to_ft(cov)
    try:
        layer_count = int(n_capas)
    except Exception:
        layer_count = 1
    layer_count = max(1, layer_count)
    layer_spacing_ft = _mm_to_ft(LAYER_SPACING_MM_DEFAULT)
    first_layer_extra_ft = _mm_to_ft(SHAFT_FIRST_LAYER_LATERAL_EXTRA_MM_DEFAULT)
    try:
        n_barras = int(n_barras_set)
    except Exception:
        n_barras = 2
    n_barras = max(1, min(5, n_barras))

    bar_type, exact_bt, delta_bt = _resolve_shaft_bar_type(doc, forced_bar_type_id)
    if bar_type is None:
        return 0, 0, [], [], u"No hay ningún RebarBarType compatible en el proyecto."

    # #region reserve_space_for_hook
    # Si esta herramienta trabaja "sin empotramientos" (bordes de losa),
    # el hook puede sobresalir del host porque la curva entregada por (q0,q1)
    # no reserva espacio para la extensión del propio gancho.
    # Reservamos espacio aumentando el recubrimiento efectivo en una cantidad igual
    # a la longitud de hook para el bar_type.
    #
    # `RebarHookType` se resuelve una sola vez aquí (no en un segundo recorrido
    # más abajo). Los fallos de GetHookLength o del cap por bbox no deben borrar
    # la resolución del tipo (antes un try/except amplio dejaba hook_type_pre=None
    # y se volvía a buscar el gancho, con posible desajuste respecto al primero).
    hook_type_resolved = None
    if hook_type_name:
        try:
            hook_type_resolved = _rebar_hook_type_by_name(doc, hook_type_name)
        except Exception:
            hook_type_resolved = None

    hook_extra_cover_ft = 0.0
    if bool(ignore_empotramientos) and hook_type_resolved is not None:
        try:
            hlen_ft = float(bar_type.GetHookLength(hook_type_resolved.Id) or 0.0)
        except Exception:
            hlen_ft = 0.0
        hook_extra_cover_ft = max(0.0, hlen_ft)
        try:
            bb = host.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is not None:
            try:
                host_thk_ft = float(bb.Max.Z - bb.Min.Z)
                max_extra = max(0.0, 0.5 * host_thk_ft - float(cover_ft))
                hook_extra_cover_ft = min(float(hook_extra_cover_ft), float(max_extra))
            except Exception:
                pass
    # #endregion

    # `cover_ft` se interpreta como "cara a superficie". Como Revit usa la curva como
    # eje/centro de la barra, para offset a eje sumamos Ø/2.
    # La longitud de gancho (hook_extra_cover_ft) no se suma al offset XY desde la cara:
    # solo ajuste vertical / consumo de espesor; de lo contrario la barra queda ~100 mm
    # más hacia el interior en planta (p. ej. ~131 mm en vez de ~31 mm con cover 25 y Ø12).
    bar_diam_ft = float(_bar_nominal_diameter_ft(bar_type) or 0.0)
    cover_axis_ft = float(cover_ft) + 0.5 * max(bar_diam_ft, 1e-6)
    cover_vertical_ft = float(cover_axis_ft) + float(hook_extra_cover_ft)

    no_forcing = (
        forced_bar_type_id is None or forced_bar_type_id == ElementId.InvalidElementId
    )
    if bool(ignore_empotramientos):
        # Herramienta bordes de losa: no consideramos empotramientos/cotas,
        # pero SÍ necesitamos traslapo/solape por longitud cuando se divide.
        # Por eso calculamos la extensión por diámetro, pero sin aplicar reglas
        # de empotramiento/cotas (guardadas con `ignore_empotramientos` mas abajo).
        if no_forcing:
            extend_each_side_mm, extend_rule_text = extension_mm_por_diametro_nominal_mm(12.0)
            _d_nom = 12.0
        else:
            extend_each_side_mm, extend_rule_text, _d_nom = extension_mm_para_bar_type(bar_type)
    elif no_forcing:
        # Cuando el usuario deja "Automático ø12", la regla de empotramiento/traslapo
        # debe responder a Ø12 exacto (860 mm/lado), aunque el tipo real más cercano
        # del proyecto no sea exactamente 12.0 mm.
        extend_each_side_mm, extend_rule_text = extension_mm_por_diametro_nominal_mm(12.0)
        _d_nom = 12.0
    else:
        extend_each_side_mm, extend_rule_text, _d_nom = extension_mm_para_bar_type(bar_type)
    extend_ft = _mm_to_ft(extend_each_side_mm)
    _overlap_split_mm = float(extend_each_side_mm)
    if lap_length_mm is not None:
        try:
            _overlap_split_mm = max(0.0, float(lap_length_mm))
        except Exception:
            _overlap_split_mm = float(extend_each_side_mm)
    # Estirón en extremos sobre la cara (anclaje/«empotramiento» en planta). Si no se
    # consideran empotramientos, no se alarga el tramo aquí. El solape al partir barras
    # >12 m usa ``_overlap_split_mm`` (``lap_length_mm`` si viene informado; si no,
    # el mismo criterio histórico basado en extensión por Ø).
    extend_ft_geom = 0.0 if bool(ignore_empotramientos) else float(extend_ft)
    if extend_rule_text and (not bool(ignore_empotramientos)):
        avisos.append(u"Extensión por lado: {0}".format(extend_rule_text))
    try:
        if (not bool(ignore_empotramientos)) and no_forcing and (not exact_bt) and delta_bt is not None:
            d_real_mm = _ft_to_mm(float(bar_type.BarNominalDiameter))
            avisos.append(
                u"Automático (objetivo Ø12 mm): tipo cercano Ø{0:.1f} mm (Δ {1:.1f} mm)."
                .format(d_real_mm, float(delta_bt))
            )
            avisos.append(
                u"Regla de empotramiento/traslapo aplicada con Ø12 objetivo."
            )
    except Exception:
        pass

    solids = _collect_solids_from_host_element(host)
    spacing_mm = float(duplex_spacing_mm if duplex_spacing_mm is not None else DUPLEX_SPACING_MM_DEFAULT)
    spacing_ft = max(_mm_to_ft(spacing_mm), _bar_nominal_diameter_ft(bar_type), _mm_to_ft(5.0))
    layout_len_ft = _layout_length_from_host_thickness(host, cover_vertical_ft, spacing_ft)

    max_layer_extra_sci = float(first_layer_extra_ft) + float(max(0, layer_count - 1)) * float(
        layer_spacing_ft
    )
    sci_layer_extra_ft = max_layer_extra_sci

    pending = []
    face_infos = []

    for i, rf in enumerate(refs):
        try:
            face = host.GetGeometryObjectFromReference(rf)
        except Exception:
            face = None
        if not isinstance(face, PlanarFace):
            avisos.append(u"Cara [{0}]: no PlanarFace.".format(i + 1))
            continue
        # Guardar referencia original y normal en planta para cotas de empotramiento
        try:
            nxy = _face_normal_xy(face)
        except Exception:
            nxy = None
        # Preferir Reference nativa de la geometría de la cara (más estable para Dimension)
        # antes que la Reference capturada por el Pick (rf), que a veces arrastra refs extra.
        face_ref = rf
        try:
            fr = getattr(face, "Reference", None)
            if fr is not None:
                face_ref = fr
        except Exception:
            face_ref = rf

        inward_dir = None
        if nxy is not None:
            try:
                probe_ft = _mm_to_ft(50.0)
                sgn_in = _infer_inward_xy_sign(_face_origin(face), nxy, solids or [], probe_ft)
            except Exception:
                sgn_in = None
            if sgn_in is None:
                sgn_in = -1.0
            try:
                inward_dir = _xyz_scale(nxy, float(sgn_in))
            except Exception:
                inward_dir = None
        n_bar_create = _face_normal_3d(face)
        face_infos.append(
            {
                "ref": face_ref,
                "face": face,
                "nxy": nxy,
                "inward_dir": inward_dir,
                "n_bar_create": n_bar_create,
            }
        )
        (fk, segments), err = _horizontal_offset_segments_for_face(
            face,
            cover_ft,
            None,
            host,
            bar_type,
            solids,
            stretch_ft_for_sci=extend_ft_geom,
            layer_extra_ft_for_sci=sci_layer_extra_ft,
            adaptive_embed_end_clip=bool(empotramiento_adaptivo_extremos),
        )
        if err:
            avisos.append(u"Cara [{0}]: {1}".format(i + 1, err))
            continue
        pending.append((i + 1, rf, face, fk, segments))

    if not pending:
        return 0, 0, [], avisos, u"No hay ninguna cara válida para crear armadura."

    if SHAFT_SYMMETRIZE_ALIGN_RECT4:
        _symmetrize_pending_shaft_segments_if_rect4(pending, avisos)
        _align_pending_hv_midpoints_rect4(pending, avisos)
    if bool(normalize_parallel_segment_direction):
        _normalize_pending_segments_direction(pending, avisos)

    creados = 0
    tags_creadas = 0
    # Solo última capa: retorno de ids (compat. con código que consume el resultado).
    rebar_ids = []
    # Todas las capas: barrido de ganchos y Armadura_Largo Total.
    rebar_ids_all = []
    # Barras creadas vía curve loop (patas como tramos de curva, sin RebarHookType API).
    _curve_loop_ids = set()
    # Misma (cara, tramo, split) × varias capas → una etiqueta multihost (AddReferences).
    multihost_groups = defaultdict(list)
    ok_tag_view, msg_tag_view = _view_ok_for_rebar_tags(tag_view) if tag_view is not None else (False, None)
    tag_map = {}
    tag_enabled = False
    if tag_view is not None:
        if ok_tag_view:
            tag_map = _collect_rebar_tag_symbol_map(doc, tag_family_name)
            if tag_map:
                tag_enabled = True
                if int(layer_count) > 1 and (
                    not _rebar_tag_family_exists_strict(doc, tag_family_name)
                ):
                    avisos.append(
                        u"Etiquetado: no hay familia con nombre exacto '{0}'. "
                        u"Se usó otra por coincidencia; la etiqueta puede mostrar solo la cantidad del set "
                        u"(p. ej. 2ø12) y no todas las capas. Cargue la familia multi-capa en el proyecto "
                        u"(p. ej. EST_A_STRUCTURAL REBAR TAG_TRIPLE QUANTITY)."
                        .format(tag_family_name)
                    )
            else:
                avisos.append(u"Etiquetado: no se encontraron tipos en familia '{}'.".format(tag_family_name))
        elif msg_tag_view:
            avisos.append(msg_tag_view)
    hook_type = None
    if hook_type_name:
        hook_type = hook_type_resolved
        if hook_type is None:
            # Mensaje mas util cuando el nombre no coincide (p. ej. deg. vs °, puntuacion, etc.).
            try:
                hook_types = list(FilteredElementCollector(doc).OfClass(RebarHookType))
            except Exception:
                hook_types = []
            def _norm_cand(s):
                try:
                    t = unicode(s or u"")
                except Exception:
                    t = u""
                try:
                    t = t.replace(u"\u00A0", u" ")
                except Exception:
                    pass
                t = t.strip().lower()
                if not t:
                    return u""
                t = t.replace(u"º", u"°")
                t = t.replace(u"°", u" deg ")
                for ch in (u".", u",", u"-", u"(", u")", u"/"):
                    t = t.replace(ch, u" ")
                parts = [p for p in t.split() if p]
                return u" ".join(parts)

            cand = []
            for ht in hook_types:
                try:
                    nm_norm = _norm_cand(getattr(ht, "Name", None))
                except Exception:
                    nm_norm = u""
                if (u"90" in nm_norm) and (u"deg" in nm_norm):
                    try:
                        nm = unicode(ht.Name or u"").strip()
                    except Exception:
                        nm = u""
                    if nm:
                        cand.append(nm)
                    if len(cand) >= 10:
                        break
            if not cand:
                # Fallback: primeros 10 nombres recortados.
                try:
                    cand = [unicode(ht.Name or u"").strip() for ht in hook_types[:10]]
                except Exception:
                    cand = []
            cand_txt = u", ".join([c for c in cand if c]) if cand else u"(sin candidatos)"

            return (
                0,
                0,
                [],
                avisos,
                u"No se encontro RebarHookType '{0}'. Hooks candidatos: {1}.".format(
                    hook_type_name,
                    cand_txt,
                ),
            )

    lap_detail_enabled = bool(place_lap_details)
    lap_detail_symbol = None
    if lap_detail_enabled:
        if lap_detail_view is None:
            avisos.append(
                u"Detalles de empalme: no hay vista activa; se omitió la colocación."
            )
            lap_detail_enabled = False
        elif lap_detail_symbol_id is None or lap_detail_symbol_id == ElementId.InvalidElementId:
            avisos.append(
                u"Detalles de empalme: no se seleccionó un Detail Component válido."
            )
            lap_detail_enabled = False
        else:
            try:
                lap_detail_symbol = doc.GetElement(lap_detail_symbol_id)
            except Exception:
                lap_detail_symbol = None
            if not isinstance(lap_detail_symbol, FamilySymbol):
                avisos.append(
                    u"Detalles de empalme: el tipo seleccionado no es un FamilySymbol válido."
                )
                lap_detail_enabled = False

    lap_details_created = 0
    lap_dims_created = 0
    lap_dim_skip_no_detail = 0
    lap_dim_skip_ref = 0
    lap_dim_skip_create = 0
    t = None
    if bool(use_transaction):
        t = Transaction(doc, u"BIMTools: enfierrado shaft (barras)")
        t.Start()
    try:
        # Cotas: solo en planta (según preferencia plan_only)
        dim_enabled = bool(tag_view is not None and _view_is_plan_only(tag_view))
        if (not bool(ignore_empotramientos)) and tag_view is not None and (not dim_enabled):
            avisos.append(u"Cotas empotramiento: solo se crean en vistas de planta.")
            avisos.append(u"Cotas traslapo: solo se crean en vistas de planta.")

        if float(SHAFT_DEBUG_FIRST_BOTTOM_EDGE_OFFSET_MM) > 1e-6 and tag_view is not None:
            for face_1based_dbg, _rf_f, face_dbg, _fk, _segs in pending:
                try:
                    if not _segs:
                        continue
                    dbg_q0, dbg_q1 = _segs[0]
                    _detail_line_first_bottom_curve_debug(
                        doc,
                        tag_view,
                        face_dbg,
                        dbg_q0,
                        dbg_q1,
                        float(SHAFT_DEBUG_FIRST_BOTTOM_EDGE_OFFSET_MM),
                        solids,
                        float(extend_ft_geom),
                        avisos,
                        face_1based_dbg,
                    )
                except Exception:
                    pass

        for face_1based, rf_face, face, fk, segments in pending:
            ok_face = 0
            fail_face = 0
            for seg_i, (q0, q1) in enumerate(segments):
                for k in range(layer_count):
                    # Capa 0: extra lateral mínimo (mm) + k× separación entre capas (100 mm por defecto).
                    layer_ex = float(first_layer_extra_ft) + float(k) * float(layer_spacing_ft)
                    stretch_kept0, stretch_kept1 = True, True
                    if bool(empotramiento_adaptivo_extremos) and extend_ft_geom > 1e-12:
                        q0e, q1e, stretch_kept0, stretch_kept1 = _shaft_probe_hit_asymmetric_extend_then_offset_into_host(
                            q0,
                            q1,
                            face,
                            host,
                            extend_ft_geom,
                            cover_axis_ft,
                            layer_ex,
                            solids,
                            vertical_cover_ft=cover_vertical_ft,
                            embed_clip_avisos=(avisos if int(k) == 0 else None),
                            clip_label=u"Cara {0}".format(face_1based),
                        )
                    else:
                        q0e, q1e = _shaft_extend_then_offset_into_host(
                            q0,
                            q1,
                            face,
                            host,
                            extend_ft_geom,
                            cover_axis_ft,
                            layer_ex,
                            solids,
                            vertical_cover_ft=cover_vertical_ft,
                        )
                    # Redondeo de largo (solo tramo recto, sin pata de gancho): tras offset al eje,
                    # subir al múltiplo de SHAFT_BAR_LENGTH_ROUND_STEP_MM. Con empotramiento adaptivo,
                    # el incremento solo en el extremo que retuvo estirón SCI; si no, simétrico.
                    # Para barras con gancho de curve loop (hook_orientation_from_face_normal=True)
                    # se omite el redondeo en (q0e,q1e) salvo un caso: empotramiento SCI en **ambos**
                    # extremos y tramo sin troceo (<12 m, sin empalmes). Ahí hace falta el mismo
                    # estirón que ``_round_up_segment_length_adaptive_embed_end`` (delta en q1) para
                    # cerrar el eje en múltiplo de 10 mm antes de ``Line.CreateBound`` / curve loop.
                    if bool(hook_orientation_from_face_normal):
                        _rounded, _delta_mm = False, 0.0
                        if (
                            bool(empotramiento_adaptivo_extremos)
                            and extend_ft_geom > 1e-12
                            and bool(stretch_kept0)
                            and bool(stretch_kept1)
                        ):
                            _len_ft_pre = _segment_length_ft(q0e, q1e)
                            if _len_ft_pre <= _mm_to_ft(float(MAX_SINGLE_BAR_LENGTH_MM)) + 1e-9:
                                q0e, q1e, _rounded, _delta_mm = (
                                    _round_up_segment_length_adaptive_embed_end(
                                        q0e,
                                        q1e,
                                        stretch_kept0,
                                        stretch_kept1,
                                    )
                                )
                    elif bool(empotramiento_adaptivo_extremos) and extend_ft_geom > 1e-12:
                        q0e, q1e, _rounded, _delta_mm = _round_up_segment_length_adaptive_embed_end(
                            q0e, q1e, stretch_kept0, stretch_kept1
                        )
                    else:
                        q0e, q1e, _rounded, _delta_mm = _round_up_segment_length_symmetric_about_mid(
                            q0e, q1e
                        )
                    hook_dev_mm = 0.0
                    if hook_type is not None:
                        hook_dev_mm = float(hook_leg_mm_para_bar_type(bar_type))
                    # Descuento de pata en cupo (primer tramo): solo si ningún extremo tiene
                    # colisión SCI con el host. Con colisión en al menos un extremo, largo
                    # nominal completo. Sin modo adaptivo no hay sonda por canto: se aplica
                    # el criterio clásico (sí descontar).
                    _adaptive_sci = bool(empotramiento_adaptivo_extremos) and extend_ft_geom > 1e-12
                    if _adaptive_sci:
                        _discount_pata_cupo = (not bool(stretch_kept0)) and (
                            not bool(stretch_kept1)
                        )
                    else:
                        _discount_pata_cupo = True
                    try:
                        _chunk_mm = float(
                            max_bar_length_mm
                            if max_bar_length_mm is not None
                            else LONG_BAR_SPLIT_CHUNK_MM
                        )
                    except Exception:
                        _chunk_mm = float(LONG_BAR_SPLIT_CHUNK_MM)
                    _chunk_mm = max(100.0, min(_chunk_mm, float(MAX_SINGLE_BAR_LENGTH_MM)))
                    # Trocear solo si este tramo supera el límite de barra recta (12 m). El
                    # valor del formulario es el paso de troceo, no el umbral; así las caras
                    # cortas no reciben traslapos cuando otra cara larga activó el modo.
                    _need_split = _segment_length_ft(q0e, q1e) > _mm_to_ft(
                        float(MAX_SINGLE_BAR_LENGTH_MM)
                    ) + 1e-9
                    _hf_split = (
                        hook_dev_mm
                        if (
                            hook_dev_mm > 0
                            and _need_split
                            and hook_type is not None
                            and _discount_pata_cupo
                        )
                        else 0.0
                    )
                    try:
                        split_lines, split_joints = _split_long_bar_with_overlap_segments(
                            q0e,
                            q1e,
                            float(MAX_SINGLE_BAR_LENGTH_MM),
                            _chunk_mm,
                            _overlap_split_mm,
                            _hf_split,
                            0.0,
                        )
                    except Exception:
                        split_lines = [(q0e, q1e)]
                        split_joints = []
                    split_rebars = []
                    ln_first = None
                    seg_count = int(len(split_lines)) if split_lines else 1

                    for sl_i, (sq0, sq1) in enumerate(split_lines):
                        try:
                            # Sentido natural del subtramo (sin invertir sq0/sq1) para
                            # CreateFromCurves y orientación de ganchos respecto a la normal.
                            ln = Line.CreateBound(sq0, sq1)
                        except Exception:
                            fail_face += 1
                            continue
                        if ln_first is None:
                            ln_first = ln
                        hook_end0 = False
                        hook_end1 = False
                        ln_work = ln
                        _created_by_curve_loop = False
                        if hook_type is not None:
                            # Troceo largo: primer subtramo → gancho al inicio de curva; último → al final.
                            # - Primer segmento: gancho al INICIO (hook_end0)
                            # - Último segmento: gancho al FINAL (hook_end1)
                            hook_end0 = bool(sl_i == 0)
                            hook_end1 = bool(sl_i == (seg_count - 1))
                            if (
                                bool(empotramiento_adaptivo_extremos)
                                and extend_ft_geom > 1e-12
                                and hook_type is not None
                            ):
                                pre_hook0, pre_hook1 = bool(hook_end0), bool(hook_end1)
                                # Gancho solo en extremo sin retención de estirón SCI (canto
                                # libre: sonda sin hormigón en la prolongación). Quitar gancho
                                # donde el canto global retuvo empotramiento (colisión).
                                if hook_end0 and _stretch_kept_at_ln_api_end(
                                    ln_work,
                                    0,
                                    q0e,
                                    q1e,
                                    stretch_kept0,
                                    stretch_kept1,
                                ):
                                    hook_end0 = False
                                if hook_end1 and _stretch_kept_at_ln_api_end(
                                    ln_work,
                                    1,
                                    q0e,
                                    q1e,
                                    stretch_kept0,
                                    stretch_kept1,
                                ):
                                    hook_end1 = False
                                if (
                                    (not hook_end0)
                                    and (not hook_end1)
                                    and (pre_hook0 or pre_hook1)
                                    and int(k) == 0
                                ):
                                    try:
                                        avisos.append(
                                            u"Cara {0}: ganchos omitidos (ambos cantos retienen estirón SCI)."
                                            .format(face_1based)
                                        )
                                    except Exception:
                                        pass
                            # Eje definitivo para Rebar: misma Line para curve loop y fallback recto.
                            if bool(hook_orientation_from_face_normal) and (
                                hook_end0 or hook_end1
                            ):
                                try:
                                    _d_half_mm = 0.0
                                    try:
                                        _d_half_mm = (
                                            _ft_to_mm(float(bar_type.BarNominalDiameter))
                                            / 2.0
                                        )
                                    except Exception:
                                        _d_half_mm = 0.0
                                    _lnr = _round_line_axis_length_mm_ceil(
                                        ln_work,
                                        hook_end0,
                                        hook_end1,
                                        d_half_mm=_d_half_mm,
                                    )
                                    if _lnr is not None:
                                        ln_work = _lnr
                                except Exception:
                                    pass
                            _face_inward_for_create = None
                            _face_bar_normal_3d = None
                            if bool(hook_orientation_from_face_normal):
                                try:
                                    _ix_fc = int(face_1based) - 1
                                    if 0 <= _ix_fc < len(face_infos):
                                        _fi = face_infos[_ix_fc]
                                        # Referencia en planta (-n_xy) para patas hacia interior.
                                        _face_inward_for_create = _inverse_face_normal_xy(
                                            _fi.get("nxy")
                                        )
                                        # -FaceNormal 3D sin normalizar (misma referencia plano + ganchos).
                                        _face_bar_normal_3d = _fi.get("n_bar_create")
                                except Exception:
                                    _face_inward_for_create = None
                                    _face_bar_normal_3d = None
                            _bordes_hi_create = bool(ignore_empotramientos) or (
                                bool(empotramiento_adaptivo_extremos)
                                and extend_ft_geom > 1e-12
                            )
                            # Modo hook_orientation_from_face_normal: patas siempre como tramos
                            # de curva (sin RebarHookType en la API — alineado con FA).
                            # Si el curve loop falla, la barra se crea recta (sin gancho API).
                            r = None
                            if bool(hook_orientation_from_face_normal):
                                if (
                                    _face_inward_for_create is not None
                                    and (hook_end0 or hook_end1)
                                ):
                                    r = _create_rebar_borde_losa_curve_loop(
                                        doc,
                                        host,
                                        bar_type,
                                        ln_work,
                                        face_inward_xy=_face_inward_for_create,
                                        hook_end0=hook_end0,
                                        hook_end1=hook_end1,
                                        bar_plane_normal_3d=_face_bar_normal_3d,
                                    )
                                    if r is not None:
                                        _created_by_curve_loop = True
                                if r is None:
                                    # Curve loop rechazado por Revit: barra recta sin gancho.
                                    # No se usa RebarHookType en este modo.
                                    if hook_end0 or hook_end1:
                                        avisos.append(
                                            u"Cara [{0}] capa {1}: curve loop rechazado por Revit; "
                                            u"barra sin gancho (sin RebarHookType).".format(
                                                face_1based, k + 1
                                            )
                                        )
                                    r = _create_single_rebar_from_line_no_hooks(
                                        doc, host, bar_type, ln_work
                                    )
                            else:
                                # Modo sin face_normal: camino original con RebarHookType.
                                if hook_end0 and hook_end1:
                                    r = _create_single_rebar_from_line_with_hooks(
                                        doc,
                                        host,
                                        bar_type,
                                        ln_work,
                                        hook_type,
                                    )
                                elif hook_end0 or hook_end1:
                                    r = _create_single_rebar_from_line_with_partial_hooks(
                                        doc,
                                        host,
                                        bar_type,
                                        ln_work,
                                        hook_type,
                                        hook_end0,
                                        hook_end1,
                                        avisos=avisos,
                                    )
                                else:
                                    r = _create_single_rebar_from_line_no_hooks(
                                        doc, host, bar_type, ln_work
                                    )
                        else:
                            r = _create_single_rebar_from_line_no_hooks(doc, host, bar_type, ln)
                        if r:
                            if _created_by_curve_loop:
                                try:
                                    _curve_loop_ids.add(r.Id)
                                except Exception:
                                    pass
                            _mh_key = (int(face_1based), int(seg_i), int(sl_i))
                            if _apply_fixed_number_layout_n_bars(r, n_barras, layout_len_ft):
                                creados += int(n_barras)
                                ok_face += 1
                                try:
                                    rebar_ids_all.append(r.Id)
                                    multihost_groups[_mh_key].append((int(k), r.Id))
                                    if k == (layer_count - 1):
                                        rebar_ids.append(r.Id)
                                except Exception:
                                    pass
                            else:
                                # Si no se pudo setear el layout, se conserva la barra simple creada.
                                creados += 1
                                ok_face += 1
                                try:
                                    rebar_ids_all.append(r.Id)
                                    multihost_groups[_mh_key].append((int(k), r.Id))
                                    if k == (layer_count - 1):
                                        rebar_ids.append(r.Id)
                                except Exception:
                                    pass
                                avisos.append(
                                    u"Cara [{0}] capa {1}: no se pudo aplicar Fixed Number={2}; se dejó barra simple."
                                    .format(face_1based, k + 1, int(n_barras))
                                )
                            try:
                                if hook_type is not None and not bool(hook_orientation_from_face_normal):
                                    if hook_end0 and hook_end1:
                                        _reapply_both_hooks_after_fixed_number_layout(
                                            r, hook_type, avisos=avisos
                                        )
                                    elif (bool(hook_end0) != bool(hook_end1)) and (
                                        bool(hook_end0) or bool(hook_end1)
                                    ):
                                        _reapply_partial_hooks_after_fixed_number_layout(
                                            r, hook_type, hook_end0, hook_end1, avisos=avisos
                                        )
                                    if hook_end0 or hook_end1:
                                        # Misma geometría que «Barras bordes de losa» (sin empotramiento o
                                        # empotramiento adaptivo): piernas 90° hacia interior.
                                        bordes_hacia_interior = bool(ignore_empotramientos) or (
                                            bool(empotramiento_adaptivo_extremos)
                                            and extend_ft_geom > 1e-12
                                        )
                                        _reapply_hook_orientations_after_layout(
                                            r,
                                            hook_end0,
                                            hook_end1,
                                            bordes_losa_hook_hacia_interior=bool(
                                                bordes_hacia_interior
                                            ),
                                            inverse_normal_xy_for_hooks=(
                                                _face_inward_for_create
                                                if bool(hook_orientation_from_face_normal)
                                                and (
                                                    (_face_inward_for_create is not None)
                                                    or (_face_bar_normal_3d is not None)
                                                )
                                                else None
                                            ),
                                            hook_curve_for_hooks=(
                                                ln_work
                                                if bool(hook_orientation_from_face_normal)
                                                and (
                                                    (_face_inward_for_create is not None)
                                                    or (_face_bar_normal_3d is not None)
                                                )
                                                else None
                                            ),
                                            bar_plane_normal_3d_for_hooks=(
                                                _face_bar_normal_3d
                                                if bool(hook_orientation_from_face_normal)
                                                and (
                                                    (_face_inward_for_create is not None)
                                                    or (_face_bar_normal_3d is not None)
                                                )
                                                else None
                                            ),
                                        )
                                    if hook_type_name and (hook_end0 or hook_end1):
                                        _enforce_rebar_hook_types_by_name(
                                            r,
                                            doc,
                                            hook_type_name,
                                            hook_end0,
                                            hook_end1,
                                            avisos=avisos,
                                        )
                            except Exception:
                                pass
                        else:
                            fail_face += 1
                        split_rebars.append(r)

                    # Detalles de empalme: por capa si hay split_joints. Cotas de traslapo solo en capa 0
                    # (primera capa); las cotas de empotramiento siguen solo en k==0 más abajo.
                    if lap_detail_enabled and split_joints:
                        for ji, jn in enumerate(split_joints):
                            if not bool(jn.get("has_lap")):
                                continue
                            ok_lap_detail, msg_lap_detail, lap_inst = _place_line_based_detail_component(
                                doc,
                                lap_detail_view,
                                lap_detail_symbol,
                                jn.get("lap_start"),
                                jn.get("lap_end"),
                            )
                            if ok_lap_detail:
                                lap_details_created += 1
                                dim_eid = None
                                if dim_enabled and int(k) == 0:
                                    ref_l, ref_r, _ref_err = _get_named_left_right_refs_from_detail_instance(
                                        lap_inst
                                    )
                                    if ref_l is None or ref_r is None:
                                        lap_dim_skip_ref += 1
                                    else:
                                        _inward_xy = None
                                        try:
                                            _fi_idx = int(face_1based) - 1
                                            if 0 <= _fi_idx < len(face_infos):
                                                _inward_xy = face_infos[_fi_idx].get(
                                                    "inward_dir"
                                                )
                                        except Exception:
                                            _inward_xy = None
                                        ok_lap_dim, _lap_dim_msg, _lap_dim_data = _create_overlap_dimension_from_detail_refs(
                                            doc,
                                            tag_view,
                                            ref_l,
                                            ref_r,
                                            jn.get("lap_start"),
                                            jn.get("lap_end"),
                                            jn.get("axis_u"),
                                            lateral_hint=_xyz_sub(q1e, q0e),
                                            line_offset_mm=450.0,
                                            inward_dir_xy=_inward_xy,
                                        )
                                        if ok_lap_dim:
                                            lap_dims_created += 1
                                            try:
                                                if (
                                                    _lap_dim_data
                                                    and _lap_dim_data.get("dim_id") is not None
                                                ):
                                                    dim_eid = ElementId(
                                                        int(_lap_dim_data["dim_id"])
                                                    )
                                            except Exception:
                                                dim_eid = None
                                        else:
                                            lap_dim_skip_create += 1
                                try:
                                    from lap_detail_link_schema import (
                                        set_lap_detail_rebar_link,
                                    )

                                    if (
                                        ji + 1 < len(split_rebars)
                                        and split_rebars[ji] is not None
                                        and split_rebars[ji + 1] is not None
                                    ):
                                        set_lap_detail_rebar_link(
                                            lap_inst,
                                            split_rebars[ji].Id,
                                            split_rebars[ji + 1].Id,
                                            dim_eid,
                                        )
                                except Exception:
                                    pass
                            else:
                                if dim_enabled:
                                    lap_dim_skip_no_detail += 1
                                avisos.append(
                                    u"Detalle de empalme: se omitió una instancia ({0}).".format(
                                        msg_lap_detail or u"sin detalle"
                                    )
                                )

                    # Cota empotramiento: solo una capa para evitar saturación (k==0),
                    # y una sola cota por tramo original (primer subtramo creado).
                    if (not bool(ignore_empotramientos)) and dim_enabled and k == 0 and split_rebars and split_rebars[0] is not None:
                        try:
                            # Dirección de barra en planta
                            dxy = _unit_xy(_xyz_sub(q1e, q0e))
                        except Exception:
                            dxy = None
                        if dxy is not None:
                            # Evaluar extremos según empotramiento por Ø (extend_each_side_mm)
                            end_idx, face_pick = _pick_end_for_anchorage(
                                q0e, q1e, dxy, face_infos, expected_mm=extend_each_side_mm, tol_mm=5.0
                            )
                            if face_pick is not None:
                                end_pt = q0e if int(end_idx) == 0 else q1e
                                face_ref = face_pick.get("ref")
                                face_nxy = face_pick.get("nxy")
                                face_o = None
                                try:
                                    face_o = _face_origin(face_pick.get("face"))
                                except Exception:
                                    face_o = None
                                base_on_plane, marker_pt = _marker_point_for_table_anchorage(
                                    end_pt,
                                    dxy,
                                    end_idx,
                                    face_o,
                                    face_nxy,
                                    anchorage_mm=extend_each_side_mm,
                                )
                                # Marcador MUY pequeño (5 mm). Para evitar que Revit "pegue"
                                # referencias de la barra, desplazamos el marcador levemente
                                # en la dirección tangente a la cara y validamos que la
                                # Dimension resultante tenga exactamente 2 referencias.
                                tdir = _rotate90_xy(face_nxy)
                                if tdir is None:
                                    tdir = XYZ(1.0, 0.0, 0.0)
                                ok_dim = False
                                for shift_mm in (2.0, 20.0, 50.0):
                                    try:
                                        marker_pt_shift = _xyz_add(
                                            marker_pt, _xyz_scale(tdir, _mm_to_ft(float(shift_mm)))
                                        )
                                    except Exception:
                                        marker_pt_shift = marker_pt
                                    dc, marker_ref = _create_marker_detailcurve(
                                        doc, tag_view, marker_pt_shift, face_nxy, length_mm=5.0
                                    )
                                    if marker_ref is None or face_ref is None:
                                        continue
                                    dim = _create_dimension_face_to_marker(
                                        doc,
                                        tag_view,
                                        face_ref,
                                        marker_ref,
                                        base_on_plane,
                                        face_nxy,
                                        face_obj=face_pick.get("face"),
                                        # Offset de la línea de cota hacia afuera para que
                                        # no quede pegada a la barra.
                                        outside_offset_mm=450.0,
                                        line_len_mm=450.0,
                                        solids=solids,
                                        dim_line_template=ln_first if ln_first is not None else ln,
                                    )
                                    nrefs = _dimension_reference_count(dim)
                                    # Deben ser EXACTAMENTE 2 referencias. Si Revit no permite leerlas
                                    # o agrega más, eliminamos y reintentamos con otro offset.
                                    if (nrefs is None) or (nrefs != 2):
                                        try:
                                            if dim is not None:
                                                doc.Delete(dim.Id)
                                        except Exception:
                                            pass
                                        # Borrar también el marcador para reintento limpio
                                        try:
                                            if dc is not None:
                                                doc.Delete(dc.Id)
                                        except Exception:
                                            pass
                                        continue
                                    if dim is not None:
                                        ok_dim = True
                                        try:
                                            st = None
                                            try:
                                                st = face_ref.ConvertToStableRepresentation(
                                                    doc
                                                )
                                            except Exception:
                                                st = None
                                            if st and dc is not None:
                                                from embed_anchorage_link_schema import (
                                                    set_embed_anchorage_link,
                                                )

                                                try:
                                                    ok_link = set_embed_anchorage_link(
                                                        dc,
                                                        split_rebars[0].Id,
                                                        dim.Id,
                                                        tag_view.Id,
                                                        st,
                                                        float(extend_each_side_mm),
                                                        int(end_idx),
                                                    )
                                                    if not ok_link:
                                                        avisos.append(
                                                            u"Cota empotramiento: no se guardó el vínculo en el marcador; "
                                                            u"la cota no se sincronizará al editar el armado."
                                                        )
                                                except Exception as ex:
                                                    avisos.append(
                                                        u"Cota empotramiento: error al guardar vínculo DMU: {0}".format(
                                                            ex
                                                        )
                                                    )
                                        except Exception:
                                            pass
                                        break
                                if not ok_dim:
                                    avisos.append(
                                        u"Cotas empotramiento: se omitió una cota porque Revit agregó referencias extra."
                                    )
            if ok_face < 1:
                avisos.append(
                    u"Cara [{0}]: CreateFromCurves no creó Rebar en ningún tramo/capa (revisar curva/host)."
                    .format(face_1based)
                )
            elif fail_face > 0:
                avisos.append(
                    u"Cara [{0}]: {1} intento(s) sin Rebar de {2} (tramos x capas)."
                    .format(face_1based, fail_face, len(segments) * layer_count)
                )

        if creados < 1:
            if t is not None:
                t.RollBack()
            return 0, 0, [], avisos, u"No se pudo crear ningún Rebar (revisar tipo y host)."

        rebar_groups_tag = []
        try:
            for _mh_k, pairs in multihost_groups.items():
                pairs_sorted = sorted(pairs, key=lambda x: int(x[0]))
                rebar_groups_tag.append([p[1] for p in pairs_sorted])
        except Exception:
            rebar_groups_tag = [[rid] for rid in rebar_ids]

        if tag_enabled and rebar_groups_tag:
            n_tag, avis_tag, err_tag = etiquetar_grupos_rebar_multihost_capas_en_vista(
                doc,
                tag_view,
                rebar_groups_tag,
                rebar_ids_all,
                family_name=tag_family_name,
                use_transaction=False,
            )
            tags_creadas += int(n_tag or 0)
            if avis_tag:
                avisos.extend(list(avis_tag))
            if err_tag:
                avisos.append(err_tag)
        if lap_detail_enabled:
            avisos.append(
                u"Detalles de empalme colocados: {0}.".format(int(lap_details_created))
            )
            if dim_enabled:
                avisos.append(
                    u"Cotas de empalme (refs detail): {0} creadas.".format(int(lap_dims_created))
                )
                if int(lap_dim_skip_no_detail) > 0:
                    avisos.append(
                        u"Cotas de empalme omitidas (sin detail): {0}.".format(
                            int(lap_dim_skip_no_detail)
                        )
                    )
                if int(lap_dim_skip_ref) > 0:
                    avisos.append(
                        u"Cotas de empalme omitidas (faltan refs Left/Right): {0}.".format(
                            int(lap_dim_skip_ref)
                        )
                    )
                if int(lap_dim_skip_create) > 0:
                    avisos.append(
                        u"Cotas de empalme omitidas (fallo al crear cota): {0}.".format(
                            int(lap_dim_skip_create)
                        )
                    )

        if hook_type_name and rebar_ids_all and not bool(hook_orientation_from_face_normal):
            try:
                _sweep_rebar_hook_types_to_name(doc, rebar_ids_all, hook_type_name, avisos=avisos)
            except Exception:
                pass

        if rebar_ids_all and not bool(hook_orientation_from_face_normal):
            try:
                _apply_rebar_hook_rotation_parameters_degrees(
                    doc,
                    rebar_ids_all,
                    SHAFT_REBAR_HOOK_ROTATION_PARAM_DEGREES,
                    avisos=avisos,
                )
            except Exception:
                pass

        if bool(apply_armadura_largo_total) and rebar_ids_all:
            try:
                _apply_armadura_largo_total_to_rebars(doc, rebar_ids_all, avisos)
            except Exception:
                pass

        if t is not None:
            t.Commit()
    except Exception as ex:
        if t is not None:
            t.RollBack()
        return 0, 0, [], avisos, u"Error en transacción:\n{0}".format(ex)

    return creados, tags_creadas, rebar_ids, avisos, None


def crear_enfierrado_bordes_losa_hook90(
    doc,
    host,
    refs,
    cover_mm=None,
    duplex_spacing_mm=None,
    n_capas=1,
    n_barras_set=2,
    forced_bar_type_id=None,
    max_bar_length_mm=12000.0,
    lap_length_mm=None,
    tag_view=None,
    tag_family_name=u"EST_A_STRUCTURAL REBAR TAG",
    place_lap_details=False,
    lap_detail_view=None,
    lap_detail_symbol_id=None,
):
    """Ruta de bordes de losa: hook fijo 90 en ambos extremos, sin empotramientos."""
    return crear_enfierrado_shaft_hashtag(
        doc,
        host,
        refs,
        cover_mm=cover_mm,
        duplex_spacing_mm=duplex_spacing_mm,
        n_capas=n_capas,
        n_barras_set=n_barras_set,
        forced_bar_type_id=forced_bar_type_id,
        max_bar_length_mm=max_bar_length_mm,
        lap_length_mm=lap_length_mm,
        tag_view=tag_view,
        tag_family_name=tag_family_name,
        place_lap_details=place_lap_details,
        lap_detail_view=lap_detail_view,
        lap_detail_symbol_id=lap_detail_symbol_id,
        hook_type_name=u"Standard - 90 deg.",
        ignore_empotramientos=True,
    )


def crear_enfierrado_bordes_losa_gancho_y_empotramiento(
    doc,
    host,
    refs,
    cover_mm=None,
    duplex_spacing_mm=None,
    n_capas=1,
    n_barras_set=2,
    forced_bar_type_id=None,
    max_bar_length_mm=12000.0,
    lap_length_mm=None,
    tag_view=None,
    tag_family_name=u"EST_A_STRUCTURAL REBAR TAG",
    place_lap_details=False,
    lap_detail_view=None,
    lap_detail_symbol_id=None,
):
    """Ruta de bordes de losa: hook 90° donde aplica, con extensiones/cotas por empotramiento."""
    return crear_enfierrado_shaft_hashtag(
        doc,
        host,
        refs,
        cover_mm=cover_mm,
        duplex_spacing_mm=duplex_spacing_mm,
        n_capas=n_capas,
        n_barras_set=n_barras_set,
        forced_bar_type_id=forced_bar_type_id,
        max_bar_length_mm=max_bar_length_mm,
        lap_length_mm=lap_length_mm,
        tag_view=tag_view,
        tag_family_name=tag_family_name,
        place_lap_details=place_lap_details,
        lap_detail_view=lap_detail_view,
        lap_detail_symbol_id=lap_detail_symbol_id,
        hook_type_name=u"Standard - 90 deg.",
        ignore_empotramientos=False,
        empotramiento_adaptivo_extremos=True,
        # No unificar sentido q0→q1 entre tramos paralelos: invertir curvas aquí rompe la
        # orientación prevista para ganchos 90° hacia interior del host.
        normalize_parallel_segment_direction=False,
        apply_armadura_largo_total=True,
        hook_orientation_from_face_normal=True,
    )


def run_pyrevit(revit, cover_mm=None, duplex_spacing_mm=None):
    import os

    try:
        _scripts_dir = os.path.dirname(os.path.abspath(__file__))
        if _scripts_dir not in sys.path:
            sys.path.insert(0, _scripts_dir)
    except Exception:
        pass
    import seleccion_caras_elemento as sel

    if (
        (getattr(sel, "HOST_CARAS_SELECCION", None) is None or not getattr(sel, "REFERENCIAS_CARAS_SELECCION", None))
        and "seleccion_caras_elemento_btn" in sys.modules
    ):
        _leg = sys.modules["seleccion_caras_elemento_btn"]
        if getattr(_leg, "HOST_CARAS_SELECCION", None) is not None and getattr(
            _leg, "REFERENCIAS_CARAS_SELECCION", None
        ):
            sel = _leg

    uidoc = revit.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(u"Enfierrado shaft", u"No hay documento activo.")
        return
    doc = uidoc.Document

    host = sel.HOST_CARAS_SELECCION
    refs = sel.REFERENCIAS_CARAS_SELECCION

    if host is None or not refs:
        TaskDialog.Show(
            u"Enfierrado shaft",
            u"Primero selecciona host y caras (herramientas de borde de losa o módulo "
            u"seleccion_caras_elemento).",
        )
        return

    view = uidoc.ActiveView
    n_dc = 0
    avisos_dc = []
    creados = 0
    tags = 0
    avisos = []
    t = Transaction(doc, u"BIMTools: shaft — detalle + barras + tags")
    t.Start()
    try:
        n_dc, avisos_dc, err_dc = crear_detail_curves_tramos_shaft_hashtag(
            doc,
            view,
            host,
            refs,
            cover_mm=cover_mm,
            ignore_empotramientos=True,
            use_transaction=False,
        )
        if err_dc:
            raise Exception(err_dc)

        creados, tags, _ids, avisos, err = crear_enfierrado_shaft_hashtag(
            doc,
            host,
            refs,
            cover_mm=cover_mm,
            duplex_spacing_mm=duplex_spacing_mm,
            tag_view=view,
            tag_family_name=u"EST_A_STRUCTURAL REBAR TAG",
            ignore_empotramientos=True,
            use_transaction=False,
        )
        if err:
            raise Exception(err)
        t.Commit()
    except Exception as ex:
        t.RollBack()
        body = unicode(ex)
        if avisos:
            body += u"\n\n" + u"\n".join(avisos[:10])
        if n_dc:
            body = u"Líneas de detalle (tramos): {0}.\n\n".format(n_dc) + body
        if avisos_dc:
            body += u"\n\n" + u"\n".join(avisos_dc[:8])
        TaskDialog.Show(u"Enfierrado shaft", body)
        return

    msg = u"Líneas de detalle (tramos): {0}.\n".format(n_dc)
    msg += u"Barras creadas: {0} (un Rebar por tramo válido).\n".format(creados)
    msg += u"Etiquetas creadas: {0}.\n".format(tags)
    todos = list(avisos_dc) + list(avisos)
    if todos:
        msg += u"\nAvisos:\n" + u"\n".join(todos[:10])
    TaskDialog.Show(u"Enfierrado shaft", msg)
