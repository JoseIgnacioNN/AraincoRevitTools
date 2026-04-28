# -*- coding: utf-8 -*-
"""
Extiende un Structural Rebar en L (pata corta al **inicio** del primer tramo o al
**final** del último tramo según ``pata_en_extremo_final``), copia la regla de reparto y asigna
ganchos de 135° en ambos extremos al Rebar resultante (salvo
``gancho_135_solo_en_extremo_pata=True``: solo 135 en el extremo de la pata L; el
opuesto queda sin gancho). Para verticales de muro, use
``pata_en_extremo_final_para_pie_por_elevacion`` / ``_cabeza_por_elevacion`` al elegir
``pata_en_extremo_final``, porque el sentido del boceto (Z) define qué extremo es pie o cabeza.

- Revit 2024
- Revit Python Shell (IronPython 3.4.x)
- No edita el boceto in-place: CreateFromCurves + borrado del original.

Lectura de layout: Rebar.MaxSpacing, Rebar.LayoutRule, RebarShapeDrivenAccessor.ArrayLength.
Largo de pata L (por defecto): **espesor del host (muro/losa) − 50 mm** (``PATA_RESTA_ESPESOR_HOST_MM``),
mínimo ``PATA_LARGO_MIN_MM``; si no se lee espesor, ``LARGO_PATA_MM``. Para **verticales** de
muro, ``largo_pata_mm_muro_vertical_entre_caras``: espesor − rec. cara ext. − rec. int. − Ø
vertical en ext. − Ø vertical en int. (p. ej. 300−25−25−8−8=234).
Ganchos: por defecto usa el RebarHookType por nombre (selector de propiedades), p. ej.
«Standard - 135 deg.» Si REBAR_HOOK_TYPE_NAME está vacío, se busca un tipo con ~135° o se crea uno.
La orientación in-plane (Left/Right) y la rotación fuera de plano (grados) se aplican vía
SetTerminationOrientation/SetHookOrientation y SetTerminationRotationAngle/SetHookRotationAngle
según la versión de Revit. Si solo Left/Right no mueve el gancho en 3D, use HOOK_ROTATION_*
e INVERTIR_NORMAL_REBAR.
"""

from __future__ import print_function

import math
import os
import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import System
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Curve,
    ElementId,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    Floor,
    IFailuresPreprocessor,
    Line,
    Transaction,
    UnitUtils,
    UnitTypeId,
    Wall,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB.Structure import (
    MultiplanarOption,
    Rebar,
    RebarBarType,
    RebarHookOrientation,
    RebarHookType,
    RebarShape,
    RebarStyle,
)
try:
    from Autodesk.Revit.DB.Structure import RebarShapeDefinitionBySegments
except System.Exception:  # noqa: BLE001
    RebarShapeDefinitionBySegments = None


class _BimToolsRebarTxnFailuresPreprocessor(IFailuresPreprocessor):
    """
    Elimina *warnings* que no requieren diálogo (evita bloqueo en ``Commit``).
    """

    def _iter_failure_msgs(self, failures_accessor):
        if failures_accessor is None:
            return
        try:
            fmsgs = failures_accessor.GetFailureMessages()
        except System.Exception:
            return
        if fmsgs is None:
            return
        try:
            n = int(fmsgs.Count)
        except System.Exception:
            n = 0
        for i in range(n):
            f = None
            try:
                f = fmsgs.get_Item(i)
            except System.Exception:
                try:
                    f = fmsgs[i]
                except System.Exception:
                    f = None
            if f is not None:
                yield f

    def PreprocessFailures(self, failures_accessor):
        if failures_accessor is None:
            return FailureProcessingResult.Continue
        for f in self._iter_failure_msgs(failures_accessor):
            try:
                if f.GetSeverity() == FailureSeverity.Warning:
                    try:
                        failures_accessor.DeleteWarning(f)
                    except System.Exception:
                        pass
            except System.Exception:
                pass
        return FailureProcessingResult.Continue


# --- Parámetros usuario ---
# Si no se puede leer el espesor del host, se usa este valor (mm) como largo de pata.
LARGO_PATA_MM = 200.0
# Muro (y losa, si aplica): largo pata = espesor del host **menos** este valor (mm).
PATA_RESTA_ESPESOR_HOST_MM = 50.0
# Largo mínimo de pata (mm) aunque espesor − resta quede por debajo.
PATA_LARGO_MIN_MM = 10.0
INVERTIR_DIRECCION_PATA = False
INDICE_POSICION = 0
# True: la pata corta del L se coloca en el **último** punto del trazado (final del sketch);
# False: en el **primer** punto (inicio del primer tramo), comportamiento histórico.
PATA_EN_EXTREMO_FINAL = False

# Ganchos 135° (ambos extremos)
# Nombre del RebarHookType (como en Propiedades / Type). Prioridad: este tipo antes que ángulo/crear.
# Dejar vacío (u"") para el comportamiento anterior: buscar ~135° o crear.
REBAR_HOOK_TYPE_NAME = u"Standard - 135 deg."
HOOK_ANGLE_DEG = 135.0
HOOK_EXTENSION_MULT = 12.0
# Longitud de gancho (mm) al crear un tipo nuevo; si ya existe 135° en el proyecto, se reutiliza
HOOK_LENGTH_MM_135 = 100.0
# Tolerancia al buscar gancho existente por ángulo (grados)
HOOK_ANGLE_MATCH_TOL_DEG = 2.0
# Longitudes (mm) a probar en el RebarBarType si CanUseHookType/Set falla (diámetro pequeño, plegado)
HOOK_LENGTH_FALLBACK_MM = (
    100.0,
    150.0,
    200.0,
    250.0,
    300.0,
    400.0,
    500.0,
    600.0,
    800.0,
    1000.0,
    1200.0,
    1500.0,
)

# Orientación in-plane (Left/Right). Si en la vista 3D «no pasa nada» al cambiar esto, Revit
# puede seguir leyendo la terminación antigua, o hace falta la rotación fuera de plano (abajo).
HOOK_ORIENT_END0 = RebarHookOrientation.Left
HOOK_ORIENT_END1 = RebarHookOrientation.Left

# Rotación fuera de plano (grados) en cada extremo con gancho: suele ser lo que alinea
# «hacia el interior del muro» cuando Left/Right no mueve el gancho 135° en 3D.
# Pruebe 0/0, 180/0, 0/180 o 180/180 según qué extremo muestre el gancho en la vista.
HOOK_ROTATION_END0_DEG = 180.0
HOOK_ROTATION_END1_DEG = 180.0

# Si True, se usa -normal (misma pata L, geometría coherente) y suele intercambiar
# el lado «interior/exterior» del plano de la barra respecto al host.
INVERTIR_NORMAL_REBAR = False

# ``extender_doble_pata_135_y_reemplazar`` siempre pasa por ``_assign_135_hooks_both_ends`` (no
# el atajo CreateFromCurves) porque la polilínea de 3+ tramos suele dejar 135° **hacia afuera**;
# Right + rot. 0/0 suele alinear hacia el interior; si en su sección aún mira al revés, pruebe
# Left, Left y/o HOOK_ROTATION_* 180,180 como en ganchos de un solo L.
DOBLE_PATA_GANCHO_ORIENT0 = RebarHookOrientation.Right
DOBLE_PATA_GANCHO_ORIENT1 = RebarHookOrientation.Right
DOBLE_PATA_GANCHO_ROT0_DEG = 0.0
DOBLE_PATA_GANCHO_ROT1_DEG = 0.0

# Muro apilado unido a la **cara superior** del host (Armado muros nudo — verticales L+135):
# Valores base; ``armado_muros_nodo_refuerzo_post`` sustituye el par (0,1) según cara
# ext./int. (gancho en ext. 0 \u2248 Hook at Start en ext., en ext. 1 en int.)
# Rot. base 0°/0°; el armado pasa 180° al extremo de pata (0 o 1) según cara. El
# extensor reaplica orient/rot en el extremo pata si la creación vino de Create+hooks.
GANCHOS_135_MURO_CARA_SUP_ORIENT0 = RebarHookOrientation.Right
GANCHOS_135_MURO_CARA_SUP_ORIENT1 = RebarHookOrientation.Right
GANCHOS_135_MURO_CARA_SUP_ROT0 = 0.0
GANCHOS_135_MURO_CARA_SUP_ROT1 = 0.0

# ``CreateFromCurves`` puede asignar un RebarShape que choca con ganchos 135°
# («Can't solve Rebar Shape» al Commit). Se usa ``CreateFromCurvesAndShape`` con un
# ``RebarShape`` cuyo **número de segmentos = número de curvas** en la polilínea (p. ej. 2
# para pata+un tramo, 4 si el boceto aportaba 3 tramos y se añade la pata).
# Nombre exacto o subcadenas priorizan formas frecuentes (11, 14, …) dentro de ese n-seg.
REBAR_SHAPE_NAME_EXACT_2SEG = u""
REBAR_SHAPE_NAME_CONTAINS_PREFER_2SEG = (u"11", u"14", u"00", u"12", u"13", u"10", u"20")
REBAR_SHAPE_MAX_TRY_2SEG = 32


def _mm_to_internal(mm):
    return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)


def _obtener_espesor_host_mm(document, host):
    """Espesor del muro o losa en mm, o None si no se pudo leer."""
    if host is None or document is None:
        return None
    try:
        if isinstance(host, Wall):
            p = host.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
            if p is not None and p.HasValue:
                return float(
                    UnitUtils.ConvertFromInternalUnits(
                        p.AsDouble(), UnitTypeId.Millimeters
                    )
                )
        if isinstance(host, Floor):
            p = host.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
            if p is not None and p.HasValue:
                return float(
                    UnitUtils.ConvertFromInternalUnits(
                        p.AsDouble(), UnitTypeId.Millimeters
                    )
                )
    except System.Exception:
        pass
    try:
        p = host.LookupParameter(u"Default Thickness")
        if p is not None and p.HasValue:
            return float(
                UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.Millimeters)
            )
    except System.Exception:
        pass
    try:
        type_id = host.GetTypeId()
        if type_id is not None and type_id != ElementId.InvalidElementId:
            wtype = document.GetElement(type_id)
            if wtype is not None:
                for pname in (u"Default Thickness", u"Thickness", u"Espesor", u"Width"):
                    try:
                        p = wtype.LookupParameter(pname)
                        if p is not None and p.HasValue:
                            return float(
                                UnitUtils.ConvertFromInternalUnits(
                                    p.AsDouble(), UnitTypeId.Millimeters
                                )
                            )
                    except System.Exception:
                        pass
    except System.Exception:
        pass
    return None


def largo_pata_mm_desde_espesor_host(
    document,
    host,
    resta_mm=None,
    fallback_mm=None,
    min_largo_mm=None,
):
    """
    Largo de la pata L = espesor(muro o losa) − resta (por defecto 50 mm).
    Si no hay espesor, se usa ``LARGO_PATA_MM`` como respaldo.
    """
    if resta_mm is None:
        resta_mm = PATA_RESTA_ESPESOR_HOST_MM
    if fallback_mm is None:
        fallback_mm = LARGO_PATA_MM
    if min_largo_mm is None:
        min_largo_mm = PATA_LARGO_MIN_MM
    th = _obtener_espesor_host_mm(document, host)
    if th is None:
        return max(float(min_largo_mm), float(fallback_mm))
    x = float(th) - float(resta_mm)
    return max(float(min_largo_mm), x)


def _documento_tipo_muro(wall):
    u"""Elemento de tipo (WallType) o None."""
    if wall is None:
        return None
    try:
        return wall.Document.GetElement(wall.GetTypeId())
    except System.Exception:
        return None


def _recubrimiento_caras_muro_mm(wall):
    u"""
    Recubrimiento (mm) cara exterior e interior (instancia o tipo de muro; parámetros
    Revit / español). Valores 0 o ausentes se tratan como no definidos (luego 25+25)
    — si Revit tiene 0 mm en el param, sin esto se obtiene 300−8−8=284 al restar
    solo diámetros.
    """
    if wall is None:
        return 25.0, 25.0
    names_out = (
        u"Rebar Cover - Exterior Face",
        u"Rebar cover exterior face",
        u"Recubrimiento rebar - Cara exte",
    )
    names_in = (
        u"Rebar Cover - Interior Face",
        u"Rebar cover interior face",
        u"Recubrimiento rebar - Cara int",
    )

    def _leer_mm_positivo(element, names):
        if element is None:
            return None
        for n in names:
            try:
                p = element.LookupParameter(n)
                if p is None or not p.HasValue:
                    continue
                v = float(
                    UnitUtils.ConvertFromInternalUnits(
                        p.AsDouble(), UnitTypeId.Millimeters
                    )
                )
                if v > 0.0:
                    return v
            except System.Exception:
                pass
        return None

    def _read_cara(names):
        u"""Instancia, luego tipo: primer valor estrictamente > 0 (mm)."""
        v = _leer_mm_positivo(wall, names)
        if v is not None:
            return v
        return _leer_mm_positivo(_documento_tipo_muro(wall), names)

    co = _read_cara(names_out)
    ci = _read_cara(names_in)
    if co is None or co <= 0.0:
        co = 25.0
    if ci is None or ci <= 0.0:
        ci = 25.0
    return co, ci


def _diametro_nominal_rebar_mm(document, rebar):
    if document is None or rebar is None:
        return None
    try:
        bt = document.GetElement(rebar.GetTypeId())
        if bt is None or not isinstance(bt, RebarBarType):
            return None
        d = float(bt.BarModelDiameter)
        return float(
            UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters)
        )
    except System.Exception:
        return None


def _rebar_es_vertical_muro_tangente(rebar, host, pos_idx=0):
    u"""Mismo criterio Z que malla vertical (cota |tangente| >= 0,45 int)."""
    if rebar is None or not isinstance(host, Wall):
        return False
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, int(pos_idx)
        )
        if crvs is None or crvs.Count < 1:
            return False
        c0 = crvs[0]
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        v = p1 - p0
        if v.GetLength() < 1e-12:
            return False
        return abs(float(v.Z)) >= 0.45
    except System.Exception:
        return False


def largo_pata_mm_muro_vertical_entre_caras(document, host, rebar, pos_idx=0):
    u"""
    Largo pata L en horizontal (entre caras) para refuerzo **vertical** de muro:
    ``espesor - rec_ext - rec_int - d_ext - d_int``, con **d** = Ø de los **verticales
    (longitudinales)** en cada cara, no el Ø del cerco horizontal (p. ej. Ø8 en planta:
    aun así los verticales suelen anotarse en sección; si no se hallan, se usa d del
    rebar en curso).

    Si no se puede calcular, devuelve ``None`` (el caller usa ``largo_pata_mm_desde_espesor_host``).
    """
    if document is None or host is None or rebar is None:
        return None
    if not isinstance(host, Wall):
        return None
    try:
        import arearein_exterior_h_l135_rps as arex
        import arearein_interior_h_l135_rps as arin
    except System.Exception:
        return None
    th = _obtener_espesor_host_mm(document, host)
    if th is None:
        return None
    cext, cint = _recubrimiento_caras_muro_mm(host)
    d_self = _diametro_nominal_rebar_mm(document, rebar)
    if d_self is None:
        return None
    wid = host.Id
    d_ext = None
    d_int = None
    try:
        for rb in FilteredElementCollector(document).OfClass(Rebar):
            if rb is None:
                continue
            try:
                if rb.GetHostId() != wid:
                    continue
            except System.Exception:
                continue
            if not _rebar_es_vertical_muro_tangente(rb, host, int(pos_idx)):
                continue
            d = _diametro_nominal_rebar_mm(document, rb)
            if d is None:
                continue
            ex = arex._rebar_solo_cara_exterior(rb, host)
            inn = arin._rebar_solo_cara_interior(rb, host)
            if ex and not inn:
                d_ext = d if d_ext is None else max(d_ext, d)
            elif inn and not ex:
                d_int = d if d_int is None else max(d_int, d)
    except System.Exception:
        return None
    if d_ext is None:
        d_ext = d_self
    if d_int is None:
        d_int = d_self
    try:
        x = float(th) - float(cext) - float(cint) - float(d_ext) - float(d_int)
    except System.Exception:
        return None
    return max(float(PATA_LARGO_MIN_MM), x)


def _rebar_normal(rebar):
    try:
        acc = rebar.GetShapeDrivenAccessor()
        if acc is not None:
            n = acc.Normal
            if n is not None and n.GetLength() > 1e-12:
                return n.Normalize()
    except System.Exception:
        pass
    return XYZ.BasisZ


def _tangent_start_first_curve(crv):
    if crv is None:
        return None
    p0 = crv.GetEndPoint(0)
    p1 = crv.GetEndPoint(1)
    v = p1 - p0
    if v.GetLength() < 1e-12:
        return None
    return v.Normalize()


def _perp_in_plane(normal, tangent):
    c = normal.CrossProduct(tangent)
    if c.GetLength() < 1e-10:
        c = tangent.CrossProduct(normal)
    if c.GetLength() < 1e-10:
        return None
    return c.Normalize()


def _rebar_shape_n_segments(sh):
    if sh is None or RebarShapeDefinitionBySegments is None:
        return -1
    try:
        rsd = sh.GetRebarShapeDefinition()
        if isinstance(rsd, RebarShapeDefinitionBySegments):
            return int(rsd.NumberOfSegments)
    except System.Exception:
        pass
    return -1


def _rebar_shape_type_label(sh):
    if sh is None:
        return u""
    for fn in (lambda: sh.Name,):
        try:
            t = fn()
            if t:
                t = (t or u"").replace(u"\u00A0", u" ").strip()
                if t:
                    return t
        except System.Exception:
            pass
    try:
        p = sh.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None and p.HasValue:
            return (p.AsString() or u"").strip()
    except System.Exception:
        pass
    return u""


def _collect_rebar_shapes_nseg_ordered(document, n_seg):
    """
    ``RebarShape`` con exactamente ``n_seg`` tramos. Orden: nombre exacto, preferidos, resto.
    Debe coincidir con len(curve_list) en ``CreateFromCurvesAndShape``.
    """
    if document is None:
        return []
    try:
        ns = int(n_seg)
    except System.Exception:
        return []
    if ns < 1:
        return []
    acc = list(FilteredElementCollector(document).OfClass(RebarShape))
    if not acc:
        return []
    with_n = [s for s in acc if _rebar_shape_n_segments(s) == ns]
    if not with_n:
        return []
    seen = set()
    out = []

    def _add(s):
        if s is None:
            return
        e = _eid_int(s.Id)
        if e and e not in seen:
            seen.add(e)
            out.append(s)

    n_ex = (REBAR_SHAPE_NAME_EXACT_2SEG or u"").strip()
    if n_ex:
        want = n_ex.lower()
        for s in with_n:
            if _rebar_shape_type_label(s).lower() == want:
                _add(s)
                break
    for sub in REBAR_SHAPE_NAME_CONTAINS_PREFER_2SEG:
        t = (sub or u"").lower().strip()
        if not t:
            continue
        for s in with_n:
            if t in _rebar_shape_type_label(s).lower():
                _add(s)
    rest = sorted(
        [s for s in with_n if _eid_int(s.Id) not in seen],
        key=lambda x: _rebar_shape_type_label(x).lower(),
    )
    for s in rest:
        _add(s)
    return out[: int(REBAR_SHAPE_MAX_TRY_2SEG)]


def _try_create_l_from_rebar_shape_2seg(
    document, curves_list, host, norm, bar_type, style, o0, o1
):
    """
    Crea con ``RebarShape`` cuyo **número de segmentos = len(curves_list)``.
    Un solo tramo en el boceto + pata = 2 curvas; varios tramos + pata = 3+ curvas.
    Sigue a ``rebar_fundacion_cara_inferior`` (mismas sobrecargas API).
    """
    if document is None or not curves_list or host is None or bar_type is None:
        return None
    n_seg = len(curves_list)
    if n_seg < 1:
        return None
    try:
        cl = List[Curve]()
        for c in curves_list:
            cl.Add(c)
    except System.Exception:
        return None
    orient_tries = [
        (o0, o1),
        (RebarHookOrientation.Right, RebarHookOrientation.Right),
        (RebarHookOrientation.Left, RebarHookOrientation.Left),
        (RebarHookOrientation.Right, RebarHookOrientation.Left),
    ]
    seen_or = set()
    pairs = []
    for a in orient_tries:
        try:
            k = (int(a[0]), int(a[1]))
        except System.Exception:
            k = (str(a[0]), str(a[1]))
        if k not in seen_or:
            seen_or.add(k)
            pairs.append(a)
    for shape in _collect_rebar_shapes_nseg_ordered(document, n_seg):
        for so, eo in pairs:
            try:
                r = Rebar.CreateFromCurvesAndShape(
                    document,
                    shape,
                    bar_type,
                    None,
                    None,
                    host,
                    norm,
                    cl,
                    so,
                    eo,
                    0.0,
                    0.0,
                    ElementId.InvalidElementId,
                    ElementId.InvalidElementId,
                )
                if r is not None:
                    _apply_rebar_style_if_writable(r, style)
                    return r
            except System.Exception:
                pass
            try:
                r = Rebar.CreateFromCurvesAndShape(
                    document, shape, bar_type, None, None, host, norm, cl, so, eo
                )
                if r is not None:
                    _apply_rebar_style_if_writable(r, style)
                    return r
            except System.Exception:
                pass
    return None


def _apply_rebar_style_if_writable(rebar, style):
    if rebar is None or style is None:
        return
    try:
        rebar.Style = style
    except System.Exception:
        pass


def _create_from_curves_no_hooks(doc, curves_list, host, norm, bar_type, style, o0, o1):
    """CreateFromCurves sin ganchos; se asignan tras el layout."""
    ct = clr.GetClrType(Line).BaseType
    n = len(curves_list)
    arr = System.Array.CreateInstance(ct, n)
    for i in range(n):
        arr[i] = curves_list[i]
    return Rebar.CreateFromCurves(
        doc, style, bar_type, None, None, host, norm, arr, o0, o1, True, True
    )


def _try_create_l_with_hook_types_both_ends(
    doc, curves_list, host, norm, bar_type, style, o0, o1, hook_id
):
    """
    Tras dejar de aplicar ganchos post-``SetLayout`` (fallos reales: ``all_hook_lengths_failed`` en
    depuración), se prueba :meth:`Rebar.CreateFromCurves` con 135° en 0 y 1 en la creación
    (misma pauta que enfierrado_shaft: combinar ``useExisting``/``createNew``).
    Si Revit acepta el trazado, el reparto (``_copy_layout``) va después, como en pata 90/area.
    """
    if (
        doc is None
        or host is None
        or bar_type is None
        or hook_id is None
        or hook_id == ElementId.InvalidElementId
    ):
        return None
    # La API (IronPython) espera instancias RebarHookType, no ElementId (ver depuración: ArgumentTypeException).
    hook_type_for_create = None
    try:
        h_el = doc.GetElement(hook_id)
        if isinstance(h_el, RebarHookType):
            hook_type_for_create = h_el
    except System.Exception:
        pass
    if hook_type_for_create is None:
        return None
    ct = clr.GetClrType(Line).BaseType
    n = len(curves_list)
    arr = System.Array.CreateInstance(ct, n)
    for i in range(n):
        arr[i] = curves_list[i]
    # Solo el mismo RebarStyle que el rebar de origen (p. ej. Standard). No StirrupTie:
    # puede disolver en «Can't solve Rebar Shape» en L + ganchos 135.
    for use_existing, create_new in ((True, True), (True, False), (False, True), (False, False)):
        try:
            r = Rebar.CreateFromCurves(
                doc,
                style,
                bar_type,
                hook_type_for_create,
                hook_type_for_create,
                host,
                norm,
                arr,
                o0,
                o1,
                use_existing,
                create_new,
            )
            if r is not None:
                return r
        except System.Exception:  # noqa: BLE001
            continue
    return None


def _layout_rule_name(rebar, acc):
    try:
        r = rebar.LayoutRule
        if r is not None:
            return r.ToString()
    except System.Exception:
        pass
    if acc is not None:
        try:
            r = acc.GetLayoutRule()
            if r is not None:
                return r.ToString()
        except System.Exception:
            pass
    return u""


def _spacing_internal(rebar):
    try:
        return float(rebar.MaxSpacing)
    except System.Exception:
        return 0.0


def _array_length_internal(acc):
    if acc is None:
        return 0.0
    try:
        return float(acc.ArrayLength)
    except System.Exception:
        try:
            return float(acc.GetArrayLength())
        except System.Exception:
            return 0.0


def _copy_layout_rebar_shape_driven(src, dst):
    a0 = src.GetShapeDrivenAccessor()
    a1 = dst.GetShapeDrivenAccessor()
    if a0 is None or a1 is None:
        return False, u"ShapeDrivenAccessor nulo (no es rebar con layout copiable)."

    rule_name = _layout_rule_name(src, a0)
    sp = _spacing_internal(src)
    alen = _array_length_internal(a0)
    b_side = bool(a0.BarsOnNormalSide)
    inc0 = bool(src.IncludeFirstBar)
    inc1 = bool(src.IncludeLastBar)
    nbars = int(src.Quantity)

    try:
        if rule_name == u"Single":
            a1.SetLayoutAsSingle()
        elif rule_name == u"MaximumSpacing":
            a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        elif rule_name in (u"Number", u"FixedNumber"):
            a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
        elif rule_name == u"NumberWithSpacing":
            a1.SetLayoutAsNumberWithSpacing(nbars, sp, alen, b_side, inc0, inc1)
        elif rule_name == u"MinimumClearSpacing":
            a1.SetLayoutAsMinimumClearSpacing(sp, alen, b_side, inc0, inc1)
        else:
            if rule_name:
                try:
                    a1.SetLayoutAsFixedNumber(nbars, alen, b_side, inc0, inc1)
                except System.Exception:
                    a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
            else:
                a1.SetLayoutAsMaximumSpacing(sp, alen, b_side, inc0, inc1)
        return True, u""
    except System.Exception as ex:
        if int(nbars) == 1:
            try:
                a1.SetLayoutAsSingle()
                return True, u""
            except System.Exception as ex2:
                return (
                    False,
                    u"{0!s} (regla: «{1}») | fallback Single: {2!s}".format(
                        ex, rule_name or u"(vacía)", ex2
                    ),
                )
        return (
            False,
            u"{0!s} (regla: «{1}»)".format(ex, rule_name or u"(vacía)"),
        )


def _hook_angle_rad(ht):
    try:
        return float(ht.HookAngle)
    except System.Exception:
        return None


def _find_rebar_hook_by_angle_deg(document, angle_deg, tol_deg):
    target = math.radians(float(angle_deg))
    tol = math.radians(float(tol_deg))
    for ht in FilteredElementCollector(document).OfClass(RebarHookType):
        a = _hook_angle_rad(ht)
        if a is None:
            continue
        if abs(a - target) < tol:
            return ht
    return None


def _hook_135_type_priority_index(ht):
    """Menor = preferido como alternativa (p. ej. estribo/sísmico)."""
    if ht is None:
        return 99
    name = u""
    try:
        name = (ht.Name or u"").lower()
    except System.Exception:
        pass
    if u"stirrup" in name or u"seismic" in name or u"tie" in name or u"lazo" in name:
        return 0
    if u"standard" in name:
        return 1
    return 2


def _ordered_135_hook_type_eids(document, primary_eid, max_total=6):
    """
    RebarBarType+gancho: el primer tipo ~135° del collector puede no ser aceptado por
    SetHookTypeId; reordena e incluye otras familias 135 del proyecto (estribo/sísmico, etc.).
    El primary (resuelto por nombre/ángulo) va siempre el primero.
    """
    if primary_eid is None or primary_eid == ElementId.InvalidElementId:
        return []
    try:
        mt = int(max_total)
    except System.Exception:
        mt = 6
    if mt < 1:
        mt = 1
    target = math.radians(float(HOOK_ANGLE_DEG))
    tol = math.radians(float(HOOK_ANGLE_MATCH_TOL_DEG))
    p_int = _eid_int(primary_eid)
    all_h = []
    try:
        for ht in FilteredElementCollector(document).OfClass(RebarHookType):
            if ht is None:
                continue
            a = _hook_angle_rad(ht)
            if a is None or abs(a - target) >= tol:
                continue
            all_h.append(ht)
    except System.Exception:
        return [primary_eid]
    all_h.sort(key=_hook_135_type_priority_index)
    out = [primary_eid]
    seen = {p_int}
    for ht in all_h:
        e = _eid_int(ht.Id)
        if e in seen:
            continue
        out.append(ht.Id)
        seen.add(e)
        if len(out) >= mt:
            break
    return out


def _crear_rebar_hook_135(document, largo_mm, mult):
    """Crea RebarHookType 135° y fija Hook Length en todos los RebarBarType."""
    bar_types = list(FilteredElementCollector(document).OfClass(RebarBarType))
    if not bar_types:
        return None
    ang = math.radians(HOOK_ANGLE_DEG)
    hook_type = RebarHookType.Create(document, ang, float(mult))
    largo_interno = UnitUtils.ConvertToInternalUnits(float(largo_mm), UnitTypeId.Millimeters)
    for bt in bar_types:
        try:
            bt.SetAutoCalcHookLengths(hook_type.Id, False)
            bt.SetHookLength(hook_type.Id, largo_interno)
        except System.Exception:
            pass
    nombre_base = u"Rebar Hook - 135 - {} mm (BIMTools L-ext)".format(int(round(largo_mm)))
    existentes = []
    for h in FilteredElementCollector(document).OfClass(RebarHookType):
        try:
            if h.Name:
                existentes.append(h.Name)
        except System.Exception:
            pass
    nombre_final = nombre_base
    if nombre_base in existentes:
        k = 1
        while u"{} ({})".format(nombre_base, k) in existentes:
            k += 1
        nombre_final = u"{} ({})".format(nombre_base, k)
    try:
        hook_type.Name = nombre_final
    except System.Exception:
        pass
    return hook_type


def _eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except System.Exception:
        return 0


def _set_hook_plane_orientation(rebar, end_idx, hook_orient):
    """
    SetHookOrientation (API clásica) o SetTerminationOrientation + RebarTerminationOrientation
    (2025+), que es la vía soportada cuando la antigua deja de actualizar la geometría.
    end_idx: 0 o 1.
    """
    if rebar is None or hook_orient is None:
        return False
    try:
        e = int(end_idx)
    except System.Exception:
        return False
    try:
        fn = getattr(rebar, "SetTerminationOrientation", None)
        if fn is not None:
            rto = None
            try:
                from Autodesk.Revit.DB.Structure import RebarTerminationOrientation
            except System.Exception:
                RebarTerminationOrientation = None
            if RebarTerminationOrientation is not None:
                for name in (u"Left", u"Right"):
                    try:
                        v = int(hook_orient) == int(getattr(RebarHookOrientation, name))
                    except System.Exception:
                        v = False
                    if v:
                        rto = getattr(RebarTerminationOrientation, name, None)
                        break
            if rto is not None:
                fn(e, rto)
                return True
    except System.Exception:
        pass
    try:
        rebar.SetHookOrientation(e, hook_orient)
        return True
    except System.Exception:
        return False


def _set_rebar_termination_or_hook_rotation_deg(rebar, end_idx, deg):
    """
    Revit 2026: SetTerminationRotationAngle(end, rad). Antes: SetHookRotationAngle(rad, end).
    Unidades internas: radianes (UnitUtils con UnitTypeId.Degrees).
    """
    if rebar is None:
        return False
    try:
        e = int(end_idx)
    except System.Exception:
        return False
    try:
        d = float(deg)
    except System.Exception:
        return False
    if abs(d) < 1e-9:
        return True
    try:
        rad = UnitUtils.ConvertToInternalUnits(d, UnitTypeId.Degrees)
    except System.Exception:
        rad = d * (math.pi / 180.0)
    try:
        fn = getattr(rebar, "SetTerminationRotationAngle", None)
        if fn is not None:
            fn(e, rad)
            return True
    except System.Exception:
        pass
    try:
        fn = getattr(rebar, "SetHookRotationAngle", None)
        if fn is not None:
            fn(rad, e)
            return True
    except System.Exception:
        pass
    return False


def _format_hook_state_debug(new_rb, base_note):
    """Resumen: orient y rotación leída (para comprobar que la API aplica)."""
    if new_rb is None:
        return base_note
    parts = [base_note] if base_note else []
    for end in (0, 1):
        try:
            o = new_rb.GetHookOrientation(end)
            parts.append(u"O{0}={1}".format(end, o))
        except System.Exception:
            try:
                o = new_rb.GetTerminationOrientation(end)
                parts.append(u"O{0}={1}".format(end, o))
            except System.Exception:
                pass
    for end in (0, 1):
        try:
            rfn = getattr(new_rb, "GetTerminationRotationAngle", None)
            if rfn is not None:
                rv = float(rfn(end))
                try:
                    degv = float(
                        UnitUtils.ConvertFromInternalUnits(rv, UnitTypeId.Degrees)
                    )
                except System.Exception:
                    degv = rv * 180.0 / math.pi
                parts.append(u"rot{0}={1:.0f}°".format(end, degv))
        except System.Exception:
            try:
                gr = getattr(new_rb, "GetHookRotationAngle", None)
                if gr is not None:
                    rads = float(gr(end))
                    degv = rads * 180.0 / math.pi
                    parts.append(u"rot{0}={1:.0f}°".format(end, degv))
            except System.Exception:
                pass
    return u" | ".join([p for p in parts if p])


def _nominal_diameter_mm(bar_type):
    try:
        d = bar_type.BarModelDiameter
        return float(UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters))
    except System.Exception:
        return 10.0


def _hook_length_candidates_mm_for_bar_type(bar_type):
    """Incluye 12×diámetro aprox. y la lista de respaldo (evita CanUseHookType false por longitud)."""
    d = _nominal_diameter_mm(bar_type)
    est = max(12.0 * d, 50.0)
    cands = [HOOK_LENGTH_MM_135, est]
    cands.extend(HOOK_LENGTH_FALLBACK_MM)
    out = []
    seen = set()
    for c in cands:
        c = max(10.0, float(c))
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _ensure_hook_length_on_bar_type(document, bar_type, hook_id, length_mm):
    if bar_type is None or hook_id is None or hook_id == ElementId.InvalidElementId:
        return
    try:
        bar_type.SetAutoCalcHookLengths(hook_id, False)
        li = UnitUtils.ConvertToInternalUnits(float(length_mm), UnitTypeId.Millimeters)
        bar_type.SetHookLength(hook_id, li)
    except System.Exception:
        pass


def _clear_both_hook_ends_rebar(new_rb, document):
    """
    Pone ganchos en ambos extremos a InvalidElementId; tras fallos 135, deja rebar L sin ganchos.
    """
    inv = ElementId.InvalidElementId
    for end in (0, 1):
        _set_rebar_hook_type_id_at_end(new_rb, end, inv, document, 3)
    try:
        document.Regenerate()
    except System.Exception:
        pass
    return True


def _set_rebar_hook_type_id_at_end(rebar, end_idx, type_id, document, max_attempts=3):
    """
    Equivalencia ligera a enfierrado_shaft: SetHookTypeId + comprobación + Regenerate.
    """
    if rebar is None:
        return False
    e = int(end_idx)
    for attempt in range(int(max_attempts)):
        try:
            rebar.SetHookTypeId(e, type_id)
        except System.Exception:
            if document is not None and attempt + 1 < int(max_attempts):
                try:
                    document.Regenerate()
                except System.Exception:
                    pass
            continue
        try:
            if _eid_int(rebar.GetHookTypeId(e)) == _eid_int(type_id):
                return True
        except System.Exception:
            pass
        if document is not None and attempt + 1 < int(max_attempts):
            try:
                document.Regenerate()
            except System.Exception:
                pass
    return False


def _hook_name_key(s):
    if s is None:
        return u""
    try:
        t = s.replace(u"\u00A0", u" ").strip()
        t = u" ".join(t.split())
    except System.Exception:
        t = u""
    return t


def _rebar_hook_type_by_name_in_document(document, name):
    """
    Resuelve RebarHookType por el nombre de tipo (Name) como en el desplegable de propiedades.
    Primero igualdad (normalizado espacios), luego misma clave sin distinguir mayúsculas.
    """
    if document is None or not name:
        return None
    want = _hook_name_key(name)
    if not want:
        return None
    try:
        hooks = list(FilteredElementCollector(document).OfClass(RebarHookType))
    except System.Exception:
        return None
    for ht in hooks:
        if ht is None:
            continue
        try:
            if _hook_name_key(ht.Name) == want:
                return document.GetElement(ht.Id) if ht.Id is not None else ht
        except System.Exception:
            pass
    want_l = want.lower()
    for ht in hooks:
        if ht is None:
            continue
        try:
            if _hook_name_key(ht.Name).lower() == want_l:
                return document.GetElement(ht.Id) if ht.Id is not None else ht
        except System.Exception:
            pass
    return None


def _rebar_hook_name_lookup_variants(name):
    n = (name or u"").strip()
    if not n:
        return []
    v = [n]
    if n.endswith(u"."):
        t = n[:-1].strip()
        if t and t not in v:
            v.append(t)
    else:
        t = n + u"."
        if t not in v:
            v.append(t)
    return v


def _resolve_rebar_hook_135_id(document, largo_mm):
    """
    Devuelve (ElementId, mensaje_error). mensaje_error es None si OK.
    Con nombre fijo: se prueban variantes (p. ej. con/sin punto final); si no hay coincidencia,
    se hace fallback por ángulo 135° o creación (mismo criterio que con nombre vacío).
    """
    n0 = (REBAR_HOOK_TYPE_NAME or u"").strip()
    if n0:
        for n in _rebar_hook_name_lookup_variants(n0):
            ht = _rebar_hook_type_by_name_in_document(document, n)
            if ht is not None:
                return ht.Id, None
    found = _find_rebar_hook_by_angle_deg(
        document, HOOK_ANGLE_DEG, HOOK_ANGLE_MATCH_TOL_DEG
    )
    if found is not None:
        return found.Id, None
    creado = _crear_rebar_hook_135(document, largo_mm, HOOK_EXTENSION_MULT)
    if creado is not None:
        return creado.Id, None
    if n0:
        return (
            ElementId.InvalidElementId,
            u"No se encontró el RebarHookType «{0}» (ni variantes) y no hay tipo 135° en el "
            u"proyecto (ni se pudo crear).".format(n0),
        )
    return (
        ElementId.InvalidElementId,
        u"No se pudo crear RebarHookType 135° (REBAR_HOOK_TYPE_NAME vacío y sin tipo por ángulo).",
    )


def _assign_135_hooks_both_ends(
    new_rb,
    document,
    rebar_bar_type,
    hook_id,
    orient0,
    orient1,
    rot0_deg,
    rot1_deg,
    extra_hook_eids=None,
):
    """
    Limpia extremos, ajusta Hook Length en el RebarBarType activo, reintenta longitudes
    (CanUseHookType) y fija 135° en 0 y 1. Patrón: enfierrado_shaft (clear + post-layout).
    Luego plano de la terminación (Left/Right) y rotación fuera de plano (grados) por extremo.
    Si extra_hook_eids: otros RebarHookType ~135° del proyecto (p. ej. estribo) por si el
    primero no lo acepta el RebarBarType.
    """
    if new_rb is None or hook_id is None or hook_id == ElementId.InvalidElementId:
        return False, u"HookId inválido."
    inv = ElementId.InvalidElementId
    fn = getattr(new_rb, "CanUseHookType", None)
    cands = _hook_length_candidates_mm_for_bar_type(rebar_bar_type)
    seq = [hook_id]
    seen = {_eid_int(hook_id)}
    for h in extra_hook_eids or ():
        if h is None or h == inv:
            continue
        e = _eid_int(h)
        if e and e not in seen:
            seen.add(e)
            seq.append(h)
    primary_eid = hook_id
    last_err = u""
    for hook_try in seq:
        hname = u""
        try:
            hel = document.GetElement(hook_try)
            if isinstance(hel, RebarHookType):
                hname = (hel.Name or u"").strip()
        except System.Exception:
            pass
        for Lmm in cands:
            _ensure_hook_length_on_bar_type(
                document, rebar_bar_type, hook_try, Lmm
            )
            try:
                document.Regenerate()
            except System.Exception:
                pass
            if fn is not None:
                try:
                    if not fn(hook_try):
                        last_err = u"CanUseHookType=False con L={0} mm (tipo: {1}); se intenta igual.".format(
                            int(Lmm), rebar_bar_type.Name if rebar_bar_type else u"?"
                        )
                except System.Exception as ex:
                    last_err = u"{0!s}".format(ex)
            for end in (0, 1):
                _set_rebar_hook_type_id_at_end(
                    new_rb, end, inv, document, 2
                )
            try:
                document.Regenerate()
            except System.Exception:
                pass
            # Sin Regenerate entre extremo 0 y 1: con solo un gancho, Revit puede validar
            # «Can't solve Rebar Shape» al regenerar; el gancha asimétrico no encaja aún.
            ok0 = _set_rebar_hook_type_id_at_end(
                new_rb, 0, hook_try, document, 3
            )
            ok1 = _set_rebar_hook_type_id_at_end(
                new_rb, 1, hook_try, document, 3
            )
            if ok0 and ok1:
                _set_hook_plane_orientation(new_rb, 0, orient0)
                _set_hook_plane_orientation(new_rb, 1, orient1)
                try:
                    document.Regenerate()
                except System.Exception:
                    pass
                rot_ok = True
                if abs(float(rot0_deg)) > 1e-9:
                    rot_ok = (
                        rot_ok
                        and _set_rebar_termination_or_hook_rotation_deg(
                            new_rb, 0, rot0_deg
                        )
                    )
                if abs(float(rot1_deg)) > 1e-9:
                    rot_ok = (
                        rot_ok
                        and _set_rebar_termination_or_hook_rotation_deg(
                            new_rb, 1, rot1_deg
                        )
                    )
                try:
                    document.Regenerate()
                except System.Exception:
                    pass
                base = u"Ganchos 135° con Hook Length \u2248{0} mm; rot(°) = ({1}, {2})".format(
                    int(Lmm), int(rot0_deg), int(rot1_deg)
                )
                if _eid_int(hook_try) != _eid_int(primary_eid) and hname:
                    base += u" (tipo gancho: «{0}»)".format(hname)
                if not rot_ok and (
                    abs(float(rot0_deg)) > 1e-9
                    or abs(float(rot1_deg)) > 1e-9
                ):
                    base += (
                        u" | AVISO: rotación fuera de plano no aplicada (API). "
                        u"Pruebe INVERTIR_NORMAL_REBAR=True o actualice Revit."
                    )
                return (True, _format_hook_state_debug(new_rb, base))
            last_err = u"SetHookTypeId no coincidió (ext.0={0}, ext.1={1}) con L={2} mm; gancho eid {3}.".format(
                ok0, ok1, int(Lmm), _eid_int(hook_try)
            )
    return (
        False,
        last_err
        or u"Revit no acepta el gancho 135° con este trazado o tipo de barra. Pruebe subir"
        u" HOOK_LENGTH_MM_135 / HOOK_LENGTH_FALLBACK o editar el RebarBarType (Hook Lengths).",
    )


def _assign_135_hook_solo_pata(
    new_rb,
    document,
    rebar_bar_type,
    hook_id,
    orient0,
    orient1,
    rot0_deg,
    rot1_deg,
    pata_en_extremo_final,
    extra_hook_eids=None,
):
    u"""
    Asigna gancho 135° **solo** en el extremo donde quedó la pata L (ext. 0 si
    ``pata_en_extremo_final`` es False, ext. 1 si es True) y deja el otro extremo sin gancho.
    """
    if new_rb is None or hook_id is None or hook_id == ElementId.InvalidElementId:
        return False, u"HookId inválido."
    inv = ElementId.InvalidElementId
    pata_end = 1 if pata_en_extremo_final else 0
    other_end = 1 - pata_end
    if pata_en_extremo_final:
        o_p = orient1
        rot_p = rot1_deg
    else:
        o_p = orient0
        rot_p = rot0_deg
    fn = getattr(new_rb, "CanUseHookType", None)
    cands = _hook_length_candidates_mm_for_bar_type(rebar_bar_type)
    seq = [hook_id]
    seen = {_eid_int(hook_id)}
    for h in extra_hook_eids or ():
        if h is None or h == inv:
            continue
        e = _eid_int(h)
        if e and e not in seen:
            seen.add(e)
            seq.append(h)
    primary_eid = hook_id
    last_err = u""
    for hook_try in seq:
        hname = u""
        try:
            hel = document.GetElement(hook_try)
            if isinstance(hel, RebarHookType):
                hname = (hel.Name or u"").strip()
        except System.Exception:
            pass
        for Lmm in cands:
            _ensure_hook_length_on_bar_type(
                document, rebar_bar_type, hook_try, Lmm
            )
            try:
                document.Regenerate()
            except System.Exception:
                pass
            if fn is not None:
                try:
                    if not fn(hook_try):
                        last_err = u"CanUseHookType=False con L={0} mm (tipo: {1}).".format(
                            int(Lmm), rebar_bar_type.Name if rebar_bar_type else u"?"
                        )
                except System.Exception as ex:
                    last_err = u"{0!s}".format(ex)
            for end in (0, 1):
                _set_rebar_hook_type_id_at_end(
                    new_rb, end, inv, document, 2
                )
            try:
                document.Regenerate()
            except System.Exception:
                pass
            okp = _set_rebar_hook_type_id_at_end(
                new_rb, pata_end, hook_try, document, 3
            )
            if not okp:
                last_err = u"Gancho 135 (solo pata) no en ext. {0}.".format(
                    pata_end
                )
                continue
            _set_rebar_hook_type_id_at_end(
                new_rb, other_end, inv, document, 3
            )
            try:
                document.Regenerate()
            except System.Exception:
                pass
            _set_hook_plane_orientation(new_rb, pata_end, o_p)
            try:
                document.Regenerate()
            except System.Exception:
                pass
            rot_ok = True
            if abs(float(rot_p)) > 1e-9:
                rot_ok = _set_rebar_termination_or_hook_rotation_deg(
                    new_rb, pata_end, rot_p
                )
            try:
                document.Regenerate()
            except System.Exception:
                pass
            base = u"Gancho 135° solo en pata (ext. {0}), L \u2248{1} mm; rot(°) pata = {2}".format(
                pata_end, int(Lmm), int(rot_p)
            )
            if _eid_int(hook_try) != _eid_int(primary_eid) and hname:
                base += u" (tipo: «{0}»)".format(hname)
            if not rot_ok and abs(float(rot_p)) > 1e-9:
                base += u" | AVISO: rot. pata no aplicada (API)."
            return (True, _format_hook_state_debug(new_rb, base))
    return (
        False,
        last_err
        or u"Revit no acepta 135° solo en un extremo; pruebe HOOK_LENGTH o tipo de barra.",
    )


def _rebar_chain_start_end_points(rebar, pos_idx):
    u"""Primer punto del 1.º tramo y último punto del último tramo (boceto de centroide)."""
    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        crvs = rebar.GetCenterlineCurves(
            False, False, False, mpo, int(pos_idx)
        )
    except System.Exception:
        return None, None
    if crvs is None or crvs.Count < 1:
        return None, None
    try:
        c0 = crvs[0]
        cN = crvs[crvs.Count - 1]
        return c0.GetEndPoint(0), cN.GetEndPoint(1)
    except System.Exception:
        return None, None


def pata_en_extremo_final_para_pie_por_elevacion(rebar, pos_idx):
    u"""
    Valor de ``pata_en_extremo_final`` para colocar el tramo L (y 135°) en el **pie** del
    trazado (cara **inferior** / menor Z) según el sentido de la polilínea en alzado.

    La API coloca el L al **inicio** del 1.º tramo (``False``) o a continuación del **último**
    punto (``True``). Si el boceto del vertical va *de arriba a abajo*, ``False`` pone el L
    arriba, no en la base: por eso se ajusta según `Z` de inicio vs fin.
    """
    p_start, p_end = _rebar_chain_start_end_points(rebar, pos_idx)
    if p_start is None or p_end is None:
        return False
    try:
        eps = UnitUtils.ConvertToInternalUnits(0.5, UnitTypeId.Millimeters)
    except System.Exception:
        eps = 1.0e-4
    try:
        dz = float(p_end.Z - p_start.Z)
        if dz > float(eps):
            return False
        if dz < -float(eps):
            return True
    except System.Exception:
        return False
    return False


def pata_en_extremo_final_para_cabeza_por_elevacion(rebar, pos_idx):
    u"""
    Igual que ``pata_en_extremo_final_para_pie_por_elevacion`` pero para el extremo
    **superior** (mayor Z), típico *sin forjado* en la cara superior del muro.
    """
    p_start, p_end = _rebar_chain_start_end_points(rebar, pos_idx)
    if p_start is None or p_end is None:
        return True
    try:
        eps = UnitUtils.ConvertToInternalUnits(0.5, UnitTypeId.Millimeters)
    except System.Exception:
        eps = 1.0e-4
    try:
        dz = float(p_end.Z - p_start.Z)
        if dz > float(eps):
            return True
        if dz < -float(eps):
            return False
    except System.Exception:
        return True
    return True


def _construir_cadena_doble_pata_cabeza_y_pie(
    chain,
    norm,
    le,
    invertir,
    pe_cab,
    pe_pie,
):
    u"""
    Polilínea con pata L en **ambos** extremos del trazado vertical (un solo ``CreateFromCurves``).
    Pares complementarios típicos: ``(pe_cab, pe_pie)`` = (False, True) o (True, False).
    """
    if chain is None or len(chain) < 1:
        return None, u"Sin tramos en cadena."
    # (False, True): cabeza = prep. 1.er tramo; pie = ap. último tramo
    if (not pe_cab) and pe_pie:
        c0 = chain[0]
        c_last = chain[-1]
        T = _tangent_start_first_curve(c0)
        if T is None:
            return None, u"Tangente nula (cabeza/pata 1)."
        B = _perp_in_plane(norm, T)
        if B is None:
            return None, u"B nula (cabeza)."
        if invertir:
            B = B.Negate()
        p0 = c0.GetEndPoint(0)
        p_leg = p0 - B.Multiply(le)
        leg_c = Line.CreateBound(p_leg, p0)
        T2 = _tangent_start_first_curve(c_last)
        if T2 is None:
            return None, u"Tangente nula (pie/pata 2)."
        B2 = _perp_in_plane(norm, T2)
        if B2 is None:
            return None, u"B nula (pie)."
        if invertir:
            B2 = B2.Negate()
        p_end = c_last.GetEndPoint(1)
        p_tip = p_end - B2.Multiply(le)
        leg_p = Line.CreateBound(p_end, p_tip)
        return [leg_c] + chain + [leg_p], None
    # (True, False): pie = prep.; cabeza = ap.
    if pe_cab and (not pe_pie):
        c0 = chain[0]
        c_last = chain[-1]
        T = _tangent_start_first_curve(c0)
        if T is None:
            return None, u"Tangente nula (pie/pata 1)."
        B = _perp_in_plane(norm, T)
        if B is None:
            return None, u"B nula (pie)."
        if invertir:
            B = B.Negate()
        p0 = c0.GetEndPoint(0)
        p_leg = p0 - B.Multiply(le)
        leg_p = Line.CreateBound(p_leg, p0)
        intermediate = [leg_p] + chain
        cL = intermediate[-1]
        T2 = _tangent_start_first_curve(cL)
        if T2 is None:
            return None, u"Tangente nula (cabeza/pata 2)."
        B2 = _perp_in_plane(norm, T2)
        if B2 is None:
            return None, u"B nula (cabeza)."
        if invertir:
            B2 = B2.Negate()
        p_end = cL.GetEndPoint(1)
        p_tip = p_end - B2.Multiply(le)
        leg_c = Line.CreateBound(p_end, p_tip)
        return intermediate + [leg_c], None
    return None, u"pe_cab/pe_pie no complementarios: {0}/{1}".format(pe_cab, pe_pie)


def extender_doble_pata_135_y_reemplazar(
    doc,
    rebar,
    largo_pata_mm,
    invertir,
    pos_idx=0,
):
    u"""
    Pata L + 135 en **cabeza y pie** en **una** transacción (polilínea continua). Evita el
    2.º ``extender_l…`` sobre rebar ya extendido, que en API devuelve InternalException (log).
    """
    if not isinstance(rebar, Rebar):
        return False, u"No es un Rebar.", None

    orig_rebar_id = rebar.Id

    host = doc.GetElement(rebar.GetHostId())
    if host is None:
        return False, u"Host inválido.", None

    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None

    try:
        style = rebar.Style
    except System.Exception:
        style = RebarStyle.Standard

    norm = _rebar_normal(rebar)
    if INVERTIR_NORMAL_REBAR:
        norm = norm.Negate()
    o0 = HOOK_ORIENT_END0
    o1 = HOOK_ORIENT_END1

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        dmm = _nominal_diameter_mm(bar_type)
        l0 = float(largo_pata_mm)
        l_min = max(12.0 * dmm, 100.0, float(PATA_LARGO_MIN_MM))
        largo_pata_mm = max(l0, l_min)
    except System.Exception:
        pass
    le = _mm_to_internal(largo_pata_mm)

    crvs = rebar.GetCenterlineCurves(
        False, False, False, mpo, int(pos_idx)
    )
    if crvs is None or crvs.Count == 0:
        return False, u"Sin GetCenterlineCurves para posición {0}.".format(pos_idx), None

    chain = [crvs[i] for i in range(crvs.Count)]
    pe_pie = pata_en_extremo_final_para_pie_por_elevacion(rebar, pos_idx)
    pe_cab = pata_en_extremo_final_para_cabeza_por_elevacion(rebar, pos_idx)
    new_chain, err_chain = _construir_cadena_doble_pata_cabeza_y_pie(
        chain, norm, le, invertir, pe_cab, pe_pie
    )
    if new_chain is None:
        return False, err_chain or u"Cadena doble pata no construida.", None

    t = Transaction(doc, u"BIMTools: Rebar doble pata L + ganchos 135°")
    out_eid = None
    try:
        _fho = t.GetFailureHandlingOptions()
        _fho.SetFailuresPreprocessor(_BimToolsRebarTxnFailuresPreprocessor())
        t.SetFailureHandlingOptions(_fho)
    except System.Exception:
        pass
    t.Start()
    try:
        hid, err_hook_resolve = _resolve_rebar_hook_135_id(doc, HOOK_LENGTH_MM_135)
        if err_hook_resolve is not None:
            t.RollBack()
            return False, err_hook_resolve, None
        if hid is None or hid == ElementId.InvalidElementId:
            t.RollBack()
            return False, u"RebarHookType 135° no resuelto.", None

        hook_eids_135 = _ordered_135_hook_type_eids(doc, hid, 6)
        h_primary = hook_eids_135[0] if hook_eids_135 else hid
        h_alts_135 = hook_eids_135[1:] if len(hook_eids_135) > 1 else []

        create_path = u""
        new_rb = _try_create_l_from_rebar_shape_2seg(
            doc, new_chain, host, norm, bar_type, style, o0, o1
        )
        if new_rb is not None:
            create_path = u"from_curves_and_shape_2seg"
        if new_rb is None:
            new_rb = _try_create_l_with_hook_types_both_ends(
                doc, new_chain, host, norm, bar_type, style, o0, o1, hid
            )
            if new_rb is not None:
                create_path = u"create_from_curves_with_hooks"
        if new_rb is None:
            new_rb = _create_from_curves_no_hooks(
                doc, new_chain, host, norm, bar_type, style, o0, o1
            )
            if new_rb is not None:
                create_path = u"create_from_curves"
        if new_rb is None:
            t.RollBack()
            return False, u"CreateFromCurves devolvió None (doble pata).", None

        ok_lay, err_lay = _copy_layout_rebar_shape_driven(rebar, new_rb)
        if not ok_lay:
            t.RollBack()
            return False, u"Layout: {0}".format(err_lay or u"error desconocido"), None

        # No usar el atajo de ganchos de CreateFromCurves: en 3+ tramos los 135° a menudo
        # apuntan hacia el «exterior»; reasignar con constantes DOBLE_PATA_GANCHO_*.
        ok_h, note_h = _assign_135_hooks_both_ends(
            new_rb,
            doc,
            bar_type,
            h_primary,
            DOBLE_PATA_GANCHO_ORIENT0,
            DOBLE_PATA_GANCHO_ORIENT1,
            DOBLE_PATA_GANCHO_ROT0_DEG,
            DOBLE_PATA_GANCHO_ROT1_DEG,
            extra_hook_eids=h_alts_135,
        )
        if not ok_h:
            fail_detail = (note_h or u"asignación 135").strip()
            try:
                _clear_both_hook_ends_rebar(new_rb, doc)
            except System.Exception as ex_c:
                t.RollBack()
                return (
                    False,
                    u"Ganchos: {0} | y no se pudo limpiar: {1!s}".format(
                        fail_detail, ex_c
                    ),
                    None,
                )
            note_h = u"(AVISO) Doble pata L; ganchos 135° no aplicados. {0}".format(
                (fail_detail)[:500]
            )

        out_eid = new_rb.Id
        t.Commit()
    except System.Exception as ex:
        t.RollBack()
        return False, u"{0!s}".format(ex), None

    t2 = Transaction(doc, u"BIMTools: Rebar doble pata — eliminar malla original")
    try:
        _fh2 = t2.GetFailureHandlingOptions()
        _fh2.SetFailuresPreprocessor(_BimToolsRebarTxnFailuresPreprocessor())
        t2.SetFailureHandlingOptions(_fh2)
    except System.Exception:
        pass
    t2.Start()
    try:
        if (
            orig_rebar_id is not None
            and orig_rebar_id != ElementId.InvalidElementId
        ):
            _old_r = doc.GetElement(orig_rebar_id)
            if _old_r is not None:
                doc.Delete(orig_rebar_id)
        t2.Commit()
    except System.Exception as ex2:
        try:
            t2.RollBack()
        except System.Exception:
            pass
        nmsgx = (note_h or u"").strip()
        note_h = (
            nmsgx
            + u" [AVISO: no se eliminó rebar malla org. (id {0}): {1!s} — ]"
        ).format(_eid_int(orig_rebar_id), ex2)

    fresh_rb = doc.GetElement(out_eid)
    if fresh_rb is None or not isinstance(fresh_rb, Rebar):
        return (
            False,
            u"Rebar doble pata no resuelto (id {0}).".format(_eid_int(out_eid)),
            None,
        )
    nnum = int(round(float(largo_pata_mm)))
    nmsg = (note_h or u"").strip()
    return (
        True,
        u"Listo (doble pata L). Nuevo Rebar id={0}. Pata L={1} mm. {2}".format(
            _eid_int(fresh_rb.Id), nnum, nmsg
        ),
        fresh_rb,
    )


def extender_l_asignar_ganchos_135_y_reemplazar(
    doc,
    rebar,
    largo_pata_mm,
    invertir,
    pos_idx=0,
    pata_en_extremo_final=False,
    gancho_135_solo_en_extremo_pata=False,
    hook_orient_end0=None,
    hook_orient_end1=None,
    hook_rot_end0_deg=None,
    hook_rot_end1_deg=None,
):
    if not isinstance(rebar, Rebar):
        return False, u"No es un Rebar.", None

    orig_rebar_id = rebar.Id

    host = doc.GetElement(rebar.GetHostId())
    if host is None:
        return False, u"Host inválido.", None

    bar_type = doc.GetElement(rebar.GetTypeId())
    if not isinstance(bar_type, RebarBarType):
        return False, u"RebarBarType no resuelto.", None

    try:
        style = rebar.Style
    except System.Exception:
        style = RebarStyle.Standard

    norm = _rebar_normal(rebar)
    if INVERTIR_NORMAL_REBAR:
        norm = norm.Negate()
    # Misma orientación en creación (CreateFromCurves) y al asignar ganchos; anulables
    # por el caller (p. ej. muro unido en cara superior — constantes *MURO_CARA_SUP*).
    o0 = (
        hook_orient_end0
        if hook_orient_end0 is not None
        else HOOK_ORIENT_END0
    )
    o1 = (
        hook_orient_end1
        if hook_orient_end1 is not None
        else HOOK_ORIENT_END1
    )
    r0h = (
        float(hook_rot_end0_deg)
        if hook_rot_end0_deg is not None
        else float(HOOK_ROTATION_END0_DEG)
    )
    r1h = (
        float(hook_rot_end1_deg)
        if hook_rot_end1_deg is not None
        else float(HOOK_ROTATION_END1_DEG)
    )

    mpo = MultiplanarOption.IncludeAllMultiplanarCurves
    try:
        dmm = _nominal_diameter_mm(bar_type)
        l0 = float(largo_pata_mm)
        l_min = max(12.0 * dmm, 100.0, float(PATA_LARGO_MIN_MM))
        largo_pata_mm = max(l0, l_min)
    except System.Exception:
        pass
    le = _mm_to_internal(largo_pata_mm)

    crvs = rebar.GetCenterlineCurves(
        False, False, False, mpo, int(pos_idx)
    )
    if crvs is None or crvs.Count == 0:
        return False, u"Sin GetCenterlineCurves para posición {0}.".format(pos_idx), None

    chain = [crvs[i] for i in range(crvs.Count)]
    if not pata_en_extremo_final:
        c0 = chain[0]
        T = _tangent_start_first_curve(c0)
        if T is None:
            return False, u"Tangente nula en el primer tramo.", None
        B = _perp_in_plane(norm, T)
        if B is None:
            return False, u"No se pudo calcular la perpendicular al plano.", None
        if invertir:
            B = B.Negate()
        p0 = c0.GetEndPoint(0)
        p_leg = p0 - B.Multiply(le)
        leg = Line.CreateBound(p_leg, p0)
        new_chain = [leg] + chain
    else:
        c_last = chain[-1]
        T = _tangent_start_first_curve(c_last)
        if T is None:
            return False, u"Tangente nula en el último tramo.", None
        B = _perp_in_plane(norm, T)
        if B is None:
            return False, u"No se pudo calcular la perpendicular al plano.", None
        if invertir:
            B = B.Negate()
        # El último tramo termina en p_end: el siguiente tramo **debe empezar en p_end** o Revit
        # rechaza el CurveLoop («curves do not form a valid CurveLoop»). Hacia el interior se usa
        # p_end - B·L (antes p_end + B·L llevaba la pata al lado opuesto, «hacia afuera»).
        p_end = c_last.GetEndPoint(1)
        p_tip = p_end - B.Multiply(le)
        leg = Line.CreateBound(p_end, p_tip)
        new_chain = chain + [leg]

    t = Transaction(doc, u"BIMTools: Rebar L + ganchos 135° (nueva + layout + gancho)")
    out_eid = None
    try:
        _fho = t.GetFailureHandlingOptions()
        _fho.SetFailuresPreprocessor(_BimToolsRebarTxnFailuresPreprocessor())
        t.SetFailureHandlingOptions(_fho)
    except System.Exception:
        pass
    t.Start()
    try:
        hid, err_hook_resolve = _resolve_rebar_hook_135_id(doc, HOOK_LENGTH_MM_135)
        if err_hook_resolve is not None:
            t.RollBack()
            return False, err_hook_resolve, None
        if hid is None or hid == ElementId.InvalidElementId:
            t.RollBack()
            return False, u"RebarHookType 135° no resuelto.", None

        hook_eids_135 = _ordered_135_hook_type_eids(doc, hid, 6)
        h_primary = hook_eids_135[0] if hook_eids_135 else hid
        h_alts_135 = hook_eids_135[1:] if len(hook_eids_135) > 1 else []

        create_path = u""
        new_rb = _try_create_l_from_rebar_shape_2seg(
            doc, new_chain, host, norm, bar_type, style, o0, o1
        )
        if new_rb is not None:
            create_path = u"from_curves_and_shape_2seg"
        if new_rb is None:
            new_rb = _try_create_l_with_hook_types_both_ends(
                doc, new_chain, host, norm, bar_type, style, o0, o1, hid
            )
            if new_rb is not None:
                create_path = u"create_from_curves_with_hooks"
        if new_rb is None:
            new_rb = _create_from_curves_no_hooks(
                doc, new_chain, host, norm, bar_type, style, o0, o1
            )
            if new_rb is not None:
                create_path = u"create_from_curves"
        if new_rb is None:
            t.RollBack()
            return False, u"CreateFromCurves devolvió None.", None

        ok_lay, err_lay = _copy_layout_rebar_shape_driven(rebar, new_rb)
        if not ok_lay:
            t.RollBack()
            return False, u"Layout: {0}".format(err_lay or u"error desconocido"), None

        # No Regenerate aquí: obligar a resolver RebarShape antes de fijar ganchos 135
        # provocaba "Can't solve Rebar Shape" / gancho incompatible con la forma.

        def _both_ends_have_hook_type(rb, hid2):
            try:
                if _eid_int(rb.GetHookTypeId(0)) != _eid_int(hid2):
                    return False
                if _eid_int(rb.GetHookTypeId(1)) != _eid_int(hid2):
                    return False
                return True
            except System.Exception:
                return False

        inv = ElementId.InvalidElementId
        created_with_builtin_hooks = create_path == u"create_from_curves_with_hooks"
        created_both_ok = (
            created_with_builtin_hooks
            and _both_ends_have_hook_type(new_rb, hid)
        )
        ok_h = True
        note_h = u""
        if created_both_ok and not gancho_135_solo_en_extremo_pata:
            note_h = u"Ganchos 135° en CreateFromCurves; layout copiado."
        elif created_both_ok and gancho_135_solo_en_extremo_pata:
            other_end = 0 if pata_en_extremo_final else 1
            pata_end = 1 - other_end
            try:
                _set_rebar_hook_type_id_at_end(
                    new_rb, other_end, inv, doc, 3
                )
            except System.Exception:
                pass
            try:
                doc.Regenerate()
            except System.Exception:
                pass
            # CreateFromCurves dejó 135 en ambos extremos; al quitar uno, la orientación y
            # la rotación del extremo pata no siempre se reflejan (p. ej. 180° en kwargs).
            o_p = o1 if pata_en_extremo_final else o0
            rot_p = float(r1h if pata_en_extremo_final else r0h)
            try:
                _set_hook_plane_orientation(new_rb, pata_end, o_p)
            except System.Exception:
                pass
            try:
                doc.Regenerate()
            except System.Exception:
                pass
            try:
                if abs(rot_p) > 1e-9:
                    _set_rebar_termination_or_hook_rotation_deg(
                        new_rb, pata_end, rot_p
                    )
            except System.Exception:
                pass
            try:
                doc.Regenerate()
            except System.Exception:
                pass
            note_h = (
                u"Gancho 135° en extremo pata; extremo {0} sin gancho 135 (CreateFromCurves)."
            ).format(other_end)
        else:
            if gancho_135_solo_en_extremo_pata:
                ok_h, note_h = _assign_135_hook_solo_pata(
                    new_rb,
                    doc,
                    bar_type,
                    h_primary,
                    o0,
                    o1,
                    r0h,
                    r1h,
                    pata_en_extremo_final,
                    extra_hook_eids=h_alts_135,
                )
                if not ok_h:
                    ok2, n2 = _assign_135_hooks_both_ends(
                        new_rb,
                        doc,
                        bar_type,
                        h_primary,
                        o0,
                        o1,
                        r0h,
                        r1h,
                        extra_hook_eids=h_alts_135,
                    )
                    if ok2:
                        other_end = 0 if pata_en_extremo_final else 1
                        try:
                            _set_rebar_hook_type_id_at_end(
                                new_rb, other_end, inv, doc, 3
                            )
                        except System.Exception:
                            pass
                        try:
                            doc.Regenerate()
                        except System.Exception:
                            pass
                        ok_h = True
                        note_h = (n2 or u"OK") + (
                            u" | 135° solo en pata: extremo {0} limpiado."
                        ).format(other_end)
                    else:
                        ok_h = False
                        note_h = n2
            else:
                ok_h, note_h = _assign_135_hooks_both_ends(
                    new_rb,
                    doc,
                    bar_type,
                    h_primary,
                    o0,
                    o1,
                    r0h,
                    r1h,
                    extra_hook_eids=h_alts_135,
                )
        if not ok_h:
            # Modo degradado: conservar pata L y quitar ganchos (evita c_l135_fail total
            # y «Can't solve Rebar Shape» por intentos 135 inválidos en polilínea).
            fail_detail = (note_h or u"asignación 135").strip()
            try:
                _clear_both_hook_ends_rebar(new_rb, doc)
            except System.Exception as ex_c:
                t.RollBack()
                return (
                    False,
                    u"Ganchos: {0} | y no se pudo limpiar: {1!s}".format(
                        fail_detail, ex_c
                    ),
                    None,
                )
            note_h = u"(AVISO) Pata L generada; ganchos 135° no aplicados por el modelo. {0}".format(
                (fail_detail)[:500]
            )

        out_eid = new_rb.Id
        t.Commit()
    except System.Exception as ex:
        t.RollBack()
        return False, u"{0!s}".format(ex), None

    t2 = Transaction(doc, u"BIMTools: Rebar L+135 — eliminar rebar (malla) original")
    try:
        _fh2 = t2.GetFailureHandlingOptions()
        _fh2.SetFailuresPreprocessor(_BimToolsRebarTxnFailuresPreprocessor())
        t2.SetFailureHandlingOptions(_fh2)
    except System.Exception:
        pass
    t2.Start()
    try:
        if (
            orig_rebar_id is not None
            and orig_rebar_id != ElementId.InvalidElementId
        ):
            _old_r = doc.GetElement(orig_rebar_id)
            if _old_r is not None:
                doc.Delete(orig_rebar_id)
        t2.Commit()
    except System.Exception as ex2:
        try:
            t2.RollBack()
        except System.Exception:
            pass
        nmsgx = (note_h or u"").strip()
        note_h = (
            nmsgx
            + u" [AVISO: no se eliminó rebar malla org. (id {0}): {1!s} — " u"revisar duplicado.]"
        ).format(_eid_int(orig_rebar_id), ex2)

    # Tras Commit el wrapper de Element puede quedar inválido; re-leer (evita
    # "The referenced object is not valid" al usar .Id / .IntegerValue).
    fresh_rb = doc.GetElement(out_eid)
    if fresh_rb is None or not isinstance(fresh_rb, Rebar):
        return (
            False,
            u"Rebar L+135 no resuelto tras transacción (id {0}).".format(
                _eid_int(out_eid)
            ),
            None,
        )
    nnum = int(round(float(largo_pata_mm)))
    nmsg = (note_h or u"").strip()
    return (
        True,
        u"Listo. Nuevo Rebar id={0}. Pata L={1} mm. {2}".format(
            _eid_int(fresh_rb.Id), nnum, nmsg
        ),
        fresh_rb,
    )


def run(uidoc):
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    if ids is None or ids.Count == 0:
        TaskDialog.Show(u"Rebar L + 135°", u"Selecciona un Rebar.")
        return
    for eid in ids:
        el = doc.GetElement(eid)
        if isinstance(el, Rebar):
            host = doc.GetElement(el.GetHostId())
            largo = largo_pata_mm_desde_espesor_host(doc, host)
            ok, msg, _nrb = extender_l_asignar_ganchos_135_y_reemplazar(
                doc,
                el,
                largo,
                INVERTIR_DIRECCION_PATA,
                INDICE_POSICION,
                PATA_EN_EXTREMO_FINAL,
            )
            print(msg)
            TaskDialog.Show(u"Rebar L + 135°", msg)
            return
    TaskDialog.Show(u"Rebar L + 135°", u"Ninguna selección es Rebar.")


if __name__ == u"__main__":
    run(__revit__.ActiveUIDocument)  # noqa: F821
