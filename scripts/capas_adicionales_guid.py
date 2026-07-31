# -*- coding: utf-8 -*-
"""
Arainco: Capas adicionales por GUID.

Flujo:
1. Al ejecutar: ``PickObject`` inmediato de Rebar → leer GUID → analizar.
   Cancel / sin GUID → aviso y **no** abrir la UI.
2. Mostrar ventana modeless ya poblada (canvas sección + cards).
3. «Cambiar selección» re-pick y refresca UI (Hide → Pick → Show).
4. Capas longitudinales solo con ``Armadura_Capa`` no vacío (trabas/estribos
   del mismo GUID no van a la tabla ni a «última capa»).
5. Crear N capas nuevas hacia el interior con distanciamiento único.
6. Heredar GUID; ``Armadura_Capa`` consecutivo (máx. + 1, +2, …).
7. Empalmes / traslapos por **paridad** + Detail Items + tags.
8. Trabas nuevas: opt-in (toggle). OFF = no copiar. ON = copias offset
   aditivas de **trabas** (nunca estribos perimetrales); el estribo existente
   no se regenera ni amplía.
9. Creación vía ``ExternalEvent`` (contexto API válido).

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

import math
import re

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    CornerRadius,
    FontWeights,
    GridLength,
    GridUnitType,
    HorizontalAlignment,
    Point as WpfPoint,
    Size as WpfSize,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    WindowState,
)
from System.Windows.Controls import (
    Border,
    Canvas,
    ColumnDefinition,
    Grid,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import (
    ArcSegment,
    Color,
    LineSegment,
    PathFigure,
    PathGeometry,
    PenLineCap,
    PenLineJoin,
    SolidColorBrush,
    SweepDirection,
    TranslateTransform,
)
from System.Windows.Shapes import Ellipse, Line, Path as WpfPath, Rectangle

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    Line as RevitLine,
    LocationCurve,
    Transaction,
    Transform,
    XYZ,
    Wall,
)
from Autodesk.Revit.DB.Structure import MultiplanarOption, Rebar, RebarBarType, RebarStyle
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler, TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from armado_muros_rebar_params import (
    collect_empalmes_por_conjunto_guid,
    collect_rebars_por_conjunto_guid,
    get_armadura_conjunto_guid,
    set_armadura_capa_desde_layer,
    stamp_armadura_capa_desde_layer,
    stamp_armadura_conjunto_guid,
)
try:
    from armado_muros_lap_detail_shared import _find_fixed_lap_detail_symbol_id
except Exception:
    _find_fixed_lap_detail_symbol_id = None
try:
    from enfierrado_shaft_hashtag import _place_line_based_detail_component
except Exception:
    _place_line_based_detail_component = None
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_ui_tokens import (
    ACCENT_PRIMARY,
    BG_APP,
    BG_GROUP_HEADER,
    BG_PANEL,
    BG_PANEL_ELEVATED,
    BORDER,
    FG_BODY,
    FG_MUTED,
    FG_TITLE,
    WINDOW_CHROME_TITLE,
)

# Preview canvas (alineado a mockup / ArmadoMurosV3)
_COLOR_CONCRETE = u"#1A2A38"
_COLOR_EXISTING = u"#3D8B6E"
_COLOR_PROPOSED = u"#5BC0DE"
_COLOR_STIRRUP = u"#00b450"
_COLOR_TIE = u"#f59e0b"
_STIRRUP_STROKE = 2.2
_TIE_STROKE = 1.6
_COVER_EXT_MM_DEFAULT = 40
_N_CAPAS_OPTIONS = tuple(range(1, 9))
_QTY_OPTIONS = tuple(range(1, 21))
from revit_wpf_window_position import (
    bind_center_wpf_on_revit_monitor,
    position_wpf_window_center_on_monitor,
    revit_main_hwnd,
)
from seleccionar_capa_armadura import (
    collect_empalmes_por_conjunto_guid_y_capa,
    get_armadura_capa,
)
try:
    from conjunto_guid import stamp_armadura_arainco
except Exception:
    stamp_armadura_arainco = None

_APPDOMAIN_WINDOW_KEY = u"Arainco_CapasAdicionalesGuid_UI"
_APPDOMAIN_CTRL_KEY = u"Arainco_CapasAdicionalesGuid_Ctrl"
_TOOL_TITLE = u"Arainco: Capas adicionales por GUID"
_TXN_NAME = u"Arainco: Capas adicionales en Muro"

_DIAMETER_OPTIONS_MM = (8, 10, 12, 16, 18, 22, 25, 28, 32)
_CAPA_RE = re.compile(ur"\((\d+)\s*[ºo°]?\s*[Cc]\.?\)")

# Familias de Structural Rebar Tag usadas en Armadura (orden de preferencia).
_TAG_FAMILY_CANDIDATES = (
    u"EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL",
    u"EST_A_STRUCTURAL REBAR TAG_HORIZONTAL",
    u"EST_A_STRUCTURAL REBAR TAG",
    u"EST_A_STRUCTURAL REBAR TAG_MALLA",
)


# ── helpers texto / avisos ───────────────────────────────────────────────────


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _mostrar_aviso(uiapp, instruction, content=u"", ok_text=u"Entendido"):
    hwnd = None
    try:
        if uiapp is not None:
            hwnd = revit_main_hwnd(uiapp)
    except Exception:
        pass
    try:
        from bimtools_instruction_dialog import show_message_dialog

        show_message_dialog(
            _TOOL_TITLE,
            instruction,
            content=content,
            ok_text=ok_text,
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        body = instruction
        if content:
            body = instruction + u"\n\n" + content
        TaskDialog.Show(_TOOL_TITLE, body)
    except Exception:
        pass


def _element_id_int(eid):
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid.Value)
        except Exception:
            return None


def _mm_to_ft(mm):
    return float(mm) / 304.8


def _ft_to_mm(ft):
    return float(ft) * 304.8


# ── lectura Rebar ────────────────────────────────────────────────────────────


def _cantidad_posiciones(rebar):
    best = 1
    for getter in (
        lambda: int(rebar.NumberOfBarPositions),
        lambda: int(rebar.GetNumberOfBarPositions()),
        lambda: int(rebar.Quantity),
    ):
        try:
            n = int(getter())
            if n > best:
                best = n
        except Exception:
            pass
    return best


def _spacing_mm(rebar):
    try:
        sp = float(rebar.MaxSpacing)
        if sp > 1e-12:
            return int(round(_ft_to_mm(sp)))
    except Exception:
        pass
    return 0


def _bar_type_of(rebar, doc):
    if rebar is None or doc is None:
        return None
    try:
        tid = rebar.GetTypeId()
        bt = doc.GetElement(tid)
        if isinstance(bt, RebarBarType):
            return bt
    except Exception:
        pass
    return None


def _nominal_diam_mm(bar_type):
    if not isinstance(bar_type, RebarBarType):
        return None
    try:
        d = int(round(float(bar_type.BarNominalDiameter) * 304.8))
        return d if d > 0 else None
    except Exception:
        return None


def _resolver_bar_type_mm(doc, target_mm):
    if doc is None:
        return None
    best = None
    best_delta = None
    target = float(target_mm)
    try:
        types = FilteredElementCollector(doc).OfClass(RebarBarType)
    except Exception:
        return None
    for bt in types:
        try:
            d_mm = float(bt.BarNominalDiameter) * 304.8
        except Exception:
            continue
        delta = abs(d_mm - target)
        if best is None or delta < best_delta:
            best = bt
            best_delta = delta
    if best is None:
        return None
    if best_delta is not None and best_delta > 0.75:
        return None
    return best


def _rebar_style(rebar):
    try:
        return rebar.Style
    except Exception:
        return None


def _is_stirrup_tie(rebar):
    try:
        return _rebar_style(rebar) == RebarStyle.StirrupTie
    except Exception:
        return False


def _parse_capa_index(capa_txt):
    """``(1ºC.)`` → 0-based index; ``None`` si no parsea."""
    if not capa_txt:
        return None
    m = _CAPA_RE.search(_as_unicode(capa_txt))
    if not m:
        return None
    try:
        n = int(m.group(1))
    except Exception:
        return None
    if n < 1:
        return None
    return n - 1


def _host_of(rebar, doc):
    if rebar is None or doc is None:
        return None
    try:
        hid = rebar.GetHostId()
        if hid is None or hid == ElementId.InvalidElementId:
            return None
        return doc.GetElement(hid)
    except Exception:
        return None


def _host_label(host, doc):
    if host is None:
        return u"(sin host)"
    name = u""
    try:
        name = _as_unicode(host.Name).strip()
    except Exception:
        pass
    cat = u""
    try:
        if host.Category is not None:
            cat = _as_unicode(host.Category.Name).strip()
    except Exception:
        pass
    level = u""
    try:
        lid = getattr(host, u"LevelId", None)
        if lid is not None and lid != ElementId.InvalidElementId and doc is not None:
            lev = doc.GetElement(lid)
            if lev is not None:
                level = _as_unicode(lev.Name).strip()
    except Exception:
        pass
    parts = []
    if cat:
        parts.append(cat)
    if name:
        parts.append(name)
    base = u" ".join(parts) if parts else u"Host Id {0}".format(
        _element_id_int(host.Id) or u"?"
    )
    if level:
        return u"{0} · {1}".format(base, level)
    return base


def _rebar_midpoint(rebar):
    """Centroide aproximado de la barra 0 (pose de layout / centerline)."""
    # Preferir origen del layout (posición del set): más estable para Δ entre capas
    # que promediar todos los tramos/hooks de la centerline.
    try:
        t = rebar.GetBarPositionTransform(0)
        if t is not None:
            return t.Origin
    except Exception:
        pass
    curves = None
    # Preferir curvas transformadas (incluyen MoveBarInSet / pose visual).
    try:
        curves = rebar.GetTransformedCenterlineCurves(
            False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0
        )
    except Exception:
        curves = None
    if curves is None:
        try:
            curves = rebar.GetCenterlineCurves(False, False, False)
        except Exception:
            curves = None
    if curves is None:
        return None
    pts = []
    try:
        n = int(curves.Count)
    except Exception:
        n = 0
    for i in range(n):
        try:
            c = curves[i]
            if c is None:
                continue
            pts.append(c.Evaluate(0.5, True))
        except Exception:
            try:
                pts.append(c.GetEndPoint(0))
                pts.append(c.GetEndPoint(1))
            except Exception:
                pass
    if not pts:
        return None
    sx = sy = sz = 0.0
    for p in pts:
        sx += float(p.X)
        sy += float(p.Y)
        sz += float(p.Z)
    npts = float(len(pts))
    return XYZ(sx / npts, sy / npts, sz / npts)


def _bbox_center(elem):
    if elem is None:
        return None
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None
        return XYZ(
            0.5 * (float(bb.Min.X) + float(bb.Max.X)),
            0.5 * (float(bb.Min.Y) + float(bb.Max.Y)),
            0.5 * (float(bb.Min.Z) + float(bb.Max.Z)),
        )
    except Exception:
        return None


def _layer_centroid(rebars):
    pts = []
    for rb in rebars or []:
        p = _rebar_midpoint(rb)
        if p is not None:
            pts.append(p)
    if not pts:
        return None
    sx = sy = sz = 0.0
    for p in pts:
        sx += float(p.X)
        sy += float(p.Y)
        sz += float(p.Z)
    n = float(len(pts))
    return XYZ(sx / n, sy / n, sz / n)


def _normalize_xyz(v):
    if v is None:
        return None
    try:
        L = float(v.GetLength())
        if L < 1e-12:
            return None
        return v.Normalize()
    except Exception:
        return None


def _xyz_sub(a, b):
    """a - b (Subtract explícito; más fiable que el operador en IronPython)."""
    if a is None or b is None:
        return None
    try:
        return a.Subtract(b)
    except Exception:
        try:
            return XYZ(
                float(a.X) - float(b.X),
                float(a.Y) - float(b.Y),
                float(a.Z) - float(b.Z),
            )
        except Exception:
            return None


def _xyz_scale(v, s):
    if v is None:
        return None
    try:
        return v.Multiply(float(s))
    except Exception:
        try:
            return XYZ(float(v.X) * float(s), float(v.Y) * float(s), float(v.Z) * float(s))
        except Exception:
            return None


def _dist_point_to_point(a, b):
    if a is None or b is None:
        return None
    try:
        return float(a.DistanceTo(b))
    except Exception:
        d = _xyz_sub(a, b)
        if d is None:
            return None
        try:
            return float(d.GetLength())
        except Exception:
            return None


def _wall_thickness_axis(host):
    """Eje unitario del espesor del muro (``Wall.Orientation``, sin signo de cara)."""
    if not isinstance(host, Wall):
        return None
    try:
        return _normalize_xyz(host.Orientation)
    except Exception:
        return None


def _sign_toward_host(axis, last_c, host_c):
    """
    Devuelve ``axis`` o ``-axis`` de modo que apunte desde ``last_c`` hacia
    el centro del host (hacia el volumen de hormigón).
    """
    if axis is None:
        return None
    if last_c is None or host_c is None:
        return axis
    toward = _xyz_sub(host_c, last_c)
    if toward is None:
        return axis
    try:
        if float(toward.GetLength()) < 1e-9:
            return axis
        if float(axis.DotProduct(toward)) < 0.0:
            return axis.Negate()
    except Exception:
        pass
    return axis


def _project_onto_unit(v, axis):
    """Proyecta ``v`` sobre el eje unitario ``axis`` (resultado colineal con axis)."""
    if v is None or axis is None:
        return None
    try:
        s = float(v.DotProduct(axis))
        return _xyz_scale(axis, s)
    except Exception:
        return None


def _remove_component_along(v, axis):
    """Quita de ``v`` la componente paralela a ``axis`` (unitario)."""
    if v is None:
        return None
    if axis is None:
        return v
    try:
        s = float(v.DotProduct(axis))
        parallel = _xyz_scale(axis, s)
        return _xyz_sub(v, parallel)
    except Exception:
        return v


def _rebar_tangent_unit(rebar):
    """Tangente unitaria aproximada de la barra 0 (eje longitudinal)."""
    if rebar is None:
        return None
    curves = None
    try:
        curves = rebar.GetTransformedCenterlineCurves(
            False, False, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0
        )
    except Exception:
        curves = None
    if curves is None:
        try:
            curves = rebar.GetCenterlineCurves(False, False, False)
        except Exception:
            curves = None
    if curves is None:
        return None
    try:
        n = int(curves.Count)
    except Exception:
        n = 0
    if n < 1:
        return None
    try:
        c0 = curves[0]
        p0 = c0.GetEndPoint(0)
        p1 = c0.GetEndPoint(1)
        return _normalize_xyz(_xyz_sub(p1, p0))
    except Exception:
        return None


def _layer_bar_tangent(layer):
    """Tangente media de las rebars de una capa (para ortogonalizar el offset)."""
    if not layer:
        return None
    rebars = layer.get(u"rebars") or []
    sx = sy = sz = 0.0
    n = 0
    for rb in rebars:
        t = _rebar_tangent_unit(rb)
        if t is None:
            continue
        if n > 0:
            try:
                ref = XYZ(sx / float(n), sy / float(n), sz / float(n))
                if float(ref.DotProduct(t)) < 0.0:
                    t = t.Negate()
            except Exception:
                pass
        sx += float(t.X)
        sy += float(t.Y)
        sz += float(t.Z)
        n += 1
    if n < 1:
        return None
    return _normalize_xyz(XYZ(sx / float(n), sy / float(n), sz / float(n)))


def _toward_len_ft(last_c, host_c):
    if last_c is None or host_c is None:
        return 0.0
    toward = _xyz_sub(host_c, last_c)
    if toward is None:
        return 0.0
    try:
        return float(toward.GetLength())
    except Exception:
        return 0.0


def _stack_on_axis(cents, axis):
    """Proyección normalizada del vector capa1→última sobre ``axis`` (±axis)."""
    if axis is None or not cents or len(cents) < 2:
        return None
    if cents[0] is None or cents[-1] is None:
        return None
    stack = _xyz_sub(cents[-1], cents[0])
    proj = _project_onto_unit(stack, axis)
    if proj is None:
        return None
    try:
        if abs(float(proj.GetLength())) < 1e-6:
            return None
    except Exception:
        return None
    return _normalize_xyz(proj)


def _inward_direction(layers_sorted, host):
    """
    Dirección unitaria para capas **después de la última** del GUID, hacia el
    interior del host y **paralela al apilado de capas existentes**.

    No usar ``host_center - last_centroid`` como eje bruto: en muros largos ese
    vector arrastra el offset a lo largo del muro (barras «en el extremo» con
    orientación aparente incorrecta). Preferir:

    1. Stack C1→Cúltima proyectado al espesor del muro (``Wall.Orientation``).
    2. Stack C1→Cúltima ortogonal a la tangente de barra (otros hosts).
    3. Una sola capa: ±Orientation (muro) o (centro−last) ortogonal a la barra.
    4. Signo: continuar el stack y, si se puede medir, hacia el centro del host.
    """
    if not layers_sorted:
        return None
    cents = []
    for ly in layers_sorted:
        c = ly.get(u"centroid")
        if c is None:
            c = _layer_centroid(ly.get(u"rebars") or [])
        cents.append(c)
    last_c = cents[-1] if cents else None
    host_c = _bbox_center(host)
    toward_len = _toward_len_ft(last_c, host_c)
    thick = _wall_thickness_axis(host)
    bar_tan = _layer_bar_tangent(layers_sorted[-1])

    direction = None

    # 1) Varias capas: continuar el apilado existente (eje de capas reales).
    if len(cents) >= 2 and cents[0] is not None and last_c is not None:
        stack_raw = _xyz_sub(last_c, cents[0])
        if thick is not None:
            # Solo componente de espesor → evita offset a lo largo del muro.
            direction = _stack_on_axis(cents, thick)
        if direction is None and stack_raw is not None:
            # Quitar componente longitudinal de la barra (empalmes / tramos).
            cleaned = _remove_component_along(stack_raw, bar_tan)
            direction = _normalize_xyz(cleaned if cleaned is not None else stack_raw)
            try:
                if direction is not None and float(direction.GetLength()) < 1e-9:
                    direction = None
            except Exception:
                pass

    # 2) Una sola capa / stack degenerado: eje de espesor o hacia el centro
    #    pero siempre ortogonal a la barra (nunca a lo largo del muro).
    if direction is None:
        if thick is not None:
            direction = thick
        elif toward_len >= 1e-4:
            toward = _xyz_sub(host_c, last_c)
            cleaned = _remove_component_along(toward, bar_tan)
            direction = _normalize_xyz(cleaned if cleaned is not None else toward)

    if direction is None:
        return None

    # Signo: hacia el interior del host desde la última capa.
    if toward_len >= 1e-4:
        direction = _sign_toward_host(direction, last_c, host_c)
    elif len(cents) >= 2 and cents[0] is not None and last_c is not None:
        # Sin centro usable: continuar C1→Cúltima.
        stack = _xyz_sub(last_c, cents[0])
        try:
            if stack is not None and float(direction.DotProduct(stack)) < 0.0:
                direction = direction.Negate()
        except Exception:
            pass

    # Validación: un paso de prueba debe acercar al centro **medido sobre el
    # mismo eje** (no distancia 3D al bbox: falla en muros largos / trabas).
    if direction is not None and last_c is not None and host_c is not None and toward_len >= 1e-4:
        try:
            toward = _xyz_sub(host_c, last_c)
            axis_comp0 = float(toward.DotProduct(direction))
            step = _xyz_scale(direction, _mm_to_ft(10.0))
            probe = last_c.Add(step) if step is not None else None
            toward1 = _xyz_sub(host_c, probe) if probe is not None else None
            axis_comp1 = (
                float(toward1.DotProduct(direction)) if toward1 is not None else None
            )
            # Acercarse al centro ⇒ componente hacia el centro disminuye.
            if axis_comp1 is not None and axis_comp1 > axis_comp0 + 1e-9:
                direction = direction.Negate()
                step2 = _xyz_scale(direction, _mm_to_ft(10.0))
                probe2 = last_c.Add(step2) if step2 is not None else None
                toward2 = _xyz_sub(host_c, probe2) if probe2 is not None else None
                axis_comp2 = (
                    float(toward2.DotProduct(direction))
                    if toward2 is not None
                    else None
                )
                if axis_comp2 is None or axis_comp2 > axis_comp0 + 1e-9:
                    # Si el stack de capas ya define el eje, confiar en él.
                    if len(cents) < 2:
                        direction = None
        except Exception:
            pass

    return direction


def _moved_along_inward(mid_before, mid_after, inward_unit, expected_ft, min_frac=0.4):
    """
    True si el desplazamiento medido va en el sentido de ``inward_unit``
    (proyección ≥ ``min_frac`` · esperado). False si va al contrario.
    None si no se puede medir.
    """
    if mid_before is None or mid_after is None or inward_unit is None:
        return None
    try:
        disp = _xyz_sub(mid_after, mid_before)
        if disp is None:
            return None
        proj = float(disp.DotProduct(inward_unit))
    except Exception:
        return None
    need = float(expected_ft) * float(min_frac)
    if proj >= need:
        return True
    if proj <= -need:
        return False
    return None


# ── análisis del conjunto GUID ───────────────────────────────────────────────


def analyze_conjunto(doc, seed_rebar):
    """
    Devuelve dict con selección, capas longitudinales, trabas e inward.

    Capas = rebars con ``Armadura_Capa`` no vacío. Style=StirrupTie o sin
    Capa van a ``ties`` (mismo GUID, fuera de la tabla de capas / última capa).

    Claves: ok, error, guid, seed, host, layers (list), ties (list),
    inward (XYZ|None), rebar_count.
    """
    out = {
        u"ok": False,
        u"error": None,
        u"guid": None,
        u"seed": seed_rebar,
        u"host": None,
        u"layers": [],
        u"ties": [],
        u"inward": None,
        u"rebar_count": 0,
    }
    if seed_rebar is None or not isinstance(seed_rebar, Rebar):
        out[u"error"] = u"Selecciona un elemento Rebar."
        return out
    gid = get_armadura_conjunto_guid(seed_rebar)
    if not gid:
        out[u"error"] = (
            u"La rebar seleccionada no tiene Armadura_Conjunto_GUID. "
            u"No se puede continuar."
        )
        return out
    out[u"guid"] = gid
    out[u"host"] = _host_of(seed_rebar, doc)

    ids = collect_rebars_por_conjunto_guid(doc, gid)
    rebars = []
    for eid in ids or []:
        try:
            el = doc.GetElement(eid)
        except Exception:
            el = None
        if isinstance(el, Rebar):
            rebars.append(el)
    out[u"rebar_count"] = len(rebars)
    if not rebars:
        out[u"error"] = u"No se encontraron rebars con ese GUID."
        return out

    layers_map = {}  # index|txt-key -> list of rebars
    ties = []
    for rb in rebars:
        # Trabas/estribos: Style StirrupTie O sin Armadura_Capa.
        # Comparten GUID del conjunto pero NUNCA cuentan como capa en la UI
        # ni en el cálculo de «última capa» / offsets longitudinales.
        if _is_stirrup_tie(rb):
            ties.append(rb)
            continue
        capa_txt = get_armadura_capa(rb)
        capa_u = _as_unicode(capa_txt).strip() if capa_txt else u""
        if not capa_u:
            ties.append(rb)
            continue
        idx = _parse_capa_index(capa_u)
        if idx is None:
            # Capa con texto no parseable: bucket por valor (sigue siendo capa).
            key = u"txt:" + capa_u
            layers_map.setdefault(key, []).append(rb)
            continue
        layers_map.setdefault(idx, []).append(rb)

    # Ordenar capas numéricas; las de texto al final.
    numeric_keys = sorted([k for k in layers_map.keys() if isinstance(k, int)])
    text_keys = sorted([k for k in layers_map.keys() if not isinstance(k, int)])

    layers = []
    display_i = 0
    for k in numeric_keys:
        display_i = int(k) + 1
        rbs = layers_map[k]
        bt = _bar_type_of(rbs[0], doc) if rbs else None
        diam = _nominal_diam_mm(bt)
        qty = sum(_cantidad_posiciones(r) for r in rbs)
        # Si hay un solo set, qty = NumberOfBarPositions; si varios sets, sumar.
        if len(rbs) == 1:
            qty = _cantidad_posiciones(rbs[0])
        else:
            # Preferir qty del set más grande (capas suelen ser un set por capa).
            qty = max(_cantidad_posiciones(r) for r in rbs)
        sp = _spacing_mm(rbs[0]) if rbs else 0
        layers.append(
            {
                u"index": int(k),
                u"display": display_i,
                u"rebars": rbs,
                u"qty": qty,
                u"diameter_mm": diam,
                u"spacing_mm": sp,
                u"capa_txt": get_armadura_capa(rbs[0]) if rbs else None,
                u"centroid": _layer_centroid(rbs),
            }
        )

    # Capas solo por texto / sin índice: asignar índices tras el máximo.
    next_idx = (numeric_keys[-1] + 1) if numeric_keys else 0
    for k in text_keys:
        rbs = layers_map[k]
        bt = _bar_type_of(rbs[0], doc) if rbs else None
        layers.append(
            {
                u"index": next_idx,
                u"display": next_idx + 1,
                u"rebars": rbs,
                u"qty": max(_cantidad_posiciones(r) for r in rbs) if rbs else 0,
                u"diameter_mm": _nominal_diam_mm(bt),
                u"spacing_mm": _spacing_mm(rbs[0]) if rbs else 0,
                u"capa_txt": get_armadura_capa(rbs[0]) if rbs else None,
                u"centroid": _layer_centroid(rbs),
            }
        )
        next_idx += 1

    if not layers:
        out[u"error"] = (
            u"El GUID no tiene rebars con Armadura_Capa (solo trabas/estribos "
            u"u otras barras sin capa). Selecciona una rebar longitudinal "
            u"con capa."
        )
        out[u"ties"] = ties
        return out

    out[u"layers"] = layers
    out[u"ties"] = ties
    out[u"inward"] = _inward_direction(layers, out[u"host"])
    if out[u"inward"] is None:
        out[u"error"] = (
            u"No se pudo determinar la dirección de capas hacia el interior. "
            u"Se usa el apilado Armadura_Capa (y el espesor del muro si aplica). "
            u"Comprueba que haya al menos una capa con posición distinta."
        )
        return out

    # Δ prev. = proyección sobre el eje de capas (inward), no longitud 3D
    # (evita Δ enormes por desfase Z/longitud que sacan barras del canvas).
    # Si inward colapsa (~0), probar espesor de muro y, en último caso, XY.
    axis = out[u"inward"]
    thick = _wall_thickness_axis(out[u"host"])
    min_ft = _mm_to_ft(1.0)
    for i, ly in enumerate(layers):
        if i == 0:
            ly[u"offset_from_prev_mm"] = None
            continue
        c0 = layers[i - 1].get(u"centroid")
        c1 = ly.get(u"centroid")
        if c0 is None or c1 is None:
            ly[u"offset_from_prev_mm"] = None
            continue
        try:
            delta = _xyz_sub(c1, c0)
            dist_ft = 0.0
            if axis is not None:
                dist_ft = abs(float(delta.DotProduct(axis)))
            if dist_ft < min_ft and thick is not None:
                dist_ft = abs(float(delta.DotProduct(thick)))
            if dist_ft < min_ft:
                # Separación en planta (ignora Z de empalmes / hooks)
                dx = float(c1.X) - float(c0.X)
                dy = float(c1.Y) - float(c0.Y)
                dist_ft = (dx * dx + dy * dy) ** 0.5
            ly[u"offset_from_prev_mm"] = int(round(_ft_to_mm(dist_ft)))
        except Exception:
            ly[u"offset_from_prev_mm"] = None

    out[u"ok"] = True
    return out


def _ties_summary(doc, ties):
    if not ties:
        return u"Sin trabas en este GUID."
    diams = []
    qty = 0
    sp_vals = []
    for t in ties:
        qty += _cantidad_posiciones(t)
        bt = _bar_type_of(t, doc)
        d = _nominal_diam_mm(bt)
        if d:
            diams.append(d)
        sp = _spacing_mm(t)
        if sp:
            sp_vals.append(sp)
    diam_txt = u"Ø{0}".format(min(diams)) if diams else u"Ø?"
    if diams and max(diams) != min(diams):
        diam_txt = u"Ø{0}–{1}".format(min(diams), max(diams))
    sp_txt = u""
    if sp_vals:
        sp_txt = u" · e={0} mm".format(int(round(sum(sp_vals) / float(len(sp_vals)))))
    return u"{0} trabas · {1}{2}".format(len(ties), diam_txt, sp_txt)


# ── creación ─────────────────────────────────────────────────────────────────


def _apply_qty_and_type(rebar, bar_type, qty):
    """Cambia tipo y cantidad (FixedNumber) preservando array length / lado."""
    if rebar is None:
        return
    if bar_type is not None:
        try:
            rebar.ChangeTypeId(bar_type.Id)
        except Exception:
            pass
    n = max(1, int(qty or 1))
    try:
        acc = rebar.GetShapeDrivenAccessor()
    except Exception:
        acc = None
    if acc is None:
        return
    alen = 0.0
    try:
        alen = float(acc.ArrayLength)
    except Exception:
        try:
            alen = float(acc.GetArrayLength())
        except Exception:
            alen = 0.0
    b_side = True
    try:
        b_side = bool(acc.BarsOnNormalSide)
    except Exception:
        pass
    try:
        inc0 = bool(rebar.IncludeFirstBar)
    except Exception:
        inc0 = True
    try:
        inc1 = bool(rebar.IncludeLastBar)
    except Exception:
        inc1 = True
    if n <= 1:
        try:
            acc.SetLayoutAsSingle()
            return
        except Exception:
            pass
    try:
        acc.SetLayoutAsFixedNumber(n, alen, b_side, inc0, inc1)
    except Exception:
        try:
            sp = float(rebar.MaxSpacing) if float(rebar.MaxSpacing) > 1e-12 else _mm_to_ft(150.0)
            acc.SetLayoutAsNumberWithSpacing(n, sp, alen, b_side, inc0, inc1)
        except Exception:
            pass


def _stamp_new_rebar(rebar, guid, layer_index, is_tie=False):
    if rebar is None:
        return
    try:
        stamp_armadura_conjunto_guid(rebar, conjunto_guid=guid)
    except Exception:
        pass
    if stamp_armadura_arainco is not None:
        try:
            stamp_armadura_arainco(rebar, yes=True)
        except Exception:
            pass
    if not is_tie and layer_index is not None:
        try:
            set_armadura_capa_desde_layer(rebar, int(layer_index))
        except Exception:
            pass


def _stamp_new_empalme_detail(element, guid, layer_index):
    """Estampa GUID + Armadura_Capa en un Detail Item de empalme copiado/colocado."""
    if element is None:
        return
    try:
        stamp_armadura_conjunto_guid(element, conjunto_guid=guid)
    except Exception:
        pass
    if stamp_armadura_arainco is not None:
        try:
            stamp_armadura_arainco(element, yes=True)
        except Exception:
            pass
    if layer_index is not None:
        try:
            stamp_armadura_capa_desde_layer(element, int(layer_index))
        except Exception:
            pass


def _capa_texts_equivalent(a, b):
    """True si ambas capas representan el mismo índice (o texto normalizado igual)."""
    ia = _parse_capa_index(a)
    ib = _parse_capa_index(b)
    if ia is not None and ib is not None:
        return int(ia) == int(ib)
    na = _as_unicode(a or u"").strip().lower()
    nb = _as_unicode(b or u"").strip().lower()
    if not na or not nb:
        return False
    return na == nb


def _detail_line_endpoints(el):
    """Extremos (XYZ, XYZ) del Detail Item line-based de empalme, o (None, None)."""
    if el is None:
        return None, None
    try:
        loc = el.Location
        if isinstance(loc, LocationCurve) and loc.Curve is not None:
            c = loc.Curve
            return c.GetEndPoint(0), c.GetEndPoint(1)
    except Exception:
        pass
    try:
        loc = el.Location
        curve = getattr(loc, u"Curve", None)
        if curve is not None:
            return curve.GetEndPoint(0), curve.GetEndPoint(1)
    except Exception:
        pass
    return None, None


def _xyz_add(a, b):
    if a is None or b is None:
        return None
    try:
        return a.Add(b)
    except Exception:
        try:
            return XYZ(
                float(a.X) + float(b.X),
                float(a.Y) + float(b.Y),
                float(a.Z) + float(b.Z),
            )
        except Exception:
            return None


def _point_near_layer(pt, src_layer, tol_ft):
    """True si ``pt`` está cerca del centroide o de algún midpoint de la capa plantilla."""
    if pt is None or not src_layer:
        return False
    c = src_layer.get(u"centroid")
    if c is not None:
        try:
            if float(pt.DistanceTo(c)) <= float(tol_ft):
                return True
        except Exception:
            pass
    for rb in src_layer.get(u"rebars") or []:
        mid = _rebar_midpoint(rb)
        if mid is None:
            continue
        try:
            if float(pt.DistanceTo(mid)) <= float(tol_ft):
                return True
        except Exception:
            continue
    return False


def _empalme_ids_for_layer(doc, guid, src_layer):
    """
    Detail Items de empalme (lap detail) del GUID asociados a la capa plantilla.

    1) Misma ``Armadura_Capa`` (índice equivalente).
    2) Fallback: empalmes del GUID cercanos a la capa plantilla.
    3) Último recurso: todos los empalmes del GUID (si hay ≤ 12).
    """
    if doc is None or not guid or not src_layer:
        return []
    capa_txt = src_layer.get(u"capa_txt")
    if not capa_txt:
        rebars = src_layer.get(u"rebars") or []
        if rebars:
            capa_txt = get_armadura_capa(rebars[0])
    if not capa_txt:
        try:
            from armado_muros_rebar_params import armadura_capa_valor_desde_layer

            capa_txt = armadura_capa_valor_desde_layer(int(src_layer.get(u"index")))
        except Exception:
            capa_txt = None

    ids = []
    try:
        ids = list(
            collect_empalmes_por_conjunto_guid_y_capa(doc, guid, capa_txt) or []
        )
    except Exception:
        ids = []

    # Recolectar todos del GUID y filtrar por capa equivalente / proximidad.
    all_emp = []
    try:
        all_emp = list(collect_empalmes_por_conjunto_guid(doc, guid) or [])
    except Exception:
        all_emp = []

    if not ids and all_emp and capa_txt:
        matched = []
        for eid in all_emp:
            try:
                el = doc.GetElement(eid)
            except Exception:
                el = None
            if el is None:
                continue
            try:
                el_capa = get_armadura_capa(el)
            except Exception:
                el_capa = None
            if el_capa and _capa_texts_equivalent(el_capa, capa_txt):
                matched.append(eid)
        if matched:
            ids = matched

    if not ids and all_emp:
        tol_ft = _mm_to_ft(350.0)
        near = []
        for eid in all_emp:
            try:
                el = doc.GetElement(eid)
            except Exception:
                el = None
            if el is None:
                continue
            p0, p1 = _detail_line_endpoints(el)
            mid = None
            if p0 is not None and p1 is not None:
                try:
                    mid = XYZ(
                        0.5 * (float(p0.X) + float(p1.X)),
                        0.5 * (float(p0.Y) + float(p1.Y)),
                        0.5 * (float(p0.Z) + float(p1.Z)),
                    )
                except Exception:
                    mid = p0
            elif p0 is not None:
                mid = p0
            if mid is not None and _point_near_layer(mid, src_layer, tol_ft):
                near.append(eid)
        if near:
            ids = near
        elif len(all_emp) <= 12:
            # GUID pequeño: heredar todos los lap details del conjunto.
            ids = list(all_emp)

    # Unique
    out = []
    seen = set()
    for eid in ids:
        try:
            key = int(eid.IntegerValue)
        except Exception:
            try:
                key = int(eid.Value)
            except Exception:
                key = id(eid)
        if key in seen:
            continue
        seen.add(key)
        out.append(eid)
    return out


def _resolve_lap_detail_symbol(doc):
    """FamilySymbol del Detail Item de empalme canónico, o None."""
    if doc is None:
        return None
    if _find_fixed_lap_detail_symbol_id is not None:
        try:
            sid, _warn = _find_fixed_lap_detail_symbol_id(doc)
            if sid is not None and sid != ElementId.InvalidElementId:
                sym = doc.GetElement(sid)
                if isinstance(sym, FamilySymbol):
                    return sym
        except Exception:
            pass
    # Respaldo: primer Empalme de categoría Detail Components
    try:
        for sym in FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(
            BuiltInCategory.OST_DetailComponents
        ):
            try:
                fam = _as_unicode(getattr(sym, u"FamilyName", u"") or u"").upper()
                typ = _as_unicode(getattr(sym, u"Name", u"") or u"").upper()
            except Exception:
                continue
            if u"EMPALME" in fam or u"EMPALME" in typ:
                return sym
    except Exception:
        pass
    return None


def _place_one_lap_detail(doc, view, symbol, p0, p1):
    """
    Coloca un lap detail line-based en ``view``.
    Returns: (instance|None, error|None)
    """
    if doc is None or view is None or symbol is None or p0 is None or p1 is None:
        return None, u"Parámetros incompletos para lap detail."
    try:
        if float(p0.DistanceTo(p1)) < _mm_to_ft(5.0):
            return None, u"Segmento de empalme demasiado corto."
    except Exception:
        pass
    if _place_line_based_detail_component is not None:
        try:
            ok, err, inst = _place_line_based_detail_component(
                doc, view, symbol, p0, p1
            )
            if ok and inst is not None:
                return inst, None
            if err:
                # Continuar a NewFamilyInstance directo
                pass
        except Exception:
            pass
    try:
        if not bool(getattr(symbol, u"IsActive", True)):
            symbol.Activate()
            try:
                doc.Regenerate()
            except Exception:
                pass
    except Exception:
        pass
    try:
        ln = RevitLine.CreateBound(p0, p1)
        inst = doc.Create.NewFamilyInstance(ln, symbol, view)
        return inst, None
    except Exception as ex:
        return None, _as_unicode(ex)


def _rebar_axis_and_mid(rebar):
    """(axis_unit, midpoint, t0, t1) de la centerline dominante, o Nones."""
    if rebar is None:
        return None, None, None, None
    try:
        from lap_detail_overlap_geom import _dominant_centerline_curve

        curve = _dominant_centerline_curve(rebar)
    except Exception:
        curve = None
    if curve is None:
        mid = _rebar_midpoint(rebar)
        return None, mid, None, None
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        axis = _normalize_xyz(_xyz_sub(p1, p0))
        mid = XYZ(
            0.5 * (float(p0.X) + float(p1.X)),
            0.5 * (float(p0.Y) + float(p1.Y)),
            0.5 * (float(p0.Z) + float(p1.Z)),
        )
        if axis is None:
            return None, mid, None, None
        t0 = 0.0
        t1 = float(_xyz_sub(p1, p0).DotProduct(axis))
        return axis, mid, min(t0, t1), max(t0, t1)
    except Exception:
        return None, _rebar_midpoint(rebar), None, None


def _cluster_collinear_rebars(rebars, lat_tol_ft=None):
    """
    Agrupa rebars colineales (misma fibra / columna de barras).

    Dentro de cada grupo se ordenan a lo largo del eje para emparejar
    tramos consecutivos del empalme.
    """
    if lat_tol_ft is None:
        lat_tol_ft = _mm_to_ft(40.0)
    items = []
    for rb in rebars or []:
        if rb is None:
            continue
        axis, mid, t0, t1 = _rebar_axis_and_mid(rb)
        if mid is None:
            continue
        items.append(
            {
                u"rebar": rb,
                u"axis": axis,
                u"mid": mid,
                u"t0": t0,
                u"t1": t1,
            }
        )
    if not items:
        return []

    clusters = []
    used = [False] * len(items)
    for i, it in enumerate(items):
        if used[i]:
            continue
        used[i] = True
        group = [it]
        axis_ref = it.get(u"axis")
        mid_ref = it.get(u"mid")
        for j in range(i + 1, len(items)):
            if used[j]:
                continue
            other = items[j]
            ax = other.get(u"axis")
            mid = other.get(u"mid")
            if mid is None or mid_ref is None:
                continue
            # Ejes paralelos (si ambos existen)
            if axis_ref is not None and ax is not None:
                try:
                    if abs(float(axis_ref.DotProduct(ax))) < 0.85:
                        continue
                except Exception:
                    continue
            # Distancia lateral al eje (o 3D si no hay eje)
            try:
                delta = _xyz_sub(mid, mid_ref)
                if axis_ref is not None and delta is not None:
                    along = float(delta.DotProduct(axis_ref))
                    parallel = _xyz_scale(axis_ref, along)
                    lat = _xyz_sub(delta, parallel)
                    lat_d = float(lat.GetLength()) if lat is not None else 1e9
                else:
                    lat_d = float(delta.GetLength())
            except Exception:
                continue
            if lat_d > float(lat_tol_ft):
                continue
            used[j] = True
            group.append(other)
        clusters.append(group)
    return clusters


def _sort_cluster_along_axis(group):
    """Ordena tramos a lo largo del eje (Z / longitud) por midpoint proyectado."""
    if not group:
        return []
    axis = None
    for it in group:
        if it.get(u"axis") is not None:
            axis = it[u"axis"]
            break
    origin = group[0].get(u"mid")
    if axis is None or origin is None:
        # Fallback: por Z
        return sorted(
            group,
            key=lambda it: float(it[u"mid"].Z) if it.get(u"mid") is not None else 0.0,
        )

    def _key(it):
        mid = it.get(u"mid")
        if mid is None:
            return 0.0
        try:
            return float(_xyz_sub(mid, origin).DotProduct(axis))
        except Exception:
            return float(mid.Z)

    return sorted(group, key=_key)


def _place_lap_details_between_new_rebars(
    doc, view, rebars, guid, layer_index, skip_regen=False
):
    """
    Coloca lap details en el **solape real** entre tramos consecutivos de las
    barras nuevas (misma fibra), no copiando details a una altura fija.

    ``skip_regen``: True si el caller ya regeneró el documento.

    Returns: (n_ok, errors_list)
    """
    created = 0
    errors = []
    if doc is None or view is None or not rebars or len(rebars) < 2:
        return 0, errors

    try:
        from lap_detail_overlap_geom import compute_lap_segment_endpoints
    except Exception as ex:
        return 0, [
            u"No se pudo cargar lap_detail_overlap_geom: {0}".format(_as_unicode(ex))
        ]

    symbol = _resolve_lap_detail_symbol(doc)
    if symbol is None:
        return 0, [u"No se encontró familia Detail Item de empalme (Empalme)."]

    if not skip_regen:
        try:
            doc.Regenerate()
        except Exception:
            pass

    # Releer elementos frescos
    fresh = []
    for rb in rebars:
        fr = _refresh_rebar(doc, rb)
        if fr is not None:
            fresh.append(fr)
    if len(fresh) < 2:
        return 0, errors

    clusters = _cluster_collinear_rebars(fresh)
    min_lap_ft = _mm_to_ft(50.0)

    for group in clusters:
        ordered = _sort_cluster_along_axis(group)
        if len(ordered) < 2:
            continue
        for i in range(len(ordered) - 1):
            ra = ordered[i].get(u"rebar")
            rb = ordered[i + 1].get(u"rebar")
            if ra is None or rb is None:
                continue
            try:
                p0, p1 = compute_lap_segment_endpoints(
                    ra, rb, view, min_len_ft=min_lap_ft
                )
            except Exception as ex:
                errors.append(
                    u"Solape C{0}: {1}".format(
                        int(layer_index) + 1 if layer_index is not None else u"?",
                        _as_unicode(ex),
                    )
                )
                continue
            if p0 is None or p1 is None:
                continue
            inst, err = _place_one_lap_detail(doc, view, symbol, p0, p1)
            if inst is None:
                if err:
                    errors.append(
                        u"Lap detail C{0}: {1}".format(
                            int(layer_index) + 1 if layer_index is not None else u"?",
                            err,
                        )
                    )
                continue
            _stamp_new_empalme_detail(inst, guid, layer_index)
            created += 1

    return created, errors


def _copy_empalme_details(doc, empalme_ids, delta_xyz, guid, layer_index, view=None):
    """
    Respaldo legado: copia/offset de Detail Items de la plantilla.

    Preferir ``_place_lap_details_between_new_rebars`` (solape entre barras nuevas).
    """
    created = 0
    errors = []
    if doc is None or not empalme_ids:
        return 0, errors
    if delta_xyz is None:
        return 0, [u"Offset nulo para Detail Items de empalme."]

    symbol = _resolve_lap_detail_symbol(doc)
    use_place = view is not None and symbol is not None

    for eid in empalme_ids:
        if eid is None or eid == ElementId.InvalidElementId:
            continue
        try:
            src = doc.GetElement(eid)
        except Exception:
            src = None
        if src is None or not isinstance(src, FamilyInstance):
            continue

        placed = None
        if use_place:
            p0, p1 = _detail_line_endpoints(src)
            if p0 is not None and p1 is not None:
                np0 = _xyz_add(p0, delta_xyz)
                np1 = _xyz_add(p1, delta_xyz)
                placed, err_p = _place_one_lap_detail(doc, view, symbol, np0, np1)
                if placed is None and err_p:
                    pass

        if placed is None:
            try:
                new_ids = ElementTransformUtils.CopyElement(doc, eid, delta_xyz)
            except Exception as ex:
                errors.append(
                    u"Empalme Id {0}: {1}".format(
                        _element_id_int(eid) or u"?", _as_unicode(ex)
                    )
                )
                continue
            count = 0
            try:
                count = int(new_ids.Count) if new_ids is not None else 0
            except Exception:
                try:
                    count = len(new_ids) if new_ids is not None else 0
                except Exception:
                    count = 0
            if count < 1:
                continue
            try:
                placed = doc.GetElement(new_ids[0])
            except Exception:
                placed = None

        if placed is None:
            continue
        _stamp_new_empalme_detail(placed, guid, layer_index)
        created += 1
    return created, errors


def _refresh_rebar(doc, rebar_or_id):
    """Relee Rebar desde el documento (evita wrappers obsoletos tras ExternalEvent)."""
    if doc is None or rebar_or_id is None:
        return None
    eid = None
    try:
        if isinstance(rebar_or_id, Rebar):
            eid = rebar_or_id.Id
        elif isinstance(rebar_or_id, ElementId):
            eid = rebar_or_id
        else:
            eid = getattr(rebar_or_id, u"Id", None)
    except Exception:
        eid = None
    if eid is None or eid == ElementId.InvalidElementId:
        return None
    try:
        el = doc.GetElement(eid)
    except Exception:
        return None
    if isinstance(el, Rebar):
        return el
    return None


def _copy_rebar_zero(doc, source_rebar):
    """
    Copia rebar en el mismo sitio (traslación cero).

    Returns: (new_rebar|None, error_msg|None)
    """
    if source_rebar is None or doc is None:
        return None, u"Fuente o documento nulo."
    src = _refresh_rebar(doc, source_rebar)
    if src is None:
        return None, u"La rebar plantilla ya no existe en el documento."
    try:
        new_ids = ElementTransformUtils.CopyElement(doc, src.Id, XYZ.Zero)
    except Exception as ex:
        return None, u"CopyElement: {0}".format(_as_unicode(ex))
    count = 0
    try:
        count = int(new_ids.Count) if new_ids is not None else 0
    except Exception:
        try:
            count = len(new_ids) if new_ids is not None else 0
        except Exception:
            count = 0
    if count < 1:
        return None, u"CopyElement no devolvió elementos."
    try:
        first_id = new_ids[0]
    except Exception:
        try:
            first_id = list(new_ids)[0]
        except Exception as ex:
            return None, u"No se pudo leer Id de la copia: {0}".format(_as_unicode(ex))
    rb2 = doc.GetElement(first_id)
    if rb2 is None or not isinstance(rb2, Rebar):
        return None, u"La copia no es un Rebar válido."
    return rb2, None


def _displacement_along(before, after, expected_dir):
    """Proyección del desplazamiento real sobre la dirección esperada (pies)."""
    if before is None or after is None or expected_dir is None:
        return None
    delta = _xyz_sub(after, before)
    if delta is None:
        return None
    try:
        return float(delta.DotProduct(expected_dir))
    except Exception:
        return None


def _try_move_bar_in_set_all(rebar, delta_xyz):
    """Aplica la misma traslación a todas las posiciones vía MoveBarInSet."""
    if rebar is None or delta_xyz is None:
        return False
    try:
        xform = Transform.CreateTranslation(delta_xyz)
    except Exception:
        return False
    n = max(1, _cantidad_posiciones(rebar))
    any_ok = False
    for i in range(n):
        try:
            rebar.MoveBarInSet(int(i), xform)
            any_ok = True
        except Exception:
            try:
                # Algunas APIs exponen MoveBarInSet solo en el accessor.
                acc = rebar.GetShapeDrivenAccessor()
                if acc is not None and hasattr(acc, u"MoveBarInSet"):
                    acc.MoveBarInSet(int(i), xform)
                    any_ok = True
            except Exception:
                pass
    return any_ok


def _offset_rebar(doc, rebar, delta_xyz, host_c=None, strict=True):
    """
    Desplaza ``rebar`` por ``delta_xyz`` (pies).

    Orden recomendado: copiar → cambiar Ø/layout → offset (SetLayout suele
    resetear la pose si se aplica *después* de MoveElement).

    Verifica el desplazamiento a lo largo de ``delta`` (no distancia 3D al
    centro del bbox: en muros largos / trabas eso rechazaba offsets correctos
    o aceptaba offsets «de lado»).

    ``strict=False`` (trabas/estribos): si ``MoveElement`` no lanza, aceptar
    aunque el midpoint multiplanar no refleje bien el desplazamiento.

    ``host_c`` se conserva por compatibilidad de firma; la validación usa el
    eje de ``delta_xyz``.

    Returns: error_msg|None (None = OK).
    """
    if rebar is None or doc is None:
        return u"Rebar o documento nulo al offset."
    if delta_xyz is None:
        return u"Vector de offset nulo."
    try:
        expected = float(delta_xyz.GetLength())
    except Exception:
        expected = 0.0
    if expected <= 1e-12:
        return u"Vector de offset ≈ 0 (revisa distanciamiento / dirección)."

    unit = _normalize_xyz(delta_xyz)
    mid0 = _rebar_midpoint(rebar)
    move_err = None
    try:
        ElementTransformUtils.MoveElement(doc, rebar.Id, delta_xyz)
    except Exception as ex:
        move_err = _as_unicode(ex)

    try:
        doc.Regenerate()
    except Exception:
        pass

    mid1 = _rebar_midpoint(rebar)
    proj = _displacement_along(mid0, mid1, unit)
    moved_ok = proj is not None and proj >= 0.4 * expected
    # Sin geometría medible pero MoveElement no lanzó: confiar (evitar doble offset).
    if not moved_ok and proj is None and move_err is None and mid0 is None:
        return None
    # Trabas/estribos multiplanares: el midpoint a menudo no se mueve proporcional
    # al MoveElement; si la API no falló, aceptar.
    if not moved_ok and not strict and move_err is None:
        return None

    moved_via_bars = False
    if not moved_ok:
        # MoveElement no movió (o falló): intentar MoveBarInSet en todas las barras.
        remaining = delta_xyz
        if proj is not None and abs(proj) > 1e-9 and unit is not None:
            remaining = _xyz_sub(delta_xyz, _xyz_scale(unit, proj))
            if remaining is None:
                remaining = delta_xyz
            try:
                if float(remaining.GetLength()) < 1e-9:
                    moved_ok = True
            except Exception:
                pass

        if not moved_ok:
            moved_via_bars = _try_move_bar_in_set_all(rebar, remaining)
            if not moved_via_bars and move_err:
                return u"MoveElement falló y MoveBarInSet no disponible: {0}".format(
                    move_err
                )

            try:
                doc.Regenerate()
            except Exception:
                pass

            mid1 = _rebar_midpoint(rebar)
            proj = _displacement_along(mid0, mid1, unit)
            if proj is not None and proj >= 0.4 * expected:
                moved_ok = True
            elif mid0 is None and moved_via_bars:
                return None
            elif proj is None and move_err is None and not moved_via_bars:
                return None
            elif not strict and (moved_via_bars or move_err is None):
                return None

    if not moved_ok:
        detail = u""
        if move_err:
            detail = u" MoveElement: {0}.".format(move_err)
        got_mm = (
            int(round(_ft_to_mm(proj))) if proj is not None else u"?"
        )
        want_mm = int(round(_ft_to_mm(expected)))
        return (
            u"La copia no se desplazó hacia el interior "
            u"(pedido {0} mm, medido ~{1} mm).{2}"
        ).format(want_mm, got_mm, detail)

    # Sanity: el move debe ir en el sentido del eje inward (proyección),
    # no «más cerca del centro 3D del bbox» (falso negativo en trabas / muros).
    along = _moved_along_inward(mid0, mid1, unit, expected, min_frac=0.35)
    if along is False:
        if not strict:
            return None
        return (
            u"El offset quedó en sentido contrario al interior del host "
            u"(eje de capas). No se creó la capa."
        )
    return None


def _delete_rebar_quiet(doc, rebar):
    if doc is None or rebar is None:
        return
    try:
        doc.Delete(rebar.Id)
    except Exception:
        pass


def _copy_and_offset(doc, source_rebar, delta_xyz, host_c=None, strict=True):
    """
    Copia rebar y la mueve ``delta_xyz``.

    Preferir el flujo de ``create_additional_layers`` (layout antes del offset).
    Returns: (new_rebar|None, error_msg|None)
    """
    rb2, err = _copy_rebar_zero(doc, source_rebar)
    if rb2 is None:
        return None, err
    err_m = _offset_rebar(doc, rb2, delta_xyz, host_c=host_c, strict=strict)
    if err_m:
        _delete_rebar_quiet(doc, rb2)
        return None, err_m
    return rb2, None


def _tie_s_along_inward(tie, origin, inward):
    """
    Proyección de la traba sobre ``inward`` desde ``origin``.

    Usa centerline planar sin ganchos (más fiable que BarPositionTransform /
    promedio con hooks, que sesgan la profundidad en espesor).
    """
    if tie is None or origin is None or inward is None:
        return None
    pts = []
    try:
        curves = tie.GetCenterlineCurves(
            False, True, False, MultiplanarOption.IncludeOnlyPlanarCurves, 0
        )
    except Exception:
        curves = None
    if curves is not None:
        try:
            n = int(curves.Count)
        except Exception:
            n = 0
        for i in range(n):
            try:
                c = curves[i]
                if c is None:
                    continue
                pts.append(c.Evaluate(0.5, True))
            except Exception:
                try:
                    pts.append(c.GetEndPoint(0))
                    pts.append(c.GetEndPoint(1))
                except Exception:
                    pass
    if not pts:
        p = _rebar_midpoint(tie)
        if p is None:
            return None
        pts = [p]
    try:
        acc = 0.0
        for p in pts:
            acc += float(_xyz_sub(p, origin).DotProduct(inward))
        return acc / float(len(pts))
    except Exception:
        return None


def _delta_tie_to_target_depth(tie, target_centroid, inward, fallback_ft=None):
    """
    Offset (pies) para llevar ``tie`` a la misma profundidad (eje inward)
    que ``target_centroid`` (centroide de la capa longitudinal nueva).
    """
    if inward is None:
        return None
    if target_centroid is None:
        if fallback_ft is None:
            return None
        return _xyz_scale(inward, float(fallback_ft))
    s_tie = _tie_s_along_inward(tie, target_centroid, inward)
    if s_tie is None:
        # Sin geometría medible: empujar solo el fallback (p. ej. k×e).
        if fallback_ft is None:
            return None
        return _xyz_scale(inward, float(fallback_ft))
    # s_tie = proyección de la traba relativa al target: hay que mover -s_tie
    # para coincidir con el centroide de la capa nueva.
    return _xyz_scale(inward, -float(s_tie))


def _delta_tie_to_new_layer(tie, last_layer, inward, k, step_ft, target_centroid=None):
    """
    Offset (pies) para llevar una copia de ``tie`` a la profundidad de la capa
    nueva k (misma profundidad inward que las longitudinales nuevas).

    Preferir ``target_centroid`` de las barras nuevas ya creadas; si no, usar
    última capa + k×e.
    """
    if inward is None or step_ft is None:
        return None
    fallback = float(step_ft) * float(k)
    if target_centroid is not None:
        return _delta_tie_to_target_depth(
            tie, target_centroid, inward, fallback_ft=fallback
        )
    last_c = None
    if last_layer is not None:
        last_c = last_layer.get(u"centroid")
        if last_c is None:
            last_c = _layer_centroid(last_layer.get(u"rebars") or [])
    if last_c is None:
        return _xyz_scale(inward, fallback)
    # Objetivo = última + k×e; anclar la medición al objetivo.
    try:
        target = last_c.Add(_xyz_scale(inward, fallback))
    except Exception:
        target = None
        try:
            d = _xyz_scale(inward, fallback)
            if d is not None:
                target = XYZ(
                    float(last_c.X) + float(d.X),
                    float(last_c.Y) + float(d.Y),
                    float(last_c.Z) + float(d.Z),
                )
        except Exception:
            target = None
    return _delta_tie_to_target_depth(
        tie, target, inward, fallback_ft=fallback
    )


def _rebar_shape_display_name(rebar, doc=None):
    if rebar is None:
        return u""
    sid = None
    try:
        sid = rebar.GetShapeId()
    except Exception:
        try:
            sid = rebar.RebarShapeId
        except Exception:
            sid = None
    if sid is None or sid == ElementId.InvalidElementId:
        return u""
    el = None
    try:
        if doc is not None:
            el = doc.GetElement(sid)
        else:
            el = rebar.Document.GetElement(sid)
    except Exception:
        el = None
    if el is None:
        return u""
    try:
        return _as_unicode(el.Name).strip()
    except Exception:
        return u""


def _rebar_centerline_curve_count(rebar, suppress_hooks=True):
    """Nº de tramos de centerline (hooks opcionalmente omitidos)."""
    if rebar is None:
        return 0
    sh = bool(suppress_hooks)
    attempts = (
        (False, sh, False, MultiplanarOption.IncludeOnlyPlanarCurves, 0),
        (False, sh, False, MultiplanarOption.IncludeAllMultiplanarCurves, 0),
        (False, sh, False),
    )
    for args in attempts:
        try:
            curves = rebar.GetCenterlineCurves(*args)
        except Exception:
            curves = None
        if curves is None:
            continue
        try:
            n = int(curves.Count)
        except Exception:
            try:
                n = len(list(curves))
            except Exception:
                n = 0
        if n > 0:
            return n
    return 0


def _tie_extent_along_axis_mm(rebar, axis):
    """Luz (mm) de la geometría proyectada sobre ``axis`` unitario."""
    if rebar is None or axis is None:
        return None
    pts = []
    try:
        curves = rebar.GetTransformedCenterlineCurves(
            False, True, False, MultiplanarOption.IncludeOnlyPlanarCurves, 0
        )
    except Exception:
        curves = None
    if curves is None:
        try:
            curves = rebar.GetCenterlineCurves(False, True, False)
        except Exception:
            curves = None
    if curves is not None:
        try:
            n = int(curves.Count)
        except Exception:
            n = 0
        for i in range(n):
            try:
                c = curves[i]
                pts.append(c.GetEndPoint(0))
                pts.append(c.GetEndPoint(1))
            except Exception:
                pass
    if len(pts) < 2:
        try:
            bb = rebar.get_BoundingBox(None)
            if bb is not None:
                pts = [bb.Min, bb.Max]
        except Exception:
            pts = []
    if len(pts) < 2:
        return None
    try:
        s_vals = [float(p.DotProduct(axis)) for p in pts]
        return abs(_ft_to_mm(max(s_vals) - min(s_vals)))
    except Exception:
        return None


def _is_perimeter_estribo(rebar, inward=None, doc=None):
    """
    True = estribo perimetral cerrado (no usar como plantilla de traba nueva).

    Heurística alineada con ArmadoMurosV3:
    - RebarShape «10» (estribo confinamiento) u otros cerrados ≥4 tramos.
    - Luz grande a lo largo del espesor (inward) → envuelve varias capas.
    Trabas = 1–2 tramos (pata ± ganchos) con poca luz en espesor.
    """
    if rebar is None:
        return False
    name = _rebar_shape_display_name(rebar, doc=doc)
    if name:
        dig = u"".join(ch for ch in name if ch.isdigit())
        # Shape «10» = estribo cerrado canónico del cabezal Arainco.
        if dig == u"10":
            return True
        low = name.lower()
        if u"stirrup" in low and u"tie" not in low:
            return True

    n_seg = _rebar_centerline_curve_count(rebar, suppress_hooks=True)
    if n_seg >= 4:
        return True

    ext_mm = _tie_extent_along_axis_mm(rebar, inward) if inward is not None else None
    if ext_mm is not None:
        # Estribo típico envuelve el apilado (≥ ~80 mm); traba es ~Ø/hooks.
        if n_seg >= 3 and ext_mm >= 80.0:
            return True
        if ext_mm >= 150.0:
            return True

    return False


def _trabas_only(ties, inward=None, doc=None):
    """Filtra estribos perimetrales; deja solo trabas aptas para copiar en capas nuevas."""
    out = []
    for t in ties or []:
        try:
            if _is_perimeter_estribo(t, inward=inward, doc=doc):
                continue
        except Exception:
            pass
        out.append(t)
    return out


def _deepest_tie(ties, inward, layers, doc=None):
    """Traba más interior (nunca estribo) según proyección sobre ``inward``."""
    lot = _innermost_tie_templates(ties, inward, layers, doc=doc)
    if lot:
        return lot[0]
    trabas = _trabas_only(ties, inward=inward, doc=doc)
    if trabas:
        return trabas[-1]
    return None


def _innermost_tie_templates(ties, inward, layers, doc=None):
    """
    Lote de **trabas** (no estribos) en la profundidad más interior del GUID.

    Se copian todas las del lote (p. ej. 3×Ø8) por cada capa nueva.
    Los estribos perimetrales del mismo GUID se excluyen siempre.
    """
    trabas = _trabas_only(ties, inward=inward, doc=doc)
    if not trabas:
        return []
    origin = None
    if layers:
        origin = layers[0].get(u"centroid")
    if origin is None:
        origin = _rebar_midpoint(trabas[0])

    scored = []
    for t in trabas:
        s = _tie_s_along_inward(t, origin, inward) if inward is not None else None
        if s is None:
            # Sin eje: incluir igual (sigue siendo traba filtrada).
            scored.append((0.0, t))
            continue
        scored.append((s, t))
    if not scored:
        return list(trabas)

    best_s = max(s for s, _t in scored)
    tol_ft = _mm_to_ft(25.0)
    lot = [t for s, t in scored if abs(s - best_s) <= tol_ft]
    return lot if lot else [scored[-1][1]]


def _max_capa_index(layers):
    """Mayor índice 0-based de ``Armadura_Capa`` entre capas detectadas."""
    best = None
    for ly in layers or []:
        try:
            idx = int(ly.get(u"index"))
        except Exception:
            continue
        if best is None or idx > best:
            best = idx
    if best is None:
        return -1
    return best


def _layer_same_parity(layers, target_index_0based):
    """
    Última capa existente con la misma paridad que ``target_index_0based``.

    Índices 0-based: 0=1ºC (impar), 1=2ºC (par), 2=3ºC (impar), …
    Así la siguiente capa tras un GUID con N capas pares hereda traslapos de
    las capas impares, y tras N impar hereda los de las pares.
    """
    try:
        want = int(target_index_0based) % 2
    except Exception:
        want = 0
    best = None
    for ly in layers or []:
        try:
            idx = int(ly.get(u"index"))
        except Exception:
            continue
        if idx % 2 == want:
            best = ly
    return best


def _offset_from_source_layer_to_new(source_layer, last_layer, inward, k, step_ft):
    """
    Vector (pies) desde la capa plantilla hasta la posición de la capa nueva k.

    Posición objetivo = última capa del GUID + k×step hacia el interior.
    Si la plantilla no es la última (paridad distinta), se suma el Δ proyectado
    plantilla→última sobre ``inward``.
    """
    if inward is None or step_ft is None:
        return None
    base = _xyz_scale(inward, float(step_ft) * float(k))
    if base is None:
        return None
    if source_layer is None or last_layer is None:
        return base
    try:
        if int(source_layer.get(u"index")) == int(last_layer.get(u"index")):
            return base
    except Exception:
        pass
    src_c = source_layer.get(u"centroid")
    if src_c is None:
        src_c = _layer_centroid(source_layer.get(u"rebars") or [])
    last_c = last_layer.get(u"centroid")
    if last_c is None:
        last_c = _layer_centroid(last_layer.get(u"rebars") or [])
    if src_c is None or last_c is None:
        return base
    to_last = _xyz_sub(last_c, src_c)
    if to_last is None:
        return base
    try:
        along = float(to_last.DotProduct(inward))
    except Exception:
        return base
    extra = _xyz_scale(inward, along)
    if extra is None:
        return base
    try:
        return base.Add(extra)
    except Exception:
        try:
            return XYZ(
                float(base.X) + float(extra.X),
                float(base.Y) + float(extra.Y),
                float(base.Z) + float(extra.Z),
            )
        except Exception:
            return base


def _max_capa_display(layers):
    """Mayor número de capa 1-based (para UI / preview)."""
    best = 0
    for ly in layers or []:
        try:
            d = int(ly.get(u"display", 0))
        except Exception:
            d = 0
        if d > best:
            best = d
    if best > 0:
        return best
    mx = _max_capa_index(layers)
    return mx + 1 if mx >= 0 else 0


def _norm_family_key(name):
    try:
        t = _as_unicode(name).strip().lower()
    except Exception:
        return u""
    for ch in (u"\xa0", u"\u200b", u"\ufeff"):
        t = t.replace(ch, u"")
    return u" ".join(t.split())


def _tag_family_name_of_symbol(sym):
    if sym is None:
        return None
    try:
        fam = sym.Family
        if fam is not None:
            return _as_unicode(fam.Name)
    except Exception:
        pass
    return None


def _tagged_ids_of_independent_tag(tag):
    """ElementIds etiquetados por un IndependentTag (API variable entre versiones)."""
    out = []
    if tag is None:
        return out
    getters = (
        lambda: tag.GetTaggedLocalElementIds(),
        lambda: tag.GetTaggedElementIds(),
    )
    for getter in getters:
        try:
            ids = getter()
        except Exception:
            ids = None
        if ids is None:
            continue
        try:
            for eid in ids:
                if eid is not None and eid != ElementId.InvalidElementId:
                    out.append(eid)
        except Exception:
            try:
                n = int(ids.Count)
                for i in range(n):
                    eid = ids[i]
                    if eid is not None and eid != ElementId.InvalidElementId:
                        out.append(eid)
            except Exception:
                pass
        if out:
            break
    return out


def _resolve_tag_family_from_rebars(doc, view, rebars):
    """
    Familia de tag ya usada en la vista para alguna rebar del conjunto.
    Si no hay, None.
    """
    if doc is None or not rebars:
        return None
    want = set()
    for rb in rebars:
        if rb is None:
            continue
        try:
            want.add(int(rb.Id.IntegerValue))
        except Exception:
            try:
                want.add(int(rb.Id.Value))
            except Exception:
                pass
    if not want:
        return None
    try:
        col = (
            FilteredElementCollector(doc, view.Id)
            .OfClass(IndependentTag)
            .OfCategory(BuiltInCategory.OST_RebarTags)
        )
    except Exception:
        try:
            col = FilteredElementCollector(doc).OfClass(IndependentTag)
        except Exception:
            return None
    for tag in col:
        try:
            if view is not None and tag.OwnerViewId != view.Id:
                continue
        except Exception:
            pass
        for eid in _tagged_ids_of_independent_tag(tag):
            try:
                key = int(eid.IntegerValue)
            except Exception:
                try:
                    key = int(eid.Value)
                except Exception:
                    continue
            if key not in want:
                continue
            try:
                sym = doc.GetElement(tag.GetTypeId())
            except Exception:
                sym = None
            fam = _tag_family_name_of_symbol(sym)
            if fam:
                return fam
    return None


def _first_available_tag_family(doc):
    """Primera familia candidata con tipos OST_RebarTags en el documento."""
    if doc is None:
        return None
    try:
        from enfierrado_shaft_hashtag import _collect_rebar_tag_symbol_map
    except Exception:
        _collect_rebar_tag_symbol_map = None
    for fam in _TAG_FAMILY_CANDIDATES:
        if _collect_rebar_tag_symbol_map is not None:
            try:
                tag_map = _collect_rebar_tag_symbol_map(doc, fam)
            except Exception:
                tag_map = None
            if tag_map:
                return fam
            continue
        want = _norm_family_key(fam)
        try:
            for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
                fn = _tag_family_name_of_symbol(sym)
                if fn and _norm_family_key(fn) == want:
                    return fam
        except Exception:
            pass
    return None


def _etiquetar_longitudinales_nuevas(doc, view, rebar_ids, seed_rebars=None):
    """
    Coloca Structural Rebar Tag en cada rebar nueva (misma Transaction abierta).

    Returns: (n_ok, aviso|None)
    """
    if doc is None or not rebar_ids:
        return 0, None
    try:
        from enfierrado_shaft_hashtag import etiquetar_rebars_creados_en_vista
    except Exception as ex:
        return 0, u"Módulo de etiquetado no disponible: {0}".format(_as_unicode(ex))

    fam = _resolve_tag_family_from_rebars(doc, view, seed_rebars or [])
    if not fam:
        fam = _first_available_tag_family(doc)
    if not fam:
        return 0, (
            u"No hay familia de Structural Rebar Tag en el proyecto "
            u"(p. ej. EST_A_STRUCTURAL REBAR TAG_WALL_HORIZONTAL). "
            u"Las capas se crearon sin etiquetas."
        )

    n_ok, avisos, err = etiquetar_rebars_creados_en_vista(
        doc,
        view,
        rebar_ids,
        family_name=fam,
        use_transaction=False,
    )
    if err:
        return int(n_ok or 0), _as_unicode(err)
    if avisos and not n_ok:
        return 0, u"; ".join([_as_unicode(a) for a in avisos[:3]])
    if avisos and n_ok:
        return int(n_ok), u"; ".join([_as_unicode(a) for a in avisos[:2]])
    return int(n_ok or 0), None


def create_additional_layers(
    doc,
    analysis,
    n_layers,
    qty_per_layer,
    diameter_mm,
    spacing_mm,
    view=None,
    add_trabas=False,
):
    """
    Crea capas longitudinales (+ trabas offset si ``add_trabas``) y etiqueta.

    Traslapos / empalmes por **paridad de capa** (1-based):
    - Capa nueva impar → última capa impar del GUID.
    - Capa nueva par → última capa par.

    Detail Items de empalme: se colocan en el **solape real** entre tramos
    de las barras nuevas (misma fibra), no como copia a altura fija.

    ``add_trabas`` (opt-in): si True, copia offset de **trabas** (no estribos)
    del lote más interior por cada capa nueva (aditivo). Nunca regenera ni
    amplía el estribo perimetral existente.

    Returns: dict ok, message, created_long, created_ties, created_tags,
    created_empalmes, errors
    """
    result = {
        u"ok": False,
        u"message": u"",
        u"created_long": 0,
        u"created_ties": 0,
        u"created_tags": 0,
        u"created_empalmes": 0,
        u"errors": [],
    }
    add_trabas = bool(add_trabas)

    if not analysis or not analysis.get(u"ok"):
        result[u"message"] = (analysis or {}).get(u"error") or u"Sin análisis válido."
        return result
    n_layers = max(0, int(n_layers or 0))
    if n_layers < 1:
        result[u"message"] = u"Indica al menos 1 capa nueva."
        return result
    spacing_mm = float(spacing_mm or 0)
    if spacing_mm <= 0:
        result[u"message"] = u"El distanciamiento debe ser mayor que 0 mm."
        return result
    qty_per_layer = max(1, int(qty_per_layer or 1))
    layers = analysis.get(u"layers") or []
    if not layers:
        result[u"message"] = u"No hay capas longitudinales de referencia."
        return result
    inward = analysis.get(u"inward")
    if inward is None:
        result[u"message"] = (
            u"No se pudo determinar la dirección hacia el interior del host. "
            u"Comprueba que la rebar tenga host válido y que la última capa "
            u"no coincida con el centro del elemento (sin eje de espesor)."
        )
        return result
    host = analysis.get(u"host")
    host_c = _bbox_center(host)
    guid = analysis.get(u"guid")
    last = layers[-1]
    bar_type = _resolver_bar_type_mm(doc, diameter_mm)
    if bar_type is None:
        result[u"message"] = u"No hay RebarBarType Ø{0} en el proyecto.".format(
            int(diameter_mm)
        )
        return result

    max_idx = _max_capa_index(layers)
    if max_idx < 0:
        max_idx = int(last.get(u"index", len(layers) - 1))
    step_ft = _mm_to_ft(spacing_mm)
    tie_templates = []
    if add_trabas:
        raw_lot = _innermost_tie_templates(
            analysis.get(u"ties") or [], inward, layers, doc=doc
        )
        for t_raw in raw_lot:
            t_fresh = _refresh_rebar(doc, t_raw)
            if t_fresh is None:
                continue
            # Doble filtro por si el refresh cambia la geometría medible.
            if _is_perimeter_estribo(t_fresh, inward=inward, doc=doc):
                continue
            tie_templates.append(t_fresh)

    t = Transaction(doc, _TXN_NAME)
    try:
        t.Start()
    except Exception as ex:
        result[u"message"] = (
            u"No se pudo iniciar la transacción (¿contexto API inválido?): {0}"
        ).format(_as_unicode(ex))
        return result
    try:
        created_long = 0
        created_ties = 0
        created_empalmes = 0
        created_long_ids = []
        parity_notes = []
        tag_seed_rebars = []
        for k in range(1, n_layers + 1):
            new_layer_index = max_idx + k
            # Plantilla de traslapo según paridad (impar↔impares, par↔pares).
            src_layer = _layer_same_parity(layers, new_layer_index)
            if src_layer is None:
                src_layer = last
            templates = [
                rb
                for rb in (
                    _refresh_rebar(doc, s)
                    for s in (src_layer.get(u"rebars") or [])
                )
                if rb is not None
            ]
            if not templates:
                result[u"errors"].append(
                    u"Capa +{0}: no hay rebars plantilla de paridad "
                    u"(capa ref. C{1}).".format(
                        k, int(src_layer.get(u"display") or (new_layer_index + 1))
                    )
                )
                continue
            if not tag_seed_rebars:
                tag_seed_rebars = list(templates)

            delta = _offset_from_source_layer_to_new(
                src_layer, last, inward, k, step_ft
            )
            if delta is None or float(delta.GetLength()) <= 1e-12:
                result[u"errors"].append(
                    u"Capa +{0}: delta de offset inválido "
                    u"(dirección o distanciamiento).".format(k)
                )
                continue

            try:
                src_disp = int(src_layer.get(u"display") or (int(src_layer.get(u"index")) + 1))
            except Exception:
                src_disp = new_layer_index + 1
            parity_notes.append(
                u"C{0}←traslapo C{1}".format(new_layer_index + 1, src_disp)
            )

            # Copiar CADA rebar de la capa de misma paridad (sets / empalmes).
            # IMPORTANTE: aplicar Ø/cantidad ANTES del offset.
            layer_long_ok = 0
            layer_new_rebars = []
            for src in templates:
                nb, err = _copy_rebar_zero(doc, src)
                if nb is None:
                    result[u"errors"].append(
                        u"Capa +{0}: {1}".format(k, err or u"falló CopyElement.")
                    )
                    continue
                _apply_qty_and_type(nb, bar_type, qty_per_layer)
                err_m = _offset_rebar(doc, nb, delta, host_c=host_c)
                if err_m:
                    _delete_rebar_quiet(doc, nb)
                    result[u"errors"].append(u"Capa +{0}: {1}".format(k, err_m))
                    continue
                _stamp_new_rebar(nb, guid, new_layer_index, is_tie=False)
                created_long += 1
                layer_long_ok += 1
                layer_new_rebars.append(nb)
                try:
                    created_long_ids.append(nb.Id)
                except Exception:
                    pass

            # Lap details en el solape real entre tramos nuevos (no altura fija).
            if layer_long_ok >= 2 and view is not None:
                n_emp, emp_errs = _place_lap_details_between_new_rebars(
                    doc, view, layer_new_rebars, guid, new_layer_index
                )
                created_empalmes += int(n_emp or 0)
                for ee in emp_errs or []:
                    result[u"errors"].append(u"Capa +{0}: {1}".format(k, ee))
            elif layer_long_ok >= 2 and view is None:
                result[u"errors"].append(
                    u"Capa +{0}: sin vista activa para colocar lap detail.".format(k)
                )
            elif layer_long_ok == 1:
                # Un solo tramo: no hay empalme entre barras nuevas.
                pass

            # Trabas aditivas (opt-in): solo trabas (no estribos) del lote interior.
            # No regenera ni amplía el estribo perimetral existente.
            if add_trabas and tie_templates and layer_long_ok > 0:
                for src_tie in tie_templates:
                    delta_tie = _delta_tie_to_new_layer(
                        src_tie, last, inward, k, step_ft
                    )
                    if delta_tie is None:
                        result[u"errors"].append(
                            u"Traba capa +{0}: delta inválido.".format(k)
                        )
                        continue
                    try:
                        if float(delta_tie.GetLength()) <= 1e-12:
                            # Ya en profundidad objetivo: copiar in-place (raro).
                            pass
                    except Exception:
                        pass
                    nt, err_t = _copy_rebar_zero(doc, src_tie)
                    if nt is None:
                        if err_t:
                            result[u"errors"].append(
                                u"Traba capa +{0}: {1}".format(k, err_t)
                            )
                        continue
                    err_tm = _offset_rebar(
                        doc, nt, delta_tie, host_c=host_c, strict=False
                    )
                    if err_tm:
                        _delete_rebar_quiet(doc, nt)
                        result[u"errors"].append(
                            u"Traba capa +{0}: {1}".format(k, err_tm)
                        )
                        continue
                    _stamp_new_rebar(nt, guid, None, is_tie=True)
                    created_ties += 1

        created_tags = 0
        tag_note = None
        if created_long_ids:
            try:
                doc.Regenerate()
            except Exception:
                pass
            try:
                created_tags, tag_note = _etiquetar_longitudinales_nuevas(
                    doc,
                    view,
                    created_long_ids,
                    seed_rebars=tag_seed_rebars or (last.get(u"rebars") or []),
                )
            except Exception as ex_tag:
                created_tags = 0
                tag_note = u"Error al etiquetar: {0}".format(_as_unicode(ex_tag))
            if tag_note and created_tags <= 0:
                result[u"errors"].append(u"Etiquetas: {0}".format(tag_note))
            elif tag_note:
                result[u"errors"].append(u"Etiquetas (avisos): {0}".format(tag_note))

        try:
            doc.Regenerate()
        except Exception:
            pass
        t.Commit()
        result[u"ok"] = created_long > 0
        result[u"created_long"] = created_long
        result[u"created_ties"] = created_ties
        result[u"created_tags"] = created_tags
        result[u"created_empalmes"] = created_empalmes
        if result[u"ok"]:
            first_capa = max_idx + 2  # display 1-based of first new
            last_capa = max_idx + 1 + n_layers
            msg = (
                u"Creadas {0} rebar(s) longitudinal(es) en {1} capa(s) "
                u"C{2}–C{3} (Ø{4}, {5} barras/capa, e={6} mm)."
            ).format(
                created_long,
                n_layers,
                first_capa,
                last_capa,
                int(diameter_mm),
                qty_per_layer,
                int(spacing_mm),
            )
            if parity_notes:
                msg += u" Traslapos por paridad: {0}.".format(
                    u", ".join(parity_notes[:6])
                )
            if created_empalmes:
                msg += u" Lap detail empalme: {0}.".format(created_empalmes)
            if created_ties:
                msg += u" + {0} traba(s) aditiva(s) (mismo GUID; estribo existente intacto).".format(
                    created_ties
                )
            elif add_trabas and not (analysis.get(u"ties") or []):
                msg += (
                    u" Opt-in trabas activo pero el GUID no tiene "
                    u"trabas/estribos de referencia."
                )
            elif add_trabas:
                n_trab = len(
                    _trabas_only(
                        analysis.get(u"ties") or [], inward=inward, doc=doc
                    )
                )
                if n_trab <= 0:
                    msg += (
                        u" Opt-in trabas activo: solo hay estribo(s) en el GUID; "
                        u"no se crearon trabas (regla: solo trabas, no estribos)."
                    )
                else:
                    msg += (
                        u" Opt-in trabas activo pero no se pudo copiar el lote "
                        u"de trabas plantilla."
                    )
            elif not add_trabas:
                msg += u" Sin trabas nuevas (opt-in off; estribo existente intacto)."
            if created_tags:
                msg += u" Etiquetas: {0}.".format(created_tags)
            elif created_long_ids:
                msg += u" Sin etiquetas colocadas."
            if result[u"errors"]:
                msg += u" Avisos: " + u"; ".join(result[u"errors"][:3])
            result[u"message"] = msg
        else:
            result[u"message"] = (
                u"No se creó ninguna longitudinal. "
                + (u"; ".join(result[u"errors"][:3]) if result[u"errors"] else u"")
            )
        return result
    except Exception as ex:
        try:
            if t.HasStarted():
                t.RollBack()
        except Exception:
            pass
        result[u"message"] = u"Error en transacción: {0}".format(_as_unicode(ex))
        return result


# ── UI ───────────────────────────────────────────────────────────────────────


def _host_thickness_mm(host):
    """Espesor del host en mm (Wall.Width o bbox)."""
    if isinstance(host, Wall):
        try:
            w = float(host.Width)
            if w > 1e-9:
                return max(1, int(round(_ft_to_mm(w))))
        except Exception:
            pass
    try:
        bb = host.get_BoundingBox(None) if host is not None else None
        if bb is not None:
            dx = abs(float(bb.Max.X - bb.Min.X))
            dy = abs(float(bb.Max.Y - bb.Min.Y))
            # Cara más delgada ≈ espesor en planta
            thick = min(dx, dy)
            if thick > 1e-9:
                return max(1, int(round(_ft_to_mm(thick))))
    except Exception:
        pass
    return 300


def _lowest_z_tie_lot(doc, ties):
    """
    Lote de estribos/trabas con Z medio más bajo (referencia canvas).
    Devuelve dict qty, diameter_mm, spacing_mm, z_mm o None.
    """
    if not ties:
        return None
    best = None
    best_z = None
    for t in ties:
        p = _rebar_midpoint(t)
        if p is None:
            continue
        try:
            z = float(p.Z)
        except Exception:
            continue
        if best_z is None or z < best_z:
            best_z = z
            best = t
    if best is None:
        best = ties[0]
        p = _rebar_midpoint(best)
        best_z = float(p.Z) if p is not None else 0.0
    # Agrupar por Z cercano (±150 mm) para qty del lote
    tol_ft = _mm_to_ft(150.0)
    lot = []
    for t in ties:
        p = _rebar_midpoint(t)
        if p is None:
            continue
        try:
            if abs(float(p.Z) - float(best_z)) <= tol_ft:
                lot.append(t)
        except Exception:
            pass
    if not lot:
        lot = [best]
    bt = _bar_type_of(lot[0], doc)
    qty = 0
    for t in lot:
        qty += max(1, _cantidad_posiciones(t))
    return {
        u"qty": qty or len(lot),
        u"diameter_mm": _nominal_diam_mm(bt),
        u"spacing_mm": _spacing_mm(lot[0]) or 0,
        u"z_mm": int(round(_ft_to_mm(best_z))) if best_z is not None else 0,
        u"count": len(lot),
    }


def _pick_rebar_element(uidoc, prompt=None):
    """PickObject Rebar-only. Devuelve elemento o None si cancela."""
    if uidoc is None:
        return None
    msg = prompt or u"Selecciona una rebar longitudinal con Armadura_Conjunto_GUID"
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _RebarSelectionFilter(),
            msg,
        )
    except OperationCanceledException:
        return None
    except Exception:
        return None
    if ref is None:
        return None
    try:
        return uidoc.Document.GetElement(ref.ElementId)
    except Exception:
        return None


def _hex_brush(hex_color):
    h = _as_unicode(hex_color or u"#000000").lstrip(u"#")
    if len(h) >= 6:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        return SolidColorBrush(Color.FromRgb(r, g, b))
    return SolidColorBrush(Color.FromRgb(0, 0, 0))


def _polar_px(cx, cy, deg, rad):
    a = float(deg) * math.pi / 180.0
    return (
        float(cx) + float(rad) * math.cos(a),
        float(cy) + float(rad) * math.sin(a),
    )


def _canvas_add_stroke_path(canv, start, ops, brush, thick, z_index=20):
    """Path abierto: ops = ('L',x,y) | ('A', r, sweep_cw, x, y)."""
    if canv is None or not ops:
        return
    fig = PathFigure()
    fig.IsClosed = False
    fig.StartPoint = WpfPoint(float(start[0]), float(start[1]))
    for op in ops:
        if not op:
            continue
        kind = op[0]
        if kind == u"L" or kind == "L":
            seg = LineSegment()
            seg.Point = WpfPoint(float(op[1]), float(op[2]))
            fig.Segments.Add(seg)
        elif kind == u"A" or kind == "A":
            r = max(0.5, float(op[1]))
            sweep_cw = bool(op[2])
            arc = ArcSegment()
            arc.Point = WpfPoint(float(op[3]), float(op[4]))
            arc.Size = WpfSize(r, r)
            arc.RotationAngle = 0.0
            arc.IsLargeArc = False
            arc.SweepDirection = (
                SweepDirection.Clockwise
                if sweep_cw
                else SweepDirection.Counterclockwise
            )
            fig.Segments.Add(arc)
    geo = PathGeometry()
    geo.Figures.Add(fig)
    path = WpfPath()
    path.Data = geo
    path.Stroke = brush
    path.StrokeThickness = float(thick)
    path.StrokeStartLineCap = PenLineCap.Round
    path.StrokeEndLineCap = PenLineCap.Round
    path.StrokeLineJoin = PenLineJoin.Round
    path.Fill = None
    try:
        from System.Windows.Controls import Panel

        Panel.SetZIndex(path, int(z_index))
    except Exception:
        pass
    canv.Children.Add(path)


def _draw_stirrup_135(
    canv, left, top, right, bot, bar_cx, bar_cy, wrap_r, brush, thick, tip_left=True
):
    """Estribo perimetral 135° (port ArmadoMurosV3 / mockup Stirrup135Path)."""
    R = max(3.0, float(wrap_r))
    cr = min(
        R,
        (float(right) - float(left)) * 0.5 - 0.5,
        (float(bot) - float(top)) * 0.5 - 0.5,
    )
    cr = max(1.0, cr)
    tail_len = max(16.0, float(thick) * 7.0)
    sqrt_half = math.sqrt(0.5)
    tx = (-sqrt_half) if tip_left else sqrt_half
    ty = sqrt_half
    l, t, rgt, b = float(left), float(top), float(right), float(bot)
    bc_x, bc_y = float(bar_cx), float(bar_cy)

    def _pol(deg, rad=None):
        return _polar_px(bc_x, bc_y, deg, R if rad is None else rad)

    north = _pol(270)
    east = _pol(0)
    west = _pol(180)
    exit_a = _pol(45) if tip_left else _pol(135)
    exit_b = _pol(225) if tip_left else _pol(315)

    if tip_left:
        body_ops = [
            (u"L", l + cr, t),
            (u"A", cr, False, l, t + cr),
            (u"L", l, b - cr),
            (u"A", cr, False, l + cr, b),
            (u"L", rgt - cr, b),
            (u"A", cr, False, rgt, b - cr),
            (u"L", east[0], east[1]),
        ]
        hook_top_ops = [
            (u"A", R, True, exit_a[0], exit_a[1]),
            (u"L", exit_a[0] + tx * tail_len, exit_a[1] + ty * tail_len),
        ]
        hook_side_ops = [
            (u"A", R, False, exit_b[0], exit_b[1]),
            (u"L", exit_b[0] + tx * tail_len, exit_b[1] + ty * tail_len),
        ]
        side_start = east
    else:
        body_ops = [
            (u"L", rgt - cr, t),
            (u"A", cr, True, rgt, t + cr),
            (u"L", rgt, b - cr),
            (u"A", cr, True, rgt - cr, b),
            (u"L", l + cr, b),
            (u"A", cr, True, l, b - cr),
            (u"L", west[0], west[1]),
        ]
        hook_top_ops = [
            (u"A", R, False, exit_a[0], exit_a[1]),
            (u"L", exit_a[0] + tx * tail_len, exit_a[1] + ty * tail_len),
        ]
        hook_side_ops = [
            (u"A", R, True, exit_b[0], exit_b[1]),
            (u"L", exit_b[0] + tx * tail_len, exit_b[1] + ty * tail_len),
        ]
        side_start = west

    _canvas_add_stroke_path(canv, north, body_ops, brush, thick, 21)
    _canvas_add_stroke_path(canv, north, hook_top_ops, brush, thick, 21)
    _canvas_add_stroke_path(canv, side_start, hook_side_ops, brush, thick, 21)


def _draw_hook_135_traba_end(
    canv, bar_cx, bar_cy, wrap_r, tip_left, end, brush, thick
):
    """Gancho 135° en extremo de traba (port Hook135TrabaEnd)."""
    R = max(3.0, float(wrap_r))
    tail_len = max(14.0, float(thick) * 6.5)
    sqrt_half = math.sqrt(0.5)
    bc_x, bc_y = float(bar_cx), float(bar_cy)

    def _pol(deg):
        return _polar_px(bc_x, bc_y, deg, R)

    end_key = _as_unicode(end or u"top")
    if end_key == u"top":
        start = _pol(0) if tip_left else _pol(180)
        exit_pt = _pol(225) if tip_left else _pol(315)
        tx = (-sqrt_half) if tip_left else sqrt_half
        ty = sqrt_half
        sweep_cw = not tip_left
    else:
        start = _pol(0) if tip_left else _pol(180)
        exit_pt = _pol(135) if tip_left else _pol(45)
        tx = (-sqrt_half) if tip_left else sqrt_half
        ty = -sqrt_half
        sweep_cw = bool(tip_left)

    ops = [
        (u"A", R, sweep_cw, exit_pt[0], exit_pt[1]),
        (u"L", exit_pt[0] + tx * tail_len, exit_pt[1] + ty * tail_len),
    ]
    _canvas_add_stroke_path(canv, start, ops, brush, thick, 22)


def _bar_ys_in_layer(qty, wall_top, wall_h, bar_r):
    n = max(1, min(int(qty or 1), 24))
    margin = float(bar_r) + 6.0
    usable = max(float(bar_r) * 2.0, float(wall_h) - margin * 2.0)
    if n == 1:
        return [float(wall_top) + float(wall_h) * 0.5]
    step = usable / float(n - 1)
    y0 = float(wall_top) + margin
    return [y0 + i * step for i in range(n)]


_MIN_CANVAS_LAYER_SEP_MM = 5.0


def _parse_hex_rgb(hex_c):
    h = _as_unicode(hex_c or u"").lstrip(u"#")
    if len(h) < 6:
        return (0x5B, 0xC0, 0xDE)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return (0x5B, 0xC0, 0xDE)


def _apply_toggle_switch(chk, label_text, parts, accent_hex=None):
    """
    Convierte un CheckBox (estilo BimToolsToggleMini) en toggle track+thumb
    con etiqueta — mismo patrón visual que ArmadoMurosV3.
    ``parts`` (dict) recibe thumb_xform / track_fill / track_border.
    """
    if chk is None or parts is None:
        return
    try:
        chk.Content = None
    except Exception:
        pass

    host = StackPanel()
    host.Orientation = Orientation.Horizontal
    host.VerticalAlignment = VerticalAlignment.Center

    track_fill = SolidColorBrush(Color.FromRgb(18, 38, 54))
    track_border = SolidColorBrush(Color.FromRgb(33, 70, 92))
    track = Border()
    track.Width = 36.0
    track.Height = 18.0
    track.CornerRadius = CornerRadius(9.0)
    track.Background = track_fill
    track.BorderBrush = track_border
    track.BorderThickness = Thickness(1)
    track.Margin = Thickness(0, 0, 8, 0)
    track.VerticalAlignment = VerticalAlignment.Center
    track.HorizontalAlignment = HorizontalAlignment.Left
    track.ClipToBounds = True
    track.SnapsToDevicePixels = True

    thumb_xform = TranslateTransform(0.0, 0.0)
    thumb = Border()
    thumb.Width = 12.0
    thumb.Height = 12.0
    thumb.CornerRadius = CornerRadius(6.0)
    thumb.Background = SolidColorBrush(Color.FromRgb(232, 244, 248))
    thumb.HorizontalAlignment = HorizontalAlignment.Left
    thumb.Margin = Thickness(2, 0, 0, 0)
    thumb.VerticalAlignment = VerticalAlignment.Center
    thumb.RenderTransform = thumb_xform
    thumb.SnapsToDevicePixels = True
    track.Child = thumb
    host.Children.Add(track)

    lbl = TextBlock()
    lbl.Text = _as_unicode(label_text or u"")
    lbl.FontSize = 11.0
    lbl.FontWeight = FontWeights.SemiBold
    lbl.VerticalAlignment = VerticalAlignment.Center
    lbl.Foreground = SolidColorBrush(Color.FromRgb(232, 244, 248))
    lbl.TextWrapping = TextWrapping.Wrap
    host.Children.Add(lbl)

    chk.Content = host
    ar, ag, ab = _parse_hex_rgb(accent_hex or ACCENT_PRIMARY)
    parts.clear()
    parts[u"thumb_xform"] = thumb_xform
    parts[u"track_fill"] = track_fill
    parts[u"track_border"] = track_border
    parts[u"accent"] = (ar, ag, ab)
    parts[u"off_fill"] = (18, 38, 54)
    parts[u"off_border"] = (33, 70, 92)
    try:
        on = bool(chk.IsChecked)
    except Exception:
        on = False
    _sync_toggle_switch_visual(parts, on)


def _sync_toggle_switch_visual(parts, checked):
    """Actualiza posición del thumb y color del track (ON/OFF)."""
    if not parts:
        return
    thumb_xform = parts.get(u"thumb_xform")
    track_fill = parts.get(u"track_fill")
    track_border = parts.get(u"track_border")
    on = bool(checked)
    try:
        if thumb_xform is not None:
            thumb_xform.X = 18.0 if on else 0.0
    except Exception:
        pass
    try:
        if on:
            ar, ag, ab = parts.get(u"accent") or (0x5B, 0xC0, 0xDE)
            if track_fill is not None:
                track_fill.Color = Color.FromRgb(ar, ag, ab)
            if track_border is not None:
                track_border.Color = Color.FromRgb(ar, ag, ab)
        else:
            fr, fg, fb = parts.get(u"off_fill") or (18, 38, 54)
            br, bg, bb = parts.get(u"off_border") or (33, 70, 92)
            if track_fill is not None:
                track_fill.Color = Color.FromRgb(fr, fg, fb)
            if track_border is not None:
                track_border.Color = Color.FromRgb(br, bg, bb)
    except Exception:
        pass


def _canvas_existing_x_positions_mm(layers, cover_mm, fallback_step_mm, inward):
    """
    X (mm desde cara exterior) de cada capa existente en el preview.

    Si la proyección de centroides sobre ``inward`` colapsa (Δ≈0 → todas las
    capas se dibujarían en la misma columna), separa esquemáticamente con
    ``fallback_step_mm`` para que C1..Cn sean visibles.
    """
    n = len(layers or [])
    if n < 1:
        return []
    cover = float(cover_mm)
    step = max(1.0, float(fallback_step_mm or 50.0))

    # 1) Posiciones absolutas sobre inward (si el span es usable).
    if inward is not None and n >= 2:
        cents = [ly.get(u"centroid") for ly in layers]
        c0 = cents[0] if cents else None
        if c0 is not None and all(c is not None for c in cents):
            try:
                s_abs = []
                for c in cents:
                    d = _xyz_sub(c, c0)
                    s_abs.append(_ft_to_mm(float(d.DotProduct(inward))))
                span = max(s_abs) - min(s_abs)
                if span >= _MIN_CANVAS_LAYER_SEP_MM:
                    s0 = s_abs[0]
                    return [cover + (s - s0) for s in s_abs]
            except Exception:
                pass

    # 2) Encadenar Δ prev.; si Δ es None o ~0, usar paso esquemático.
    xs = [cover]
    for i in range(1, n):
        off = layers[i].get(u"offset_from_prev_mm")
        try:
            ov = float(off) if off is not None else None
        except Exception:
            ov = None
        if ov is None or abs(ov) < _MIN_CANVAS_LAYER_SEP_MM:
            ov = step
        xs.append(xs[-1] + abs(ov))
    return xs


class _RebarSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Rebar)

    def AllowReference(self, reference, point):
        return True


class _CrearCapasHandler(IExternalEventHandler):
    """
    Handler único de proceso (AppDomain).

    No crear/Dispose ExternalEvent por cada apertura de ventana: en IronPython
    eso tumba Revit en la 2ª ejecución.
    """

    def GetName(self):
        return _TXN_NAME

    def Execute(self, uiapp):
        ctrl = _CREATE_TARGET
        if ctrl is None:
            return
        try:
            ctrl._execute_create(uiapp)
        except Exception as ex:
            try:
                ctrl._set_status(_as_unicode(ex))
            except Exception:
                pass
            _mostrar_aviso(
                uiapp,
                u"Error al crear capas adicionales.",
                content=_as_unicode(ex),
            )


_CREATE_TARGET = None
_CREATE_EVENT = None
_CREATE_HANDLER = None


def _ensure_create_event():
    """Un solo ExternalEvent para toda la vida del AppDomain. Nunca Dispose."""
    global _CREATE_EVENT, _CREATE_HANDLER
    if _CREATE_EVENT is None:
        _CREATE_HANDLER = _CrearCapasHandler()
        _CREATE_EVENT = ExternalEvent.Create(_CREATE_HANDLER)
    return _CREATE_EVENT


def _set_create_target(ctrl):
    global _CREATE_TARGET
    _CREATE_TARGET = ctrl


def _clear_create_target(ctrl=None):
    global _CREATE_TARGET
    if ctrl is None or _CREATE_TARGET is ctrl:
        _CREATE_TARGET = None


def _get_active_window():
    ctrl = _get_active_controller()
    if ctrl is not None:
        try:
            return ctrl._win
        except Exception:
            pass
    try:
        win = AppDomain.CurrentDomain.GetData(_APPDOMAIN_WINDOW_KEY)
    except Exception:
        return None
    if win is None:
        return None
    try:
        _ = win.Title
    except Exception:
        _clear_active_window()
        return None
    try:
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_active_window()
            return None
    except Exception:
        pass
    return win


def _get_active_controller():
    try:
        ctrl = AppDomain.CurrentDomain.GetData(_APPDOMAIN_CTRL_KEY)
    except Exception:
        return None
    if ctrl is None:
        return None
    try:
        win = getattr(ctrl, u"_win", None)
        if win is None:
            _clear_active_controller()
            return None
        if hasattr(win, "IsLoaded") and (not win.IsLoaded):
            _clear_active_controller()
            return None
        _ = win.Title
    except Exception:
        _clear_active_controller()
        return None
    return ctrl


def _set_active_window(win):
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, win)
    except Exception:
        pass


def _set_active_controller(ctrl):
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CTRL_KEY, ctrl)
    except Exception:
        pass
    try:
        if ctrl is not None:
            _set_active_window(getattr(ctrl, u"_win", None))
    except Exception:
        pass


def _clear_active_window():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_WINDOW_KEY, None)
    except Exception:
        pass


def _clear_active_controller():
    try:
        AppDomain.CurrentDomain.SetData(_APPDOMAIN_CTRL_KEY, None)
    except Exception:
        pass
    _clear_active_window()


# Shell: canvas sección (izq) + cards SELECCION/CAPAS/NUEVAS (der) + footer.
_WINDOW_XAML = u"""
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__CHROME_TITLE__"
  Width="1040" MinWidth="960" MinHeight="720"
  SizeToContent="Height"
  ResizeMode="CanResizeWithGrip"
  WindowStartupLocation="Manual"
  Background="__BG_APP__"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__BIMTOOLS_DARK_STYLES__
  </Window.Resources>
  <Border Background="__BG_APP__" BorderBrush="__BORDER__" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <!-- Auto (no *): SizeToContent=Height + * colapsa el canvas a 0 px -->
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock x:Name="TxtTitle" Text="__TOOL_TITLE__"
                   Foreground="__FG_TITLE__" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0"
                   Foreground="__FG_BODY__" FontSize="11" TextWrapping="Wrap"
                   Text="Pick rebar → GUID → capas · canvas sección · capas adicionales hacia el interior."/>
      </StackPanel>

      <TextBlock x:Name="TxtInfoHint" Grid.Row="1" Foreground="__FG_MUTED__" FontSize="10"
                 Margin="0,0,0,10" TextWrapping="Wrap"
                 Text="Solo Armadura_Capa cuenta como capa. Estribo regenerable C1…Cn. Trabas ⊥ (sin última si regen) · long. Tipo 3 opt-in."/>

      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="1.15*"/>
          <ColumnDefinition Width="12"/>
          <ColumnDefinition Width="*" MinWidth="360"/>
        </Grid.ColumnDefinitions>

        <!-- LEFT: section canvas -->
        <Border Grid.Column="0" Background="__BG_PANEL__" BorderBrush="__BORDER__"
                BorderThickness="1" CornerRadius="4" Padding="10">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <DockPanel Grid.Row="0" Margin="0,0,0,8" LastChildFill="True">
              <Border DockPanel.Dock="Right" Background="__BG_ELEV__"
                      BorderBrush="__ACCENT__" BorderThickness="1" CornerRadius="3"
                      Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
                <TextBlock Text="PREVIEW" Foreground="__ACCENT__" FontSize="10" FontWeight="SemiBold"/>
              </Border>
              <TextBlock Text="Sección · barras por capa" Foreground="__FG_TITLE__"
                         FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
            </DockPanel>
            <Border x:Name="BrdSectionCanvas" Grid.Row="1" Background="__BG_ELEV__"
                    BorderBrush="__BORDER__" BorderThickness="1" CornerRadius="4"
                    Padding="6" MinHeight="480" MinWidth="320"
                    VerticalAlignment="Stretch">
              <Canvas x:Name="SectionCanvas" ClipToBounds="True" MinHeight="480"
                      HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
            </Border>
            <TextBlock x:Name="TxtCanvasMeta" Grid.Row="2" Margin="0,8,0,0"
                       Foreground="__FG_MUTED__" FontSize="10" TextWrapping="Wrap"
                       Text=""/>
          </Grid>
        </Border>

        <!-- RIGHT: cards — sin MaxHeight: SizeToContent muestra SELECCION/CAPAS/NUEVAS al abrir -->
        <ScrollViewer x:Name="ScrollRightRail" Grid.Column="2"
                      VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled">
          <StackPanel x:Name="PanelRightRail">

            <Border x:Name="BrdCardSeleccion" Background="__BG_PANEL__" BorderBrush="__BORDER__"
                    BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,10"
                    Opacity="0.96">
              <StackPanel>
                <DockPanel Margin="0,0,0,6" LastChildFill="True">
                  <Border DockPanel.Dock="Right" Background="__BG_ELEV__"
                          BorderBrush="__ACCENT__" BorderThickness="1" CornerRadius="3"
                          Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
                    <TextBlock Text="SELECCION" Foreground="__ACCENT__" FontSize="10" FontWeight="SemiBold"/>
                  </Border>
                  <TextBlock Text="Selección de rebar" Foreground="__FG_TITLE__"
                             FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBlock Foreground="__FG_MUTED__" FontSize="10" Margin="0,0,0,8"
                           TextWrapping="Wrap"
                           Text="Ya seleccionada al lanzar · plantilla de capas nuevas"/>

                <Grid Margin="0,0,0,8">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0" Margin="0,0,6,0">
                    <TextBlock Style="{StaticResource LabelSmall}" Text="Elemento"/>
                    <TextBlock x:Name="TxtElement" Text="—" Foreground="__FG_TITLE__"
                               FontSize="11" FontWeight="SemiBold" TextWrapping="Wrap"/>
                  </StackPanel>
                  <StackPanel Grid.Column="1" Margin="6,0,0,0">
                    <TextBlock Style="{StaticResource LabelSmall}" Text="Host"/>
                    <TextBlock x:Name="TxtHost" Text="—" Foreground="__FG_TITLE__"
                               FontSize="11" FontWeight="SemiBold" TextWrapping="Wrap"/>
                  </StackPanel>
                </Grid>
                <Border Background="__BG_ELEV__" BorderBrush="__BORDER__" BorderThickness="1"
                        CornerRadius="4" Padding="8,8" Margin="0,0,0,8">
                  <StackPanel>
                    <DockPanel LastChildFill="True" Margin="0,0,0,2">
                      <TextBlock Text="Se hereda" DockPanel.Dock="Right"
                                 Foreground="__ACCENT__" FontSize="10" Margin="8,0,0,0"/>
                      <TextBlock Text="Armadura_Conjunto_GUID" Foreground="__FG_MUTED__" FontSize="10"/>
                    </DockPanel>
                    <TextBlock x:Name="TxtGuid" Text="—" Foreground="__ACCENT__"
                               FontSize="11" FontWeight="SemiBold" FontFamily="Consolas"
                               TextWrapping="Wrap"/>
                  </StackPanel>
                </Border>
                <StackPanel Orientation="Horizontal">
                  <Button x:Name="BtnRepick" Content="Cambiar selección"
                          Style="{StaticResource BtnSelectOutline}"
                          MinHeight="28" Padding="12,0" Margin="0,0,8,0"/>
                  <TextBlock x:Name="TxtRunInfo" VerticalAlignment="Center"
                             Foreground="__FG_BODY__" FontSize="10" TextWrapping="Wrap"/>
                </StackPanel>
              </StackPanel>
            </Border>

            <Border x:Name="PanelLayers" Background="__BG_PANEL__" BorderBrush="__BORDER__"
                    BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,10"
                    Opacity="0.96">
              <StackPanel>
                <DockPanel Margin="0,0,0,6" LastChildFill="True">
                  <Border DockPanel.Dock="Right" Background="__BG_ELEV__"
                          BorderBrush="__ACCENT__" BorderThickness="1" CornerRadius="3"
                          Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
                    <TextBlock Text="CAPAS" Foreground="__ACCENT__" FontSize="10" FontWeight="SemiBold"/>
                  </Border>
                  <TextBlock Text="Capas longitudinales del GUID" Foreground="__FG_TITLE__"
                             FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBlock Foreground="__FG_MUTED__" FontSize="10" Margin="0,0,0,8"
                           TextWrapping="Wrap"
                           Text="Solo Armadura_Capa · última capa = origen de offset"/>

                <Border BorderBrush="__BORDER__" BorderThickness="1" CornerRadius="4">
                  <StackPanel>
                    <Grid Background="__BG_HEADER__">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="48"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="56"/>
                        <ColumnDefinition Width="72"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Grid.Column="0" Text="Capa" Foreground="__FG_MUTED__"
                                 FontSize="10" Margin="10,6,0,6"/>
                      <TextBlock Grid.Column="1" Text="Barras" Foreground="__FG_MUTED__"
                                 FontSize="10" Margin="0,6,0,6"/>
                      <TextBlock Grid.Column="2" Text="Ø" Foreground="__FG_MUTED__"
                                 FontSize="10" Margin="0,6,0,6"/>
                      <TextBlock Grid.Column="3" Text="Δ prev." Foreground="__FG_MUTED__"
                                 FontSize="10" Margin="0,6,10,6"/>
                    </Grid>
                    <StackPanel x:Name="PanelLayerRows"/>
                  </StackPanel>
                </Border>
                <TextBlock x:Name="TxtLastLayer" Foreground="__FG_MUTED__" FontSize="10"
                           Margin="0,8,0,0" TextWrapping="Wrap"/>

                <Border Background="__BG_ELEV__" BorderBrush="__BORDER__" BorderThickness="1"
                        CornerRadius="4" Padding="8,8" Margin="0,10,0,0">
                  <StackPanel>
                    <DockPanel Margin="0,0,0,4" LastChildFill="True">
                      <Border DockPanel.Dock="Right" Background="__BG_PANEL__"
                              BorderBrush="__BORDER__" BorderThickness="1" CornerRadius="3"
                              Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
                        <TextBlock Text="TRABAS" Foreground="__FG_BODY__" FontSize="10" FontWeight="SemiBold"/>
                      </Border>
                      <TextBlock Text="Trabas / estribos (estilo V3 · 135°)" Foreground="__FG_TITLE__"
                                 FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    </DockPanel>
                    <TextBlock x:Name="TxtTies" Foreground="__FG_BODY__" FontSize="11"
                               TextWrapping="Wrap"/>
                    <TextBlock x:Name="TxtTiesLot" Foreground="__FG_MUTED__" FontSize="10"
                               Margin="0,4,0,0" TextWrapping="Wrap"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </Border>

            <Border x:Name="PanelForm" Background="__BG_PANEL__" BorderBrush="__ACCENT__"
                    BorderThickness="1.5" CornerRadius="4" Padding="10" Margin="0,0,0,0">
              <StackPanel>
                <DockPanel Margin="0,0,0,6" LastChildFill="True">
                  <Border DockPanel.Dock="Right" Background="__BG_ELEV__"
                          BorderBrush="__ACCENT__" BorderThickness="1" CornerRadius="3"
                          Padding="6,2" Margin="8,0,0,0" VerticalAlignment="Center">
                    <TextBlock Text="NUEVAS" Foreground="__ACCENT__" FontSize="10" FontWeight="SemiBold"/>
                  </Border>
                  <TextBlock Text="Capas adicionales" Foreground="__FG_TITLE__"
                             FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center"/>
                </DockPanel>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,0,8">
                    <TextBlock Style="{StaticResource LabelSmall}" Text="Capas nuevas"/>
                    <ComboBox x:Name="CmbNCapas" Style="{StaticResource ComboStretch}"
                              Height="26" HorizontalAlignment="Stretch"
                              MaxWidth="10000" Width="Auto"/>
                  </StackPanel>
                  <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,8">
                    <TextBlock Style="{StaticResource LabelSmall}" Text="Diámetro"/>
                    <ComboBox x:Name="CmbDiam" Style="{StaticResource ComboStretch}"
                              Height="26" HorizontalAlignment="Stretch"
                              MaxWidth="10000" Width="Auto"/>
                  </StackPanel>
                  <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,0,4">
                    <TextBlock Style="{StaticResource LabelSmall}" Text="Barras / capa"/>
                    <ComboBox x:Name="CmbQty" Style="{StaticResource ComboStretch}"
                              Height="26" HorizontalAlignment="Stretch"
                              MaxWidth="10000" Width="Auto"/>
                  </StackPanel>
                  <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,0,0,4">
                    <TextBlock Style="{StaticResource LabelSmall}" Text="Distanciamiento"/>
                    <DockPanel>
                      <TextBlock DockPanel.Dock="Right" Text="mm" Foreground="__FG_MUTED__"
                                 FontSize="11" VerticalAlignment="Center" Margin="6,0,0,0"/>
                      <TextBox x:Name="TxtSpacing" Text="50"
                               Style="{StaticResource BimToolsTextBoxDark}" Height="26"/>
                    </DockPanel>
                  </StackPanel>
                </Grid>
                <TextBlock Text="Offset interior · traslapo por paridad · GUID + Detail Items · tags"
                           Foreground="__FG_MUTED__" FontSize="9" Margin="0,8,0,6"
                           TextWrapping="Wrap"/>
                <CheckBox x:Name="ChkSame"
                          Content="Misma qty y Ø en todas"
                          Style="{StaticResource BimToolsToggleMini}" IsChecked="True"
                          Foreground="__FG_BODY__" FontSize="11" Margin="0,0,0,10"/>
                <Border Background="__BG_ELEV__" BorderBrush="__BORDER__" BorderThickness="1"
                        CornerRadius="4" Padding="10,8" Margin="0,0,0,4">
                  <StackPanel>
                    <CheckBox x:Name="ChkAddTrabas"
                              Style="{StaticResource BimToolsToggleMini}" IsChecked="False"
                              VerticalAlignment="Center"/>
                    <TextBlock x:Name="TxtTrabasHint" Foreground="__FG_MUTED__" FontSize="9"
                               Margin="0,6,0,0" TextWrapping="Wrap"
                               Text="OFF: no se crean trabas en capas nuevas (estribo existente intacto)."/>
                    <TextBlock x:Name="TxtTrabasPlantilla" Foreground="__ACCENT__" FontSize="10"
                               Margin="0,4,0,0" TextWrapping="Wrap" Visibility="Collapsed"
                               Text=""/>
                  </StackPanel>
                </Border>

                <Border Background="__BG_ELEV__" BorderBrush="__BORDER__" BorderThickness="1"
                        CornerRadius="4" Margin="0,4,0,0">
                  <StackPanel x:Name="PanelPreview"/>
                </Border>
                <TextBlock x:Name="TxtGuidHint" Foreground="__FG_MUTED__" FontSize="10"
                           Margin="0,8,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>

          </StackPanel>
        </ScrollViewer>
      </Grid>

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,12,0">
          <TextBlock x:Name="TxtStatus" Foreground="__FG_MUTED__" FontSize="10" TextWrapping="Wrap"
                     Text=""/>
          <TextBlock x:Name="TxtFooterHint" Foreground="__FG_MUTED__" FontSize="10"
                     TextWrapping="Wrap" Margin="0,4,0,0"
                     Text="Creación: plantilla por paridad → offset → GUID/Capa → etiquetar · trabas solo si opt-in."/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnClose" Content="Cerrar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnCreate" Content="Crear capas"
                  Style="{StaticResource BtnPrimary}" MinWidth="160" IsEnabled="False"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
"""


def _build_xaml():
    xaml = _WINDOW_XAML.replace(u"__BIMTOOLS_DARK_STYLES__", BIMTOOLS_DARK_STYLES_XML)
    repl = {
        u"__CHROME_TITLE__": WINDOW_CHROME_TITLE,
        u"__TOOL_TITLE__": _TOOL_TITLE,
        u"__BG_APP__": BG_APP,
        u"__BG_PANEL__": BG_PANEL,
        u"__BG_ELEV__": BG_PANEL_ELEVATED,
        u"__BG_HEADER__": BG_GROUP_HEADER,
        u"__BORDER__": BORDER,
        u"__ACCENT__": ACCENT_PRIMARY,
        u"__FG_TITLE__": FG_TITLE,
        u"__FG_BODY__": FG_BODY,
        u"__FG_MUTED__": FG_MUTED,
    }
    for k, v in repl.items():
        xaml = xaml.replace(k, v)
    return xaml


def _attach_revit_owner(win, uiapp):
    """
    Owner opcional. Desactivado: fijar Owner al hwnd de Revit con WPF modeless
    ha correlacionado con cierres en reaperturas; la ventana sigue usable.
    """
    return
    # if win is None or uiapp is None:
    #     return
    # try:
    #     from System.Windows.Interop import WindowInteropHelper
    #     hwnd = revit_main_hwnd(uiapp)
    #     if hwnd is not None:
    #         WindowInteropHelper(win).Owner = hwnd
    # except Exception:
    #     pass


def _prepare_window(win, uiapp):
    if win is None:
        return
    try:
        hwnd = revit_main_hwnd(uiapp)
        bind_center_wpf_on_revit_monitor(win, hwnd)
        position_wpf_window_center_on_monitor(win, hwnd)
    except Exception:
        pass
    # No Owner hwnd (ver _attach_revit_owner).



def _parse_int_box(tb, default=0, minimum=None):
    try:
        raw = _as_unicode(tb.Text).strip()
        v = int(float(raw.replace(u",", u".")))
    except Exception:
        v = default
    if minimum is not None and v < minimum:
        v = minimum
    return v


def _combo_int(cmb, default=1):
    if cmb is None:
        return default
    try:
        item = cmb.SelectedItem
        if item is not None:
            return int(float(_as_unicode(item).strip()))
    except Exception:
        pass
    try:
        return int(float(_as_unicode(cmb.Text).strip()))
    except Exception:
        return default


def _select_combo_int(cmb, value, options):
    if cmb is None:
        return
    try:
        v = int(value)
    except Exception:
        v = None
    if v is None or v not in options:
        try:
            cmb.SelectedIndex = 0
        except Exception:
            pass
        return
    label = _as_unicode(v)
    try:
        for i in range(int(cmb.Items.Count)):
            if _as_unicode(cmb.Items[i]) == label:
                cmb.SelectedIndex = i
                return
    except Exception:
        pass


class CapasAdicionalesGuidWindow(object):
    def __init__(self, uiapp, analysis=None):
        self._uiapp = uiapp
        self._uidoc = uiapp.ActiveUIDocument
        self._doc = self._uidoc.Document if self._uidoc else None
        self._analysis = analysis if analysis and analysis.get(u"ok") else None
        self._pending_create = None
        self._create_done = False
        self._section_canvas = None
        self._trabas_toggle_parts = {}
        self._win = XamlReader.Parse(_build_xaml())
        # ExternalEvent de proceso: Create una vez, nunca Dispose.
        self._create_event = _ensure_create_event()
        _set_create_target(self)
        self._wire()
        _prepare_window(self._win, uiapp)
        _set_active_controller(self)

        try:
            self._win.Closed += EventHandler(self._on_window_closed)
        except Exception:
            pass

        if self._analysis:
            self._populate_selection(self._analysis)
        else:
            self._set_status(u"Sin selección válida.")

    def _on_window_closed(self, sender, args):
        """Solo limpia singleton; no Dispose del ExternalEvent compartido."""
        _clear_create_target(self)
        _clear_active_controller()

    def _schedule_close(self):
        """Cerrar fuera de Execute (si se usa). Preferir dejar la ventana abierta."""
        win = self._win
        if win is None:
            return

        def _do_close():
            try:
                win.Close()
            except Exception:
                pass

        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            win.Dispatcher.BeginInvoke(
                Action(_do_close),
                DispatcherPriority.Background,
            )
        except Exception:
            try:
                _do_close()
            except Exception:
                pass

    def _wire(self):
        w = self._win

        cmb_n = w.FindName(u"CmbNCapas")
        if cmb_n is not None:
            for n in _N_CAPAS_OPTIONS:
                cmb_n.Items.Add(_as_unicode(n))
            try:
                cmb_n.SelectedIndex = 1  # 2
            except Exception:
                pass

        cmb_q = w.FindName(u"CmbQty")
        if cmb_q is not None:
            for n in _QTY_OPTIONS:
                cmb_q.Items.Add(_as_unicode(n))
            try:
                cmb_q.SelectedIndex = 3  # 4
            except Exception:
                pass

        cmb = w.FindName(u"CmbDiam")
        if cmb is not None:
            for d in _DIAMETER_OPTIONS_MM:
                cmb.Items.Add(u"Ø{0}".format(d))
            try:
                cmb.SelectedIndex = 3  # Ø16
            except Exception:
                pass

        w.FindName(u"BtnRepick").Click += self._on_pick
        w.FindName(u"BtnCreate").Click += self._on_create
        w.FindName(u"BtnClose").Click += self._on_close

        tb_sp = w.FindName(u"TxtSpacing")
        if tb_sp is not None:
            try:
                tb_sp.TextChanged += self._on_form_changed
            except Exception:
                pass
        for name in (u"CmbNCapas", u"CmbQty", u"CmbDiam"):
            c = w.FindName(name)
            if c is not None:
                try:
                    c.SelectionChanged += self._on_form_changed
                except Exception:
                    pass
        chk = w.FindName(u"ChkAddTrabas")
        if chk is not None:
            try:
                _apply_toggle_switch(
                    chk,
                    u"Agregar trabas en las capas nuevas",
                    self._trabas_toggle_parts,
                    ACCENT_PRIMARY,
                )
            except Exception:
                pass
            try:
                chk.Checked += self._on_trabas_toggle
                chk.Unchecked += self._on_trabas_toggle
            except Exception:
                pass

        self._section_canvas = w.FindName(u"SectionCanvas")
        try:
            canv = self._section_canvas
            if canv is not None:
                canv.SizeChanged += self._on_canvas_size
        except Exception:
            pass
        try:
            self._win.ContentRendered += EventHandler(self._on_content_rendered)
        except Exception:
            pass
        try:
            self._win.Loaded += EventHandler(self._on_content_rendered)
        except Exception:
            pass

    def _on_content_rendered(self, sender, args):
        # Tras layout real (SizeToContent / Show): igualar canvas al rail y redibujar.
        try:
            self._fit_canvas_height_to_rail()
        except Exception:
            pass
        try:
            self._redraw_section_canvas()
        except Exception:
            pass

    def _fit_canvas_height_to_rail(self):
        """Sube el canvas al alto del rail derecho para no dejar tarjetas cortadas visualmente."""
        w = self._win
        canv = w.FindName(u"SectionCanvas")
        brd = w.FindName(u"BrdSectionCanvas")
        rail = w.FindName(u"PanelRightRail")
        if canv is None:
            return
        target = 480.0
        try:
            if rail is not None:
                # DesiredSize tras measure; ActualHeight si ya hay layout
                h = float(rail.ActualHeight or 0.0)
                if h != h or h < 40.0:
                    rail.Measure(WpfSize(4000.0, 8000.0))
                    h = float(rail.DesiredSize.Height)
                if h == h and h > target:
                    # Restar padding del border del canvas (~12) + cabecera/meta aprox.
                    target = max(480.0, h - 36.0)
        except Exception:
            target = 480.0
        try:
            canv.Height = target
            canv.MinHeight = target
        except Exception:
            pass
        if brd is not None:
            try:
                brd.MinHeight = target
            except Exception:
                pass

    def _on_canvas_size(self, sender, args):
        self._redraw_section_canvas()

    def _canvas_draw_size(self, canv):
        """Tamaño de dibujo sin fijar Width/Height del Canvas (rompe el layout)."""
        width = 0.0
        height = 0.0
        try:
            width = float(canv.ActualWidth)
        except Exception:
            width = 0.0
        try:
            height = float(canv.ActualHeight)
        except Exception:
            height = 0.0
        # NaN → 0
        if width != width:
            width = 0.0
        if height != height:
            height = 0.0
        if width < 40 or height < 40:
            try:
                parent = canv.Parent
                if parent is not None:
                    pw = float(parent.ActualWidth)
                    ph = float(parent.ActualHeight)
                    if pw == pw and pw > 48:
                        width = max(width, pw - 12.0)
                    if ph == ph and ph > 48:
                        height = max(height, ph - 12.0)
            except Exception:
                pass
        if width < 40:
            width = 460.0
        if height < 40:
            height = 480.0
        return width, height

    def _set_card_accent(self, border, active):
        if border is None:
            return
        try:
            hex_c = ACCENT_PRIMARY if active else BORDER
            r = int(hex_c[1:3], 16)
            g = int(hex_c[3:5], 16)
            b = int(hex_c[5:7], 16)
            border.BorderBrush = SolidColorBrush(Color.FromRgb(r, g, b))
            border.BorderThickness = Thickness(1.5 if active else 1.0)
            border.Opacity = 1.0 if active else 0.96
        except Exception:
            pass

    def _set_status(self, text):
        st = self._win.FindName(u"TxtStatus")
        if st is not None:
            st.Text = _as_unicode(text)

    def _on_close(self, sender, args):
        try:
            self._win.Close()
        except Exception:
            pass

    def _on_form_changed(self, sender, args):
        self._refresh_preview()
        self._update_trabas_hint()

    def _on_trabas_toggle(self, sender, args):
        try:
            _sync_toggle_switch_visual(
                self._trabas_toggle_parts, self._add_trabas_checked()
            )
        except Exception:
            pass
        self._on_form_changed(sender, args)

    def _add_trabas_checked(self):
        chk = self._win.FindName(u"ChkAddTrabas")
        if chk is None:
            return False
        try:
            return bool(chk.IsChecked)
        except Exception:
            return False

    def _update_trabas_hint(self):
        tb = self._win.FindName(u"TxtTrabasHint")
        plant = self._win.FindName(u"TxtTrabasPlantilla")
        add_on = self._add_trabas_checked()
        try:
            _sync_toggle_switch_visual(self._trabas_toggle_parts, add_on)
        except Exception:
            pass
        n_ties = 0
        lot_n = 0
        n_trabas = 0
        if self._analysis and self._analysis.get(u"ok"):
            ties = self._analysis.get(u"ties") or []
            n_ties = len(ties)
            try:
                trabas = _trabas_only(
                    ties,
                    inward=self._analysis.get(u"inward"),
                    doc=self._doc,
                )
                n_trabas = len(trabas)
                lot = _innermost_tie_templates(
                    ties,
                    self._analysis.get(u"inward"),
                    self._analysis.get(u"layers") or [],
                    doc=self._doc,
                )
                lot_n = len(lot or [])
            except Exception:
                lot_n = 0
                n_trabas = 0
        if tb is not None:
            if add_on:
                if lot_n > 0:
                    tb.Text = (
                        u"ON: se agregan trabas ⊥ naranja solo en capas nuevas "
                        u"(mismo GUID). El estribo existente no se regenera ni amplía."
                    )
                else:
                    tb.Text = (
                        u"ON: opt-in activo, pero no hay trabas plantilla "
                        u"(solo estribo o vacío)."
                    )
            else:
                tb.Text = (
                    u"OFF: no se crean trabas en capas nuevas "
                    u"(estribo existente intacto)."
                )
        if plant is not None:
            try:
                from System.Windows import Visibility

                if add_on and lot_n > 0:
                    plant.Visibility = Visibility.Visible
                    plant.Text = (
                        u"Plantilla: {0} traba(s) del lote interior "
                        u"({1} trabas / {2} ties en GUID)."
                    ).format(lot_n, n_trabas, n_ties)
                elif add_on:
                    plant.Visibility = Visibility.Visible
                    plant.Text = (
                        u"Sin lote de trabas (n_trabas={0}, ties={1})."
                    ).format(n_trabas, n_ties)
                else:
                    plant.Visibility = Visibility.Collapsed
            except Exception:
                pass

    def _on_pick(self, sender, args):
        if self._uidoc is None:
            _mostrar_aviso(self._uiapp, u"No hay documento activo.")
            return
        prev = self._analysis
        try:
            self._win.Hide()
        except Exception:
            pass
        el = _pick_rebar_element(
            self._uidoc,
            u"Selecciona una rebar longitudinal con Armadura_Conjunto_GUID",
        )
        try:
            self._win.Show()
            self._win.Activate()
        except Exception:
            pass
        if el is None:
            self._set_status(u"Selección cancelada.")
            return
        analysis = analyze_conjunto(self._doc, el)
        if not analysis.get(u"ok"):
            msg = analysis.get(u"error") or u"Selección inválida."
            _mostrar_aviso(self._uiapp, msg)
            self._set_status(msg)
            if prev and prev.get(u"ok"):
                self._analysis = prev
            return
        self._analysis = analysis
        self._populate_selection(analysis)
        self._set_status(
            u"{0} capas longitudinales leídas · GUID válido".format(
                len(analysis.get(u"layers") or [])
            )
        )

    def _populate_selection(self, analysis):
        w = self._win
        seed = analysis.get(u"seed")
        doc = self._doc
        bt = _bar_type_of(seed, doc)
        diam = _nominal_diam_mm(bt)
        qty = _cantidad_posiciones(seed)
        eid = _element_id_int(seed.Id) if seed else u"?"
        diam_txt = u"Ø{0}".format(diam) if diam else u"Ø?"
        w.FindName(u"TxtElement").Text = u"Rebar Id {0} · {1}×{2}".format(
            eid, qty, diam_txt
        )
        w.FindName(u"TxtHost").Text = _host_label(analysis.get(u"host"), doc)
        gid = analysis.get(u"guid") or u""
        w.FindName(u"TxtGuid").Text = gid

        n_rebars = analysis.get(u"rebar_count") or 0
        n_ties = len(analysis.get(u"ties") or [])
        info = u"{0} rebars en corrida".format(n_rebars)
        if n_ties:
            info += u" · Incluye trabas"
        w.FindName(u"TxtRunInfo").Text = info

        rows_host = w.FindName(u"PanelLayerRows")
        if rows_host is not None:
            rows_host.Children.Clear()
            for ly in analysis.get(u"layers") or []:
                rows_host.Children.Add(self._make_layer_row(ly))

        last = (analysis.get(u"layers") or [None])[-1]
        max_disp = _max_capa_display(analysis.get(u"layers") or [])
        if last:
            w.FindName(u"TxtLastLayer").Text = (
                u"Última capa detectada: C{0} · origen de offset; "
                u"capas nuevas continúan desde C{1}."
            ).format(max_disp, max_disp + 1)

        ties = analysis.get(u"ties") or []
        w.FindName(u"TxtTies").Text = _ties_summary(doc, ties)
        lot = _lowest_z_tie_lot(doc, ties)
        tb_lot = w.FindName(u"TxtTiesLot")
        if tb_lot is not None:
            if lot:
                dmm = lot.get(u"diameter_mm")
                diam_l = u"Ø{0}".format(dmm) if dmm else u"Ø?"
                tb_lot.Text = (
                    u"Canvas: lote Z más bajo → {0}×{1} · e={2} mm · Z={3} mm. "
                    u"Activa el toggle de trabas para copiar solo trabas "
                    u"(no estribos) en capas nuevas."
                ).format(
                    lot.get(u"qty") or 0,
                    diam_l,
                    lot.get(u"spacing_mm") or 0,
                    lot.get(u"z_mm") or 0,
                )
            else:
                tb_lot.Text = u"Sin estribos/trabas en este GUID para el canvas."

        self._set_card_accent(w.FindName(u"BrdCardSeleccion"), False)
        self._set_card_accent(w.FindName(u"PanelLayers"), False)
        self._set_card_accent(w.FindName(u"PanelForm"), True)
        self._update_trabas_hint()

        if last:
            _select_combo_int(
                w.FindName(u"CmbQty"), last.get(u"qty") or 4, _QTY_OPTIONS
            )
            dmm = last.get(u"diameter_mm")
            cmb = w.FindName(u"CmbDiam")
            if cmb is not None and dmm:
                label = u"Ø{0}".format(int(dmm))
                try:
                    for i in range(int(cmb.Items.Count)):
                        if _as_unicode(cmb.Items[i]) == label:
                            cmb.SelectedIndex = i
                            break
                except Exception:
                    pass

        self._refresh_preview()
        btn = w.FindName(u"BtnCreate")
        if btn is not None:
            btn.IsEnabled = True
        try:
            # Filas de capas añadidas: recalcular alto para SizeToContent
            self._win.InvalidateMeasure()
            self._win.UpdateLayout()
            self._fit_canvas_height_to_rail()
            self._redraw_section_canvas()
        except Exception:
            pass

    def _make_layer_row(self, ly):
        g = Grid()
        g.Background = _hex_brush(u"#0E1B32")
        widths = (
            GridLength(48.0),
            GridLength(1, GridUnitType.Star),
            GridLength(56.0),
            GridLength(72.0),
        )
        for wth in widths:
            cd = ColumnDefinition()
            cd.Width = wth
            g.ColumnDefinitions.Add(cd)

        def _cell(col, text, muted=False):
            tb = TextBlock()
            tb.Text = _as_unicode(text)
            tb.FontSize = 11
            tb.Foreground = _hex_brush(FG_MUTED if muted else FG_TITLE)
            if col == 0:
                tb.Margin = Thickness(10, 7, 0, 7)
            elif col == 3:
                tb.Margin = Thickness(0, 7, 10, 7)
            else:
                tb.Margin = Thickness(0, 7, 0, 7)
            Grid.SetColumn(tb, col)
            g.Children.Add(tb)

        diam = ly.get(u"diameter_mm")
        diam_txt = u"Ø{0}".format(diam) if diam else u"—"
        sp = ly.get(u"spacing_mm") or 0
        bars = u"{0} barras".format(ly.get(u"qty") or 0)
        if sp:
            bars += u" · e={0} mm".format(sp)
        off = ly.get(u"offset_from_prev_mm")
        off_txt = u"—" if off is None else u"{0} mm".format(off)
        _cell(0, u"C{0}".format(ly.get(u"display")))
        _cell(1, bars)
        _cell(2, diam_txt)
        _cell(3, off_txt, muted=True)
        return g

    def _read_form(self):
        w = self._win
        n = _combo_int(w.FindName(u"CmbNCapas"), 2)
        qty = _combo_int(w.FindName(u"CmbQty"), 4)
        spacing = _parse_int_box(w.FindName(u"TxtSpacing"), 0, minimum=0)
        diam = 16
        cmb = w.FindName(u"CmbDiam")
        if cmb is not None and cmb.SelectedItem is not None:
            m = re.search(r"(\d+)", _as_unicode(cmb.SelectedItem))
            if m:
                diam = int(m.group(1))
        add_trabas = self._add_trabas_checked()
        return n, qty, diam, spacing, add_trabas

    def _refresh_preview(self):
        w = self._win
        host = w.FindName(u"PanelPreview")
        if host is not None:
            host.Children.Clear()
        if not self._analysis or not self._analysis.get(u"ok"):
            self._redraw_section_canvas()
            return
        n, qty, diam, spacing, add_trabas = self._read_form()
        layers = self._analysis.get(u"layers") or []
        base_display = _max_capa_display(layers)
        gid = self._analysis.get(u"guid") or u""
        hint = w.FindName(u"TxtGuidHint")
        if hint is not None:
            short = gid[:8] + u"…" if len(gid) > 8 else gid
            hint.Text = u"GUID heredado de la selección → {0}".format(short)

        btn = w.FindName(u"BtnCreate")
        if btn is not None:
            btn.IsEnabled = n >= 1 and spacing > 0
            try:
                if add_trabas:
                    btn.Content = u"Crear capas + trabas"
                else:
                    btn.Content = u"Crear capas"
            except Exception:
                pass

        self._update_trabas_hint()

        meta = w.FindName(u"TxtCanvasMeta")
        if meta is not None:
            meta.Text = (
                u"e = {0} mm · N = {1} · qty nuevas = {2}"
                u"{3}"
            ).format(
                spacing,
                n,
                qty,
                u" · +trabas" if add_trabas else u"",
            )

        if host is not None and n >= 1:
            col_widths = (
                GridLength(56.0),
                GridLength(90.0),
                GridLength(56.0),
                GridLength(1, GridUnitType.Star),
            )
            for i in range(n):
                g = Grid()
                g.Background = _hex_brush(u"#0E1B32")
                for wth in col_widths:
                    cd = ColumnDefinition()
                    cd.Width = wth
                    g.ColumnDefinitions.Add(cd)
                off_txt = (
                    u"{0} mm desde última".format(spacing)
                    if i == 0
                    else u"{0} mm entre capas".format(spacing)
                )
                if add_trabas:
                    off_txt += u" · +traba"
                texts = (
                    u"C{0}".format(base_display + i + 1),
                    u"{0} barras".format(qty),
                    u"Ø{0}".format(diam),
                    off_txt,
                )
                colors = (
                    _COLOR_PROPOSED,
                    FG_TITLE,
                    FG_TITLE,
                    FG_MUTED,
                )
                for col in range(4):
                    tb = TextBlock()
                    tb.Text = texts[col]
                    tb.FontSize = 11 if col < 3 else 10
                    if col == 0:
                        tb.FontWeight = FontWeights.SemiBold
                    tb.Foreground = _hex_brush(colors[col])
                    tb.Margin = Thickness(
                        10 if col == 0 else 0, 6, 10 if col == 3 else 0, 6
                    )
                    Grid.SetColumn(tb, col)
                    g.Children.Add(tb)
                if i > 0:
                    b = Border()
                    b.BorderBrush = _hex_brush(BORDER)
                    b.BorderThickness = Thickness(0, 1, 0, 0)
                    b.Child = g
                    host.Children.Add(b)
                else:
                    host.Children.Add(g)

        self._update_trabas_hint()
        self._redraw_section_canvas()

    def _redraw_section_canvas(self):
        w = self._win
        canv = self._section_canvas
        if canv is None:
            try:
                canv = w.FindName(u"SectionCanvas")
                self._section_canvas = canv
            except Exception:
                canv = None
        if canv is None:
            self._set_status(u"Canvas de sección no disponible (SectionCanvas).")
            return

        try:
            canv.Children.Clear()
        except Exception as ex:
            self._set_status(
                u"No se pudo limpiar el canvas: {0}".format(_as_unicode(ex))
            )
            return

        try:
            self._paint_section_canvas(canv)
        except Exception as ex:
            msg = u"Error al dibujar sección: {0}".format(_as_unicode(ex))
            self._set_status(msg)
            try:
                print(msg)
            except Exception:
                pass
            try:
                tb = TextBlock()
                tb.Text = msg
                tb.Foreground = _hex_brush(u"#F87171")
                tb.FontSize = 11
                tb.TextWrapping = TextWrapping.Wrap
                tb.Width = 360
                Canvas.SetLeft(tb, 12)
                Canvas.SetTop(tb, 12)
                canv.Children.Add(tb)
            except Exception:
                pass

    def _paint_section_canvas(self, canv):
        """Dibuja hormigón + capas existentes/propuestas + estribo Z↓ + trabas opt-in."""
        width, height = self._canvas_draw_size(canv)

        if not self._analysis or not self._analysis.get(u"ok"):
            tb = TextBlock()
            tb.Text = u"Sin datos de sección."
            tb.Foreground = _hex_brush(FG_MUTED)
            tb.FontSize = 11
            Canvas.SetLeft(tb, 12)
            Canvas.SetTop(tb, 12)
            canv.Children.Add(tb)
            return

        n, qty, diam, spacing, add_trabas = self._read_form()
        layers = self._analysis.get(u"layers") or []
        if not layers:
            tb = TextBlock()
            tb.Text = u"Sin capas longitudinales para dibujar."
            tb.Foreground = _hex_brush(FG_MUTED)
            tb.FontSize = 11
            Canvas.SetLeft(tb, 12)
            Canvas.SetTop(tb, 12)
            canv.Children.Add(tb)
            return

        host = self._analysis.get(u"host")
        wall_mm = float(_host_thickness_mm(host) or 300)
        cover = float(_COVER_EXT_MM_DEFAULT)
        # Espaciado esquemático si Δ prev. ≈0 (evita apilar C1..Cn en una columna)
        fallback_step = float(spacing) if spacing and spacing > 0 else 50.0
        inward = self._analysis.get(u"inward")

        pad_l, pad_r = 48.0, 56.0
        pad_bottom = 48.0
        wall_top = 48.0
        wall_h = max(160.0, height - wall_top - pad_bottom)
        usable_w = max(40.0, width - pad_l - pad_r)

        # Posiciones mm (existentes + propuestas)
        items = []
        exist_xs = _canvas_existing_x_positions_mm(
            layers, cover, fallback_step, inward
        )
        for i, ly in enumerate(layers):
            x_mm = exist_xs[i] if i < len(exist_xs) else cover
            items.append(
                {
                    u"x_mm": x_mm,
                    u"kind": u"existing",
                    u"label": u"C{0}".format(ly.get(u"display")),
                    u"qty": max(1, int(ly.get(u"qty") or 1)),
                    u"meta": u"{0}×Ø{1}".format(
                        ly.get(u"qty") or 0, ly.get(u"diameter_mm") or u"?"
                    ),
                }
            )
        x_mm = items[-1][u"x_mm"] if items else cover
        step = max(0.0, float(spacing or 0))
        base_disp = _max_capa_display(layers)
        for i in range(max(0, int(n))):
            # k×e desde la última (e=0 → coincide con última; sigue visible en cyan)
            x_mm += step
            items.append(
                {
                    u"x_mm": x_mm,
                    u"kind": u"proposed",
                    u"label": u"C{0}".format(base_disp + i + 1),
                    u"qty": max(1, int(qty)),
                    u"meta": u"{0}×Ø{1}".format(qty, diam),
                }
            )

        # Escala: que quepan host y última capa (evita ClipToBounds vacío)
        max_x = max([it[u"x_mm"] for it in items] + [cover])
        span_mm = max(float(wall_mm), max_x + cover, 1.0)

        def mm_to_x(mm):
            return pad_l + (float(mm) / span_mm) * usable_w

        max_bars = max([it[u"qty"] for it in items] + [1])
        bar_r = 5.5 if max_bars > 10 else 8.0
        margin_px = bar_r + 4.0

        # Hormigón
        rect = Rectangle()
        rect.Width = usable_w
        rect.Height = wall_h
        rect.Fill = _hex_brush(_COLOR_CONCRETE)
        rect.Stroke = _hex_brush(BORDER)
        rect.StrokeThickness = 1.5
        Canvas.SetLeft(rect, pad_l)
        Canvas.SetTop(rect, wall_top)
        canv.Children.Add(rect)

        for i in range(8):
            x0 = pad_l + 8 + i * (usable_w / 7.0)
            ln = Line()
            ln.X1, ln.Y1 = x0, wall_top + 4
            ln.X2, ln.Y2 = x0 - 28, wall_top + wall_h - 4
            ln.Stroke = _hex_brush(BORDER)
            ln.StrokeThickness = 1
            ln.Opacity = 0.35
            canv.Children.Add(ln)

        def _label(text, x, y, anchor=u"left", color=FG_MUTED, size=9):
            tb = TextBlock()
            tb.Text = _as_unicode(text)
            tb.Foreground = _hex_brush(color)
            tb.FontSize = size
            canv.Children.Add(tb)
            try:
                tb.Measure(WpfSize(4000.0, 400.0))
                tw = float(tb.DesiredSize.Width)
            except Exception:
                tw = 40.0
            if anchor == u"end":
                Canvas.SetLeft(tb, x - tw)
            elif anchor == u"middle":
                Canvas.SetLeft(tb, x - tw * 0.5)
            else:
                Canvas.SetLeft(tb, x)
            Canvas.SetTop(tb, y)

        face_y = wall_top + wall_h + 12
        thick_y = wall_top + wall_h + 28
        _label(u"Exterior", pad_l, face_y)
        _label(u"Interior", pad_l + usable_w, face_y, anchor=u"end")
        dim = Line()
        dim.X1, dim.Y1 = pad_l, thick_y
        dim.X2, dim.Y2 = pad_l + usable_w, thick_y
        dim.Stroke = _hex_brush(FG_MUTED)
        dim.StrokeThickness = 1
        canv.Children.Add(dim)
        host_label = u"{0} mm (host)".format(int(wall_mm))
        if abs(span_mm - wall_mm) > 1.0:
            host_label = u"{0} mm host · vista {1:.0f} mm".format(
                int(wall_mm), span_mm
            )
        _label(host_label, pad_l + usable_w * 0.5, thick_y + 4, anchor=u"middle")

        tip_left = True
        existing_pts = []
        all_pts = []
        for it in items:
            cx = mm_to_x(it[u"x_mm"])
            for cy in _bar_ys_in_layer(it[u"qty"], wall_top, wall_h, bar_r):
                all_pts.append((cx, cy))
                if it[u"kind"] == u"existing":
                    existing_pts.append((cx, cy))

        # Barras primero (si el estribo falla, las capas siguen visibles)
        bars_drawn = 0
        for idx, it in enumerate(items):
            if idx > 0:
                prev = items[idx - 1]
                x1 = mm_to_x(prev[u"x_mm"])
                x2 = mm_to_x(it[u"x_mm"])
                # Distancia mostrada = separación real en canvas (mm)
                dist_mm = abs(float(it[u"x_mm"]) - float(prev[u"x_mm"]))
                if it[u"kind"] == u"proposed":
                    dist = spacing if spacing else int(round(dist_mm))
                else:
                    # Preferir Δ medido; si colapsó (~0), usar la separación del preview
                    dist = None
                    if idx < len(layers):
                        dist = layers[idx].get(u"offset_from_prev_mm")
                    if dist is None or abs(float(dist or 0)) < _MIN_CANVAS_LAYER_SEP_MM:
                        dist = int(round(dist_mm)) if dist_mm >= 1.0 else int(fallback_step)
                if dist is not None and abs(x2 - x1) > 8:
                    mid = (x1 + x2) * 0.5
                    ln = Line()
                    ln.X1, ln.Y1 = x1, wall_top - 14
                    ln.X2, ln.Y2 = x2, wall_top - 14
                    ln.Stroke = _hex_brush(
                        _COLOR_PROPOSED
                        if it[u"kind"] == u"proposed"
                        else FG_MUTED
                    )
                    ln.StrokeThickness = 1
                    canv.Children.Add(ln)
                    _label(
                        u"{0}".format(int(dist)),
                        mid,
                        wall_top - 28,
                        anchor=u"middle",
                        size=8,
                        color=(
                            _COLOR_PROPOSED
                            if it[u"kind"] == u"proposed"
                            else FG_MUTED
                        ),
                    )

            cx = mm_to_x(it[u"x_mm"])
            color = (
                _COLOR_EXISTING if it[u"kind"] == u"existing" else _COLOR_PROPOSED
            )
            for cy in _bar_ys_in_layer(it[u"qty"], wall_top, wall_h, bar_r):
                el = Ellipse()
                el.Width = bar_r * 2
                el.Height = bar_r * 2
                el.Fill = _hex_brush(color)
                el.Stroke = _hex_brush(FG_TITLE)
                el.StrokeThickness = 0.6
                Canvas.SetLeft(el, cx - bar_r)
                Canvas.SetTop(el, cy - bar_r)
                canv.Children.Add(el)
                bars_drawn += 1
            _label(
                it[u"label"],
                cx,
                wall_top + wall_h + 2,
                anchor=u"middle",
                color=color,
                size=9,
            )

        # Estribo 135° — solo capas existentes (no amplía a propuestas)
        lot = _lowest_z_tie_lot(self._doc, self._analysis.get(u"ties") or [])
        if len(existing_pts) >= 2:
            try:
                cxs = [p[0] for p in existing_pts]
                cys = [p[1] for p in existing_pts]
                left = min(cxs) - margin_px
                right = max(cxs) + margin_px
                top = min(cys) - margin_px
                bot = max(cys) + margin_px
                bar_cx = max(cxs) if tip_left else min(cxs)
                bar_cy = min(cys)
                brush = _hex_brush(_COLOR_STIRRUP)
                _draw_stirrup_135(
                    canv,
                    left,
                    top,
                    right,
                    bot,
                    bar_cx,
                    bar_cy,
                    margin_px,
                    brush,
                    _STIRRUP_STROKE,
                    tip_left=tip_left,
                )
                if lot:
                    dmm = lot.get(u"diameter_mm")
                    label = u"Estribo 135° Z↓ {0}×Ø{1} e={2}".format(
                        lot.get(u"qty") or 0,
                        dmm if dmm else u"?",
                        lot.get(u"spacing_mm") or 0,
                    )
                    _label(
                        label,
                        (left + right) * 0.5,
                        top - 14,
                        anchor=u"middle",
                        color=_COLOR_STIRRUP,
                        size=8,
                    )
            except Exception as ex_st:
                try:
                    print(
                        u"Estribo canvas: {0}".format(_as_unicode(ex_st))
                    )
                except Exception:
                    pass

        # Trabas propuestas (opt-in aditivo; no modifica estribo existente)
        if add_trabas:
            try:
                tie_brush = _hex_brush(_COLOR_TIE)
                tie_side = 1.0 if tip_left else -1.0
                for it in items:
                    if it[u"kind"] != u"proposed":
                        continue
                    cx = mm_to_x(it[u"x_mm"])
                    ys = _bar_ys_in_layer(it[u"qty"], wall_top, wall_h, bar_r)
                    if not ys:
                        continue
                    y_min, y_max = min(ys), max(ys)
                    x_tie = cx + tie_side * margin_px
                    ln = Line()
                    ln.X1, ln.Y1 = x_tie, y_min
                    ln.X2, ln.Y2 = x_tie, y_max
                    ln.Stroke = tie_brush
                    ln.StrokeThickness = _TIE_STROKE
                    canv.Children.Add(ln)
                    _draw_hook_135_traba_end(
                        canv,
                        cx,
                        y_min,
                        margin_px,
                        tip_left,
                        u"top",
                        tie_brush,
                        _TIE_STROKE,
                    )
                    _draw_hook_135_traba_end(
                        canv,
                        cx,
                        y_max,
                        margin_px,
                        tip_left,
                        u"bottom",
                        tie_brush,
                        _TIE_STROKE,
                    )
            except Exception as ex_tie:
                try:
                    print(u"Trabas canvas: {0}".format(_as_unicode(ex_tie)))
                except Exception:
                    pass

        if bars_drawn < 1:
            self._set_status(
                u"Canvas: no se dibujaron barras ({0} capas en análisis)."
                .format(len(layers))
            )
        else:
            # No pisar un status de creación; solo si está vacío o de preview
            st = self._win.FindName(u"TxtStatus")
            cur = _as_unicode(st.Text) if st is not None else u""
            if (not cur) or (u"capas longitudinales" in cur.lower()) or (
                u"canvas" in cur.lower()
            ):
                n_ex = sum(1 for it in items if it[u"kind"] == u"existing")
                n_pr = sum(1 for it in items if it[u"kind"] == u"proposed")
                self._set_status(
                    u"Canvas: {0} capa(s) existente(s) · {1} propuesta(s) · "
                    u"{2} barras dibujadas."
                    .format(n_ex, n_pr, bars_drawn)
                )

    def _on_create(self, sender, args):
        if not self._analysis or not self._analysis.get(u"ok"):
            _mostrar_aviso(self._uiapp, u"Selecciona primero una rebar con GUID.")
            return
        n, qty, diam, spacing, add_trabas = self._read_form()
        if n < 1:
            _mostrar_aviso(self._uiapp, u"Indica al menos 1 capa nueva.")
            return
        if spacing <= 0:
            _mostrar_aviso(
                self._uiapp,
                u"El distanciamiento entre capas debe ser mayor que 0 mm.",
            )
            return
        if self._analysis.get(u"inward") is None:
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo determinar la dirección hacia el interior del host.",
                content=(
                    u"Comprueba que la rebar tenga un host válido (muro/viga/columna) "
                    u"y que la última capa no coincida con el centro del elemento."
                ),
            )
            return

        self._pending_create = {
            u"n": n,
            u"qty": qty,
            u"diam": diam,
            u"spacing": spacing,
            u"add_trabas": add_trabas,
        }
        self._set_status(u"Creando capas adicionales…")
        btn = self._win.FindName(u"BtnCreate")
        if btn is not None:
            try:
                btn.IsEnabled = False
            except Exception:
                pass
            try:
                evt = self._create_event or _ensure_create_event()
                evt.Raise()
            except Exception as ex:
                if btn is not None:
                    try:
                        btn.IsEnabled = True
                    except Exception:
                        pass
                self._pending_create = None
                _mostrar_aviso(
                    self._uiapp,
                    u"No se pudo encolar la creación en Revit.",
                    content=_as_unicode(ex),
                )

    def _execute_create(self, uiapp):
        btn = None
        try:
            btn = self._win.FindName(u"BtnCreate")
        except Exception:
            btn = None
        pending = self._pending_create
        self._pending_create = None
        try:
            if not pending:
                self._set_status(u"Sin parámetros de creación.")
                _mostrar_aviso(uiapp, u"Sin parámetros de creación pendientes.")
                return
            uidoc = uiapp.ActiveUIDocument if uiapp is not None else None
            if uidoc is None:
                msg = u"No hay documento activo."
                self._set_status(msg)
                _mostrar_aviso(uiapp, msg)
                return
            doc = uidoc.Document
            self._uiapp = uiapp
            self._uidoc = uidoc
            self._doc = doc

            if not self._analysis or not self._analysis.get(u"ok"):
                msg = u"Selecciona primero una rebar con GUID."
                self._set_status(msg)
                _mostrar_aviso(uiapp, msg)
                return

            seed = self._analysis.get(u"seed")
            seed_fresh = _refresh_rebar(doc, seed)
            if seed_fresh is None:
                msg = u"La rebar de selección ya no existe en el documento."
                self._set_status(msg)
                _mostrar_aviso(uiapp, msg)
                return
            analysis = analyze_conjunto(doc, seed_fresh)
            if not analysis.get(u"ok"):
                msg = analysis.get(u"error") or u"Análisis inválido."
                self._set_status(msg)
                _mostrar_aviso(uiapp, msg)
                return
            self._analysis = analysis

            result = create_additional_layers(
                doc,
                analysis,
                n_layers=pending[u"n"],
                qty_per_layer=pending[u"qty"],
                diameter_mm=pending[u"diam"],
                spacing_mm=pending[u"spacing"],
                view=uidoc.ActiveView if uidoc is not None else None,
                add_trabas=bool(pending.get(u"add_trabas")),
            )
            self._set_status(result.get(u"message") or u"")
            if result.get(u"ok"):
                # No cerrar la ventana desde Execute. Marcar éxito para que
                # finally no re-habilite «Crear».
                self._create_done = True
                try:
                    btn_c = self._win.FindName(u"BtnCreate")
                    if btn_c is not None:
                        btn_c.IsEnabled = False
                        btn_c.Content = u"Capas creadas"
                except Exception:
                    pass
                return
            _mostrar_aviso(
                uiapp,
                result.get(u"message") or u"No se pudieron crear las capas.",
            )
        except Exception as ex:
            msg = u"Error al crear capas: {0}".format(_as_unicode(ex))
            self._set_status(msg)
            _mostrar_aviso(
                uiapp, u"Error al crear capas adicionales.", content=_as_unicode(ex)
            )
        finally:
            try:
                if self._win is None or not self._win.IsLoaded:
                    return
            except Exception:
                return
            if getattr(self, u"_create_done", False):
                return
            if btn is not None:
                try:
                    n, _qty, _d, spacing, _at = self._read_form()
                    btn.IsEnabled = bool(
                        self._analysis
                        and self._analysis.get(u"ok")
                        and n >= 1
                        and spacing > 0
                    )
                except Exception:
                    try:
                        btn.IsEnabled = True
                    except Exception:
                        pass

    def show(self):
        try:
            self._win.Show()
        except Exception:
            self._win.ShowDialog()


def run(revit):
    """Punto de entrada pyRevit: pick inmediato → analizar → UI poblada."""
    existing_ctrl = _get_active_controller()
    existing = None
    if existing_ctrl is not None:
        try:
            existing = existing_ctrl._win
        except Exception:
            existing = None
    if existing is None:
        existing = _get_active_window()
    if existing is not None:
        try:
            if existing.WindowState == WindowState.Minimized:
                existing.WindowState = WindowState.Normal
        except Exception:
            pass
        try:
            existing.Activate()
            existing.Focus()
        except Exception:
            pass
        _mostrar_aviso(revit, u"La herramienta ya esta en ejecucion.")
        return

    # Por si quedó un ExternalEvent/huérfano de una sesión previa mal cerrada.
    _clear_active_controller()

    uidoc = getattr(revit, u"ActiveUIDocument", None)
    if uidoc is None:
        _mostrar_aviso(revit, u"No hay documento activo.")
        return
    doc = uidoc.Document

    el = _pick_rebar_element(
        uidoc,
        u"Selecciona una rebar longitudinal con Armadura_Conjunto_GUID",
    )
    if el is None:
        _mostrar_aviso(revit, u"Operación cancelada.")
        return

    analysis = analyze_conjunto(doc, el)
    if not analysis.get(u"ok"):
        _mostrar_aviso(
            revit,
            analysis.get(u"error") or u"No se pudo analizar el conjunto GUID.",
        )
        return

    win = CapasAdicionalesGuidWindow(revit, analysis=analysis)
    win.show()
