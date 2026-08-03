# -*- coding: utf-8 -*-
"""
Arainco: Area Rein. losa (Sketch) — herramienta nueva.

Flujo:
  1. Pick Floor
  2. Contorno desde ``Floor.SketchId`` → ``Sketch.Profile`` (loop exterior)
  3. Detectar AreaReinforcement existentes en el Floor (visualización)
  4. UI WPF (planta a escala real): tabs Superior / Inferior; paños por cara
     (2 clics → rectángulo). Sin polígonos en una cara → no se crea AR ahí.
     — clic en paño = activo; cards de la cara activa mutan por paño
  5. ``AreaReinforcement.Create`` por cada paño de su cara
     (Major = luz menor del paño)
     — settings por paño (layer_cfg + ahorro); «Ahorro de fierro»
       por cara: si ON → 2 AreaReinforcement intercalados @ 2e con polígonos
       recortados ~10% en extremos opuestos; luz de distribución
       redondeada a floor(L/2e)·2e; si OFF → 1 AR completo a e
       (Superior e Inferior independientes; polígonos por cara)
     — post-create: RemoveAreaReinforcementSystem → Rebar libres →
       Show Middle + IndependentTag EST_A_STRUCTURAL REBAR TAG_FLOOR
       (tipo = RebarShape, fallback «01») + MRA «Recorrido Barras» en cada Rebar

Solo vistas de planta (``ViewPlan``; no alzado, sección, 3D ni plantilla).
Independiente de ``area_reinforcement_losa`` / Malla en losa existente.
Revit 2024+ | IronPython (pyRevit).
"""

from __future__ import print_function

import math
import os
import weakref

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System import AppDomain, EventHandler, TimeSpan
from System.Collections.Generic import List
from System.IO import MemoryStream
from System.Windows import (
    HorizontalAlignment,
    RoutedEventHandler,
    SizeChangedEventHandler,
    Thickness,
    WindowState,
)
from System.Windows.Controls import Canvas as WpfCanvas
from System.Windows.Controls import ComboBoxItem
from System.Windows.Controls import SelectionChangedEventHandler
from System.Windows.Media import (
    Color,
    CombinedGeometry,
    DoubleCollection,
    FillRule,
    GeometryCombineMode,
    Matrix,
    MatrixTransform,
    PathFigure,
    PathGeometry,
    PointCollection,
    SolidColorBrush,
    TranslateTransform,
)
from System.Windows.Threading import DispatcherTimer
from System.Windows.Media import LineSegment as WpfLineSegment
from System.Windows.Shapes import Line as WpfLine
from System.Windows.Shapes import Path as WpfPath
from System.Windows.Shapes import Polygon as WpfPolygon
from System.Windows.Controls import TextBlock
from System.Windows import FontWeights, Point as WpfPoint, VerticalAlignment
from System.Windows.Input import (
    Cursor,
    Cursors,
    Key,
    KeyEventHandler,
    Keyboard,
    KeyboardFocusChangedEventHandler,
    ModifierKeys,
    MouseButton,
    MouseButtonEventHandler,
    MouseEventHandler,
    MouseWheelEventHandler,
)
from System.Windows.Markup import XamlReader
from System.Windows.Shapes import Ellipse as WpfEllipse
from System.Windows.Shapes import Rectangle as WpfRectangle

from Autodesk.Revit.DB import (
    BoundingBoxIntersectsFilter,
    BuiltInCategory,
    BuiltInParameter,
    Curve,
    ElementId,
    ElementTypeGroup,
    FilteredElementCollector,
    Floor,
    Grid,
    IndependentTag,
    Line,
    LocationCurve,
    Opening,
    Outline,
    Reference,
    Sketch,
    StorageType,
    TagMode,
    TagOrientation,
    Transaction,
    TransactionGroup,
    UnitTypeId,
    UnitUtils,
    View,
    View3D,
    ViewPlan,
    Wall,
    XYZ,
)
try:
    from bimtools_joined_geometry import get_joined_element_ids
except Exception:
    get_joined_element_ids = None

try:
    from area_rein_losa_panos import (
        shoelace_area_m2,
        rect_from_two_points_mm,
        span_direction_from_polygon_mm,
        luz_menor_mm_from_polygon,
        point_in_polygon_mm,
        union_polygons_mm,
        polygons_form_single_component_mm,
        inset_polygon_mm,
        ahorro_fierro_polygons_mm,
        _line_seg_intersection as _pano_line_seg_intersection,
    )
except Exception:
    shoelace_area_m2 = None
    point_in_polygon_mm = None
    union_polygons_mm = None
    polygons_form_single_component_mm = None
    inset_polygon_mm = None
    ahorro_fierro_polygons_mm = None
    _pano_line_seg_intersection = None

    def rect_from_two_points_mm(p1, p2):
        if p1 is None or p2 is None:
            return None
        x0 = min(float(p1[0]), float(p2[0]))
        x1 = max(float(p1[0]), float(p2[0]))
        y0 = min(float(p1[1]), float(p2[1]))
        y1 = max(float(p1[1]), float(p2[1]))
        if (x1 - x0) < 1.0 or (y1 - y0) < 1.0:
            return None
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]

    def luz_menor_mm_from_polygon(pts):
        # Fallback local: AABB min(width, height) como area_rein_losa_panos.
        if not pts or len(pts) < 3:
            return None
        xs = []
        ys = []
        for p in pts:
            try:
                xs.append(float(p[0]))
                ys.append(float(p[1]))
            except Exception:
                continue
        if len(xs) < 3:
            return None
        width = max(xs) - min(xs)
        height = max(ys) - min(ys)
        if width < 1e-6 or height < 1e-6:
            return None
        return min(width, height)

    def span_direction_from_polygon_mm(pts, equal_tol_mm=1.0):
        # Fallback local: misma regla AABB / luz menor que area_rein_losa_panos.
        if not pts or len(pts) < 3:
            return None
        tol = float(equal_tol_mm)
        if tol < 0.0:
            tol = 0.0
        xs = []
        ys = []
        for p in pts:
            try:
                xs.append(float(p[0]))
                ys.append(float(p[1]))
            except Exception:
                continue
        if len(xs) < 3:
            return None
        width = max(xs) - min(xs)
        height = max(ys) - min(ys)
        if width < 1e-6 or height < 1e-6:
            return None
        if abs(width - height) <= tol:
            return (1.0, 0.0)
        if width < height:
            return (1.0, 0.0)
        return (0.0, 1.0)
from Autodesk.Revit.DB.Structure import (
    AreaReinforcement,
    AreaReinforcementLayerType,
    AreaReinforcementType,
    Rebar,
    RebarBarType,
    RebarInSystem,
    RebarPresentationMode,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from bimtools_instruction_dialog import show_message_dialog
from bimtools_rebar_3d_visibility import apply_reinforcement_unobscured_in_view
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_ui_tokens import WINDOW_CHROME_TITLE
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)

try:
    from conjunto_guid import (
        ARMADURA_UBICACION_INFERIOR,
        ARMADURA_UBICACION_SUPERIOR,
        finalizar_armadura_conjunto_guid_ejecucion,
        iniciar_armadura_conjunto_guid_ejecucion,
        stamp_armadura_arainco,
        stamp_armadura_conjunto_guid,
        stamp_armadura_malla,
        stamp_armadura_nivel,
        stamp_armadura_posicion,
        stamp_armadura_ubicacion,
    )
except Exception:
    ARMADURA_UBICACION_INFERIOR = u"F"
    ARMADURA_UBICACION_SUPERIOR = u"F'"
    finalizar_armadura_conjunto_guid_ejecucion = None
    iniciar_armadura_conjunto_guid_ejecucion = None
    stamp_armadura_arainco = None
    stamp_armadura_conjunto_guid = None
    stamp_armadura_malla = None
    stamp_armadura_nivel = None
    stamp_armadura_posicion = None
    stamp_armadura_ubicacion = None

_DIALOG_TITLE = u"Arainco: Area Rein. losa"
_SINGLETON_KEY = u"Arainco.AreaReinLosaSketch.ActiveWindow"
_ARROW_PLUS_CURSOR_KEY = u"Arainco.AreaReinLosaSketch.ArrowPlusCursor"
_ARROW_PLUS_STREAM_KEY = u"Arainco.AreaReinLosaSketch.ArrowPlusCursorStream"
_FT_TO_MM = 304.8

# CUR 32x32 flecha + plus (hotspot 0,0). Respaldo si falta arrow_plus.cur junto al script.
_ARROW_PLUS_CUR_B64 = (
    u"AAACAAEAICAAAAAAAACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/AAAA/wAA"
    u"AP8AAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////////wAAAP8AAAD/AAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAD/AAAA/////////////////wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAD/AAAA/wAAAP8AAAAAAAAAAAAAAP//////////////"
    u"//8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AP//////AAAA/wAAAP8AAAD/AAAA/////////////////wAAAP8AAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////AAAA/wAA"
    u"AP////////////////8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAD/////////////////AAAA/////////////////wAA"
    u"AP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AP//////////////////////////////////////AAAA/wAAAP8AAAD/AAAA/wAA"
    u"AP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////////////"
    u"////////////////////////////////////////AAAA/wAAAAAAAAAAAAAAAAAA"
    u"AAAAAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA"
    u"AAAAAAAAAAAAAAAAAAAAAAD/////////////////////////////////////////"
    u"/////////////wAAAP8AAAD/AAAAAAAAAAAAAAAAAAAAAAAAAP//////////////"
    u"//////////////////////////////////8AAAD/AAAAAAAAAAAAAAAAAAAAAAAA"
    u"AP////////////////////////////////////////////////8AAAD/AAAA/wAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////////////////8AAAD/////////"
    u"/////////////wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAA////////////////////"
    u"////////////////////////AAAA/wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAD/////////////////AAAA/wAAAP//////////////////////AAAA/wAA"
    u"AAAAAAAAAAAAAAAAAAAAAAD//////////////////////////////////////wAA"
    u"AP8AAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////////"
    u"//8AAAD/AAAA//////////////////////8AAAD/AAAAAAAAAAAAAAAAAAAAAAAA"
    u"AP////////////////////////////////8AAAD/AAAA/wAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAA//////8AAAD/AAAA/wAAAP8AAAD/AAAA/wAA"
    u"AP8AAAD//////wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAA////////////////////"
    u"////////AAAA/wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAD///////////8AAAD/AAAA/wAAAP8AAAD/AAAA////////////AAAA/wAA"
    u"AAAAAAAAAAAAAAAAAAAAAAD//////////////////////wAAAP8AAAD/AAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////////"
    u"//8AAAD/AAAA//////////////////////8AAAD/AAAAAAAAAAAAAAAAAAAAAAAA"
    u"AP////////////////8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////////////////8AAAD/////////"
    u"/////////////wAAAP8AAAAAAAAAAAAAAAAAAAAAAAAA////////////AAAA/wAA"
    u"AP8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAD/////////////////////////////////////////////////AAAA/wAA"
    u"AAAAAAAAAAAAAAAAAAAAAAD//////wAAAP8AAAD/AAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8AAAD/AAAA/wAA"
    u"AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAAAAAAAAAAAAAAAAAAAAAA"
    u"AP8AAAD/AAAA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    u"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////////////////"
    u"/////////////////////////////////////////D////wf///4H///GB///wA/"
    u"//8AP///AH///wAH//8AB4APAAeADwAPgA8AH4APAD+ADwB/gA8A/4APAf+ADwP/"
    u"gA8H/4APD/+ADx////8="
)


def _ctrl_modifier_down():
    try:
        mods = Keyboard.Modifiers
        return (mods & ModifierKeys.Control) == ModifierKeys.Control
    except Exception:
        return False


def _get_arrow_plus_cursor():
    """Cursor flecha+plus para Ctrl (fusión). Cache AppDomain; fallback Arrow/Cross."""
    try:
        cached = AppDomain.CurrentDomain.GetData(_ARROW_PLUS_CURSOR_KEY)
        if cached is not None:
            return cached
    except Exception:
        pass

    cur = None
    # 1) Archivo junto al script
    try:
        path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), u"arrow_plus.cur"
        )
        if os.path.isfile(path):
            cur = Cursor(path)
    except Exception:
        cur = None

    # 2) Bytes embebidos (MemoryStream debe vivir con el Cursor)
    if cur is None:
        try:
            from System import Convert as _SysConvert
            arr = _SysConvert.FromBase64String(_ARROW_PLUS_CUR_B64)
            ms = MemoryStream(arr)
            cur = Cursor(ms)
            try:
                AppDomain.CurrentDomain.SetData(_ARROW_PLUS_STREAM_KEY, ms)
            except Exception:
                pass
        except Exception:
            cur = None

    if cur is None:
        try:
            cur = Cursors.Arrow
        except Exception:
            cur = Cursors.Cross

    try:
        AppDomain.CurrentDomain.SetData(_ARROW_PLUS_CURSOR_KEY, cur)
    except Exception:
        pass
    return cur

_LAYER_KEYS = (
    u"exterior_major",
    u"exterior_minor",
    u"interior_major",
    u"interior_minor",
)

_LAYER_LABELS = {
    u"exterior_major": u"Principal (luz menor)",
    u"exterior_minor": u"Secundaria",
    u"interior_major": u"Principal (luz menor)",
    u"interior_minor": u"Secundaria",
}

_LAYER_LABEL_TOOLTIPS = {
    u"exterior_major": u"Dirección Principal (Major en Revit) — luz menor del paño",
    u"exterior_minor": u"Dirección Secundaria (Minor en Revit)",
    u"interior_major": u"Dirección Principal (Major en Revit) — luz menor del paño",
    u"interior_minor": u"Dirección Secundaria (Minor en Revit)",
}

# Grupos visuales en el rail Capas (tabs: Inferior izq., Superior der.).
# Toggle malla OFF por defecto; al agregar un polígono se activa en esa cara.
# Superior → exterior_* (TopOrFront); Inferior → interior_* (BottomOrBack).
_LAYER_GROUPS = (
    {
        u"id": u"inferior",
        u"title": u"Malla Inferior",
        u"pill": u"INF",
        u"color": u"#4ade80",
        u"keys": (u"interior_major", u"interior_minor"),
    },
    {
        u"id": u"superior",
        u"title": u"Malla Superior",
        u"pill": u"SUP",
        u"color": u"#5BC0DE",
        u"keys": (u"exterior_major", u"exterior_minor"),
    },
)

_FACE_KEYS = {
    u"superior": (u"exterior_major", u"exterior_minor"),
    u"inferior": (u"interior_major", u"interior_minor"),
}
_FACE_PILL = {u"superior": u"SUP", u"inferior": u"INF"}
_FACE_AHORRO_KEY = {
    u"superior": u"ahorro_superior",
    u"inferior": u"ahorro_inferior",
}


def _normalize_face_id(face_id):
    if face_id in (u"superior", u"inferior"):
        return face_id
    return u"inferior"


def _keys_for_face(face_id):
    return _FACE_KEYS.get(_normalize_face_id(face_id)) or _FACE_KEYS[u"inferior"]

_LAYER_COLORS = {
    u"exterior_major": u"#5BC0DE",
    u"exterior_minor": u"#7dd3e8",
    u"interior_major": u"#4ade80",
    u"interior_minor": u"#86efac",
}

_LAYER_TYPE = {
    u"exterior_major": AreaReinforcementLayerType.TopOrFrontMajor,
    u"exterior_minor": AreaReinforcementLayerType.TopOrFrontMinor,
    u"interior_major": AreaReinforcementLayerType.BottomOrBackMajor,
    u"interior_minor": AreaReinforcementLayerType.BottomOrBackMinor,
}

# TopOrFront = superior (F') · BottomOrBack = inferior (F)
_LAYER_TOP = (
    AreaReinforcementLayerType.TopOrFrontMajor,
    AreaReinforcementLayerType.TopOrFrontMinor,
)
_LAYER_BOTTOM = (
    AreaReinforcementLayerType.BottomOrBackMajor,
    AreaReinforcementLayerType.BottomOrBackMinor,
)
_LAYER_MAJOR = (
    AreaReinforcementLayerType.TopOrFrontMajor,
    AreaReinforcementLayerType.BottomOrBackMajor,
)
# Inferior Major/Minor → i/s; Superior Major/Minor → s/i
_LAYER_POSICION = {
    u"interior_major": u"i",
    u"interior_minor": u"s",
    u"exterior_major": u"s",
    u"exterior_minor": u"i",
}
_LAYER_TYPE_POSICION = {
    AreaReinforcementLayerType.BottomOrBackMajor: u"i",
    AreaReinforcementLayerType.BottomOrBackMinor: u"s",
    AreaReinforcementLayerType.TopOrFrontMajor: u"s",
    AreaReinforcementLayerType.TopOrFrontMinor: u"i",
}

_LAYER_BAR_TYPE_PARAM_NAMES = {
    u"exterior_major": (u"Exterior Major Bar Type", u"Exterior Major Rebar Type"),
    u"exterior_minor": (u"Exterior Minor Bar Type", u"Exterior Minor Rebar Type"),
    u"interior_major": (u"Interior Major Bar Type", u"Interior Major Rebar Type"),
    u"interior_minor": (u"Interior Minor Bar Type", u"Interior Minor Rebar Type"),
}

_SPACING_OPTS_MM = (100, 125, 150, 200, 250, 300)

# Overlay canvas (magnitud real)
_CTX_WALL = u"wall"
_CTX_BEAM = u"beam"
_CTX_PASADA = u"pasada"
_CTX_GRID = u"grid"
_CTX_COLORS = {
    # Muro: terracota (distinto del azul-gris de la losa)
    _CTX_WALL: (u"#2a1814", u"#d4957a"),
    _CTX_BEAM: (u"#2a2818", u"#c4b86a"),
    # Mismo trazo que borde de losa (#95B8CC); sin estilo «cruz»/dash
    _CTX_PASADA: (u"#071018", u"#95B8CC"),
    _CTX_GRID: (u"#1a1510", u"#d4a574"),
}
_GRID_CLIP_PAD_MM = 5000.0  # margen alrededor de la losa (ejes → snap)
_BEAM_WIDTH_FALLBACK_MM = 300.0

_PANO_COLORS = (
    u"#5BC0DE",
    u"#4ade80",
    u"#fbbf24",
    u"#a78bfa",
    u"#f87171",
    u"#2dd4bf",
    u"#fb923c",
    u"#e879f9",
)

# AreaReinforcement existentes en la losa (solo lectura / visualización)
_EXISTING_AR_FILL = u"#1e1333"
_EXISTING_AR_STROKE = u"#c084fc"
_EXISTING_AR_LABEL = u"#e9d5ff"
_MERGE_STROKE = u"#fbbf24"
_MERGE_FILL = u"#f59e0b"
# Paño activo (edición de cards) — distinto del multi-select de fusión
_ACTIVE_STROKE = u"#38bdf8"
_ACTIVE_FILL = u"#0ea5e9"

# Snap magnético en canvas (radio en píxeles de pantalla)
_SNAP_PX = 8.0
_SNAP_TAG = u"snapOverlay"
_HUD_SCALE_TAG = u"hudScaleBar"
_SNAP_VERTEX_PREF = 1.15  # umbral relativo: vértice gana sobre arista
_SNAP_CELL_MM = 1000.0  # índice espacial snap (mm); misma prioridad vértice/arista

# Cache de SolidColorBrush congelados: clave (hex6, alpha)
_BRUSH_CACHE = {}

# Inset solo al crear AreaReinforcement (no afecta canvas / pts UI)
_AR_INSET_MM = 25.0

# Ahorro de fierro: dos series intercaladas @ 2e, recorte ~10% en extremos opuestos
_AHORRO_CUTBACK_PCT = 10.0
_AHORRO_TOGGLE_ACCENT = u"#fbbf24"
# Card Inferior — barras en borde de losa (UI; comportamiento pendiente)
_BORDE_LOSA_ACCENT = u"#f59e0b"
_BORDE_LOSA_TIP = (
    u"Barras en borde de losa (cara inferior). "
    u"Comportamiento por definir."
)
# Temporalmente desactivados — poner True para reactivar UI / lógica
_FEATURE_AHORRO_FIERRO = False
_FEATURE_PATA_L = False
# Patas L: al crear (borde / shafts / huecos → extremos de barra)
try:
    from area_rein_losa_sketch_pata import (
        aplicar_patas_l_por_outline,
    )
except Exception:
    aplicar_patas_l_por_outline = None
try:
    from armado_muros_txn import attach_rebar_outside_host_swallower
except Exception:
    attach_rebar_outside_host_swallower = None
_AHORRO_TIP = (
    u"2 series intercaladas a 2×espaciado, extremos recortados ~10%"
)

# Multi-Rebar Annotation (categoría Multi-Rebar Annotations) — mismo tipo que zapata de muro
_MRA_TYPE_NAME_RECORRIDO_BARRAS = u"Recorrido Barras"
# Offset lateral MRA respecto al array (mm), además de ½ proyección bbox
_MRA_OFFSET_EXTRA_MM = 300.0

# IndependentTag de barras (Show Middle) — convención corporativa EST_A
_REBAR_TAG_FAMILY_NAME = u"EST_A_STRUCTURAL REBAR TAG_FLOOR"
_REBAR_TAG_FALLBACK_TYPE = u"01"


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def _as_unicode(text):
    if text is None:
        return u""
    try:
        return unicode(text)
    except NameError:
        return str(text)


def _brush(hex_color, alpha=255):
    h = (_as_unicode(hex_color) or u"#95B8CC").lstrip(u"#")
    if len(h) != 6:
        h = u"95B8CC"
    try:
        a = int(alpha)
    except Exception:
        a = 255
    key = (h, a)
    cached = _BRUSH_CACHE.get(key)
    if cached is not None:
        return cached
    brush = SolidColorBrush(
        Color.FromArgb(a, int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    )
    try:
        if brush.CanFreeze:
            brush.Freeze()
    except Exception:
        pass
    _BRUSH_CACHE[key] = brush
    return brush


def _polygon_to_path_geometry(pts_mm, to_px):
    """PathGeometry cerrado desde puntos mm (vía to_px)."""
    if not pts_mm or len(pts_mm) < 3:
        return None
    try:
        fig = PathFigure()
        fig.IsClosed = True
        fig.IsFilled = True
        x0, y0 = to_px(pts_mm[0][0], pts_mm[0][1])
        fig.StartPoint = WpfPoint(x0, y0)
        for xmm, ymm in pts_mm[1:]:
            px, py = to_px(xmm, ymm)
            fig.Segments.Add(WpfLineSegment(WpfPoint(px, py), True))
        geo = PathGeometry()
        geo.Figures.Add(fig)
        geo.FillRule = FillRule.Nonzero
        return geo
    except Exception:
        return None


def _union_polygons_geometry(polygons_mm, to_px):
    """
    Unión booleana WPF de polígonos (encuentros monolíticos).
    Elimina aristas internas en solapes / L / T.
    """
    combined = None
    for pts in polygons_mm or []:
        g = _polygon_to_path_geometry(pts, to_px)
        if g is None:
            continue
        if combined is None:
            combined = g
        else:
            try:
                combined = CombinedGeometry(
                    GeometryCombineMode.Union, combined, g
                )
            except Exception:
                continue
    return combined


def _geometry_exclude(base_geo, mask_geo):
    """base − mask (enmascara). Si falla, devuelve base."""
    if base_geo is None:
        return None
    if mask_geo is None:
        return base_geo
    try:
        return CombinedGeometry(GeometryCombineMode.Exclude, base_geo, mask_geo)
    except Exception:
        return base_geo


def _collect_pasada_rings_mm(overlays, sketch_holes=None):
    """
    Anillos mm de pasadas/shafts (overlays) + huecos del Sketch de losa.

    Usados para restar del relleno del paño en canvas.
    """
    rings = []
    for ov in overlays or []:
        try:
            if ov.get(u"kind") != _CTX_PASADA:
                continue
        except Exception:
            continue
        pts = ov.get(u"pts") or []
        if pts and len(pts) >= 3:
            rings.append(pts)
    for hole in sketch_holes or []:
        if hole and len(hole) >= 3:
            rings.append(hole)
    return rings


def _pano_geometry_cut_by_pasadas(pts_mm, pasada_mask_geo, to_px):
    """
    PathGeometry del paño con pasadas/huecos restados (Exclude).

    Si no hay máscara o falla, geometría del polígono completo.
    """
    base = _polygon_to_path_geometry(pts_mm, to_px)
    if base is None:
        return None
    if pasada_mask_geo is None:
        return base
    return _geometry_exclude(base, pasada_mask_geo)


def _add_path_geometry(
    cv, geo, fill_hex, stroke_hex, stroke_w=1.4, fill_a=255, dashed=False
):
    if cv is None or geo is None:
        return
    path = WpfPath()
    path.Data = geo
    path.Fill = _brush(fill_hex, fill_a)
    path.Stroke = _brush(stroke_hex)
    path.StrokeThickness = stroke_w
    if dashed:
        try:
            dashes = DoubleCollection()
            dashes.Add(4)
            dashes.Add(3)
            path.StrokeDashArray = dashes
        except Exception:
            pass
    try:
        path.IsHitTestVisible = False
    except Exception:
        pass
    cv.Children.Add(path)


def _to_mm_identity(xmm, ymm):
    """Identity mapper: PathGeometry en coordenadas mm (sin proyección a px)."""
    return float(xmm), float(ymm)


def _flatten_path_geometry(geo):
    """Aplana CombinedGeometry → PathGeometry cacheable (Freeze si es posible)."""
    if geo is None:
        return None
    try:
        flat = geo.GetFlattenedPathGeometry()
    except Exception:
        flat = geo
    if flat is None:
        return None
    try:
        if not flat.IsFrozen:
            flat.Freeze()
    except Exception:
        pass
    return flat


def _view_matrix_mm_to_px(ox, oy, min_x, max_y, scale):
    """
    Matriz WPF: mm → px canvas.
    px = ox + (xmm - min_x) * scale
    py = oy + (max_y - ymm) * scale
    """
    s = float(scale)
    return Matrix(s, 0.0, 0.0, -s, float(ox) - float(min_x) * s, float(oy) + float(max_y) * s)


def _geo_transformed_for_view(geo_mm, matrix):
    """Clone + Geometry.Transform (stroke en px no escala con el zoom)."""
    if geo_mm is None:
        return None
    try:
        g = geo_mm.Clone()
    except Exception:
        g = geo_mm
    if g is None:
        return None
    try:
        g.Transform = MatrixTransform(matrix)
    except Exception:
        return g
    return g


def _wall_beam_pts_lists(overlays):
    """Listas de polígonos muro/viga (sin CombinedGeometry)."""
    wall_pts_list = [
        ov.get(u"pts")
        for ov in (overlays or [])
        if ov.get(u"kind") == _CTX_WALL and ov.get(u"pts")
    ]
    beam_pts_list = [
        ov.get(u"pts")
        for ov in (overlays or [])
        if ov.get(u"kind") == _CTX_BEAM and ov.get(u"pts")
    ]
    return wall_pts_list, beam_pts_list


def _build_wall_beam_geo_mm(overlays):
    """
    Unión muros / vigas en mm (una vez). Evita CombinedGeometry en cada zoom.
    No restar muros de vigas (Exclude dejaba cuñas/slivers en bordes).
    Returns: (wall_geo_mm, beam_geo_mm, wall_pts_list, beam_pts_list)
    """
    wall_pts_list, beam_pts_list = _wall_beam_pts_lists(overlays)
    wall_geo = None
    beam_geo = None
    wall_raw = None
    try:
        wall_raw = _union_polygons_geometry(wall_pts_list, _to_mm_identity)
    except Exception:
        wall_raw = None
    try:
        beam_raw = _union_polygons_geometry(beam_pts_list, _to_mm_identity)
        beam_geo = _flatten_path_geometry(beam_raw)
    except Exception:
        beam_geo = None
    try:
        wall_geo = _flatten_path_geometry(wall_raw)
    except Exception:
        wall_geo = None
    return wall_geo, beam_geo, wall_pts_list, beam_pts_list


def _element_id_int(eid):
    if eid is None:
        return None
    try:
        v = getattr(eid, "Value", None)
        if v is not None:
            return int(v)
    except Exception:
        pass
    try:
        return int(eid.IntegerValue)
    except Exception:
        return None


def _mostrar_aviso(uiapp, instruction, content=u""):
    try:
        hwnd = revit_main_hwnd(uiapp) if uiapp is not None else None
        show_message_dialog(
            _DIALOG_TITLE,
            instruction=_as_unicode(instruction),
            content=_as_unicode(content) if content else None,
            ok_text=u"Entendido",
            hwnd_revit=hwnd,
            uiapp=uiapp,
        )
        return
    except Exception:
        pass
    try:
        from Autodesk.Revit.UI import TaskDialog

        TaskDialog.Show(
            _DIALOG_TITLE,
            u"{0}\n{1}".format(
                _as_unicode(instruction), _as_unicode(content)
            ).strip(),
        )
    except Exception:
        pass


def _focus_existing(uiapp):
    try:
        win = AppDomain.CurrentDomain.GetData(_SINGLETON_KEY)
    except Exception:
        win = None
    if win is None:
        return False
    try:
        if not bool(win.IsLoaded):
            return False
    except Exception:
        return False
    try:
        if int(win.WindowState) == int(WindowState.Minimized):
            win.WindowState = WindowState.Normal
    except Exception:
        pass
    try:
        win.Activate()
        win.Focus()
    except Exception:
        pass
    _mostrar_aviso(uiapp, u"La herramienta ya esta en ejecucion.")
    return True


def _register_singleton(win):
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass


def _unregister_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Geometría Sketch
# ---------------------------------------------------------------------------


def obtener_loops_sketch(floor, document):
    """
    Devuelve lista de loops: cada loop es ``list`` de ``Curve``.
    Índice 0 = perímetro exterior; resto = huecos.
    """
    if floor is None or document is None or not isinstance(floor, Floor):
        return None
    try:
        sketch_id = floor.SketchId
        if sketch_id is None or sketch_id == ElementId.InvalidElementId:
            return None
        sketch = document.GetElement(sketch_id)
        if sketch is None or not isinstance(sketch, Sketch):
            return None
        profile = sketch.Profile
        if profile is None:
            return None
        n_loops = int(profile.Size)
        if n_loops < 1:
            return None
        loops = []
        for i in range(n_loops):
            curve_array = profile.get_Item(i)
            if curve_array is None:
                continue
            curves = []
            n_curves = int(curve_array.Size)
            for j in range(n_curves):
                c = curve_array.get_Item(j)
                if c is not None and c.IsBound:
                    try:
                        curves.append(c.Clone())
                    except Exception:
                        curves.append(c)
            if curves:
                loops.append(curves)
        return loops if loops else None
    except Exception:
        return None


def _plane_from_curves(curves):
    if not curves:
        return None
    try:
        from Autodesk.Revit.DB import CurveLoop

        loop = CurveLoop.Create(List[Curve](curves))
        if loop is not None and loop.HasPlane():
            return loop.GetPlane()
    except Exception:
        pass
    try:
        p0 = curves[0].GetEndPoint(0)
        return type(
            "P",
            (),
            {
                "Origin": p0,
                "XVec": XYZ(1, 0, 0),
                "YVec": XYZ(0, 1, 0),
                "Normal": XYZ(0, 0, 1),
            },
        )()
    except Exception:
        return None


def _xyz_to_plane_mm(pt, plane):
    """Proyecta XYZ al plano → (x_mm, y_mm) en ejes del Sketch/plane."""
    o = plane.Origin
    xv = plane.XVec
    yv = plane.YVec
    vx = float(pt.X) - float(o.X)
    vy = float(pt.Y) - float(o.Y)
    vz = float(pt.Z) - float(o.Z)
    x_ft = vx * float(xv.X) + vy * float(xv.Y) + vz * float(xv.Z)
    y_ft = vx * float(yv.X) + vy * float(yv.Y) + vz * float(yv.Z)
    return x_ft * _FT_TO_MM, y_ft * _FT_TO_MM


def _sample_curve_mm(curve, plane, n_arc=12):
    """Lista de (x_mm, y_mm) a lo largo de la curva (para dibujo)."""
    pts = []
    try:
        if hasattr(curve, "Tessellate"):
            for p in curve.Tessellate():
                pts.append(_xyz_to_plane_mm(p, plane))
            if pts:
                return pts
    except Exception:
        pass
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        pts.append(_xyz_to_plane_mm(p0, plane))
        # Arcos / no-lineas: muestrear parámetro
        is_line = False
        try:
            from Autodesk.Revit.DB import Line

            is_line = isinstance(curve, Line)
        except Exception:
            pass
        if not is_line:
            for k in range(1, n_arc):
                t = float(k) / float(n_arc)
                try:
                    p = curve.Evaluate(t, True)
                    pts.append(_xyz_to_plane_mm(p, plane))
                except Exception:
                    pass
        pts.append(_xyz_to_plane_mm(p1, plane))
    except Exception:
        pass
    return pts


def _loop_to_polyline_mm(curves, plane):
    pts = []
    for c in curves:
        seg = _sample_curve_mm(c, plane)
        if not seg:
            continue
        if pts and abs(pts[-1][0] - seg[0][0]) < 0.5 and abs(pts[-1][1] - seg[0][1]) < 0.5:
            pts.extend(seg[1:])
        else:
            pts.extend(seg)
    return pts


def _bbox_mm(pts):
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def _ring_closed_pts(pts):
    """Quita vértice de cierre duplicado si existe."""
    if not pts:
        return []
    out = list(pts)
    if len(out) > 1:
        a, b = out[0], out[-1]
        if abs(a[0] - b[0]) < 0.5 and abs(a[1] - b[1]) < 0.5:
            out = out[:-1]
    return out


def _append_ring_snap(verts, segs, pts, include_midpoints=True):
    """Añade vértices, aristas y (opc.) puntos medios de un anillo cerrado."""
    ring = _ring_closed_pts(pts)
    n = len(ring)
    if n < 2:
        return
    for p in ring:
        verts.append((float(p[0]), float(p[1])))
    for i in range(n):
        a = ring[i]
        b = ring[(i + 1) % n]
        ax, ay = float(a[0]), float(a[1])
        bx, by = float(b[0]), float(b[1])
        segs.append(((ax, ay), (bx, by)))
        if include_midpoints:
            verts.append(((ax + bx) * 0.5, (ay + by) * 0.5))


def _append_polyline_snap(verts, segs, pts, include_midpoints=True):
    """Añade vértices/aristas de una polilínea abierta (p. ej. ejes Grid)."""
    if not pts or len(pts) < 2:
        return
    poly = [(float(p[0]), float(p[1])) for p in pts]
    n = len(poly)
    for p in poly:
        verts.append(p)
    for i in range(n - 1):
        a = poly[i]
        b = poly[i + 1]
        segs.append((a, b))
        if include_midpoints:
            verts.append(((a[0] + b[0]) * 0.5, (a[1] + b[1]) * 0.5))


def _ring_edges_mm(pts):
    """Aristas ((ax,ay),(bx,by)) de un anillo cerrado en mm."""
    ring = _ring_closed_pts(pts)
    n = len(ring)
    if n < 2:
        return []
    out = []
    for i in range(n):
        a = ring[i]
        b = ring[(i + 1) % n]
        out.append(
            ((float(a[0]), float(a[1])), (float(b[0]), float(b[1])))
        )
    return out


def _seg_intersect_point_mm(a, b, c, d):
    """Punto de intersección de segmentos ab y cd, o None."""
    if _pano_line_seg_intersection is not None:
        try:
            hit = _pano_line_seg_intersection(a, b, c, d)
            if hit is None:
                return None
            return hit[0]
        except Exception:
            pass
    # Fallback local (misma lógica que area_rein_losa_panos)
    ax, ay = float(a[0]), float(a[1])
    bx, by = float(b[0]), float(b[1])
    cx, cy = float(c[0]), float(c[1])
    dx, dy = float(d[0]), float(d[1])
    r_x, r_y = bx - ax, by - ay
    s_x, s_y = dx - cx, dy - cy
    den = r_x * s_y - r_y * s_x
    if abs(den) < 1e-9:
        return None
    qp_x, qp_y = cx - ax, cy - ay
    t = (qp_x * s_y - qp_y * s_x) / den
    u = (qp_x * r_y - qp_y * r_x) / den
    if t < -1e-9 or t > 1.0 + 1e-9 or u < -1e-9 or u > 1.0 + 1e-9:
        return None
    t = max(0.0, min(1.0, t))
    return (ax + t * r_x, ay + t * r_y)


def _add_edge_pair_intersections_snap(verts, edges_a, edges_b, seen, tol):
    """Añade a ``verts`` intersecciones entre dos listas de aristas (con dedupe)."""
    if not edges_a or not edges_b:
        return
    for wa, wb in edges_a:
        try:
            wminx = min(wa[0], wb[0]) - tol
            wmaxx = max(wa[0], wb[0]) + tol
            wminy = min(wa[1], wb[1]) - tol
            wmaxy = max(wa[1], wb[1]) + tol
        except Exception:
            continue
        for ba, bb in edges_b:
            try:
                if max(ba[0], bb[0]) < wminx or min(ba[0], bb[0]) > wmaxx:
                    continue
                if max(ba[1], bb[1]) < wminy or min(ba[1], bb[1]) > wmaxy:
                    continue
            except Exception:
                pass
            pt = _seg_intersect_point_mm(wa, wb, ba, bb)
            if pt is None:
                continue
            try:
                key = (int(round(pt[0] / tol)), int(round(pt[1] / tol)))
            except Exception:
                continue
            if key in seen:
                continue
            seen.add(key)
            verts.append((float(pt[0]), float(pt[1])))


def _append_wall_beam_intersection_snap(verts, overlays, tol_mm=1.0):
    """
    Vértices de snap en cruces muro × viga y muro × muro
    (contornos del contexto planta).
    """
    if verts is None:
        return
    wall_rings = []  # lista de listas de aristas (un muro = un anillo)
    beam_edges = []
    for ov in overlays or []:
        try:
            kind = ov.get(u"kind")
            pts = ov.get(u"pts") or []
        except Exception:
            continue
        if kind == _CTX_WALL:
            edges = _ring_edges_mm(pts)
            if edges:
                wall_rings.append(edges)
        elif kind == _CTX_BEAM:
            beam_edges.extend(_ring_edges_mm(pts))
    if not wall_rings:
        return
    try:
        tol = max(0.25, float(tol_mm))
    except Exception:
        tol = 1.0
    seen = set()
    # Muro × viga
    if beam_edges:
        for w_edges in wall_rings:
            _add_edge_pair_intersections_snap(
                verts, w_edges, beam_edges, seen, tol
            )
    # Muro × muro (sin auto-intersección del mismo anillo)
    n_w = len(wall_rings)
    for i in range(n_w):
        for j in range(i + 1, n_w):
            _add_edge_pair_intersections_snap(
                verts, wall_rings[i], wall_rings[j], seen, tol
            )


def _closest_on_segment_mm(px, py, a, b):
    """Punto más cercano en segmento ab; retorna ((x,y), dist2)."""
    ax, ay = a[0], a[1]
    bx, by = b[0], b[1]
    dx = bx - ax
    dy = by - ay
    len2 = dx * dx + dy * dy
    if len2 < 1e-12:
        qx, qy = ax, ay
    else:
        t = ((px - ax) * dx + (py - ay) * dy) / len2
        if t < 0.0:
            t = 0.0
        elif t > 1.0:
            t = 1.0
        qx = ax + t * dx
        qy = ay + t * dy
    ddx = px - qx
    ddy = py - qy
    return (qx, qy), ddx * ddx + ddy * ddy


def _snap_cells_for_radius(px, py, radius, cell):
    """Celdas (ix, iy) que cubren el disco de radio ``radius`` en mm."""
    cell = float(cell)
    if cell < 1e-9:
        return
    r = float(radius)
    if r < 0.0:
        r = 0.0
    ix0 = int(math.floor((px - r) / cell))
    ix1 = int(math.floor((px + r) / cell))
    iy0 = int(math.floor((py - r) / cell))
    iy1 = int(math.floor((py + r) / cell))
    for ix in range(ix0, ix1 + 1):
        for iy in range(iy0, iy1 + 1):
            yield (ix, iy)


def _build_snap_cell_index(verts, segs, cell_mm=None):
    """Índice espacial por celdas para candidatos de snap (vértice / arista)."""
    cell = float(cell_mm if cell_mm is not None else _SNAP_CELL_MM)
    if cell < 1.0:
        cell = float(_SNAP_CELL_MM)
    vmap = {}
    smap = {}
    for v in verts or []:
        try:
            vx, vy = float(v[0]), float(v[1])
        except Exception:
            continue
        key = (int(math.floor(vx / cell)), int(math.floor(vy / cell)))
        bucket = vmap.get(key)
        if bucket is None:
            vmap[key] = [(vx, vy)]
        else:
            bucket.append((vx, vy))
    for seg in segs or []:
        try:
            a, b = seg[0], seg[1]
            ax, ay = float(a[0]), float(a[1])
            bx, by = float(b[0]), float(b[1])
        except Exception:
            continue
        seg_t = ((ax, ay), (bx, by))
        minx = ax if ax < bx else bx
        maxx = bx if ax < bx else ax
        miny = ay if ay < by else by
        maxy = by if ay < by else ay
        ix0 = int(math.floor(minx / cell))
        ix1 = int(math.floor(maxx / cell))
        iy0 = int(math.floor(miny / cell))
        iy1 = int(math.floor(maxy / cell))
        for ix in range(ix0, ix1 + 1):
            for iy in range(iy0, iy1 + 1):
                key = (ix, iy)
                bucket = smap.get(key)
                if bucket is None:
                    smap[key] = [seg_t]
                else:
                    bucket.append(seg_t)
    return {u"cell": cell, u"verts": vmap, u"segs": smap}


def _snap_point_mm(pt, verts, segs, thresh_mm, cell_index=None):
    """
    Snap a vértice (prioridad) o proyección sobre arista.
    Retorna (pt_snapped, kind) con kind in ('vertex','edge',None).
    ``cell_index`` opcional: mismos criterios, menos candidatos.
    """
    if pt is None or thresh_mm is None or thresh_mm <= 0:
        return pt, None
    px, py = float(pt[0]), float(pt[1])
    thresh2 = float(thresh_mm) * float(thresh_mm)
    cand_verts = verts
    cand_segs = segs
    if cell_index is not None:
        try:
            cell = float(cell_index[u"cell"])
            vmap = cell_index[u"verts"]
            smap = cell_index[u"segs"]
            if cell > 1e-9:
                cv = []
                r_v = float(thresh_mm) * float(_SNAP_VERTEX_PREF)
                for key in _snap_cells_for_radius(px, py, r_v, cell):
                    bucket = vmap.get(key)
                    if bucket:
                        cv.extend(bucket)
                cand_verts = cv
                cs = []
                for key in _snap_cells_for_radius(px, py, float(thresh_mm), cell):
                    bucket = smap.get(key)
                    if bucket:
                        cs.extend(bucket)
                cand_segs = cs
        except Exception:
            cand_verts = verts
            cand_segs = segs
    best_v = None
    best_vd2 = thresh2 * (_SNAP_VERTEX_PREF * _SNAP_VERTEX_PREF)
    for v in cand_verts or []:
        dx = px - float(v[0])
        dy = py - float(v[1])
        d2 = dx * dx + dy * dy
        if d2 <= best_vd2:
            best_vd2 = d2
            best_v = (float(v[0]), float(v[1]))
    if best_v is not None:
        return best_v, u"vertex"
    best_e = None
    best_ed2 = thresh2
    for seg in cand_segs or []:
        try:
            a, b = seg[0], seg[1]
        except Exception:
            continue
        q, d2 = _closest_on_segment_mm(px, py, a, b)
        if d2 <= best_ed2:
            best_ed2 = d2
            best_e = q
    if best_e is not None:
        return best_e, u"edge"
    return (px, py), None


def _curve_length_ft(curve):
    try:
        return float(curve.Length)
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return p0.DistanceTo(p1)
        except Exception:
            return 0.0


def direccion_arista_mas_larga(curves):
    """XYZ normalizado de la arista más larga del loop exterior."""
    best_len = -1.0
    best = XYZ(1, 0, 0)
    if not curves:
        return best
    for c in curves:
        try:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            d = XYZ(p1.X - p0.X, p1.Y - p0.Y, p1.Z - p0.Z)
            L = d.GetLength()
            if L > best_len and L > 1e-9:
                best_len = L
                best = XYZ(d.X / L, d.Y / L, d.Z / L)
        except Exception:
            continue
    return best


# ---------------------------------------------------------------------------
# Contexto planta: muros, vigas, pasadas (magnitud real)
# ---------------------------------------------------------------------------


def _category_id_int(elem):
    try:
        return int(elem.Category.Id.IntegerValue)
    except Exception:
        return None


def _es_muro(elem):
    if elem is None:
        return False
    try:
        if isinstance(elem, Wall):
            return True
    except Exception:
        pass
    return _category_id_int(elem) == int(BuiltInCategory.OST_Walls)


def _es_viga(elem):
    if elem is None:
        return False
    return _category_id_int(elem) == int(BuiltInCategory.OST_StructuralFraming)


def _shaft_category():
    for attr in (u"OST_ShaftOpening", u"OST_Opening"):
        try:
            return getattr(BuiltInCategory, attr)
        except Exception:
            continue
    return None


def _ft_to_mm(val):
    try:
        return float(
            UnitUtils.ConvertFromInternalUnits(float(val), UnitTypeId.Millimeters)
        )
    except Exception:
        return float(val) * _FT_TO_MM


def _wall_width_mm(wall):
    try:
        w = float(wall.Width)
        if w > 1e-9:
            return _ft_to_mm(w)
    except Exception:
        pass
    try:
        p = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        if p is not None and p.HasValue:
            return _ft_to_mm(p.AsDouble())
    except Exception:
        pass
    return 200.0


def _beam_width_mm(beam):
    names = (
        u"b",
        u"B",
        u"Width",
        u"Ancho",
        u"bw",
        u"BF",
        u"bf",
    )
    for src in (beam, getattr(beam, u"Symbol", None)):
        if src is None:
            continue
        for name in names:
            try:
                p = src.LookupParameter(name)
                if p is not None and p.HasValue:
                    try:
                        v = float(p.AsDouble())
                    except Exception:
                        continue
                    if v > 1e-9:
                        mm = _ft_to_mm(v)
                        if 50.0 <= mm <= 2500.0:
                            return mm
            except Exception:
                continue
    try:
        p = beam.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH)
        if p is not None and p.HasValue:
            mm = _ft_to_mm(p.AsDouble())
            if mm > 1e-9:
                return mm
    except Exception:
        pass
    # Bbox en planta: lado menor
    try:
        bb = beam.get_BoundingBox(None)
        if bb is not None:
            dx = abs(float(bb.Max.X) - float(bb.Min.X)) * _FT_TO_MM
            dy = abs(float(bb.Max.Y) - float(bb.Min.Y)) * _FT_TO_MM
            side = min(dx, dy)
            if 50.0 <= side <= 2500.0:
                return side
    except Exception:
        pass
    return _BEAM_WIDTH_FALLBACK_MM


def _location_curve(elem):
    try:
        loc = elem.Location
    except Exception:
        return None
    if not isinstance(loc, LocationCurve):
        return None
    try:
        return loc.Curve
    except Exception:
        return None


def _strip_polygon_mm(center_mm, width_mm):
    """Huella rectangular (o polilínea offset) desde eje en mm del plano."""
    if not center_mm or len(center_mm) < 2 or width_mm <= 1e-6:
        return None
    half = float(width_mm) * 0.5
    left = []
    right = []
    n = len(center_mm)
    for i in range(n):
        if i == 0:
            tx = center_mm[1][0] - center_mm[0][0]
            ty = center_mm[1][1] - center_mm[0][1]
        elif i == n - 1:
            tx = center_mm[-1][0] - center_mm[-2][0]
            ty = center_mm[-1][1] - center_mm[-2][1]
        else:
            tx = center_mm[i + 1][0] - center_mm[i - 1][0]
            ty = center_mm[i + 1][1] - center_mm[i - 1][1]
        ln = math.hypot(tx, ty)
        if ln < 1e-9:
            nx, ny = 0.0, 1.0
        else:
            nx = -ty / ln
            ny = tx / ln
        x, y = center_mm[i]
        left.append((x + nx * half, y + ny * half))
        right.append((x - nx * half, y - ny * half))
    return left + list(reversed(right))


def _xy_dist(p0, p1):
    return math.hypot(float(p1[0]) - float(p0[0]), float(p1[1]) - float(p0[1]))


def _line_line_hit(p1, v1, p2, v2):
    cross = float(v1[0]) * float(v2[1]) - float(v1[1]) * float(v2[0])
    if abs(cross) < 1e-12:
        return None
    dx = float(p2[0]) - float(p1[0])
    dy = float(p2[1]) - float(p1[1])
    t = (dx * float(v2[1]) - dy * float(v2[0])) / cross
    return (
        float(p1[0]) + t * float(v1[0]),
        float(p1[1]) + t * float(v1[1]),
    )


def _wall_normal_at(center_pts, idx, ref_nx, ref_ny):
    n = len(center_pts)
    if idx <= 0:
        tx = center_pts[1][0] - center_pts[0][0]
        ty = center_pts[1][1] - center_pts[0][1]
    elif idx >= n - 1:
        tx = center_pts[-1][0] - center_pts[-2][0]
        ty = center_pts[-1][1] - center_pts[-2][1]
    else:
        tx = center_pts[idx + 1][0] - center_pts[idx - 1][0]
        ty = center_pts[idx + 1][1] - center_pts[idx - 1][1]
    ln = math.hypot(tx, ty)
    if ln < 1e-12:
        return ref_nx, ref_ny
    nx = -ty / ln
    ny = tx / ln
    if nx * ref_nx + ny * ref_ny < 0.0:
        nx = -nx
        ny = -ny
    return nx, ny


def _wall_dir_into(center_pts, at_start):
    if at_start:
        tx = center_pts[1][0] - center_pts[0][0]
        ty = center_pts[1][1] - center_pts[0][1]
    else:
        tx = center_pts[-2][0] - center_pts[-1][0]
        ty = center_pts[-2][1] - center_pts[-1][1]
    ln = math.hypot(tx, ty)
    if ln < 1e-12:
        return (1.0, 0.0)
    return (tx / ln, ty / ln)


def _miter_offset(center_pts, idx, offset, ref_nx, ref_ny, side):
    n = len(center_pts)
    cx, cy = center_pts[idx]
    if offset <= 1e-12:
        return cx, cy
    sign = 1.0 if side == u"ext" else -1.0
    if idx <= 0 or idx >= n - 1:
        nx, ny = _wall_normal_at(center_pts, idx, ref_nx, ref_ny)
        return cx + sign * nx * offset, cy + sign * ny * offset
    tx_in = center_pts[idx][0] - center_pts[idx - 1][0]
    ty_in = center_pts[idx][1] - center_pts[idx - 1][1]
    tx_out = center_pts[idx + 1][0] - center_pts[idx][0]
    ty_out = center_pts[idx + 1][1] - center_pts[idx][1]
    ln_in = math.hypot(tx_in, ty_in)
    ln_out = math.hypot(tx_out, ty_out)
    if ln_in < 1e-12 or ln_out < 1e-12:
        nx, ny = _wall_normal_at(center_pts, idx, ref_nx, ref_ny)
        return cx + sign * nx * offset, cy + sign * ny * offset
    t_in = (tx_in / ln_in, ty_in / ln_in)
    t_out = (tx_out / ln_out, ty_out / ln_out)
    n_in = (-t_in[1], t_in[0])
    n_out = (-t_out[1], t_out[0])
    if n_in[0] * ref_nx + n_in[1] * ref_ny < 0.0:
        n_in = (-n_in[0], -n_in[1])
    if n_out[0] * ref_nx + n_out[1] * ref_ny < 0.0:
        n_out = (-n_out[0], -n_out[1])
    p_in = (cx + sign * n_in[0] * offset, cy + sign * n_in[1] * offset)
    p_out = (cx + sign * n_out[0] * offset, cy + sign * n_out[1] * offset)
    hit = _line_line_hit(p_in, t_in, p_out, (-t_out[0], -t_out[1]))
    if hit is None:
        nx, ny = _wall_normal_at(center_pts, idx, ref_nx, ref_ny)
        return cx + sign * nx * offset, cy + sign * ny * offset
    if math.hypot(hit[0] - cx, hit[1] - cy) > offset * 4.0:
        nx, ny = _wall_normal_at(center_pts, idx, ref_nx, ref_ny)
        return cx + sign * nx * offset, cy + sign * ny * offset
    return hit


def _wall_ref_normal_mm(wall, center_mm, plane):
    """Normal de referencia en plano (Orientation proyectada o ⟂ al eje)."""
    try:
        orient = wall.Orientation
        ox = (
            float(orient.X) * float(plane.XVec.X)
            + float(orient.Y) * float(plane.XVec.Y)
            + float(orient.Z) * float(plane.XVec.Z)
        )
        oy = (
            float(orient.X) * float(plane.YVec.X)
            + float(orient.Y) * float(plane.YVec.Y)
            + float(orient.Z) * float(plane.YVec.Z)
        )
        ln = math.hypot(ox, oy)
        if ln > 1e-9:
            return ox / ln, oy / ln
    except Exception:
        pass
    if len(center_mm) >= 2:
        tx = center_mm[1][0] - center_mm[0][0]
        ty = center_mm[1][1] - center_mm[0][1]
        ln = math.hypot(tx, ty)
        if ln > 1e-9:
            return -ty / ln, tx / ln
    return 0.0, 1.0


def _collect_axis_infos_mm(elements, plane, width_fn):
    """Ejes en planta (mm) para muros o vigas: LocationCurve + ancho."""
    infos = []
    for el in elements or []:
        curve = _location_curve(el)
        if curve is None:
            continue
        center = _curve_to_centerline_mm(curve, plane)
        if not center or len(center) < 2:
            continue
        try:
            width = float(width_fn(el))
        except Exception:
            width = 0.0
        if width <= 1e-6:
            continue
        half = width * 0.5
        # Normal ref: Orientation de muro si existe; si no, ⟂ al eje
        ref_nx, ref_ny = _wall_ref_normal_mm(el, center, plane)
        infos.append(
            {
                u"center_pts": center,
                u"off_ext": half,
                u"off_int": half,
                u"ref_nx": ref_nx,
                u"ref_ny": ref_ny,
                u"width": width,
                u"eid": _element_id_int(el.Id),
            }
        )
    return infos


def _collect_wall_axis_infos_mm(walls, plane):
    return _collect_axis_infos_mm(walls, plane, _wall_width_mm)


def _junction_tol_mm(axis_infos):
    # Holgado: ejes de muro a menudo no coinciden (LocationLine ≠ center).
    tol = 250.0  # mm
    for info in axis_infos or []:
        try:
            tol = max(tol, float(info.get(u"width", 0.0)) * 1.25 + 100.0)
        except Exception:
            pass
    return tol


def _spatial_cell_xy(x, y, cell):
    c = float(cell) if cell else 1.0
    if c < 1e-9:
        c = 1.0
    return int(math.floor(float(x) / c)), int(math.floor(float(y) / c))


def _spatial_cells_for_aabb(xmin, ymin, xmax, ymax, cell):
    ix0, iy0 = _spatial_cell_xy(xmin, ymin, cell)
    ix1, iy1 = _spatial_cell_xy(xmax, ymax, cell)
    if ix1 < ix0:
        ix0, ix1 = ix1, ix0
    if iy1 < iy0:
        iy0, iy1 = iy1, iy0
    for ix in range(ix0, ix1 + 1):
        for iy in range(iy0, iy1 + 1):
            yield ix, iy


def _cluster_wall_junctions_mm(axis_infos):
    tol = _junction_tol_mm(axis_infos)
    cell = max(float(tol), 100.0)
    clusters = []
    buckets = {}  # (ix,iy) -> [cluster_index]

    def _bucket_add(ix, iy, ci):
        key = (ix, iy)
        lst = buckets.get(key)
        if lst is None:
            buckets[key] = [ci]
        else:
            lst.append(ci)

    for wall_idx, info in enumerate(axis_infos):
        pts = info[u"center_pts"]
        for end_key, pt in ((u"start", pts[0]), (u"end", pts[-1])):
            matched = None
            ix0, iy0 = _spatial_cell_xy(pt[0], pt[1], cell)
            for dix in (-1, 0, 1):
                for diy in (-1, 0, 1):
                    for ci in buckets.get((ix0 + dix, iy0 + diy), ()):
                        cluster = clusters[ci]
                        if _xy_dist(cluster[u"pt"], pt) <= tol:
                            matched = cluster
                            break
                    if matched is not None:
                        break
                if matched is not None:
                    break
            if matched is None:
                matched = {u"pt": pt, u"members": []}
                ci = len(clusters)
                clusters.append(matched)
                _bucket_add(ix0, iy0, ci)
            matched[u"members"].append((wall_idx, end_key))
    return [c for c in clusters if len(c[u"members"]) >= 2]


def _junction_corner_pair_mm(cluster, axis_infos):
    members = cluster[u"members"]
    if len(members) != 2:
        return None
    (idx_a, end_a), (idx_b, end_b) = members[0], members[1]
    info_a = axis_infos[idx_a]
    info_b = axis_infos[idx_b]
    px, py = cluster[u"pt"]
    dir_a = _wall_dir_into(info_a[u"center_pts"], end_a == u"start")
    dir_b = _wall_dir_into(info_b[u"center_pts"], end_b == u"start")
    idx_a_pt = 0 if end_a == u"start" else len(info_a[u"center_pts"]) - 1
    idx_b_pt = 0 if end_b == u"start" else len(info_b[u"center_pts"]) - 1
    na = _wall_normal_at(
        info_a[u"center_pts"], idx_a_pt, info_a[u"ref_nx"], info_a[u"ref_ny"]
    )
    nb = _wall_normal_at(
        info_b[u"center_pts"], idx_b_pt, info_b[u"ref_nx"], info_b[u"ref_ny"]
    )
    p_ext_a = (px + na[0] * info_a[u"off_ext"], py + na[1] * info_a[u"off_ext"])
    p_ext_b = (px + nb[0] * info_b[u"off_ext"], py + nb[1] * info_b[u"off_ext"])
    p_int_a = (px - na[0] * info_a[u"off_int"], py - na[1] * info_a[u"off_int"])
    p_int_b = (px - nb[0] * info_b[u"off_int"], py - nb[1] * info_b[u"off_int"])
    outer = _line_line_hit(p_ext_a, dir_a, p_ext_b, dir_b)
    inner = _line_line_hit(p_int_a, dir_a, p_int_b, dir_b)
    if outer is None or inner is None:
        return None
    return {u"outer": outer, u"inner": inner}


def _junction_corner_lookup_mm(axis_infos):
    lookup = {}
    for cluster in _cluster_wall_junctions_mm(axis_infos):
        corners = _junction_corner_pair_mm(cluster, axis_infos)
        if corners is None:
            continue
        for wall_idx, end_key in cluster[u"members"]:
            lookup[(wall_idx, end_key)] = corners
    return lookup


def _wall_footprint_one_mm(info, wall_idx, junction_lookup):
    center_pts = info[u"center_pts"]
    if len(center_pts) < 2:
        return None
    off_ext = info[u"off_ext"]
    off_int = info[u"off_int"]
    ref_nx = info[u"ref_nx"]
    ref_ny = info[u"ref_ny"]
    n = len(center_pts)
    start_corner = junction_lookup.get((wall_idx, u"start"))
    end_corner = junction_lookup.get((wall_idx, u"end"))
    ext_side = []
    int_side = []
    for i in range(n):
        if i == 0 and start_corner is not None:
            ext_side.append(start_corner[u"outer"])
            int_side.append(start_corner[u"inner"])
            continue
        if i == n - 1 and end_corner is not None:
            ext_side.append(end_corner[u"outer"])
            int_side.append(end_corner[u"inner"])
            continue
        ext_side.append(
            _miter_offset(center_pts, i, off_ext, ref_nx, ref_ny, u"ext")
        )
        int_side.append(
            _miter_offset(center_pts, i, off_int, ref_nx, ref_ny, u"int")
        )
    poly = list(ext_side) + list(reversed(int_side))
    if len(poly) < 3:
        return None
    return poly


def _junction_partner_mm(wall_idx, end_key, clusters):
    for cluster in clusters or []:
        members = cluster.get(u"members") or []
        if (wall_idx, end_key) not in members:
            continue
        for member in members:
            if member != (wall_idx, end_key):
                return member
    return None


def _oriented_side_chains_mm(info, wall_idx, from_start, junction_lookup):
    center_pts = info[u"center_pts"]
    n = len(center_pts)
    off_ext = info[u"off_ext"]
    off_int = info[u"off_int"]
    ref_nx = info[u"ref_nx"]
    ref_ny = info[u"ref_ny"]
    start_corner = junction_lookup.get((wall_idx, u"start"))
    end_corner = junction_lookup.get((wall_idx, u"end"))
    indices = list(range(n)) if from_start else list(range(n - 1, -1, -1))
    ext_chain = []
    int_chain = []
    for i in indices:
        if i == 0 and start_corner is not None:
            ext_chain.append(start_corner[u"outer"])
            int_chain.append(start_corner[u"inner"])
            continue
        if i == n - 1 and end_corner is not None:
            ext_chain.append(end_corner[u"outer"])
            int_chain.append(end_corner[u"inner"])
            continue
        ext_chain.append(
            _miter_offset(center_pts, i, off_ext, ref_nx, ref_ny, u"ext")
        )
        int_chain.append(
            _miter_offset(center_pts, i, off_int, ref_nx, ref_ny, u"int")
        )
    return ext_chain, int_chain


def _append_chain(path, chain, tol):
    if not chain:
        return path
    if not path:
        return list(chain)
    if _xy_dist(path[-1], chain[0]) <= tol:
        return path + list(chain[1:])
    return path + list(chain)


def _walk_wall_path_mm(indices, clusters):
    free_ends = []
    for wall_idx in indices:
        for end_key in (u"start", u"end"):
            if _junction_partner_mm(wall_idx, end_key, clusters) is None:
                free_ends.append((wall_idx, end_key))
    if len(free_ends) != 2:
        return None
    path = []
    visited = set()
    cur_wall, cur_end = free_ends[0]
    while True:
        from_start = cur_end == u"start"
        if cur_wall in visited:
            break
        visited.add(cur_wall)
        path.append((cur_wall, from_start))
        other_end = u"end" if cur_end == u"start" else u"start"
        partner = _junction_partner_mm(cur_wall, other_end, clusters)
        if partner is None:
            break
        p_wall, p_end = partner
        if p_wall in visited:
            break
        cur_wall = p_wall
        cur_end = p_end
        if len(path) > max(len(indices) * 2, 8):
            break
    if len(visited) != len(indices):
        return None
    return path


def _merge_wall_path_mm(path, axis_infos, junction_lookup, tol):
    if not path:
        return None
    if len(path) == 1:
        wall_idx, _fs = path[0]
        return _wall_footprint_one_mm(axis_infos[wall_idx], wall_idx, junction_lookup)
    ext_path = []
    int_chains = []
    for wall_idx, from_start in path:
        ext_chain, int_chain = _oriented_side_chains_mm(
            axis_infos[wall_idx], wall_idx, from_start, junction_lookup
        )
        ext_path = _append_chain(ext_path, ext_chain, tol)
        int_chains.append(int_chain)
    int_back = []
    for int_chain in reversed(int_chains):
        int_back = _append_chain(int_back, list(reversed(int_chain)), tol)
    poly = _append_chain(ext_path, int_back, tol)
    if len(poly) < 3:
        return None
    return poly


def _wall_union_components_mm(indices, clusters):
    parent = {i: i for i in indices}

    def _find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def _union(a, b):
        ra = _find(a)
        rb = _find(b)
        if ra != rb:
            parent[rb] = ra

    for cluster in clusters or []:
        members = [m[0] for m in cluster.get(u"members") or []]
        for j in range(1, len(members)):
            if members[0] in parent and members[j] in parent:
                _union(members[0], members[j])
    groups = {}
    for idx in indices:
        groups.setdefault(_find(idx), []).append(idx)
    return list(groups.values())


def _dist_point_to_segment(pt, a, b):
    ax, ay = float(a[0]), float(a[1])
    bx, by = float(b[0]), float(b[1])
    px, py = float(pt[0]), float(pt[1])
    abx, aby = bx - ax, by - ay
    ab2 = abx * abx + aby * aby
    if ab2 < 1e-18:
        return _xy_dist(pt, a), a
    t = ((px - ax) * abx + (py - ay) * aby) / ab2
    if t < 0.0:
        t = 0.0
    elif t > 1.0:
        t = 1.0
    cx, cy = ax + t * abx, ay + t * aby
    return math.hypot(px - cx, py - cy), (cx, cy)


def _extend_centerlines_at_encounters(axis_infos):
    """
    En encuentros Revit el LocationCurve suele cortar en la cara del muro
    vecino: la huella rectangular no cubre el cuadrante de la esquina exterior
    (el «cuadrito» vacío). Estira el extremo del eje hacia afuera del encuentro
    en ``ancho_vecino`` mm para que el strip cubra esa esquina antes del Union.

    Usa hash espacial (celdas) para evitar O(n²) puro con muchas piezas.
    """
    n = len(axis_infos or [])
    if n < 2:
        return
    for info in axis_infos:
        info[u"center_pts"] = [tuple(p) for p in info[u"center_pts"]]

    max_r = 400.0
    for info in axis_infos:
        try:
            max_r = max(max_r, float(info.get(u"width", 0.0) or 0.0) * 1.35 + 80.0)
        except Exception:
            pass
    cell = max(max_r * 0.55, 200.0)

    # buckets[(ix,iy)] -> list of (j, "end"|"seg", payload)
    buckets = {}

    def _bucket_add(ix, iy, item):
        key = (ix, iy)
        lst = buckets.get(key)
        if lst is None:
            buckets[key] = [item]
        else:
            lst.append(item)

    for j, info_j in enumerate(axis_infos):
        pts_j = info_j[u"center_pts"]
        if len(pts_j) < 2:
            continue
        for ep in (pts_j[0], pts_j[-1]):
            ix, iy = _spatial_cell_xy(ep[0], ep[1], cell)
            _bucket_add(ix, iy, (j, u"end", ep))
        for k in range(len(pts_j) - 1):
            a = pts_j[k]
            b = pts_j[k + 1]
            xmin = min(a[0], b[0])
            xmax = max(a[0], b[0])
            ymin = min(a[1], b[1])
            ymax = max(a[1], b[1])
            for ix, iy in _spatial_cells_for_aabb(xmin, ymin, xmax, ymax, cell):
                _bucket_add(ix, iy, (j, u"seg", (a, b)))

    for i in range(n):
        info_i = axis_infos[i]
        pts = info_i[u"center_pts"]
        if len(pts) < 2:
            continue
        wi = float(info_i.get(u"width", 0.0) or 0.0)
        for end_key in (u"start", u"end"):
            end_pt = pts[0] if end_key == u"start" else pts[-1]
            best_j = None
            best_d = 1e18
            # Radio de consulta holgado (cubre lim_end típico)
            q_r = max(wi * 1.35 + 80.0, max_r)
            ix0, iy0 = _spatial_cell_xy(end_pt[0], end_pt[1], cell)
            n_cells = int(math.ceil(q_r / cell)) + 1
            seen_end = set()
            seen_seg = set()
            for dix in range(-n_cells, n_cells + 1):
                for diy in range(-n_cells, n_cells + 1):
                    for item in buckets.get((ix0 + dix, iy0 + diy), ()):
                        j, kind, payload = item
                        if j == i:
                            continue
                        info_j = axis_infos[j]
                        wj = float(info_j.get(u"width", 0.0) or 0.0)
                        if kind == u"end":
                            if (j, payload) in seen_end:
                                continue
                            seen_end.add((j, payload))
                            lim_end = max(wi, wj) * 1.35 + 80.0
                            d = _xy_dist(end_pt, payload)
                            if d <= lim_end and d < best_d:
                                best_d = d
                                best_j = j
                        else:
                            a, b = payload
                            sk = (j, a, b)
                            if sk in seen_seg:
                                continue
                            seen_seg.add(sk)
                            lim_seg = max(wi, wj) * 0.85 + 40.0
                            d, _cl = _dist_point_to_segment(end_pt, a, b)
                            if d <= lim_seg and d < best_d:
                                best_d = d
                                best_j = j
            if best_j is None:
                continue
            # Hacia afuera del muro (opuesto a «entrar» al tramo)
            din = _wall_dir_into(pts, end_key == u"start")
            dout = (-float(din[0]), -float(din[1]))
            # Solo ½ ancho del vecino: rellena el cuadrante exterior sin
            # sobresalir de la cara lejana (antes: ancho completo → «nub»).
            dist = float(axis_infos[best_j].get(u"width", 0.0) or 0.0) * 0.5
            if dist < 1.0:
                dist = max(wi * 0.25, 40.0)
            new_pt = (end_pt[0] + dout[0] * dist, end_pt[1] + dout[1] * dist)
            if end_key == u"start":
                pts[0] = new_pt
            else:
                pts[-1] = new_pt


def _linear_monolithic_footprints_mm(elements, plane, width_fn):
    """
    Contornos monolíticos (muros o vigas): estira ½ ancho en encuentros,
    miter/merge de cadenas. Misma regla para uniones entre vigas.
    """
    axis_infos = _collect_axis_infos_mm(elements, plane, width_fn)
    if not axis_infos:
        return []
    _extend_centerlines_at_encounters(axis_infos)
    clusters = _cluster_wall_junctions_mm(axis_infos)
    junction_lookup = _junction_corner_lookup_mm(axis_infos)
    tol = _junction_tol_mm(axis_infos)
    all_indices = list(range(len(axis_infos)))
    polygons = []
    merged = set()
    for component in _wall_union_components_mm(all_indices, clusters):
        path = _walk_wall_path_mm(component, clusters)
        if path is not None and len(component) >= 2:
            poly = _merge_wall_path_mm(path, axis_infos, junction_lookup, tol)
            if poly is not None:
                eids = [
                    axis_infos[wi][u"eid"]
                    for wi, _ in path
                    if axis_infos[wi].get(u"eid") is not None
                ]
                polygons.append({u"pts": poly, u"eids": eids})
                for wi, _ in path:
                    merged.add(wi)
                continue
        for wall_idx in component:
            if wall_idx in merged:
                continue
            poly = _wall_footprint_one_mm(
                axis_infos[wall_idx], wall_idx, junction_lookup
            )
            if poly is not None:
                eid = axis_infos[wall_idx].get(u"eid")
                polygons.append(
                    {u"pts": poly, u"eids": [eid] if eid is not None else []}
                )
                merged.add(wall_idx)
    for wall_idx in all_indices:
        if wall_idx in merged:
            continue
        poly = _wall_footprint_one_mm(
            axis_infos[wall_idx], wall_idx, junction_lookup
        )
        if poly is not None:
            eid = axis_infos[wall_idx].get(u"eid")
            polygons.append(
                {u"pts": poly, u"eids": [eid] if eid is not None else []}
            )
    return polygons


def _wall_monolithic_footprints_mm(walls, plane):
    return _linear_monolithic_footprints_mm(walls, plane, _wall_width_mm)


def _beam_monolithic_footprints_mm(beams, plane):
    """
    Huella de viga: franja del LocationCurve + estirado ½ ancho solo entre
    vigas (rellena el cuadrito exterior en L/T). Sin miter/merge de muros
    (ese camino torcía ejes y dejaba cuñas en el canvas).
    """
    axis_infos = _collect_axis_infos_mm(beams, plane, _beam_width_mm)
    if not axis_infos:
        return []
    # Solo entre vigas: cubre el cuadrante exterior del encuentro.
    _extend_centerlines_at_encounters(axis_infos)
    polygons = []
    for info in axis_infos:
        center = info.get(u"center_pts") or []
        try:
            width = float(info.get(u"width", 0.0) or 0.0)
        except Exception:
            width = 0.0
        if width <= 1e-6:
            continue
        poly = _strip_polygon_mm(center, width)
        if not poly or len(poly) < 3:
            continue
        eid = info.get(u"eid")
        polygons.append(
            {u"pts": poly, u"eids": [eid] if eid is not None else []}
        )
    return polygons


def _curve_to_centerline_mm(curve, plane):
    return _sample_curve_mm(curve, plane, n_arc=16)


def _elements_intersecting_floor_bbox(document, floor, category):
    """Collector por categoría que intersecta bbox de la losa (ampliado)."""
    out = []
    if document is None or floor is None or category is None:
        return out
    try:
        bb = floor.get_BoundingBox(None)
        if bb is None:
            return out
        mn = bb.Min
        mx = bb.Max
        pad = 0.5  # ft
        outline = Outline(
            XYZ(mn.X - pad, mn.Y - pad, mn.Z - pad),
            XYZ(mx.X + pad, mx.Y + pad, mx.Z + pad),
        )
        filt = BoundingBoxIntersectsFilter(outline)
        for el in (
            FilteredElementCollector(document)
            .OfCategory(category)
            .WherePasses(filt)
            .WhereElementIsNotElementType()
        ):
            if el is not None:
                out.append(el)
    except Exception:
        pass
    return out


def _joined_context_elements(document, floor):
    walls, beams = [], []
    ids = []
    if get_joined_element_ids is not None:
        try:
            ids = list(get_joined_element_ids(document, floor) or [])
        except Exception:
            ids = []
    if not ids:
        try:
            from Autodesk.Revit.DB import JoinGeometryUtils

            raw = JoinGeometryUtils.GetJoinedElements(document, floor)
            if raw is not None:
                for eid in raw:
                    if eid is not None and eid != ElementId.InvalidElementId:
                        ids.append(eid)
        except Exception:
            pass
    seen = set()
    for eid in ids:
        ni = _element_id_int(eid)
        if ni is not None and ni in seen:
            continue
        if ni is not None:
            seen.add(ni)
        el = document.GetElement(eid)
        if _es_muro(el):
            walls.append(el)
        elif _es_viga(el):
            beams.append(el)
    return walls, beams


def _curve_has_graphics_style(curve):
    """True si la curva tiene estilo gráfico (p. ej. línea simbólica de cruz)."""
    try:
        gsid = curve.GraphicsStyleId
        if gsid is not None and gsid != ElementId.InvalidElementId:
            return True
    except Exception:
        pass
    return False


def _filter_out_shaft_cross_curves(curves):
    """
    Excluye diagonales tipo cruz simbólica del shaft:
    - curvas con GraphicsStyleId
    - segmentos que unen esquinas opuestas del bbox XY
    """
    if not curves:
        return []
    segs = []
    for c in curves:
        if c is None:
            continue
        try:
            if not c.IsBound:
                continue
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
        except Exception:
            continue
        segs.append((c, p0, p1))
    if not segs:
        return []

    xs = [float(p.X) for _, p0, p1 in segs for p in (p0, p1)]
    ys = [float(p.Y) for _, p0, p1 in segs for p in (p0, p1)]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    span = max(abs(max_x - min_x), abs(max_y - min_y), 1e-9)
    tol = max(span * 0.06, 1e-4)
    corners = (
        (min_x, min_y),
        (max_x, min_y),
        (max_x, max_y),
        (min_x, max_y),
    )

    def _near(p, corner):
        return (
            abs(float(p.X) - corner[0]) <= tol
            and abs(float(p.Y) - corner[1]) <= tol
        )

    out = []
    for c, p0, p1 in segs:
        if _curve_has_graphics_style(c):
            continue
        is_diag = False
        for a, b in ((0, 2), (1, 3)):
            if (
                (_near(p0, corners[a]) and _near(p1, corners[b]))
                or (_near(p0, corners[b]) and _near(p1, corners[a]))
            ):
                is_diag = True
                break
        if is_diag:
            continue
        out.append(c)
    return out if out else [c for c, _, _ in segs]


def _curvas_pasada_opening(opening):
    """
    Contorno de Shaft Opening / Opening (sin líneas simbólicas de cruz).

    Preferir ``Sketch.Profile`` (solo loops de boundary); si no,
    BoundaryRect / BoundaryCurves filtrando diagonales simbólicas.
    """
    # 1) Sketch.Profile — no incluye líneas simbólicas (cruz X)
    try:
        doc = opening.Document
        sid = getattr(opening, "SketchId", None)
        if (
            doc is not None
            and sid is not None
            and sid != ElementId.InvalidElementId
        ):
            sketch = doc.GetElement(sid)
            if sketch is not None and isinstance(sketch, Sketch):
                profile = sketch.Profile
                if profile is not None and int(profile.Size) >= 1:
                    # Un loop cerrado por pasada (Profile no trae la cruz simbólica)
                    for i in range(int(profile.Size)):
                        curve_array = profile.get_Item(i)
                        if curve_array is None:
                            continue
                        curvas = []
                        for j in range(int(curve_array.Size)):
                            c = curve_array.get_Item(j)
                            if c is not None and c.IsBound:
                                try:
                                    curvas.append(c.Clone())
                                except Exception:
                                    curvas.append(c)
                        if len(curvas) >= 3:
                            ordered = _order_curves_closed(curvas)
                            return ordered if ordered else curvas
    except Exception:
        pass

    # 2) Rectángulo explícito
    curvas = []
    try:
        if getattr(opening, "IsRectBoundary", False):
            rect = getattr(opening, "BoundaryRect", None)
            if rect is not None:
                mn = getattr(rect, "Min", None) or getattr(rect, "Minimum", None)
                mx = getattr(rect, "Max", None) or getattr(rect, "Maximum", None)
                if mn is not None and mx is not None:
                    z = float(mn.Z)
                    pts = [
                        XYZ(mn.X, mn.Y, z),
                        XYZ(mx.X, mn.Y, z),
                        XYZ(mx.X, mx.Y, z),
                        XYZ(mn.X, mx.Y, z),
                    ]
                    for i in range(4):
                        c = Line.CreateBound(pts[i], pts[(i + 1) % 4])
                        if c is not None:
                            curvas.append(c)
                    return curvas
    except Exception:
        pass

    # 3) BoundaryCurves sin cruz simbólica
    try:
        boundary = getattr(opening, "BoundaryCurves", None)
        if boundary is not None:
            raw = []
            for c in boundary:
                if c is not None and c.IsBound:
                    raw.append(c)
            curvas = _filter_out_shaft_cross_curves(raw)
    except Exception:
        pass
    if not curvas:
        return []
    ordered = _order_curves_closed(curvas)
    return ordered if ordered else curvas


def _hosted_floor_openings(document, floor):
    """Opening cuyo Host es la losa (filtrado por bbox de la losa)."""
    out = []
    if document is None or floor is None:
        return out
    fid = floor.Id
    try:
        bb = floor.get_BoundingBox(None)
        col = FilteredElementCollector(document).OfClass(Opening)
        if bb is not None:
            pad = 0.5  # ft
            mn, mx = bb.Min, bb.Max
            outline = Outline(
                XYZ(mn.X - pad, mn.Y - pad, mn.Z - pad),
                XYZ(mx.X + pad, mx.Y + pad, mx.Z + pad),
            )
            col = col.WherePasses(BoundingBoxIntersectsFilter(outline))
        for op in col:
            try:
                host = op.Host
                if host is not None and host.Id == fid:
                    out.append(op)
            except Exception:
                continue
    except Exception:
        pass
    return out


def _floor_plane_bbox_mm(floor, plane, pad_mm=0.0):
    """BBox de la losa proyectada al plano Sketch, con margen en mm."""
    if floor is None or plane is None:
        return None
    try:
        bb = floor.get_BoundingBox(None)
        if bb is None:
            return None
        mn = bb.Min
        mx = bb.Max
        corners = []
        for x in (float(mn.X), float(mx.X)):
            for y in (float(mn.Y), float(mx.Y)):
                for z in (float(mn.Z), float(mx.Z)):
                    corners.append(XYZ(x, y, z))
        pts = [_xyz_to_plane_mm(c, plane) for c in corners]
        if not pts:
            return None
        min_x, min_y, max_x, max_y = _bbox_mm(pts)
        pad = float(pad_mm or 0.0)
        return (min_x - pad, min_y - pad, max_x + pad, max_y + pad)
    except Exception:
        return None


def _clip_segment_bbox_mm(p0, p1, bbox):
    """Liang-Barsky: recorta segmento al bbox (xmin,ymin,xmax,ymax)."""
    if p0 is None or p1 is None or bbox is None:
        return None
    xmin, ymin, xmax, ymax = bbox
    x0, y0 = float(p0[0]), float(p0[1])
    x1, y1 = float(p1[0]), float(p1[1])
    dx = x1 - x0
    dy = y1 - y0
    u1, u2 = 0.0, 1.0
    pq = (
        (-dx, x0 - xmin),
        (dx, xmax - x0),
        (-dy, y0 - ymin),
        (dy, ymax - y0),
    )
    for p, q in pq:
        if abs(p) < 1e-15:
            if q < 0.0:
                return None
            continue
        t = q / p
        if p < 0.0:
            if t > u2:
                return None
            if t > u1:
                u1 = t
        else:
            if t < u1:
                return None
            if t < u2:
                u2 = t
    if u1 > u2:
        return None
    return (
        (x0 + u1 * dx, y0 + u1 * dy),
        (x0 + u2 * dx, y0 + u2 * dy),
    )


def _clip_polyline_to_bbox_mm(pts, bbox):
    """Recorta polilínea abierta al bbox; devuelve piezas con ≥2 puntos."""
    if not pts or len(pts) < 2 or bbox is None:
        return []
    pieces = []
    current = []
    for i in range(len(pts) - 1):
        clipped = _clip_segment_bbox_mm(pts[i], pts[i + 1], bbox)
        if clipped is None:
            if len(current) >= 2:
                pieces.append(current)
            current = []
            continue
        a, b = clipped
        if not current:
            current = [a, b]
            continue
        lx, ly = current[-1]
        if abs(lx - a[0]) < 0.5 and abs(ly - a[1]) < 0.5:
            current.append(b)
        else:
            if len(current) >= 2:
                pieces.append(current)
            current = [a, b]
    if len(current) >= 2:
        pieces.append(current)
    return pieces


def _grid_name(grid):
    try:
        name = grid.Name
        if name:
            return _as_unicode(name)
    except Exception:
        pass
    try:
        p = grid.get_Parameter(BuiltInParameter.DATUM_TEXT)
        if p is not None:
            return _as_unicode(p.AsString() or u"")
    except Exception:
        pass
    return u""


def _collect_grid_overlays_mm(document, floor, plane, pad_mm=None):
    """
    Grids del proyecto cuya curva proyectada cruza el entorno de la losa.
    Segmentos recortados al bbox ampliado (evita zoom a toda la ciudad).
    """
    overlays = []
    if document is None or floor is None or plane is None:
        return overlays
    if pad_mm is None:
        pad_mm = _GRID_CLIP_PAD_MM
    bbox = _floor_plane_bbox_mm(floor, plane, pad_mm)
    if bbox is None:
        return overlays

    grids = []
    try:
        for g in FilteredElementCollector(document).OfClass(Grid):
            if g is not None:
                grids.append(g)
    except Exception:
        try:
            for g in (
                FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_Grids)
                .WhereElementIsNotElementType()
            ):
                if g is not None:
                    grids.append(g)
        except Exception:
            return overlays

    seen = set()
    for grid in grids:
        ni = _element_id_int(grid.Id)
        if ni is not None and ni in seen:
            continue
        if ni is not None:
            seen.add(ni)
        curve = None
        try:
            curve = grid.Curve
        except Exception:
            curve = None
        if curve is None:
            continue
        raw = _sample_curve_mm(curve, plane)
        if not raw or len(raw) < 2:
            continue
        pieces = _clip_polyline_to_bbox_mm(raw, bbox)
        if not pieces:
            continue
        label = _grid_name(grid) or u"Eje"
        for pts in pieces:
            overlays.append(
                {
                    u"kind": _CTX_GRID,
                    u"pts": pts,
                    u"eid": ni,
                    u"label": label,
                    u"closed": False,
                }
            )
    return overlays


def recolectar_contexto_planta(document, floor, plane, include_grids=True):
    """
    Muros, vigas, pasadas y (opcional) ejes (Grid) en mm del plano del Sketch.
    Returns: (overlays, walls_list, beams_list)

    ``include_grids=False``: omite ejes (solo snap; no se pintan). Útil para
    diferirlos tras el primer paint de muros/vigas.
    """
    overlays = []
    if document is None or floor is None or plane is None:
        return overlays, [], []

    walls_j, beams_j = _joined_context_elements(document, floor)
    walls_b = _elements_intersecting_floor_bbox(
        document, floor, BuiltInCategory.OST_Walls
    )
    beams_b = _elements_intersecting_floor_bbox(
        document, floor, BuiltInCategory.OST_StructuralFraming
    )

    walls_by_id = {}
    for w in walls_j + walls_b:
        ni = _element_id_int(w.Id)
        if ni is not None:
            walls_by_id[ni] = w
    beams_by_id = {}
    for b in beams_j + beams_b:
        ni = _element_id_int(b.Id)
        if ni is not None:
            beams_by_id[ni] = b

    walls_list = list(walls_by_id.values())
    for mono in _wall_monolithic_footprints_mm(walls_list, plane):
        pts = mono.get(u"pts")
        if not pts or len(pts) < 3:
            continue
        eids = mono.get(u"eids") or []
        overlays.append(
            {
                u"kind": _CTX_WALL,
                u"pts": pts,
                u"eid": eids[0] if eids else None,
                u"eids": eids,
                u"label": u"Muro",
            }
        )

    beams_list = list(beams_by_id.values())
    for mono in _beam_monolithic_footprints_mm(beams_list, plane):
        pts = mono.get(u"pts")
        if not pts or len(pts) < 3:
            continue
        eids = mono.get(u"eids") or []
        overlays.append(
            {
                u"kind": _CTX_BEAM,
                u"pts": pts,
                u"eid": eids[0] if eids else None,
                u"eids": eids,
                u"label": u"Viga",
            }
        )

    # Pasadas: shafts por bbox + openings hospedados
    pasadas = []
    shaft_cat = _shaft_category()
    if shaft_cat is not None:
        pasadas.extend(_elements_intersecting_floor_bbox(document, floor, shaft_cat))
    pasadas.extend(_hosted_floor_openings(document, floor))
    seen_p = set()
    for op in pasadas:
        ni = _element_id_int(op.Id)
        if ni is not None and ni in seen_p:
            continue
        if ni is not None:
            seen_p.add(ni)
        curvas = _curvas_pasada_opening(op)
        if not curvas:
            continue
        pts = _ring_closed_pts(_loop_to_polyline_mm(curvas, plane))
        if pts and len(pts) >= 3:
            overlays.append(
                {
                    u"kind": _CTX_PASADA,
                    u"pts": pts,
                    u"eid": ni,
                    u"label": u"Pasada",
                }
            )

    if include_grids:
        overlays.extend(_collect_grid_overlays_mm(document, floor, plane))

    # Huecos del Sketch (loops interiores) como pasadas de perfil
    return overlays, walls_list, beams_list


def _plane_mm_to_xyz(pt_mm, plane):
    x_ft = float(pt_mm[0]) / _FT_TO_MM
    y_ft = float(pt_mm[1]) / _FT_TO_MM
    o = plane.Origin
    xv = plane.XVec
    yv = plane.YVec
    return XYZ(
        float(o.X) + x_ft * float(xv.X) + y_ft * float(yv.X),
        float(o.Y) + x_ft * float(xv.Y) + y_ft * float(yv.Y),
        float(o.Z) + x_ft * float(xv.Z) + y_ft * float(yv.Z),
    )


def _plane_mm_dir_to_xyz(dx_mm, dy_mm, plane):
    """Convierte dirección unitaria (dx, dy) en plano-mm a XYZ Revit."""
    if plane is None:
        return XYZ(1, 0, 0)
    xv = plane.XVec
    yv = plane.YVec
    x = float(dx_mm) * float(xv.X) + float(dy_mm) * float(yv.X)
    y = float(dx_mm) * float(xv.Y) + float(dy_mm) * float(yv.Y)
    z = float(dx_mm) * float(xv.Z) + float(dy_mm) * float(yv.Z)
    v = XYZ(x, y, z)
    try:
        return v.Normalize()
    except Exception:
        L = (x * x + y * y + z * z) ** 0.5
        if L > 1e-9:
            return XYZ(x / L, y / L, z / L)
        return XYZ(1, 0, 0)


def _poly_mm_to_curves(pts_mm, plane):
    """Lista de Line en el plano del Sketch a partir de polígono mm."""
    if not pts_mm or len(pts_mm) < 3 or plane is None:
        return None
    curves = []
    n = len(pts_mm)
    for i in range(n):
        p0 = _plane_mm_to_xyz(pts_mm[i], plane)
        p1 = _plane_mm_to_xyz(pts_mm[(i + 1) % n], plane)
        if p0.DistanceTo(p1) < 1e-9:
            continue
        try:
            ln = Line.CreateBound(p0, p1)
            if ln is not None and ln.IsBound:
                curves.append(ln)
        except Exception:
            continue
    if len(curves) < 3:
        return None
    return curves


def _count_ctx(overlays):
    n_w = n_b = n_p = 0
    for o in overlays or []:
        k = o.get(u"kind")
        if k == _CTX_WALL:
            eids = o.get(u"eids")
            if eids:
                n_w += len(eids)
            else:
                n_w += 1
        elif k == _CTX_BEAM:
            eids = o.get(u"eids")
            if eids:
                n_b += len(eids)
            else:
                n_b += 1
        elif k == _CTX_PASADA:
            n_p += 1
    return n_w, n_b, n_p


# ---------------------------------------------------------------------------
# AreaReinforcement existentes (detección read-only)
# ---------------------------------------------------------------------------


def _xyz_dist2(a, b):
    try:
        dx = float(a.X) - float(b.X)
        dy = float(a.Y) - float(b.Y)
        dz = float(a.Z) - float(b.Z)
        return dx * dx + dy * dy + dz * dz
    except Exception:
        return 1.0e30


def _order_curves_closed(curves, tol_ft=0.01):
    """Encadena curvas por extremos (tol en pies) → lista ordenada (posible reverse)."""
    if not curves:
        return []
    remaining = list(curves)
    ordered = [remaining.pop(0)]
    tol2 = float(tol_ft) * float(tol_ft)
    guard = 0
    while remaining and guard < 500:
        guard += 1
        last_end = ordered[-1].GetEndPoint(1)
        found_i = -1
        reverse = False
        for i, c in enumerate(remaining):
            try:
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
            except Exception:
                continue
            if _xyz_dist2(last_end, p0) <= tol2:
                found_i = i
                reverse = False
                break
            if _xyz_dist2(last_end, p1) <= tol2:
                found_i = i
                reverse = True
                break
        if found_i < 0:
            break
        c = remaining.pop(found_i)
        if reverse:
            try:
                c = c.CreateReversed()
            except Exception:
                pass
        ordered.append(c)
    return ordered


def _area_reinforcement_sketch(document, area_rein):
    """Sketch dependiente del AR, si existe."""
    if document is None or area_rein is None:
        return None
    try:
        from Autodesk.Revit.DB import ElementClassFilter

        dep = area_rein.GetDependentElements(
            ElementClassFilter(clr.GetClrType(Sketch))
        )
        if dep:
            for sid in dep:
                if sid is None or sid == ElementId.InvalidElementId:
                    continue
                sk = document.GetElement(sid)
                if sk is not None and isinstance(sk, Sketch):
                    return sk
    except Exception:
        pass
    try:
        dep_all = area_rein.GetDependentElements(None)
        if dep_all:
            for sid in dep_all:
                if sid is None or sid == ElementId.InvalidElementId:
                    continue
                sk = document.GetElement(sid)
                if sk is not None and isinstance(sk, Sketch):
                    return sk
    except Exception:
        pass
    try:
        ids = area_rein.GetBoundaryCurveIds()
    except Exception:
        ids = None
    if ids:
        for cid in ids:
            try:
                ce = document.GetElement(cid)
                sid = getattr(ce, "SketchId", None)
                if sid and sid != ElementId.InvalidElementId:
                    sk = document.GetElement(sid)
                    if sk is not None and isinstance(sk, Sketch):
                        return sk
            except Exception:
                continue
    return None


def _area_reinforcement_loops_mm(document, area_rein, plane):
    """
    Contornos del AR en mm sobre ``plane`` (mismo sistema que el Sketch de la losa).
    Devuelve lista de polilíneas (exterior + huecos si hay Sketch.Profile).
    """
    if document is None or area_rein is None or plane is None:
        return []
    loops_mm = []

    # Preferir Sketch.Profile (loops ordenados)
    sketch = _area_reinforcement_sketch(document, area_rein)
    if sketch is not None:
        try:
            profile = sketch.Profile
            if profile is not None:
                n_loops = int(profile.Size)
                for i in range(n_loops):
                    curve_array = profile.get_Item(i)
                    if curve_array is None:
                        continue
                    curves = []
                    n_curves = int(curve_array.Size)
                    for j in range(n_curves):
                        c = curve_array.get_Item(j)
                        if c is not None and c.IsBound:
                            curves.append(c)
                    if curves:
                        pts = _loop_to_polyline_mm(curves, plane)
                        ring = _ring_closed_pts(pts)
                        if len(ring) >= 3:
                            loops_mm.append(ring)
        except Exception:
            loops_mm = []
        if loops_mm:
            return loops_mm

    # Fallback: CurveElements del boundary
    curves = []
    try:
        ids = area_rein.GetBoundaryCurveIds()
    except Exception:
        ids = None
    if ids:
        for cid in ids:
            try:
                ce = document.GetElement(cid)
                if ce is None:
                    continue
                gc = getattr(ce, "GeometryCurve", None)
                if gc is not None and gc.IsBound:
                    curves.append(gc)
            except Exception:
                continue
    if not curves:
        return []
    ordered = _order_curves_closed(curves)
    if not ordered:
        ordered = curves
    pts = _loop_to_polyline_mm(ordered, plane)
    ring = _ring_closed_pts(pts)
    if len(ring) >= 3:
        return [ring]
    return []


def collect_existing_area_rein_on_floor(document, floor, plane):
    """
    AreaReinforcement cuyo host es ``floor`` (GetHostId).
    Read-only: lista de dicts ``{id, label, pts, loops}``.
    ``pts`` = loop exterior; ``loops`` = todos los anillos.
    """
    out = []
    if document is None or floor is None or plane is None:
        return out
    floor_id = _element_id_int(floor.Id)
    if floor_id is None:
        return out
    try:
        collector = (
            FilteredElementCollector(document)
            .OfClass(AreaReinforcement)
            .WhereElementIsNotElementType()
        )
    except Exception:
        return out
    for ar in collector:
        if ar is None or not isinstance(ar, AreaReinforcement):
            continue
        try:
            hid = ar.GetHostId()
        except Exception:
            hid = None
        if _element_id_int(hid) != floor_id:
            continue
        ar_id = _element_id_int(ar.Id)
        if ar_id is None:
            continue
        loops = _area_reinforcement_loops_mm(document, ar, plane)
        pts = loops[0] if loops else []
        if len(pts) < 3:
            # Sin geometría dibujable: aún listar en panel
            out.append(
                {
                    u"id": ar_id,
                    u"label": u"AR Id {0}".format(ar_id),
                    u"pts": [],
                    u"loops": [],
                }
            )
            continue
        out.append(
            {
                u"id": ar_id,
                u"label": u"AR Id {0}".format(ar_id),
                u"pts": pts,
                u"loops": loops,
            }
        )
    out.sort(key=lambda d: int(d.get(u"id") or 0))
    return out


# ---------------------------------------------------------------------------
# Rebar types / Create
# ---------------------------------------------------------------------------


def _bar_types_sorted(document):
    out = []
    try:
        rts = list(FilteredElementCollector(document).OfClass(RebarBarType))
    except Exception:
        return out
    for bt in rts:
        try:
            diam_ft = getattr(bt, "BarNominalDiameter", None)
            dmm = 0
            if diam_ft is not None:
                dmm = int(round(float(diam_ft) * _FT_TO_MM))
            label = u"Ø{} mm".format(dmm) if dmm > 0 else _as_unicode(bt.Name)
            out.append((dmm, label, bt))
        except Exception:
            continue
    out.sort(key=lambda x: x[0])
    return out



def _default_area_type_id(document):
    try:
        eid = document.GetDefaultElementTypeId(ElementTypeGroup.AreaReinforcementType)
        if eid and eid != ElementId.InvalidElementId:
            return eid
    except Exception:
        pass
    try:
        for el in FilteredElementCollector(document).OfClass(AreaReinforcementType):
            return el.Id
    except Exception:
        pass
    return ElementId.InvalidElementId


def _mm_to_internal(mm):
    try:
        return UnitUtils.ConvertToInternalUnits(float(mm), UnitTypeId.Millimeters)
    except Exception:
        return float(mm) / _FT_TO_MM


# BuiltInParameter (multiidioma) + LookupParameter (inglés) por capa Floor.
# DIR_1 = Major, DIR_2 = Minor; Top = exterior/superior; Bottom = interior/inferior.
_LAYER_BIP_SPACING = {
    u"exterior_major": (
        u"REBAR_SYSTEM_SPACING_TOP_DIR_1",
        u"REBAR_SYSTEM_SPACING_TOP_DIR_1_GENERIC",
    ),
    u"exterior_minor": (
        u"REBAR_SYSTEM_SPACING_TOP_DIR_2",
        u"REBAR_SYSTEM_SPACING_TOP_DIR_2_GENERIC",
    ),
    u"interior_major": (
        u"REBAR_SYSTEM_SPACING_BOTTOM_DIR_1",
        u"REBAR_SYSTEM_SPACING_BOTTOM_DIR_1_GENERIC",
    ),
    u"interior_minor": (
        u"REBAR_SYSTEM_SPACING_BOTTOM_DIR_2",
        u"REBAR_SYSTEM_SPACING_BOTTOM_DIR_2_GENERIC",
    ),
}
_LAYER_BIP_BAR_TYPE = {
    u"exterior_major": (
        u"REBAR_SYSTEM_BAR_TYPE_TOP_DIR_1",
        u"REBAR_SYSTEM_BAR_TYPE_TOP_DIR_1_GENERIC",
    ),
    u"exterior_minor": (
        u"REBAR_SYSTEM_BAR_TYPE_TOP_DIR_2",
        u"REBAR_SYSTEM_BAR_TYPE_TOP_DIR_2_GENERIC",
    ),
    u"interior_major": (
        u"REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_1",
        u"REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_1_GENERIC",
    ),
    u"interior_minor": (
        u"REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_2",
        u"REBAR_SYSTEM_BAR_TYPE_BOTTOM_DIR_2_GENERIC",
    ),
}

# LookupParameter names (fallback si BIP no aplica) — una sola vez a nivel módulo.
_LAYER_DIR_PARAM_NAMES = {
    u"exterior_major": (
        u"Exterior Major Direction",
        u"Top Major Direction",
        u"Top Mayor Direction",
    ),
    u"exterior_minor": (u"Exterior Minor Direction", u"Top Minor Direction"),
    u"interior_major": (
        u"Interior Major Direction",
        u"Bottom Major Direction",
        u"Bottom Mayor Direction",
    ),
    u"interior_minor": (u"Interior Minor Direction", u"Bottom Minor Direction"),
}
_LAYER_SPACING_PARAM_NAMES = {
    u"exterior_major": (u"Exterior Major Spacing", u"Top Major Spacing"),
    u"exterior_minor": (u"Exterior Minor Spacing", u"Top Minor Spacing"),
    u"interior_major": (u"Interior Major Spacing", u"Bottom Major Spacing"),
    u"interior_minor": (u"Interior Minor Spacing", u"Bottom Minor Spacing"),
}
_LAYER_BAR_TYPE_APPLY_PARAM_NAMES = {
    u"exterior_major": (
        u"Exterior Major Bar Type",
        u"Exterior Major Rebar Type",
        u"Top Major Bar Type",
    ),
    u"exterior_minor": (
        u"Exterior Minor Bar Type",
        u"Exterior Minor Rebar Type",
        u"Top Minor Bar Type",
    ),
    u"interior_major": (
        u"Interior Major Bar Type",
        u"Interior Major Rebar Type",
        u"Bottom Major Bar Type",
    ),
    u"interior_minor": (
        u"Interior Minor Bar Type",
        u"Interior Minor Rebar Type",
        u"Bottom Minor Bar Type",
    ),
}

# BuiltInParameter enums resueltos una vez (evita getattr string por AR/capa).
_LAYER_BIP_SPACING_ENUMS = None
_LAYER_BIP_BAR_TYPE_ENUMS = None


def _resolved_layer_bip_enums():
    """Cache de enums BuiltInParameter por capa (spacing / bar type)."""
    global _LAYER_BIP_SPACING_ENUMS, _LAYER_BIP_BAR_TYPE_ENUMS
    if _LAYER_BIP_SPACING_ENUMS is not None and _LAYER_BIP_BAR_TYPE_ENUMS is not None:
        return _LAYER_BIP_SPACING_ENUMS, _LAYER_BIP_BAR_TYPE_ENUMS
    spacing = {}
    bar_type = {}
    for key in _LAYER_KEYS:
        spacing[key] = []
        for bip_name in _LAYER_BIP_SPACING.get(key, ()):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
            except Exception:
                bip = None
            if bip is not None:
                spacing[key].append(bip)
        bar_type[key] = []
        for bip_name in _LAYER_BIP_BAR_TYPE.get(key, ()):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
            except Exception:
                bip = None
            if bip is not None:
                bar_type[key].append(bip)
    _LAYER_BIP_SPACING_ENUMS = spacing
    _LAYER_BIP_BAR_TYPE_ENUMS = bar_type
    return spacing, bar_type


def _set_param_double(param, value):
    if param is None or param.IsReadOnly:
        return False
    try:
        if param.StorageType != StorageType.Double:
            return False
        param.Set(float(value))
        return True
    except Exception:
        return False


def _set_param_element_id(param, eid):
    if param is None or param.IsReadOnly:
        return False
    if eid is None or eid == ElementId.InvalidElementId:
        return False
    try:
        if param.StorageType != StorageType.ElementId:
            return False
        param.Set(eid)
        return True
    except Exception:
        return False


def _set_param_int01(param, value):
    if param is None or param.IsReadOnly:
        return False
    try:
        param.Set(1 if value else 0)
        return True
    except Exception:
        return False


def _aplicar_capas(area_rein, layer_cfg):
    """
    layer_cfg[key] = dict(active, bar_type_id, spacing_mm).

    Espaciado/tipo: BuiltInParameter primero (UI localizada), luego LookupParameter.
    Nota: el espaciado de AreaReinforcement es *Maximum Spacing* (≤ valor pedido);
    Revit puede ajustar el paso real al ancho del contorno.
    """
    if area_rein is None:
        return
    bip_spacing, bip_bar_type = _resolved_layer_bip_enums()

    for key in _LAYER_KEYS:
        cfg = layer_cfg.get(key) or {}
        active = bool(cfg.get(u"active", True))
        spacing_mm = cfg.get(u"spacing_mm", 150)
        bar_id = cfg.get(u"bar_type_id")
        layer_type = _LAYER_TYPE.get(key)
        if layer_type is not None:
            try:
                area_rein.SetLayerActive(layer_type, active)
            except Exception:
                pass
        for pname in _LAYER_DIR_PARAM_NAMES.get(key, ()):
            try:
                if _set_param_int01(area_rein.LookupParameter(pname), active):
                    break
            except Exception:
                continue
        spacing_int = _mm_to_internal(spacing_mm)
        spacing_ok = False
        for bip in bip_spacing.get(key, ()):
            try:
                if _set_param_double(area_rein.get_Parameter(bip), spacing_int):
                    spacing_ok = True
                    break
            except Exception:
                continue
        if not spacing_ok:
            for pname in _LAYER_SPACING_PARAM_NAMES.get(key, ()):
                try:
                    if _set_param_double(
                        area_rein.LookupParameter(pname), spacing_int
                    ):
                        break
                except Exception:
                    pass
        if bar_id and bar_id != ElementId.InvalidElementId:
            bar_ok = False
            for bip in bip_bar_type.get(key, ()):
                try:
                    if _set_param_element_id(area_rein.get_Parameter(bip), bar_id):
                        bar_ok = True
                        break
                except Exception:
                    continue
            if not bar_ok:
                for pname in _LAYER_BAR_TYPE_APPLY_PARAM_NAMES.get(key, ()):
                    try:
                        if _set_param_element_id(
                            area_rein.LookupParameter(pname), bar_id
                        ):
                            break
                    except Exception:
                        pass


def _collect_rebars_de_area_reinforcement(document, area_rein):
    """``Rebar`` (OST_Rebar) hijos de un ``AreaReinforcement`` tras ``Regenerate``."""
    rebars = []
    if area_rein is None or document is None:
        return rebars
    try:
        from Autodesk.Revit.DB import ElementCategoryFilter

        flt = ElementCategoryFilter(BuiltInCategory.OST_Rebar)
        dep = area_rein.GetDependentElements(flt)
        if dep is None:
            return rebars
        try:
            nd = int(dep.Count)
        except Exception:
            nd = 0
        for i in range(nd):
            try:
                el = document.GetElement(dep[i])
                if isinstance(el, Rebar):
                    rebars.append(el)
            except Exception:
                continue
    except Exception:
        pass
    return rebars


def _collect_rebar_in_system_de_area_reinforcement(document, area_rein):
    """``RebarInSystem`` hijos de un ``AreaReinforcement`` tras ``Regenerate``."""
    barras = []
    if area_rein is None or document is None:
        return barras
    seen = set()
    try:
        sys_ids = area_rein.GetRebarInSystemIds()
    except Exception:
        sys_ids = None
    if sys_ids is not None:
        try:
            nd = int(sys_ids.Count)
        except Exception:
            nd = 0
        for i in range(nd):
            try:
                eid = sys_ids[i]
                eid_int = _element_id_int(eid)
                if eid_int is None or eid_int in seen:
                    continue
                el = document.GetElement(eid)
                if isinstance(el, RebarInSystem):
                    barras.append(el)
                    seen.add(eid_int)
            except Exception:
                continue
    if barras:
        return barras
    try:
        from Autodesk.Revit.DB import ElementCategoryFilter

        flt = ElementCategoryFilter(BuiltInCategory.OST_RebarInSystem)
        dep = area_rein.GetDependentElements(flt)
        if dep is None:
            return barras
        try:
            nd = int(dep.Count)
        except Exception:
            nd = 0
        for i in range(nd):
            try:
                eid_int = _element_id_int(dep[i])
                if eid_int is None or eid_int in seen:
                    continue
                el = document.GetElement(dep[i])
                if isinstance(el, RebarInSystem):
                    barras.append(el)
                    seen.add(eid_int)
            except Exception:
                continue
    except Exception:
        pass
    return barras


def _collect_barras_params_de_area_reinforcement(document, area_rein):
    """
    Barras a estampar (mismo criterio que ``area_reinforcement_losa``):
    ``Rebar`` estructural y ``RebarInSystem``.
    """
    barras = []
    seen = set()
    for rb in _collect_rebars_de_area_reinforcement(document, area_rein):
        eid_int = _element_id_int(rb.Id)
        if eid_int is None or eid_int in seen:
            continue
        barras.append(rb)
        seen.add(eid_int)
    for rb in _collect_rebar_in_system_de_area_reinforcement(document, area_rein):
        eid_int = _element_id_int(rb.Id)
        if eid_int is None or eid_int in seen:
            continue
        barras.append(rb)
        seen.add(eid_int)
    return barras


def _resolver_vista_para_show_middle(document, view):
    """
    Vista de modelo válida para ``SetPresentationMode`` (no plantilla).
    Relee el elemento desde ``document`` (mismo criterio que Unobscured).
    """
    if document is None or view is None:
        return None
    try:
        if not isinstance(view, View):
            return None
    except Exception:
        return None
    try:
        if bool(view.IsTemplate):
            return None
    except Exception:
        pass
    try:
        resolved = document.GetElement(view.Id)
        if isinstance(resolved, View):
            view = resolved
    except Exception:
        pass
    try:
        if bool(view.IsTemplate):
            return None
    except Exception:
        pass
    return view


def _presentacion_show_middle_en_vista(barra, view):
    """
    En la vista dada, presentación **Middle** del conjunto (UI *Show Middle*).
    Aplica a ``Rebar`` y ``RebarInSystem``.
    Si ``CanApplyPresentationMode`` es False o falla, intenta ``SetPresentationMode``
    como respaldo (como ``vista_seccion_enfierrado_vigas``).
    """
    if barra is None or view is None:
        return False
    if not isinstance(barra, (Rebar, RebarInSystem)):
        return False
    try:
        can = False
        try:
            can = bool(barra.CanApplyPresentationMode(view))
        except Exception:
            can = False
        if can:
            barra.SetPresentationMode(view, RebarPresentationMode.Middle)
            return True
        # Respaldo: CanApply denegó o no disponible
        barra.SetPresentationMode(view, RebarPresentationMode.Middle)
        return True
    except Exception:
        return False


def _element_ids_to_rebars(document, id_list):
    """Convierte IEnumerable/list de ElementId → ``Rebar`` / ``RebarInSystem``."""
    out = []
    if document is None or id_list is None:
        return out
    try:
        iterable = list(id_list)
    except Exception:
        try:
            nd = int(id_list.Count)
            iterable = [id_list[i] for i in range(nd)]
        except Exception:
            return out
    for eid in iterable:
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if isinstance(el, (Rebar, RebarInSystem)):
            out.append(el)
    return out


def _remove_area_reinforcement_system(document, area_rein):
    """
    UI «Remove Area Reinforcement System»: elimina el AR y deja ``Rebar`` libres.

    MRA («Recorrido Barras») requiere Rebar individuales; ``RebarInSystem`` no
    admite MultiReferenceAnnotation.
    """
    if document is None or area_rein is None:
        return []
    try:
        new_ids = AreaReinforcement.RemoveAreaReinforcementSystem(
            document, area_rein
        )
    except Exception:
        return []
    return _element_ids_to_rebars(document, new_ids)


def _bar_center_xyz(barra, view=None):
    """Centro bbox (vista → global) o Show Middle como respaldo."""
    if barra is None:
        return None
    for v in (view, None):
        try:
            bb = barra.get_BoundingBox(v)
            if bb is not None:
                return (bb.Min + bb.Max) * 0.5
        except Exception:
            pass
    try:
        return _punto_insercion_tag_show_middle(barra, view)
    except Exception:
        return None


def _stamp_snapshots_from_barras(document, area_rein, barras, view=None):
    """
    Antes de RemoveAreaSystem: captura ubicación/posición + centro por barra.

    Tras el dissolve los Id cambian; el match se hace por centro geométrico.
    """
    snaps = []
    if not barras:
        return snaps
    ubicacion_por_id = {}
    posicion_por_id = {}
    try:
        ubicacion_por_id = _resolver_ubicacion_barras_area_reinforcement(
            document, area_rein, barras
        )
    except Exception:
        ubicacion_por_id = {}
    try:
        posicion_por_id = _resolver_posicion_barras_area_reinforcement(
            document, area_rein, barras
        )
    except Exception:
        posicion_por_id = {}
    for barra in barras:
        try:
            eid = _element_id_int(barra.Id)
        except Exception:
            eid = None
        c = _bar_center_xyz(barra, view)
        if c is None:
            continue
        snaps.append(
            {
                u"center": c,
                u"ubicacion": ubicacion_por_id.get(eid) if eid is not None else None,
                u"posicion": posicion_por_id.get(eid) if eid is not None else None,
            }
        )
    return snaps


def _match_snap_nearest(barra, snaps, view=None, used=None):
    """Asigna el snapshot más cercano (cada snap como máximo una vez)."""
    if not snaps or barra is None:
        return None
    c = _bar_center_xyz(barra, view)
    if c is None:
        return None
    best = None
    best_d2 = None
    best_i = None
    for i, snap in enumerate(snaps):
        if used is not None and i in used:
            continue
        sc = snap.get(u"center")
        if sc is None:
            continue
        d2 = _xyz_dist2(c, sc)
        if best_d2 is None or d2 < best_d2:
            best_d2 = d2
            best = snap
            best_i = i
    if best is not None and used is not None and best_i is not None:
        used.add(best_i)
    return best


def _stamp_armadura_params_en_rebars(
    document, rebars, snaps=None, conjunto_guid=None, nivel_valor=None, view=None
):
    """Estampa params corporativos en Rebar libres (post RemoveAreaSystem)."""
    if document is None or not rebars:
        return 0
    if stamp_armadura_conjunto_guid is None:
        return 0
    used = set()
    n_ok = 0
    for barra in rebars:
        try:
            barra = document.GetElement(barra.Id)
        except Exception:
            pass
        if barra is None or not isinstance(barra, (Rebar, RebarInSystem)):
            continue
        snap = _match_snap_nearest(barra, snaps, view=view, used=used)
        ubicacion = (snap or {}).get(u"ubicacion")
        posicion = (snap or {}).get(u"posicion")
        try:
            if stamp_armadura_arainco is not None:
                stamp_armadura_arainco(barra, yes=True)
        except Exception:
            pass
        try:
            if stamp_armadura_malla is not None:
                stamp_armadura_malla(barra, yes=True)
        except Exception:
            pass
        try:
            if stamp_armadura_ubicacion is not None and ubicacion:
                stamp_armadura_ubicacion(barra, ubicacion)
        except Exception:
            pass
        try:
            if stamp_armadura_posicion is not None and posicion:
                stamp_armadura_posicion(barra, posicion)
        except Exception:
            pass
        try:
            if stamp_armadura_nivel is not None and nivel_valor:
                stamp_armadura_nivel(barra, nivel_valor)
        except Exception:
            pass
        try:
            if stamp_armadura_conjunto_guid(barra, conjunto_guid=conjunto_guid):
                n_ok += 1
        except Exception:
            pass
    return n_ok


def _aplicar_show_middle_barras(document, barras, view):
    """Show Middle en una lista de ``Rebar`` / ``RebarInSystem``."""
    if document is None or view is None or not barras:
        return 0
    view = _resolver_vista_para_show_middle(document, view)
    if view is None:
        return 0
    n_ok = 0
    for barra in barras:
        try:
            barra = document.GetElement(barra.Id)
        except Exception:
            pass
        if _presentacion_show_middle_en_vista(barra, view):
            n_ok += 1
    return n_ok


def _project_dir_on_view(v, view):
    """Proyecta dirección al plano de vista (sin componente ViewDirection)."""
    if v is None or view is None:
        return None
    try:
        vd = view.ViewDirection
        if vd is None or float(vd.GetLength()) < 1e-12:
            return None
        vd = vd.Normalize()
        proj = v - vd.Multiply(float(v.DotProduct(vd)))
        if float(proj.GetLength()) < 1e-9:
            return None
        return proj.Normalize()
    except Exception:
        return None


def _spacing_dir_rebar_en_vista(barra, view):
    """
    Dirección de distribución del set en el plano de vista (recorrido MRA).

    Preferencia: pos0→posN; si no, bar_dir × vd; si no, RightDirection.
    """
    if barra is None or view is None:
        return None
    try:
        rd = view.RightDirection.Normalize()
    except Exception:
        rd = None
    try:
        vd = view.ViewDirection.Normalize()
    except Exception:
        return rd
    try:
        v_up = view.UpDirection.Normalize()
    except Exception:
        v_up = None

    npos = _numero_posiciones_barra(barra)
    if npos > 1:
        for mpo in _mpo_centerline_options():
            try:
                cs0 = list(barra.GetCenterlineCurves(False, False, False, mpo, 0))
                csn = list(
                    barra.GetCenterlineCurves(False, False, False, mpo, npos - 1)
                )
            except Exception:
                cs0 = csn = []
            if not cs0 or not csn:
                continue
            try:
                c0 = cs0[0].Evaluate(0.5, True)
                cn = csn[0].Evaluate(0.5, True)
                vec = cn - c0
                if float(vec.GetLength()) > 1e-6:
                    proj = _project_dir_on_view(vec.Normalize(), view)
                    if proj is not None:
                        return proj
            except Exception:
                pass

    bar_xy = _direccion_barra_xy(barra)
    if bar_xy is not None:
        try:
            bar_dir = XYZ(float(bar_xy.X), float(bar_xy.Y), float(bar_xy.Z))
            if float(bar_dir.GetLength()) > 1e-9:
                bar_dir = bar_dir.Normalize()
                try:
                    if abs(float(bar_dir.DotProduct(vd))) > 0.8 and v_up is not None:
                        return v_up
                except Exception:
                    pass
                try:
                    spacing = bar_dir.CrossProduct(vd)
                    if float(spacing.GetLength()) > 1e-9:
                        proj = _project_dir_on_view(spacing.Normalize(), view)
                        if proj is not None:
                            return proj
                except Exception:
                    pass
        except Exception:
            pass
    return rd


def _mra_margen_lateral_ft():
    """Solo margen pequeño (mm→ft). No usar ½ bbox: eso lleva el MRA al extremo."""
    try:
        from Autodesk.Revit.DB import UnitTypeId, UnitUtils

        return float(
            UnitUtils.ConvertToInternalUnits(
                float(_MRA_OFFSET_EXTRA_MM), UnitTypeId.Millimeters
            )
        )
    except Exception:
        return float(_MRA_OFFSET_EXTRA_MM) / 304.8


def _punto_mitad_longitud_barra_mra(barra, view):
    """
    Mitad de la longitud de la barra Show Middle (centerline), no centro del set.

    Orden: arco path de la posición Middle → curva más larga → transform desde
    pos. 0. Sin bbox del conjunto (desplaza el ancla fuera del medio de barra).
    """
    if barra is None:
        return None
    mid_idx = _indice_barra_show_middle(barra)
    curvas = _curvas_centerline_posicion_barra(barra, mid_idx)
    pt = _midpoint_arco_path_curvas(curvas)
    if pt is None:
        pt = _midpoint_curva_mas_larga(curvas)
    if pt is not None:
        return _proyectar_punto_plano_vista(pt, view)
    if mid_idx != 0:
        try:
            base = []
            for mpo in _mpo_centerline_options():
                try:
                    raw = barra.GetCenterlineCurves(False, False, False, mpo, 0)
                    base = list(raw) if raw is not None else []
                    if base:
                        break
                except Exception:
                    base = []
            p0 = _midpoint_arco_path_curvas(base)
            if p0 is None:
                p0 = _midpoint_curva_mas_larga(base)
            tr = _combinar_transforms_barra(barra, mid_idx)
            if p0 is not None and tr is not None:
                return _proyectar_punto_plano_vista(tr.OfPoint(p0), view)
        except Exception:
            pass
    # Último recurso: helper de tag (puede caer a bbox)
    try:
        return _punto_insercion_tag_show_middle(barra, view)
    except Exception:
        return None


def _crear_mra_rebar_uno(document, view, barra, mrat_type, avisos):
    """
    Una MRA «Recorrido Barras» en Rebar libre.

    Ancla en la **mitad de la longitud** de la barra Show Middle.
    ``DimensionLineDirection`` = distribución; offset lateral mínimo (no a lo
    largo de la barra).
    """
    if avisos is None:
        avisos = []
    if document is None or view is None or barra is None or mrat_type is None:
        return False
    try:
        from Autodesk.Revit.DB import (
            DimensionStyleType,
            MultiReferenceAnnotation,
            MultiReferenceAnnotationOptions,
        )
    except Exception as ex:
        avisos.append(u"MRA: imports fallaron ({0}).".format(_as_unicode(ex)))
        return False

    try:
        rid = int(barra.Id.IntegerValue)
    except Exception:
        rid = 0

    # Mitad de longitud de la barra del set (Show Middle), no extremo ni bbox set.
    p_mid = _punto_mitad_longitud_barra_mra(barra, view)
    p_mid = _proyectar_punto_plano_vista(p_mid, view)
    if p_mid is None:
        avisos.append(u"MRA Id {0}: sin mitad de longitud de barra.".format(rid))
        return False
    try:
        vd = view.ViewDirection.Normalize()
    except Exception:
        avisos.append(u"MRA Id {0}: ViewDirection inválida.".format(rid))
        return False

    spacing_dir = _spacing_dir_rebar_en_vista(barra, view)
    if spacing_dir is None:
        avisos.append(u"MRA Id {0}: sin dirección de distribución.".format(rid))
        return False

    # Offset lateral = a lo largo de la distribución (⟂ a la barra en planta).
    # Así se conserva la coordenada a media longitud de la barra.
    try:
        off_ft = _mra_margen_lateral_ft()
        p_line = p_mid + spacing_dir.Multiply(float(off_ft))
    except Exception:
        p_line = p_mid

    try:
        opts = MultiReferenceAnnotationOptions(mrat_type)
    except Exception:
        try:
            opts = MultiReferenceAnnotationOptions()
            opts.MultiReferenceAnnotationType = mrat_type.Id
        except Exception as ex:
            avisos.append(
                u"MRA Id {0}: options ({1}).".format(rid, _as_unicode(ex))
            )
            return False
    try:
        opts.DimensionStyleType = DimensionStyleType.Linear
    except Exception:
        pass
    try:
        opts.DimensionPlaneNormal = vd
        opts.DimensionLineDirection = spacing_dir
        opts.DimensionLineOrigin = p_line
        opts.TagHeadPosition = p_mid  # cabecera = mitad de longitud
        opts.TagHasLeader = False
    except Exception as ex:
        avisos.append(
            u"MRA Id {0}: config ({1}).".format(rid, _as_unicode(ex))
        )
        return False

    ids = List[ElementId]()
    ids.Add(barra.Id)
    try:
        opts.SetElementsToDimension(ids)
    except Exception as ex:
        avisos.append(
            u"MRA Id {0}: SetElements ({1}).".format(rid, _as_unicode(ex))
        )
        return False
    try:
        if hasattr(opts, u"ElementsMatchReferenceCategory"):
            if not opts.ElementsMatchReferenceCategory(document):
                avisos.append(
                    u"MRA Id {0}: no válido para el tipo «{1}».".format(
                        rid, _MRA_TYPE_NAME_RECORRIDO_BARRAS
                    )
                )
                return False
    except Exception:
        pass
    try:
        mra = MultiReferenceAnnotation.Create(document, view.Id, opts)
        if mra is None:
            avisos.append(u"MRA Id {0}: Create retornó None.".format(rid))
            return False
    except Exception as ex:
        avisos.append(
            u"MRA Id {0}: Create falló ({1}).".format(rid, _as_unicode(ex))
        )
        return False
    # Reafirmar cabecera en la mitad de la barra (Create puede desplazarla).
    try:
        from Autodesk.Revit.DB import ElementCategoryFilter, IndependentTag

        flt = ElementCategoryFilter(BuiltInCategory.OST_RebarTags)
        for did in mra.GetDependentElements(flt) or []:
            el = document.GetElement(did)
            if isinstance(el, IndependentTag):
                try:
                    if el.HasLeader:
                        el.HasLeader = False
                except Exception:
                    pass
                _aplicar_estilo_tag_rebar_sin_leader(el, p_mid)
    except Exception:
        pass
    return True


def _aplicar_mra_rebars(
    document, view, rebars, avisos
):
    """
    Una MRA «Recorrido Barras» por cada ``Rebar`` libre en ``rebars``.

    Omite ``RebarInSystem`` (no admiten MRA). Requiere transacción abierta.
    """
    if document is None or view is None:
        return 0
    if avisos is None:
        avisos = []
    libres = [
        rb
        for rb in (rebars or [])
        if rb is not None and isinstance(rb, Rebar)
    ]
    if not libres:
        if rebars:
            avisos.append(
                u"MRA: se requieren Rebar libres (no RebarInSystem). "
                u"Use Remove Area Reinforcement System."
            )
        return 0
    try:
        from geometria_estribos_viga import (
            _multi_reference_annotation_type_by_name,
            _vista_permite_multi_rebar_annotation,
        )
    except Exception:
        avisos.append(
            u"Multi-Rebar Annotation: no se pudo cargar el helper compartido."
        )
        return 0
    try:
        if not _vista_permite_multi_rebar_annotation(view):
            avisos.append(
                u"MRA «{0}»: use planta/alzado/sección (no plantilla ni 3D).".format(
                    _MRA_TYPE_NAME_RECORRIDO_BARRAS
                )
            )
            return 0
    except Exception:
        pass
    mrat_type = _multi_reference_annotation_type_by_name(
        document, _MRA_TYPE_NAME_RECORRIDO_BARRAS
    )
    if mrat_type is None:
        avisos.append(
            u"Multi-Rebar Annotation: no existe el tipo «{0}» en el proyecto.".format(
                _MRA_TYPE_NAME_RECORRIDO_BARRAS
            )
        )
        return 0
    n_ok = 0
    for rb in libres:
        try:
            rb = document.GetElement(rb.Id)
        except Exception:
            pass
        if not isinstance(rb, Rebar):
            continue
        if _crear_mra_rebar_uno(document, view, rb, mrat_type, avisos):
            n_ok += 1
    return int(n_ok)


def _aplicar_mra_barras_area_reinforcement(
    document, area_rein, view, avisos, barras=None
):
    """
    Compat: MRA sobre barras del AR.

    Preferir ``_aplicar_mra_rebars`` tras ``RemoveAreaReinforcementSystem``.
    """
    if barras is None:
        barras = _collect_barras_params_de_area_reinforcement(document, area_rein)
    return _aplicar_mra_rebars(document, view, barras, avisos)


def _vista_es_planta(view):
    """True si la vista activa es planta (``ViewPlan``) y no plantilla."""
    if view is None:
        return False
    try:
        if bool(view.IsTemplate):
            return False
    except Exception:
        pass
    try:
        return isinstance(view, ViewPlan)
    except Exception:
        return False


def _vista_ok_para_etiquetas_rebar(view):
    """Solo planta (no plantilla ni 3D / alzado / sección)."""
    if view is None:
        return False, u"Etiqueta rebar: no hay vista activa."
    if not _vista_es_planta(view):
        return False, u"Etiqueta rebar: se requiere una vista de planta."
    return True, None


def _numero_posiciones_barra(barra):
    if barra is None:
        return 0
    try:
        return int(barra.NumberOfBarPositions)
    except Exception:
        pass
    try:
        if hasattr(barra, "GetNumberOfBarPositions"):
            return int(barra.GetNumberOfBarPositions())
    except Exception:
        pass
    return 0


def _indice_barra_show_middle(barra):
    """
    Índice de la barra visible con ``RebarPresentationMode.Middle``.

    Usa ``NumberOfBarPositions // 2``; si esa posición está excluida
    (``IncludeFirstBar`` / ``IncludeLastBar``), busca la existente más cercana.
    """
    npos = _numero_posiciones_barra(barra)
    if npos <= 0:
        return 0
    mid = int(npos / 2)
    if not hasattr(barra, "DoesBarExistAtPosition"):
        return mid
    try:
        if bool(barra.DoesBarExistAtPosition(mid)):
            return mid
    except Exception:
        return mid
    for delta in range(1, npos):
        for idx in (mid - delta, mid + delta):
            if idx < 0 or idx >= npos:
                continue
            try:
                if bool(barra.DoesBarExistAtPosition(idx)):
                    return int(idx)
            except Exception:
                continue
    return mid


def _longitud_curva_safe(curve):
    if curve is None:
        return 0.0
    try:
        return float(curve.Length)
    except Exception:
        pass
    try:
        return float(curve.GetEndPoint(0).DistanceTo(curve.GetEndPoint(1)))
    except Exception:
        return 0.0


def _midpoint_curva_mas_larga(curvas):
    """Punto a param 0.5 de la centerline más larga; None si no hay curvas útiles."""
    if not curvas:
        return None
    best = None
    best_len = -1.0
    for c in curvas:
        if c is None:
            continue
        ln = _longitud_curva_safe(c)
        if ln > best_len:
            best = c
            best_len = ln
    if best is None:
        return None
    try:
        return best.Evaluate(0.5, True)
    except Exception:
        return None


def _midpoint_arco_path_curvas(curvas):
    """
    Midpoint por longitud de arco del path completo (varias curvas en orden).

    Más fiable que param 0.5 de la curva más larga en formas multi-tramo.
    """
    segs = []
    total = 0.0
    for c in curvas or []:
        if c is None:
            continue
        ln = _longitud_curva_safe(c)
        if ln <= 1e-12:
            continue
        segs.append((c, ln))
        total += ln
    if not segs:
        return None
    if len(segs) == 1 or total <= 1e-12:
        try:
            return segs[0][0].Evaluate(0.5, True)
        except Exception:
            return None
    target = total * 0.5
    acc = 0.0
    for c, ln in segs:
        if acc + ln >= target - 1e-12:
            try:
                return c.Evaluate((target - acc) / ln, True)
            except Exception:
                try:
                    return c.Evaluate(0.5, True)
                except Exception:
                    return None
        acc += ln
    try:
        return segs[-1][0].Evaluate(1.0, True)
    except Exception:
        return None


def _xyz_dist2(a, b):
    if a is None or b is None:
        return 1.0e99
    try:
        dx = float(a.X) - float(b.X)
        dy = float(a.Y) - float(b.Y)
        dz = float(a.Z) - float(b.Z)
        return dx * dx + dy * dy + dz * dz
    except Exception:
        return 1.0e99


def _transform_is_identity(tr):
    if tr is None:
        return True
    try:
        return bool(tr.IsIdentity)
    except Exception:
        pass
    try:
        o = tr.Origin
        return (
            abs(float(o.X)) < 1e-9
            and abs(float(o.Y)) < 1e-9
            and abs(float(o.Z)) < 1e-9
        )
    except Exception:
        return False


def _bar_position_transform(barra, bar_idx):
    """``GetBarPositionTransform`` en Rebar / RebarInSystem / ShapeDrivenAccessor."""
    if barra is None:
        return None
    bi = int(bar_idx)
    try:
        if hasattr(barra, "GetBarPositionTransform"):
            return barra.GetBarPositionTransform(bi)
    except Exception:
        pass
    try:
        acc = barra.GetShapeDrivenAccessor()
        if acc is not None and hasattr(acc, "GetBarPositionTransform"):
            return acc.GetBarPositionTransform(bi)
    except Exception:
        pass
    return None


def _moved_bar_transform(barra, bar_idx):
    if barra is None or not hasattr(barra, "GetMovedBarTransform"):
        return None
    try:
        return barra.GetMovedBarTransform(int(bar_idx))
    except Exception:
        return None


def _combinar_transforms_barra(barra, bar_idx):
    """
    Transform modelo = MovedBar × BarPosition (como ``GetTransformedCenterlineCurves``).
    """
    tr = _bar_position_transform(barra, bar_idx)
    moved = _moved_bar_transform(barra, bar_idx)
    if tr is None:
        return moved
    if moved is None or _transform_is_identity(moved):
        return tr
    try:
        return moved.Multiply(tr)
    except Exception:
        try:
            return tr.Multiply(moved)
        except Exception:
            return tr


def _aplicar_transform_a_curvas(curvas, tr):
    if not curvas or tr is None:
        return list(curvas or [])
    if _transform_is_identity(tr):
        return list(curvas)
    moved = []
    for c in curvas:
        if c is None:
            continue
        try:
            moved.append(c.CreateTransformed(tr))
        except Exception:
            continue
    return moved if moved else list(curvas)


def _mpo_centerline_options():
    """Preferir path completo; planar como respaldo (planta / hooks)."""
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption
    except Exception:
        return []
    out = []
    for mpo_name in (
        u"IncludeAllMultiplanarCurves",
        u"IncludeOnlyPlanarCurves",
    ):
        mpo = getattr(MultiplanarOption, mpo_name, None)
        if mpo is not None:
            out.append(mpo)
    return out


def _curvas_centerline_posicion_barra(barra, bar_idx):
    """
    Centerlines en la posición visual ``bar_idx`` (Show Middle).

    1. ``GetTransformedCenterlineCurves`` (BarPosition + MovedBar).
    2. Si el API devolvió geometría de manejo (sigue en barra 0), reaplicar transform.
    3. Respaldo BuildingCoder: curvas índice 0 + ``CreateTransformed``.
    """
    if barra is None:
        return []
    bi = int(bar_idx)
    mpo_list = _mpo_centerline_options()
    # (adjustSelf, suppressHooks, suppressBend)
    flag_sets = (
        (False, False, False),
        (False, True, False),
    )

    def _as_list(raw):
        if raw is None:
            return []
        try:
            return list(raw)
        except Exception:
            return []

    # 1) Curvas ya en posición modelo
    if hasattr(barra, "GetTransformedCenterlineCurves"):
        for adj, sh, sb in flag_sets:
            for mpo in mpo_list:
                try:
                    curvas = _as_list(
                        barra.GetTransformedCenterlineCurves(adj, sh, sb, mpo, bi)
                    )
                except Exception:
                    curvas = []
                if not curvas:
                    continue
                # ¿GetTransformed devolvió geometría aún en barra 0?
                if bi != 0:
                    tr = _combinar_transforms_barra(barra, bi)
                    if tr is not None and not _transform_is_identity(tr):
                        pt_t = _midpoint_arco_path_curvas(curvas)
                        if pt_t is None:
                            pt_t = _midpoint_curva_mas_larga(curvas)
                        pt_0 = None
                        try:
                            c0 = _as_list(
                                barra.GetTransformedCenterlineCurves(
                                    adj, sh, sb, mpo, 0
                                )
                            )
                            if not c0:
                                c0 = _as_list(
                                    barra.GetCenterlineCurves(adj, sh, sb, mpo, 0)
                                )
                            pt_0 = _midpoint_arco_path_curvas(c0)
                            if pt_0 is None:
                                pt_0 = _midpoint_curva_mas_larga(c0)
                        except Exception:
                            pt_0 = None
                        if (
                            pt_t is not None
                            and pt_0 is not None
                            and _xyz_dist2(pt_t, pt_0) < 1e-8
                        ):
                            try:
                                expected = tr.OfPoint(pt_0)
                            except Exception:
                                expected = None
                            if (
                                expected is not None
                                and _xyz_dist2(expected, pt_0) > 1e-6
                            ):
                                fixed = _aplicar_transform_a_curvas(curvas, tr)
                                if fixed:
                                    return fixed
                return curvas

    # 2) Manejo (barra 0) + transform a ``bi`` — Rebar / RebarInSystem
    handle = []
    for adj, sh, sb in flag_sets:
        for mpo in mpo_list:
            try:
                handle = _as_list(
                    barra.GetCenterlineCurves(adj, sh, sb, mpo, 0)
                )
            except Exception:
                handle = []
            if handle:
                break
        if handle:
            break
    if not handle:
        try:
            handle = _as_list(barra.GetCenterlineCurves(False, False, False))
        except Exception:
            handle = []
    if not handle:
        # Último intento: GetCenterlineCurves en el propio índice (free-form)
        for mpo in mpo_list:
            try:
                handle = _as_list(
                    barra.GetCenterlineCurves(False, False, False, mpo, bi)
                )
            except Exception:
                handle = []
            if handle:
                break
    if not handle:
        return []

    if bi == 0:
        return handle
    tr = _combinar_transforms_barra(barra, bi)
    moved = _aplicar_transform_a_curvas(handle, tr)
    return moved if moved else handle


def _punto_insercion_tag_show_middle(barra, view):
    """
    Cabecera = midpoint (arco) de la centerline de la barra Show Middle.

    Prefiere path transformado; ``OfPoint`` desde barra 0; bbox del set solo
    como último recurso (centra el reparto, no la barra media).
    """
    if barra is None:
        return None
    mid_idx = _indice_barra_show_middle(barra)

    curvas = _curvas_centerline_posicion_barra(barra, mid_idx)
    pt = _midpoint_arco_path_curvas(curvas)
    if pt is None:
        pt = _midpoint_curva_mas_larga(curvas)
    if pt is not None:
        return _proyectar_punto_plano_vista(pt, view)

    # Respaldo: midpoint barra 0 + transform de posición media
    if mid_idx != 0:
        try:
            base = []
            for mpo in _mpo_centerline_options():
                try:
                    raw = barra.GetCenterlineCurves(False, False, False, mpo, 0)
                    base = list(raw) if raw is not None else []
                    if base:
                        break
                except Exception:
                    base = []
            p0 = _midpoint_arco_path_curvas(base)
            if p0 is None:
                p0 = _midpoint_curva_mas_larga(base)
            tr = _combinar_transforms_barra(barra, mid_idx)
            if p0 is not None and tr is not None:
                try:
                    return _proyectar_punto_plano_vista(tr.OfPoint(p0), view)
                except Exception:
                    pass
        except Exception:
            pass

    for v in (view, None):
        try:
            bb = barra.get_BoundingBox(v)
            if bb is not None:
                return _proyectar_punto_plano_vista((bb.Min + bb.Max) * 0.5, view)
        except Exception:
            pass
    return None


def _proyectar_punto_plano_vista(p, view):
    """Proyecta ``p`` al plano de la vista (evita offset por Z fuera del corte)."""
    if p is None or view is None:
        return p
    try:
        vd = view.ViewDirection
        if vd is None or float(vd.GetLength()) < 1e-12:
            return p
        vd = vd.Normalize()
        vo = view.Origin
        if vo is None:
            return p
        d = float((p - vo).DotProduct(vd))
        return p - vd.Multiply(d)
    except Exception:
        return p


def _resolver_tag_type_id_por_shape(document, barra, tag_map, fallback_type):
    """
    Tipo de ``EST_A_STRUCTURAL REBAR TAG_FLOOR`` = nombre del RebarShape;
    si no hay homónimo, ``fallback_type`` (p. ej. «01»).
    """
    if not tag_map:
        return None
    try:
        from enfierrado_shaft_hashtag import (
            _norm_text,
            _primary_rebar_shape_tag_key,
            _rebar_shape_name_candidates,
        )
    except Exception:
        return tag_map.get(_as_unicode(fallback_type).strip().lower()) if fallback_type else None

    primary = _primary_rebar_shape_tag_key(document, barra)
    if primary and primary in tag_map:
        return tag_map[primary]
    for sk in _rebar_shape_name_candidates(document, barra) or []:
        if sk in tag_map:
            return tag_map[sk]
        # «1» ↔ «01»
        try:
            if sk and all(u"0" <= c <= u"9" for c in sk):
                for alt in (sk.zfill(2), type(u"")(int(sk)), type(u"")(int(sk)).zfill(2)):
                    an = _norm_text(alt)
                    if an and an in tag_map:
                        return tag_map[an]
        except Exception:
            pass
    if fallback_type:
        fb = _norm_text(fallback_type)
        if fb and fb in tag_map:
            return tag_map[fb]
    return None


def _aplicar_estilo_tag_rebar_sin_leader(tag, head):
    """Sin leader; fuerza cabecera en el midpoint Show Middle (después de HasLeader)."""
    if tag is None or head is None:
        return
    try:
        tag.TagHeadPosition = head
    except Exception:
        pass
    try:
        tag.HasLeader = False
    except Exception:
        pass
    # Revit a veces reposiciona al quitar el leader: reaplicar cabecera.
    try:
        tag.TagHeadPosition = head
    except Exception:
        pass


def _crear_independent_tag_rebar(document, view, barra, tag_type_id, point):
    """
    Crea ``IndependentTag`` sin leader, cabecera en ``point``.
    Reintenta referencia y orientación (Horizontal/Vertical); no re-habilita leader.
    """
    if document is None or view is None or barra is None or tag_type_id is None:
        return None
    if point is None:
        return None
    try:
        from enfierrado_shaft_hashtag import _rebar_reference_candidates_for_tag

        refs = _rebar_reference_candidates_for_tag(document, view, barra)
    except Exception:
        refs = []
    if not refs:
        # Preferir referencia a la posición Show Middle
        npos = _numero_posiciones_barra(barra)
        mid_idx = _indice_barra_show_middle(barra)
        for idx in (mid_idx, 0, max(0, npos - 1)):
            try:
                if hasattr(barra, "GetReferenceToBarPosition"):
                    r = barra.GetReferenceToBarPosition(idx)
                elif hasattr(barra, "GetReferenceForBarPosition"):
                    r = barra.GetReferenceForBarPosition(idx)
                else:
                    r = None
                if r is not None:
                    refs.append(r)
                    break
            except Exception:
                continue
        try:
            refs.append(Reference(barra))
        except Exception:
            pass
    if not refs:
        return None

    try:
        sym = document.GetElement(tag_type_id)
        if sym is not None and hasattr(sym, "IsActive") and not bool(sym.IsActive):
            sym.Activate()
    except Exception:
        pass

    add_leader = False
    created = None
    for ref in refs:
        for orient in (TagOrientation.Horizontal, TagOrientation.Vertical):
            try:
                created = IndependentTag.Create(
                    document,
                    tag_type_id,
                    view.Id,
                    ref,
                    add_leader,
                    orient,
                    point,
                )
            except Exception:
                created = None
            if created is not None:
                _aplicar_estilo_tag_rebar_sin_leader(created, point)
                return created
            try:
                created = IndependentTag.Create(
                    document,
                    view.Id,
                    ref,
                    add_leader,
                    TagMode.TM_ADDBY_CATEGORY,
                    orient,
                    point,
                )
                if created is not None:
                    try:
                        created.SetTypeId(tag_type_id)
                    except Exception:
                        pass
            except Exception:
                created = None
            if created is not None:
                _aplicar_estilo_tag_rebar_sin_leader(created, point)
                return created
    return created


def _aplicar_etiquetas_show_middle_barras(
    document, view, barras, avisos, tag_map=None
):
    """
    Una ``IndependentTag`` (``EST_A_STRUCTURAL REBAR TAG_FLOOR``) por barra,
    anclada a Show Middle. Tipo = RebarShape; fallback ``01``.
    """
    if document is None or view is None or not barras:
        return 0
    if avisos is None:
        avisos = []
    ok_view, msg_view = _vista_ok_para_etiquetas_rebar(view)
    if not ok_view:
        avisos.append(msg_view or u"Etiqueta rebar: vista no válida.")
        return 0

    if tag_map is None:
        try:
            from enfierrado_shaft_hashtag import _collect_rebar_tag_symbol_map
        except Exception as ex:
            avisos.append(
                u"Etiqueta rebar: no se pudo cargar helper compartido ({0}).".format(
                    _as_unicode(ex)
                )
            )
            return 0
        tag_map = _collect_rebar_tag_symbol_map(document, _REBAR_TAG_FAMILY_NAME)
    if not tag_map:
        avisos.append(
            u"Etiqueta rebar: no hay tipos de familia «{0}» en el documento "
            u"(cargue la familia e intente de nuevo). Las AR se crearon sin etiquetas.".format(
                _REBAR_TAG_FAMILY_NAME
            )
        )
        return 0

    n_ok = 0
    for barra in barras:
        try:
            barra = document.GetElement(barra.Id)
        except Exception:
            pass
        if barra is None:
            continue
        if not isinstance(barra, (Rebar, RebarInSystem)):
            continue
        tag_type_id = _resolver_tag_type_id_por_shape(
            document, barra, tag_map, _REBAR_TAG_FALLBACK_TYPE
        )
        if tag_type_id is None:
            try:
                rid = _element_id_int(barra.Id)
            except Exception:
                rid = u"?"
            avisos.append(
                u"Etiqueta Id {0}: sin tipo por shape ni fallback «{1}» en «{2}».".format(
                    rid, _REBAR_TAG_FALLBACK_TYPE, _REBAR_TAG_FAMILY_NAME
                )
            )
            continue
        p_raw = _punto_insercion_tag_show_middle(barra, view)
        p = _proyectar_punto_plano_vista(p_raw, view)
        if p is None:
            try:
                rid = _element_id_int(barra.Id)
            except Exception:
                rid = u"?"
            avisos.append(
                u"Etiqueta Id {0}: sin punto Show Middle para insertar.".format(rid)
            )
            continue
        created = _crear_independent_tag_rebar(
            document, view, barra, tag_type_id, p
        )
        if created is not None:
            _aplicar_estilo_tag_rebar_sin_leader(created, p)
            n_ok += 1
        else:
            try:
                rid = _element_id_int(barra.Id)
            except Exception:
                rid = u"?"
            avisos.append(
                u"Etiqueta Id {0}: no se pudo crear IndependentTag.".format(rid)
            )
    return n_ok


def _aplicar_etiquetas_show_middle_area_reinforcement(
    document, area_rein, view, avisos, barras=None, tag_map=None
):
    """
    Una ``IndependentTag`` por cada barra del AR (Show Middle).

    Soft-fail con avisos. Preferir ``_aplicar_etiquetas_show_middle_barras``
    tras RemoveAreaSystem.
    """
    if document is None or area_rein is None or view is None:
        return 0
    if avisos is None:
        avisos = []
    try:
        area_rein = document.GetElement(area_rein.Id)
    except Exception:
        pass
    if area_rein is None or not isinstance(area_rein, AreaReinforcement):
        return 0
    if barras is None:
        barras = _collect_barras_params_de_area_reinforcement(document, area_rein)
    if not barras:
        return 0
    return _aplicar_etiquetas_show_middle_barras(
        document, view, barras, avisos, tag_map=tag_map
    )


def _aplicar_show_middle_barras_area_reinforcement(
    document, area_rein, view, barras=None, allow_retry_regenerate=True
):
    """
    Tras ``Regenerate``, aplica ``RebarPresentationMode.Middle`` a las barras
    (``Rebar`` / ``RebarInSystem``) del ``AreaReinforcement`` en ``view``.

    ``barras`` opcional: si se pasa, no vuelve a recolectar (salvo lista vacía
    y ``allow_retry_regenerate``). Tras un Regenerate de batch, pasar
    ``allow_retry_regenerate=False`` para evitar un segundo Regenerate.
    """
    if document is None or area_rein is None or view is None:
        return 0
    view = _resolver_vista_para_show_middle(document, view)
    if view is None:
        return 0
    try:
        area_rein = document.GetElement(area_rein.Id)
    except Exception:
        area_rein = None
    if area_rein is None or not isinstance(area_rein, AreaReinforcement):
        return 0
    if barras is None:
        barras = _collect_barras_params_de_area_reinforcement(document, area_rein)
    if not barras and allow_retry_regenerate:
        # A veces hace falta un segundo Regenerate para materializar hijos
        try:
            document.Regenerate()
        except Exception:
            pass
        try:
            area_rein = document.GetElement(area_rein.Id)
        except Exception:
            area_rein = None
        if area_rein is None or not isinstance(area_rein, AreaReinforcement):
            return 0
        barras = _collect_barras_params_de_area_reinforcement(document, area_rein)
    n_ok = 0
    for barra in barras or []:
        try:
            barra = document.GetElement(barra.Id)
        except Exception:
            pass
        if _presentacion_show_middle_en_vista(barra, view):
            n_ok += 1
    return n_ok


def _host_losa_de_area_reinforcement(document, area_rein):
    """Floor host del AreaReinforcement (``GetHostId``), o ``None``."""
    if document is None or area_rein is None:
        return None
    try:
        hid = area_rein.GetHostId()
    except Exception:
        return None
    if hid is None or hid == ElementId.InvalidElementId:
        return None
    try:
        host = document.GetElement(hid)
    except Exception:
        return None
    return host if isinstance(host, Floor) else None


def _nivel_losa_como_string(document, floor):
    """Nombre del nivel de la losa host, como texto (misma convención que Malla en losa)."""
    if document is None or floor is None:
        return None
    lid = None
    try:
        lid = floor.LevelId
        if lid is None or lid == ElementId.InvalidElementId:
            lid = None
    except Exception:
        lid = None
    if lid is None:
        for bip_name in (
            u"INSTANCE_REFERENCE_LEVEL_PARAM",
            u"LEVEL_PARAM",
            u"SCHEDULE_LEVEL_PARAM",
        ):
            try:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                p = floor.get_Parameter(bip)
                if p is None or not p.HasValue or p.StorageType != StorageType.ElementId:
                    continue
                eid = p.AsElementId()
                if eid is not None and eid != ElementId.InvalidElementId:
                    lid = eid
                    break
            except Exception:
                pass
    if lid is None:
        return None
    try:
        level = document.GetElement(lid)
        if level is None:
            return None
        name = level.Name
        if name is None:
            return None
        return _as_unicode(name)
    except Exception:
        return None


def _nivel_losa_area_reinforcement(document, area_rein):
    """Nivel de la losa que hospeda el Area Reinforcement, como string."""
    floor = _host_losa_de_area_reinforcement(document, area_rein)
    return _nivel_losa_como_string(document, floor)


def _element_type_id_int(element):
    if element is None:
        return None
    try:
        return _element_id_int(element.GetTypeId())
    except Exception:
        return None


def _ubicacion_por_capa_area_reinforcement(layer_type):
    """
    TopOrFront (exterior / superior) → ``F'``;
    BottomOrBack (interior / inferior) → ``F``.
    """
    if layer_type in _LAYER_TOP:
        return ARMADURA_UBICACION_SUPERIOR
    if layer_type in _LAYER_BOTTOM:
        return ARMADURA_UBICACION_INFERIOR
    return None


def _bar_type_id_de_capa_area_reinforcement(area_rein, layer_key):
    """ElementId int del RebarBarType configurado en la capa del AR, o ``None``."""
    if area_rein is None or not layer_key:
        return None
    for pname in _LAYER_BAR_TYPE_PARAM_NAMES.get(layer_key, ()):
        try:
            p = area_rein.LookupParameter(pname)
            if p is None or p.StorageType != StorageType.ElementId:
                continue
            eid = p.AsElementId()
            tid = _element_id_int(eid)
            if tid is not None and eid != ElementId.InvalidElementId:
                return tid
        except Exception:
            continue
    return None


def _capas_activas_area_reinforcement(area_rein):
    """
    Capas activas del AR: ``[(layer_key, layer_type, bar_type_id_int), ...]``.
    """
    out = []
    if area_rein is None:
        return out
    for key in _LAYER_KEYS:
        layer_type = _LAYER_TYPE.get(key)
        if layer_type is None:
            continue
        try:
            if not bool(area_rein.IsLayerActive(layer_type)):
                continue
        except Exception:
            pass
        out.append(
            (key, layer_type, _bar_type_id_de_capa_area_reinforcement(area_rein, key))
        )
    return out


def _z_caras_losa(floor):
    """Z superior / inferior de la losa (bbox; suficiente para mitad de espesor)."""
    if floor is None:
        return None, None
    try:
        bb = floor.get_BoundingBox(None)
        if bb is not None:
            return float(bb.Max.Z), float(bb.Min.Z)
    except Exception:
        pass
    return None, None


def _ubicacion_por_geometria_cara_losa(barra, z_top, z_bottom):
    """
    Respaldo geométrico: mitad superior → ``F'``; inferior → ``F``.
    Prioriza BoundingBox (fiable en RebarInSystem).
    """
    if z_top is None or z_bottom is None or barra is None:
        return None
    z_mid = (float(z_top) + float(z_bottom)) * 0.5
    try:
        bb = barra.get_BoundingBox(None)
        if bb is not None:
            z_max = float(bb.Max.Z)
            z_min = float(bb.Min.Z)
            if z_min >= z_mid:
                return ARMADURA_UBICACION_SUPERIOR
            if z_max <= z_mid:
                return ARMADURA_UBICACION_INFERIOR
            z_centro = (z_max + z_min) * 0.5
            if z_centro >= z_mid:
                return ARMADURA_UBICACION_SUPERIOR
            return ARMADURA_UBICACION_INFERIOR
    except Exception:
        pass
    return None


def _resolver_ubicacion_barras_area_reinforcement(document, area_rein, barras):
    """
    ``{bar_id: F|F'}`` según capa AreaReinforcementLayerType.

    1. Tipo de barra único en capas de una sola cara → esa ubicación.
    2. Si ambas caras comparten tipo / ambigüedad → geometría Z vs losa.
    """
    out = {}
    if not barras or area_rein is None:
        return out
    capas = _capas_activas_area_reinforcement(area_rein)
    floor = _host_losa_de_area_reinforcement(document, area_rein)
    z_top, z_bottom = _z_caras_losa(floor)

    # Solo una cara activa → todas las barras de esa cara.
    faces = set()
    for _k, lt, _bt in capas:
        ub = _ubicacion_por_capa_area_reinforcement(lt)
        if ub:
            faces.add(ub)
    unica_cara = list(faces)[0] if len(faces) == 1 else None

    for barra in barras:
        eid = _element_id_int(barra.Id)
        if eid is None:
            continue
        if unica_cara is not None:
            out[eid] = unica_cara
            continue
        tid = _element_type_id_int(barra)
        matching = []
        if tid is not None:
            for _k, lt, bt in capas:
                if bt is not None and bt == tid:
                    ub = _ubicacion_por_capa_area_reinforcement(lt)
                    if ub and ub not in matching:
                        matching.append(ub)
        if len(matching) == 1:
            out[eid] = matching[0]
            continue
        geo = _ubicacion_por_geometria_cara_losa(barra, z_top, z_bottom)
        if geo:
            out[eid] = geo
    return out


def _posicion_por_capa_key(layer_key):
    """``Armadura_Posicion`` (``i``/``s``) según clave de capa UI."""
    if not layer_key:
        return None
    return _LAYER_POSICION.get(layer_key)


def _posicion_por_capa_type(layer_type):
    """``Armadura_Posicion`` (``i``/``s``) según AreaReinforcementLayerType."""
    if layer_type is None:
        return None
    return _LAYER_TYPE_POSICION.get(layer_type)


def _direccion_barra_xy(barra):
    """Vector XY de la dirección de la barra (centerline; distribución como respaldo)."""
    if barra is None:
        return None
    try:
        from Autodesk.Revit.DB.Structure import MultiplanarOption

        for mpo_name in (
            u"IncludeOnlyPlanarCurves",
            u"IncludeAllMultiplanarCurves",
        ):
            mpo = getattr(MultiplanarOption, mpo_name, None)
            if mpo is None:
                continue
            try:
                curves = barra.GetCenterlineCurves(
                    False, False, False, mpo, 0,
                )
            except Exception:
                curves = None
            if curves is None:
                continue
            try:
                n = int(curves.Count)
            except Exception:
                try:
                    n = len(curves)
                except Exception:
                    n = 0
            if n <= 0:
                continue
            c = curves[0]
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            return XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
    except Exception:
        pass
    try:
        path = barra.GetDistributionPath()
        if path is not None and path.IsBound:
            p0 = path.GetEndPoint(0)
            p1 = path.GetEndPoint(1)
            return XYZ(float(p1.X - p0.X), float(p1.Y - p0.Y), 0.0)
    except Exception:
        pass
    return None


def _vectores_paralelos_xy(vec_a, vec_b, tol=0.95):
    if vec_a is None or vec_b is None:
        return None
    try:
        ax, ay = float(vec_a.X), float(vec_a.Y)
        bx, by = float(vec_b.X), float(vec_b.Y)
        la = math.hypot(ax, ay)
        lb = math.hypot(bx, by)
        if la < 1e-9 or lb < 1e-9:
            return None
        dot = abs((ax / la) * (bx / lb) + (ay / la) * (by / lb))
        return dot >= float(tol)
    except Exception:
        return None


def _barra_es_direccion_major(area_rein, barra):
    """True si la barra sigue la dirección Major del AreaReinforcement."""
    if area_rein is None or barra is None:
        return None
    major_dir = None
    for lt in _LAYER_MAJOR:
        try:
            if bool(area_rein.IsLayerActive(lt)):
                major_dir = area_rein.GetLayerDirection(lt)
                if major_dir is not None:
                    break
        except Exception:
            pass
    if major_dir is None:
        try:
            major_dir = area_rein.GetLayerDirection(
                AreaReinforcementLayerType.TopOrFrontMajor
            )
        except Exception:
            major_dir = None
    if major_dir is None:
        try:
            major_dir = area_rein.Direction
        except Exception:
            major_dir = None
    paralelo = _vectores_paralelos_xy(_direccion_barra_xy(barra), major_dir)
    return paralelo


def _resolver_posicion_barras_area_reinforcement(document, area_rein, barras):
    """
    ``{bar_id: i|s}`` según cara × dirección (Major/Minor) del AreaReinforcement.

    Inferior Major→i / Minor→s; Superior Major→s / Minor→i.

    Jobs partidos por cara / Major-Minor / ahorro: si el AR solo tiene una
    combinación activa, todas las barras reciben ese valor.
    """
    out = {}
    if not barras or area_rein is None:
        return out
    capas = _capas_activas_area_reinforcement(area_rein)
    if not capas:
        return out

    valores = set()
    for key, lt, _bt in capas:
        val = _posicion_por_capa_key(key) or _posicion_por_capa_type(lt)
        if val:
            valores.add(val)
    # Un solo valor posible en este AR (p. ej. solo Major inferior, o ahorro Maj).
    if len(valores) == 1:
        unico = list(valores)[0]
        for barra in barras:
            eid = _element_id_int(barra.Id)
            if eid is not None:
                out[eid] = unico
        return out

    floor = _host_losa_de_area_reinforcement(document, area_rein)
    z_top, z_bottom = _z_caras_losa(floor)

    faces = set()
    dirs_major = set()
    for _k, lt, _bt in capas:
        ub = _ubicacion_por_capa_area_reinforcement(lt)
        if ub:
            faces.add(ub)
        dirs_major.add(lt in _LAYER_MAJOR)
    unica_cara = list(faces)[0] if len(faces) == 1 else None
    unica_dir_major = list(dirs_major)[0] if len(dirs_major) == 1 else None

    for barra in barras:
        eid = _element_id_int(barra.Id)
        if eid is None:
            continue
        tid = _element_type_id_int(barra)
        matching = []
        if tid is not None:
            for key, lt, bt in capas:
                if bt is not None and bt == tid:
                    val = _posicion_por_capa_key(key) or _posicion_por_capa_type(lt)
                    if val and val not in matching:
                        matching.append(val)
        if len(matching) == 1:
            out[eid] = matching[0]
            continue

        es_top = None
        if unica_cara is not None:
            es_top = unica_cara == ARMADURA_UBICACION_SUPERIOR
        else:
            geo = _ubicacion_por_geometria_cara_losa(barra, z_top, z_bottom)
            if geo == ARMADURA_UBICACION_SUPERIOR:
                es_top = True
            elif geo == ARMADURA_UBICACION_INFERIOR:
                es_top = False

        es_major = unica_dir_major
        if es_major is None:
            es_major = _barra_es_direccion_major(area_rein, barra)

        if es_top is None or es_major is None:
            continue
        if es_top:
            lt = (
                AreaReinforcementLayerType.TopOrFrontMajor
                if es_major
                else AreaReinforcementLayerType.TopOrFrontMinor
            )
        else:
            lt = (
                AreaReinforcementLayerType.BottomOrBackMajor
                if es_major
                else AreaReinforcementLayerType.BottomOrBackMinor
            )
        val = _posicion_por_capa_type(lt)
        if val:
            out[eid] = val
    return out


def stamp_armadura_params_en_area_reinforcement(
    document, area_rein, conjunto_guid=None, barras=None
):
    """
    Estampa en las barras del AreaReinforcement:
      - ``Armadura_Arainco`` = Yes
      - ``Armadura_Malla`` = Yes
      - ``Armadura_Ubicacion`` = F (inferior) / F' (superior) según capa
      - ``Armadura_Posicion`` = i/s según cara × Major/Minor
      - ``Armadura_Nivel`` = nombre del nivel de la losa host
      - ``Armadura_Conjunto_GUID``

    Requiere ``document.Regenerate()`` previo.
    Usa ``scripts/conjunto_guid.py`` (mismo helper que Malla en losa / columnas).
    Devuelve el número de barras con GUID estampado OK.
    ``barras`` opcional: reutiliza la colección del post-create.
    """
    if (
        document is None
        or area_rein is None
        or stamp_armadura_conjunto_guid is None
    ):
        return 0
    nivel_valor = None
    try:
        nivel_valor = _nivel_losa_area_reinforcement(document, area_rein)
    except Exception:
        nivel_valor = None
    if barras is None:
        barras = _collect_barras_params_de_area_reinforcement(document, area_rein)
    ubicacion_por_id = {}
    try:
        ubicacion_por_id = _resolver_ubicacion_barras_area_reinforcement(
            document, area_rein, barras,
        )
    except Exception:
        ubicacion_por_id = {}
    posicion_por_id = {}
    try:
        posicion_por_id = _resolver_posicion_barras_area_reinforcement(
            document, area_rein, barras,
        )
    except Exception:
        posicion_por_id = {}
    n_ok = 0
    for barra in barras or []:
        try:
            if stamp_armadura_arainco is not None:
                stamp_armadura_arainco(barra, yes=True)
        except Exception:
            pass
        try:
            if stamp_armadura_malla is not None:
                stamp_armadura_malla(barra, yes=True)
        except Exception:
            pass
        try:
            ubicacion = ubicacion_por_id.get(_element_id_int(barra.Id))
            if stamp_armadura_ubicacion is not None and ubicacion:
                stamp_armadura_ubicacion(barra, ubicacion)
        except Exception:
            pass
        try:
            posicion = posicion_por_id.get(_element_id_int(barra.Id))
            if stamp_armadura_posicion is not None and posicion:
                stamp_armadura_posicion(barra, posicion)
        except Exception:
            pass
        try:
            if stamp_armadura_nivel is not None and nivel_valor:
                stamp_armadura_nivel(barra, nivel_valor)
        except Exception:
            pass
        try:
            if stamp_armadura_conjunto_guid(barra, conjunto_guid=conjunto_guid):
                n_ok += 1
        except Exception:
            pass
    return n_ok


def _xyz_to_plane_mm_dir(dir_xyz, plane):
    """Proyecta XYZ unitario al plano Sketch → (dx, dy) unitario en mm-space."""
    if dir_xyz is None or plane is None:
        return (1.0, 0.0)
    try:
        xv = plane.XVec
        yv = plane.YVec
        dx = (
            float(dir_xyz.X) * float(xv.X)
            + float(dir_xyz.Y) * float(xv.Y)
            + float(dir_xyz.Z) * float(xv.Z)
        )
        dy = (
            float(dir_xyz.X) * float(yv.X)
            + float(dir_xyz.Y) * float(yv.Y)
            + float(dir_xyz.Z) * float(yv.Z)
        )
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return (1.0, 0.0)
        return (dx / L, dy / L)
    except Exception:
        return (1.0, 0.0)


def _layer_cfg_for_keys(layer_cfg, keys, spacing_mm=None):
    """Copia de layer_cfg con solo ``keys`` activas (opcional spacing forzado)."""
    key_set = set(keys or [])
    out = {}
    for key in _LAYER_KEYS:
        src = layer_cfg.get(key) or {}
        esp = int(src.get(u"spacing_mm") or 150)
        if spacing_mm is not None and key in key_set:
            esp = int(spacing_mm)
        out[key] = {
            u"active": key in key_set,
            u"bar_type_id": src.get(u"bar_type_id"),
            u"spacing_mm": esp,
        }
    return out


def _ahorro_layer_groups(layer_cfg):
    """
    Agrupa capas activas por (es_major, spacing_mm).

    Cada grupo genera 2 AreaReinforcement (Set A / Set B) con spacing 2e.
    Capas Major y Minor no comparten AR: el recorte es según el eje de barra.
    """
    groups = {}
    for key in _LAYER_KEYS:
        cfg = layer_cfg.get(key) or {}
        if not cfg.get(u"active"):
            continue
        is_major = key.endswith(u"_major")
        esp = int(cfg.get(u"spacing_mm") or 150)
        gkey = (is_major, esp)
        if gkey not in groups:
            groups[gkey] = []
        groups[gkey].append(key)
    # Orden estable: major antes que minor, luego spacing, luego keys
    items = []
    for (is_major, esp) in sorted(groups.keys(), key=lambda t: (0 if t[0] else 1, t[1])):
        items.append((is_major, esp, list(groups[(is_major, esp)])))
    return items


def _face_layer_partitions(layer_cfg, ahorro_superior, ahorro_inferior):
    """
    Particiones activas por cara (Inferior / Superior) con flag de ahorro.

    Returns:
        list of (face_keys, ahorro, tag) — solo caras con ≥1 capa activa.
    """
    faces = (
        (
            (u"interior_major", u"interior_minor"),
            bool(ahorro_inferior),
            u"Inf",
        ),
        (
            (u"exterior_major", u"exterior_minor"),
            bool(ahorro_superior),
            u"Sup",
        ),
    )
    out = []
    for keys, ahorro, tag in faces:
        face_keys = [
            k for k in keys if (layer_cfg.get(k) or {}).get(u"active")
        ]
        if face_keys:
            out.append((face_keys, ahorro, tag))
    return out


def _append_ahorro_create_jobs(
    create_jobs, pts_ar, layer_cfg, major_mm, minor_mm, label, errores
):
    """
    Añade jobs ahorro (2 AR @ 2e por grupo dir/e) a ``create_jobs``.

    Si el módulo no está disponible o no hay grupos, registra en ``errores``.
    """
    if ahorro_fierro_polygons_mm is None:
        errores.append(
            u"{0}: módulo ahorro de fierro no disponible.".format(label)
        )
        return
    groups = _ahorro_layer_groups(layer_cfg)
    if not groups:
        errores.append(
            u"{0}: no hay capas activas para ahorro.".format(label)
        )
        return
    for is_major, esp, keys in groups:
        bar_dir = major_mm if is_major else minor_mm
        sets = ahorro_fierro_polygons_mm(
            pts_ar,
            bar_dir[0],
            bar_dir[1],
            esp,
            cutback_pct=_AHORRO_CUTBACK_PCT,
        )
        if not sets:
            errores.append(
                u"{0}: no se pudo aplicar ahorro "
                u"(paño estrecho para e={1:g} / 2e={2:g} "
                u"/ recorte {3:g}% o luz dist. < 2e).".format(
                    label,
                    esp,
                    int(esp) * 2,
                    _AHORRO_CUTBACK_PCT,
                )
            )
            continue
        cfg_set = _layer_cfg_for_keys(
            layer_cfg, keys, spacing_mm=int(esp) * 2
        )
        dir_tag = u"Maj" if is_major else u"Min"
        for set_tag, pts_set in sets:
            create_jobs.append(
                {
                    u"pts": pts_set,
                    u"layer_cfg": cfg_set,
                    u"tx_label": u"{0} {1}-{2}".format(
                        label, dir_tag, set_tag
                    ),
                    u"pbar_label": u"{0} {1}-{2}".format(
                        label, dir_tag, set_tag
                    ),
                    u"ahorro": True,
                }
            )


def _build_pano_create_jobs(
    pts_ar,
    layer_cfg,
    major,
    plane,
    label,
    ahorro_superior,
    ahorro_inferior,
    errores,
):
    """
    Jobs Create de un paño según ahorro por cara.

    - Ambas caras activas sin ahorro → 1 AR con todas las capas activas.
    - Ambas caras activas con ahorro → jobs ahorro sobre todas las capas.
    - Flags distintos → jobs separados por cara (una puede ser ahorro y la otra no).
    - Cara OFF (sin capas activas) → se omite.
    """
    create_jobs = []
    partitions = _face_layer_partitions(
        layer_cfg, ahorro_superior, ahorro_inferior
    )
    if not partitions:
        return create_jobs

    major_mm = _xyz_to_plane_mm_dir(major, plane)
    minor_mm = (-major_mm[1], major_mm[0])

    any_ahorro = any(a for _, a, _ in partitions)
    all_ahorro = all(a for _, a, _ in partitions)

    if not any_ahorro:
        create_jobs.append(
            {
                u"pts": pts_ar,
                u"layer_cfg": layer_cfg,
                u"tx_label": label,
                u"pbar_label": label,
                u"ahorro": False,
            }
        )
        return create_jobs

    if all_ahorro and len(partitions) > 1:
        _append_ahorro_create_jobs(
            create_jobs,
            pts_ar,
            layer_cfg,
            major_mm,
            minor_mm,
            label,
            errores,
        )
        return create_jobs

    # Una sola cara con ahorro, o flags mixtos → por cara
    for face_keys, ahorro, tag in partitions:
        face_cfg = _layer_cfg_for_keys(layer_cfg, face_keys)
        face_label = (
            u"{0} {1}".format(label, tag)
            if len(partitions) > 1
            else label
        )
        if ahorro:
            _append_ahorro_create_jobs(
                create_jobs,
                pts_ar,
                face_cfg,
                major_mm,
                minor_mm,
                face_label,
                errores,
            )
        else:
            create_jobs.append(
                {
                    u"pts": pts_ar,
                    u"layer_cfg": face_cfg,
                    u"tx_label": face_label,
                    u"pbar_label": face_label,
                    u"ahorro": False,
                }
            )
    return create_jobs


def _estimate_pano_job_count(layer_cfg, ahorro_superior, ahorro_inferior):
    """Estimación de jobs Create por paño (progress bar)."""
    partitions = _face_layer_partitions(
        layer_cfg, ahorro_superior, ahorro_inferior
    )
    if not partitions:
        return 0
    any_ahorro = any(a for _, a, _ in partitions)
    all_ahorro = all(a for _, a, _ in partitions)
    if not any_ahorro:
        return 1
    if all_ahorro and len(partitions) > 1:
        return max(1, len(_ahorro_layer_groups(layer_cfg))) * 2
    n = 0
    for face_keys, ahorro, _tag in partitions:
        if ahorro:
            face_cfg = _layer_cfg_for_keys(layer_cfg, face_keys)
            n += max(1, len(_ahorro_layer_groups(face_cfg))) * 2
        else:
            n += 1
    return max(1, n)


def crear_area_reinforcement(
    doc, floor, curves, major_dir, layer_cfg, area_type_id=None, bars=None
):
    """Crea AR y aplica capas. Devuelve (area_rein, error_msg).

    ``area_type_id`` / ``bars`` opcionales: reutilizar lookups del batch Create.
    """
    if doc is None or floor is None or not curves:
        return None, u"Datos incompletos."
    if area_type_id is None:
        area_type_id = _default_area_type_id(doc)
    if area_type_id == ElementId.InvalidElementId:
        return None, u"No hay AreaReinforcementType en el proyecto."
    if bars is None:
        bars = _bar_types_sorted(doc)
    if not bars:
        return None, u"No hay RebarBarType en el proyecto."
    first_bar_id = bars[0][2].Id
    # Preferir diámetro de la primera capa activa
    for key in _LAYER_KEYS:
        cfg = layer_cfg.get(key) or {}
        if cfg.get(u"active") and cfg.get(u"bar_type_id"):
            first_bar_id = cfg[u"bar_type_id"]
            break
    try:
        curve_list = List[Curve](curves)
        ar = AreaReinforcement.Create(
            doc,
            floor,
            curve_list,
            major_dir,
            area_type_id,
            first_bar_id,
            ElementId.InvalidElementId,
        )
        _aplicar_capas(ar, layer_cfg)
        return ar, None
    except Exception as ex:
        return None, _as_unicode(ex)


def _active_view(doc, uidoc):
    view = None
    try:
        if uidoc is not None:
            view = uidoc.ActiveView
    except Exception:
        view = None
    if view is None and doc is not None:
        try:
            view = doc.ActiveView
        except Exception:
            view = None
    return view


def _post_create_area_reinforcement(
    doc,
    ar,
    uidoc,
    mra_avisos,
    mra_ok_total,
    mra_tipo_aviso_hecho,
    tag_avisos=None,
    tag_ok_total=0,
    tag_familia_aviso_hecho=False,
    tag_map=None,
    skip_regenerate=False,
    allow_retry_regenerate=True,
    out_free_rebars=None,
    pata_ctx=None,
    pata_avisos=None,
):
    """
    Tras Create: Regenerate → snapshots stamps → RemoveAreaSystem → Rebar libres
    → Show Middle → stamps → [patas L + re-stamp] → Unobscured → etiquetas → MRA.

    ``pata_ctx`` opcional: floor, plane, face, outline_pts, hole_rings,
    pano_pts, enabled. MRA exige ``Rebar`` libres (no ``RebarInSystem``).
    """
    if tag_avisos is None:
        tag_avisos = []
    if mra_avisos is None:
        mra_avisos = []
    if pata_avisos is None:
        pata_avisos = []
    view = _active_view(doc, uidoc)
    if not skip_regenerate:
        try:
            doc.Regenerate()
        except Exception:
            pass
    barras = _collect_barras_params_de_area_reinforcement(doc, ar)
    if not barras and allow_retry_regenerate:
        try:
            doc.Regenerate()
        except Exception:
            pass
        barras = _collect_barras_params_de_area_reinforcement(doc, ar)

    nivel_valor = None
    try:
        nivel_valor = _nivel_losa_area_reinforcement(doc, ar)
    except Exception:
        nivel_valor = None

    snaps = []
    try:
        snaps = _stamp_snapshots_from_barras(doc, ar, barras, view=view)
    except Exception:
        snaps = []

    conjunto_guid = None
    if iniciar_armadura_conjunto_guid_ejecucion is not None:
        try:
            conjunto_guid = iniciar_armadura_conjunto_guid_ejecucion()
        except Exception:
            conjunto_guid = None

    # Convertir a Rebar libres (requisito MRA).
    free_rebars = _remove_area_reinforcement_system(doc, ar)
    if not free_rebars:
        # Respaldo: seguir con barras de sistema (MRA probablemente 0).
        free_rebars = list(barras or [])
        try:
            mra_avisos.append(
                u"RemoveAreaReinforcementSystem no devolvió Rebar; "
                u"MRA puede no aplicarse a RebarInSystem."
            )
        except Exception:
            pass

    try:
        if view is not None:
            _aplicar_show_middle_barras(doc, free_rebars, view)
    except Exception:
        pass

    try:
        _stamp_armadura_params_en_rebars(
            doc,
            free_rebars,
            snaps=snaps,
            conjunto_guid=conjunto_guid,
            nivel_valor=nivel_valor,
            view=view,
        )
    except Exception:
        pass

    # Patas L: aristas paño ∩ (outline / shafts / huecos) (antes de tags/MRA)
    if (
        aplicar_patas_l_por_outline is not None
        and pata_ctx
        and bool(pata_ctx.get(u"enabled"))
        and (pata_ctx.get(u"pano_pts") or [])
        and (
            (pata_ctx.get(u"outline_pts") or [])
            or (pata_ctx.get(u"hole_rings") or [])
        )
    ):
        try:
            free_rebars = aplicar_patas_l_por_outline(
                doc,
                free_rebars,
                pata_ctx.get(u"floor"),
                pata_ctx.get(u"plane"),
                pata_ctx.get(u"face"),
                pata_ctx.get(u"outline_pts"),
                avisos=pata_avisos,
                hole_rings=pata_ctx.get(u"hole_rings"),
                pano_pts=pata_ctx.get(u"pano_pts"),
            )
            # Barras nuevas: Show Middle + stamps (match por geometría)
            try:
                if view is not None:
                    _aplicar_show_middle_barras(doc, free_rebars, view)
            except Exception:
                pass
            try:
                _stamp_armadura_params_en_rebars(
                    doc,
                    free_rebars,
                    snaps=snaps,
                    conjunto_guid=conjunto_guid,
                    nivel_valor=nivel_valor,
                    view=view,
                )
            except Exception:
                pass
        except Exception as ex_pata:
            try:
                pata_avisos.append(
                    u"Pata L: {0}".format(_as_unicode(ex_pata))
                )
            except Exception:
                pass

    if out_free_rebars is not None:
        try:
            del out_free_rebars[:]
            out_free_rebars.extend(free_rebars)
        except Exception:
            try:
                out_free_rebars.extend(free_rebars)
            except Exception:
                pass

    try:
        if view is not None:
            apply_reinforcement_unobscured_in_view(
                doc, free_rebars, view, unobscured=True
            )
    except Exception:
        pass

    try:
        if view is not None:
            avisos_tag = []
            n_tag = _aplicar_etiquetas_show_middle_barras(
                doc, view, free_rebars, avisos_tag, tag_map=tag_map
            )
            if n_tag > 0:
                tag_ok_total = int(tag_ok_total) + int(n_tag)
            for av in avisos_tag or []:
                if (
                    u"no hay tipos de familia" in (av or u"")
                    and tag_familia_aviso_hecho
                ):
                    continue
                if u"no hay tipos de familia" in (av or u""):
                    tag_familia_aviso_hecho = True
                tag_avisos.append(av)
    except Exception as ex_tag:
        try:
            tag_avisos.append(
                u"Etiqueta rebar: {0}".format(_as_unicode(ex_tag))
            )
        except Exception:
            pass

    try:
        if view is not None:
            avisos_mra = []
            n_mra = _aplicar_mra_rebars(doc, view, free_rebars, avisos_mra)
            if n_mra > 0:
                mra_ok_total = int(mra_ok_total) + int(n_mra)
            for av in avisos_mra or []:
                if u"no existe el tipo" in (av or u"") and mra_tipo_aviso_hecho:
                    continue
                if u"no existe el tipo" in (av or u""):
                    mra_tipo_aviso_hecho = True
                mra_avisos.append(av)
    except Exception as ex_mra:
        try:
            mra_avisos.append(u"MRA: {0}".format(_as_unicode(ex_mra)))
        except Exception:
            pass

    return (
        mra_ok_total,
        mra_tipo_aviso_hecho,
        tag_ok_total,
        tag_familia_aviso_hecho,
    )


# ---------------------------------------------------------------------------
# Selección Floor
# ---------------------------------------------------------------------------


class _FloorSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, Floor)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return True


def _pick_floor(uidoc, doc, uiapp):
    # Preselección
    try:
        ids = list(uidoc.Selection.GetElementIds())
        for eid in ids:
            el = doc.GetElement(eid)
            if isinstance(el, Floor):
                return el
    except Exception:
        pass
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            _FloorSelectionFilter(),
            u"Seleccione una losa (Floor)",
        )
        if ref is None:
            return None
        el = doc.GetElement(ref.ElementId)
        if isinstance(el, Floor):
            return el
    except Exception:
        return None
    _mostrar_aviso(uiapp, u"El elemento seleccionado no es una losa (Floor).")
    return None


# ---------------------------------------------------------------------------
# Pintado de contexto de planta (compartido con Suples losa)
# ---------------------------------------------------------------------------


def paint_planta_context_layers(
    scene,
    hud,
    to_px,
    add_polygon,
    loop_polylines,
    overlays,
    existing_ars,
    curves_outer,
    plane,
    ctx_geo_cache,
    ox0,
    oy0,
    min_x,
    max_x,
    min_y,
    max_y,
    fit_scale,
    bw,
    bh,
    cw,
    ch,
    major_xyz=None,
    mid_layer_callback=None,
    header_tb=None,
    header_text=None,
    context_line_scale=1.0,
):
    """
    Pinta capas de contexto de planta (mismo orden visual que Area Rein. Sketch):

    grid 1 m → losa/huecos → AR existentes → vigas → muros → pasadas →
    [mid_layer encima] → header → leyenda HUD.

    ``mid_layer_callback(scene, to_px, add_polygon[, sw_fn])`` opcional
    (p. ej. paños). Se pinta **después** de muros/vigas/pasadas para quedar
    por encima. ``sw_fn(base)`` es el mismo escalado de grosor del contexto.
    ``ctx_geo_cache``: dict con wall_geo/beam_geo/wall_pts/beam_pts (o None).
    ``context_line_scale``: multiplica el grosor de losa/huecos/muros/vigas/pasadas
    (1.0 = grosor actual del sketch; p. ej. 0.5 en Suples).
    """
    try:
        _cls = float(context_line_scale)
    except Exception:
        _cls = 1.0
    if _cls < 0.15:
        _cls = 0.15
    if _cls > 3.0:
        _cls = 3.0

    def _sw(base):
        try:
            v = float(base) * _cls
        except Exception:
            v = float(base)
        return max(0.35, v)

    # Grid 1 m
    try:
        x0 = int(math.floor(min_x / 1000.0)) * 1000
        y0 = int(math.floor(min_y / 1000.0)) * 1000
        x1 = int(math.ceil(max_x / 1000.0)) * 1000
        y1 = int(math.ceil(max_y / 1000.0)) * 1000
        gx = x0
        while gx <= x1:
            px0, py0 = to_px(gx, min_y)
            px1, py1 = to_px(gx, max_y)
            ln = WpfLine()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = px0, py0, px1, py1
            ln.Stroke = _brush(u"#21465C", 90)
            ln.StrokeThickness = 0.5
            scene.Children.Add(ln)
            gx += 1000
        gy = y0
        while gy <= y1:
            px0, py0 = to_px(min_x, gy)
            px1, py1 = to_px(max_x, gy)
            ln = WpfLine()
            ln.X1, ln.Y1, ln.X2, ln.Y2 = px0, py0, px1, py1
            ln.Stroke = _brush(u"#21465C", 90)
            ln.StrokeThickness = 0.5
            scene.Children.Add(ln)
            gy += 1000
    except Exception:
        pass

    # Contorno losa + huecos Sketch
    if loop_polylines:
        add_polygon(
            loop_polylines[0],
            u"#1a3544",
            u"#95B8CC",
            stroke_w=_sw(1.5),
            fill_a=220,
        )
        for hole in loop_polylines[1:]:
            add_polygon(
                hole,
                u"#071018",
                u"#95B8CC",
                stroke_w=_sw(1.2),
                dashed=False,
                fill_a=255,
            )

    # AreaReinforcement existentes (solo visualización)
    try:
        for ar in existing_ars or []:
            loops = ar.get(u"loops") or []
            if not loops and ar.get(u"pts"):
                loops = [ar.get(u"pts")]
            if not loops:
                continue
            for hi, ring in enumerate(loops):
                if not ring or len(ring) < 3:
                    continue
                if hi == 0:
                    add_polygon(
                        ring,
                        _EXISTING_AR_FILL,
                        _EXISTING_AR_STROKE,
                        stroke_w=1.6,
                        dashed=True,
                        fill_a=70,
                    )
                else:
                    add_polygon(
                        ring,
                        u"#071018",
                        _EXISTING_AR_STROKE,
                        stroke_w=1.2,
                        dashed=True,
                        fill_a=180,
                    )
            try:
                pts0 = loops[0]
                cx = sum(q[0] for q in pts0) / float(len(pts0))
                cy = sum(q[1] for q in pts0) / float(len(pts0))
                px, py = to_px(cx, cy)
                tb = TextBlock()
                tb.Text = ar.get(u"label") or u"AR"
                tb.Foreground = _brush(_EXISTING_AR_LABEL)
                tb.FontSize = 9
                tb.FontWeight = FontWeights.SemiBold
                WpfCanvas.SetLeft(tb, px - 22)
                WpfCanvas.SetTop(tb, py - 8)
                scene.Children.Add(tb)
            except Exception:
                pass
    except Exception:
        pass

    # Vigas/muros: PathGeometry en coords fit
    fill_w, stroke_w_col = _CTX_COLORS[_CTX_WALL]
    fill_b, stroke_b = _CTX_COLORS[_CTX_BEAM]
    cache = ctx_geo_cache
    if cache is None:
        wall_geo, beam_geo, wall_pts_list, beam_pts_list = _build_wall_beam_geo_mm(
            overlays
        )
        cache = {
            u"wall_geo": wall_geo,
            u"beam_geo": beam_geo,
            u"wall_pts": wall_pts_list,
            u"beam_pts": beam_pts_list,
        }
    wall_pts_list = cache.get(u"wall_pts") or []
    beam_pts_list = cache.get(u"beam_pts") or []
    try:
        mtx = _view_matrix_mm_to_px(ox0, oy0, min_x, max_y, fit_scale)
    except Exception:
        mtx = None
    wall_geo = None
    beam_geo = None
    if mtx is not None:
        try:
            wall_geo = _geo_transformed_for_view(cache.get(u"wall_geo"), mtx)
        except Exception:
            wall_geo = None
        try:
            beam_geo = _geo_transformed_for_view(cache.get(u"beam_geo"), mtx)
        except Exception:
            beam_geo = None
        for g in (wall_geo, beam_geo):
            if g is None:
                continue
            try:
                if not g.IsFrozen:
                    g.Freeze()
            except Exception:
                pass
    try:
        if beam_geo is not None:
            _add_path_geometry(
                scene, beam_geo, fill_b, stroke_b, stroke_w=_sw(1.4), fill_a=220
            )
        else:
            for pts in beam_pts_list:
                add_polygon(pts, fill_b, stroke_b, stroke_w=_sw(1.2), fill_a=220)
    except Exception:
        for pts in beam_pts_list:
            add_polygon(pts, fill_b, stroke_b, stroke_w=_sw(1.2), fill_a=220)
    try:
        if wall_geo is not None:
            _add_path_geometry(
                scene, wall_geo, fill_w, stroke_w_col, stroke_w=_sw(1.4), fill_a=255
            )
        else:
            for pts in wall_pts_list:
                add_polygon(pts, fill_w, stroke_w_col, stroke_w=_sw(1.2), fill_a=255)
    except Exception:
        for pts in wall_pts_list:
            add_polygon(pts, fill_w, stroke_w_col, stroke_w=_sw(1.2), fill_a=255)

    # Pasadas / shafts (mismo color que borde losa; sin dash)
    fill_p, stroke_p = _CTX_COLORS[_CTX_PASADA]
    for ov in overlays or []:
        if ov.get(u"kind") != _CTX_PASADA:
            continue
        add_polygon(
            ov.get(u"pts"),
            fill_p,
            stroke_p,
            stroke_w=_sw(1.2),
            dashed=False,
            fill_a=230,
        )

    # Paños / mid-layer al final → por encima de muros, vigas y pasadas
    if mid_layer_callback is not None:
        try:
            mid_layer_callback(scene, to_px, add_polygon, _sw)
        except TypeError:
            try:
                mid_layer_callback(scene, to_px, add_polygon)
            except Exception:
                pass
        except Exception:
            pass

    if header_tb is not None and header_text is not None:
        try:
            header_tb.Text = _as_unicode(header_text)
        except Exception:
            pass

    # Leyenda HUD
    try:
        lx, ly = cw - 118.0, 10.0
        items = (
            (u"Muro", _CTX_COLORS[_CTX_WALL][1], False),
            (u"Viga", _CTX_COLORS[_CTX_BEAM][1], False),
            (u"Pasada", _CTX_COLORS[_CTX_PASADA][1], False),
            (u"AR exist.", _EXISTING_AR_STROKE, True),
            (u"Losa", u"#95B8CC", False),
        )
        for i, (lab, col, dashed) in enumerate(items):
            yy = ly + i * 16.0
            sw = WpfLine()
            sw.X1, sw.Y1, sw.X2, sw.Y2 = lx, yy + 6, lx + 14, yy + 6
            sw.Stroke = _brush(col)
            sw.StrokeThickness = 3
            sw.IsHitTestVisible = False
            if dashed:
                try:
                    dashes = DoubleCollection()
                    dashes.Add(3)
                    dashes.Add(2)
                    sw.StrokeDashArray = dashes
                except Exception:
                    pass
            hud.Children.Add(sw)
            tb = TextBlock()
            tb.Text = lab
            tb.Foreground = _brush(u"#95B8CC")
            tb.FontSize = 10
            tb.IsHitTestVisible = False
            WpfCanvas.SetLeft(tb, lx + 18)
            WpfCanvas.SetTop(tb, yy)
            hud.Children.Add(tb)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# ProgressBar (pyRevit — mismo patrón que Armado vigas / Numerar marcas)
# ---------------------------------------------------------------------------


def _pbar_enabled():
    try:
        from pyrevit import forms as _forms  # noqa: F401
    except Exception:
        return False
    return True


class _AreaReinLosaCrearProgress(object):
    """Context manager no-op si ``pyrevit.forms.ProgressBar`` no está disponible."""

    def __init__(self, total, title_prefix=None):
        self._total = max(1, int(total or 1))
        self._pb = None
        self._open = False
        self._title_prefix = title_prefix or _DIALOG_TITLE

    def __enter__(self):
        if not _pbar_enabled():
            return self
        try:
            from pyrevit import forms as _pyrevit_forms

            self._pb = _pyrevit_forms.ProgressBar(
                title=self._title(0),
                cancellable=False,
            )
            try:
                from System.Windows.Media import Color, SolidColorBrush

                self._pb.Resources[u"pyRevitAccentBrush"] = SolidColorBrush(
                    Color.FromRgb(91, 192, 222),
                )
            except Exception:
                pass
            self._pb.__enter__()
            self._open = True
        except Exception:
            self._pb = None
            self._open = False
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._open and self._pb is not None:
            try:
                self._pb.__exit__(exc_type, exc_val, exc_tb)
            except Exception:
                pass
        self._open = False
        self._pb = None
        return False

    def _title(self, current):
        # current 0 = arranque; 1..N = paño en curso
        cur = max(0, int(current))
        if cur < 1:
            return u"{0} — Creando 0/{1}…".format(
                self._title_prefix, int(self._total)
            )
        return u"{0} — Creando {1}/{2}…".format(
            self._title_prefix, cur, int(self._total)
        )

    def update(self, current, label=None):
        """Actualiza barra al paño *current* (1-based) en curso."""
        if self._pb is None:
            return
        c = max(1, min(int(current), int(self._total)))
        base = self._title(c)
        if label:
            base = u"{0} ({1})".format(base, _as_unicode(label))
        try:
            if hasattr(self._pb, u"update_progress"):
                try:
                    self._pb.update_progress(c, max_value=self._total)
                except TypeError:
                    try:
                        self._pb.update_progress(c, max=self._total)
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            self._pb.title = base
        except Exception:
            pass


def _set_wait_cursor(on):
    """Cursor de espera estilo Windows hasta que la UI esté lista para mostrar."""
    try:
        from System.Windows.Input import Mouse

        Mouse.OverrideCursor = Cursors.Wait if on else None
    except Exception:
        pass
    try:
        clr.AddReference(u"System.Windows.Forms")
        from System.Windows.Forms import (
            Application as WinFormsApp,
            Cursor as WinFormsCursor,
            Cursors as WinFormsCursors,
        )

        WinFormsApp.UseWaitCursor = bool(on)
        WinFormsCursor.Current = (
            WinFormsCursors.WaitCursor if on else WinFormsCursors.Default
        )
    except Exception:
        pass


# ---------------------------------------------------------------------------
# ExternalEvent
# ---------------------------------------------------------------------------


class _CrearHandler(IExternalEventHandler):
    def __init__(self, ctrl_ref):
        self._ctrl_ref = ctrl_ref
        # Strong refs while ExternalEvent is queued / running (weakref alone
        # is not enough after the WPF window closes).
        self._pending_ctrl = None
        self._pending_request = None

    def GetName(self):
        return u"AraincoAreaReinLosaSketchCrear"

    def Execute(self, app):
        ctrl = self._pending_ctrl
        if ctrl is None and self._ctrl_ref:
            ctrl = self._ctrl_ref()
        req = self._pending_request
        if ctrl is None and req is None:
            return
        try:
            if ctrl is not None:
                ctrl._execute_crear(req)
            else:
                uiapp = req.get(u"uiapp") if req else None
                if uiapp is not None:
                    _mostrar_aviso(
                        uiapp,
                        u"No se pudo crear AreaReinforcement.",
                        content=u"Controlador no disponible.",
                    )
        except Exception as ex:
            uiapp = None
            if req:
                uiapp = req.get(u"uiapp")
            if uiapp is None and ctrl is not None:
                uiapp = getattr(ctrl, u"_uiapp", None)
            if uiapp is not None:
                _mostrar_aviso(
                    uiapp,
                    u"Error al crear AreaReinforcement.",
                    content=_as_unicode(ex),
                )
        finally:
            self._pending_ctrl = None
            self._pending_request = None
            if ctrl is not None:
                try:
                    ctrl._crear_request = None
                    ctrl._dispose_crear_event()
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

_XAML = u"""
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__CHROME__"
  Height="720" Width="980"
  MinHeight="640" MinWidth="900"
  ResizeMode="CanResize"
  WindowStartupLocation="Manual"
  Background="#071018"
  FontFamily="Segoe UI"
  FontSize="12"
  ShowInTaskbar="False">
  <Window.Resources>
__STYLES__
  </Window.Resources>
  <Border Background="#071018" BorderBrush="#21465C" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <StackPanel Grid.Row="0" Margin="0,0,0,8">
        <TextBlock x:Name="TxtTitle" Text="Arainco: Area Rein. losa"
                   Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC"
                   FontSize="11" TextWrapping="Wrap"
                   Text="Contorno desde Sketch del Floor · planta a escala real."/>
      </StackPanel>

      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="360"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="#0a1620" BorderBrush="#21465C"
                BorderThickness="1" CornerRadius="4,0,0,4" Padding="0">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#0a1620" BorderBrush="#21465C"
                    BorderThickness="0,0,0,1" Padding="8,6,8,4">
              <TextBlock x:Name="TxtCanvasHeader" Foreground="#64748b"
                         FontSize="10" FontWeight="SemiBold"
                         Text="PLANTA · SKETCH (mm)"/>
            </Border>
            <Border Grid.Row="1" Background="#050E18" BorderBrush="Transparent"
                    BorderThickness="0" Padding="8,4,8,8">
              <Border Background="#050E18" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4">
                <Canvas x:Name="CvPlan" ClipToBounds="True"/>
              </Border>
            </Border>
          </Grid>
        </Border>

        <Border Grid.Column="1" Background="#0a1620" BorderBrush="#21465C"
                BorderThickness="1" CornerRadius="0,4,4,0" Padding="8,8">
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="PnlSectionRail">

              <Border Background="#0a1620" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,10">
                <StackPanel>
                  <TextBlock Text="Capas Area Reinforcement" Foreground="#E8F4F8"
                             FontSize="12" FontWeight="SemiBold"
                             Margin="0,0,0,6" TextWrapping="NoWrap"/>
                  <TextBlock x:Name="TxtLayersHint" Foreground="#64748b" FontSize="10"
                             Margin="0,0,0,8" TextWrapping="Wrap"
                             Text="Cree o seleccione un paño en la planta para editar mallas."/>
                  <StackPanel x:Name="PnlLayers"/>
                </StackPanel>
              </Border>

            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <TextBlock Grid.Row="2" x:Name="TxtHint" Foreground="#64748b" FontSize="10"
                 TextWrapping="Wrap" Margin="0,8,0,0"
                 Text="Tabs Superior/Inferior: paños por cara (2 clics). Sin polígonos en una cara → no se crea AR ahí. Clic = activo · Principal = luz menor · Ctrl+clic fusión · Supr elimina · Esc cancela."/>

      <Grid Grid.Row="3" Margin="0,14,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnManual" Grid.Column="0" Content="Manual"
                Style="{StaticResource BtnSelectOutline}" MinWidth="96"
                Margin="0,0,12,0" ToolTip="Abrir manual de usuario"
                VerticalAlignment="Center"/>
        <TextBlock x:Name="TxtStatus" Grid.Column="1" VerticalAlignment="Center"
                   Foreground="#64748b" FontSize="10" TextWrapping="Wrap" Margin="0,0,12,0"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button x:Name="BtnCancelar" Content="Cancelar"
                  Style="{StaticResource BtnSelectOutline}" MinWidth="110" Margin="0,0,10,0"/>
          <Button x:Name="BtnCrear" Content="Crear Area Reinf."
                  Style="{StaticResource BtnPrimary}" MinWidth="180"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
""".replace(u"__CHROME__", WINDOW_CHROME_TITLE).replace(
    u"__STYLES__", BIMTOOLS_DARK_STYLES_XML
)


class AreaReinLosaSketchController(object):
    def __init__(self, uiapp, uidoc, doc, floor, curves_outer, loops_all, plane):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._floor = floor
        self._curves = list(curves_outer)
        self._loops = loops_all or [curves_outer]
        self._plane = plane
        self._bar_types = _bar_types_sorted(doc)
        self._layer_ui = {}
        self._face_ui = {}
        self._active_face = u"inferior"
        self._win = None
        self._crear_event = None
        self._crear_request = None  # snapshot para ExternalEvent (sin WPF)
        # Contexto completo antes del reveal (muros/vigas/pasadas/ejes snap).
        overlays, walls_list, beams_list = recolectar_contexto_planta(
            doc, floor, plane, include_grids=True
        )
        self._overlays = overlays or []
        self._walls_list = walls_list or []
        self._beams_list = beams_list or []
        self._ctx_loaded = True
        self._ctx_load_scheduled = True
        self._ctx_grids_loaded = True
        self._ctx_grids_scheduled = True
        self._sketch_holes = max(0, len(self._loops) - 1)
        self._panos = []
        self._pano_selected = set()
        self._pano_merge = set()  # ids para fusión (Ctrl+clic)
        self._active_pano_id = None  # paño cuyas cards se editan
        self._ui_syncing = False  # True al empujar settings → UI (sin writeback)
        self._pano_seq = 0
        self._pano_seq_by_face = {u"superior": 0, u"inferior": 0}
        self._borde_inf_ui = {}  # card Barras en borde (Inferior) — UI only
        self._pick_pt1 = None  # (x_mm, y_mm) o None
        self._view_xform = None  # dict para clic→mm
        # Zoom/pan de usuario (persisten entre redraw; zoom=1 + pan=0 = fit)
        self._view_zoom = 1.0
        self._view_pan_x = 0.0  # mm: desplazamiento del centro de vista vs bbox
        self._view_pan_y = 0.0
        self._panning = False
        self._pan_last_x = 0.0  # px canvas durante arrastre botón medio
        self._pan_last_y = 0.0
        self._cursor_ctrl = False  # último estado Ctrl aplicado al cursor del canvas
        self._snap_verts = []
        self._snap_segs = []
        self._snap_cell_index = None
        self._snap_geo_dirty = True
        self._hover_snap = None  # (x_mm, y_mm, kind) o None
        # Cache geometría muros/vigas en mm (CombinedGeometry solo al crear)
        self._ctx_geo_cache = None
        # Polylines Sketch en mm (loops); invalidar solo si cambia contexto losa
        self._sketch_loop_polylines_mm = None
        # Refs WPF nombradas (FindName una vez tras load)
        self._ui_cv_plan = None
        self._ui_txt_canvas_header = None
        self._last_canvas_cw = 0.0
        self._last_canvas_ch = 0.0
        # Capas canvas: escena (mm→px fit) + HUD; zoom/pan = MatrixTransform
        self._scene_layer = None
        self._hud_layer = None
        self._scene_base = None  # ox0/oy0/fit_scale/cw/ch/bbox
        self._scene_matrix_transform = None
        # Fallback coalesce si aún no hay capa de escena
        self._view_redraw_timer = None
        self._view_redraw_pending = False
        # SizeChanged coalescido (~12 ms); full redraw solo si cw/ch cambian
        self._size_redraw_timer = None
        self._size_redraw_pending = False
        self._ui_revealed = False  # True tras el primer Show() con UI lista
        self._ui_prepare_done = False  # contenido montado antes de Show()
        self._existing_ars = collect_existing_area_rein_on_floor(doc, floor, plane)
        self._build_sketch_polylines_cache()

        self._win = XamlReader.Parse(_XAML)
        # No Show() aquí: la ventana solo se levanta en show() con el canvas ya pintado.
        self._wire()
        self._cache_ui_refs()
        self._build_layer_panels()
        self._sync_cards_from_active()
        self._update_create_button()
        n_ar = len(self._existing_ars or [])
        if n_ar > 0:
            self._set_status(
                u"Floor Id {0} · {1} Area Rein. existentes · dibuje paños: 2 clics.".format(
                    _element_id_int(floor.Id), n_ar
                )
            )
        else:
            self._set_status(
                u"Floor Id {0} · dibuje paños: 2 clics en el canvas (esquinas opuestas).".format(
                    _element_id_int(floor.Id)
                )
            )

        self._handler = _CrearHandler(weakref.ref(self))
        self._crear_event = ExternalEvent.Create(self._handler)

        def _on_closed(sender, args):
            try:
                self._stop_view_redraw_timer()
            except Exception:
                pass
            try:
                self._stop_size_redraw_timer()
            except Exception:
                pass
            _unregister_singleton()
            self._win = None
            self._ui_cv_plan = None
            self._ui_txt_canvas_header = None
            self._ui_revealed = False
            # Creación pendiente: no Dispose del ExternalEvent hasta Execute.
            if self._crear_request is not None:
                return
            self._dispose_crear_event()

        self._win.Closed += EventHandler(_on_closed)

        def _on_size(sender, args):
            self._schedule_size_redraw()

        def _on_loaded(sender, args):
            # Contenido ya se monta en show() antes de Show(); solo fallback.
            if not getattr(self, u"_ui_prepare_done", False):
                self._prepare_ui_content()

        self._win.SizeChanged += SizeChangedEventHandler(_on_size)
        self._win.Loaded += RoutedEventHandler(_on_loaded)

    def _wire(self):
        win = self._win
        btn_c = win.FindName(u"BtnCancelar")
        btn_ok = win.FindName(u"BtnCrear")
        btn_man = win.FindName(u"BtnManual")
        cv = win.FindName(u"CvPlan")
        self._ui_cv_plan = cv
        try:
            self._ui_txt_canvas_header = win.FindName(u"TxtCanvasHeader")
        except Exception:
            self._ui_txt_canvas_header = None

        def _cancel(s, e):
            try:
                win.Close()
            except Exception:
                pass

        def _manual(s, e):
            self._open_manual()

        def _crear(s, e):
            req, err = self._snapshot_crear_request()
            if err:
                _mostrar_aviso(self._uiapp, err)
                return
            self._crear_request = req
            try:
                self._handler._pending_ctrl = self
                self._handler._pending_request = req
            except Exception:
                pass
            ev = self._crear_event
            if ev is None:
                self._crear_request = None
                try:
                    self._handler._pending_ctrl = None
                    self._handler._pending_request = None
                except Exception:
                    pass
                _mostrar_aviso(
                    self._uiapp,
                    u"No se pudo iniciar la creación.",
                    content=u"ExternalEvent no disponible.",
                )
                return
            try:
                ev.Raise()
            except Exception as ex:
                self._crear_request = None
                try:
                    self._handler._pending_ctrl = None
                    self._handler._pending_request = None
                except Exception:
                    pass
                _mostrar_aviso(
                    self._uiapp,
                    u"No se pudo iniciar la creación.",
                    content=_as_unicode(ex),
                )
                return
            # Cerrar UI de inmediato; Execute usa el snapshot (sin controles WPF).
            try:
                win.Close()
            except Exception:
                pass

        def _on_key(s, e):
            try:
                key = e.Key
            except Exception:
                return
            # Actualizar cursor flecha+plus al pulsar Ctrl
            try:
                if key == Key.LeftCtrl or key == Key.RightCtrl:
                    self._update_canvas_cursor()
            except Exception:
                pass
            # Ctrl+0: reset zoom/pan al fit
            try:
                if key == Key.D0 or key == Key.NumPad0:
                    mods = Keyboard.Modifiers
                    if (mods & ModifierKeys.Control) == ModifierKeys.Control:
                        self._reset_canvas_view()
                        try:
                            e.Handled = True
                        except Exception:
                            pass
                        return
            except Exception:
                pass
            # Supr/Delete (y Back): elimina paños en selección Ctrl (_pano_merge).
            # Prioridad sobre fusión: no auto-merge en esta tecla.
            try:
                is_delete = (key == Key.Delete)
                try:
                    if key == Key.Back:
                        is_delete = True
                except Exception:
                    pass
            except Exception:
                is_delete = False
            if is_delete:
                ids = list(self._pano_merge or [])
                if ids:
                    self._remove_panos(ids)
                    try:
                        e.Handled = True
                    except Exception:
                        pass
                return
            try:
                if key != Key.Escape:
                    return
            except Exception:
                return
            # Esc: limpia fusión / cancela paño en curso
            cleared_merge = bool(self._pano_merge)
            if cleared_merge:
                self._pano_merge.clear()
            if self._pick_pt1 is not None or self._hover_snap is not None:
                self._cancel_current_pano_pick()
                try:
                    e.Handled = True
                except Exception:
                    pass
            elif cleared_merge:
                self._set_status(u"Selección de fusión cancelada.")
                self._redraw_canvas()
                try:
                    e.Handled = True
                except Exception:
                    pass

        def _on_key_up(s, e):
            try:
                key = e.Key
            except Exception:
                return
            try:
                if key == Key.LeftCtrl or key == Key.RightCtrl:
                    self._update_canvas_cursor()
            except Exception:
                pass

        def _on_lost_focus(s, e):
            try:
                self._update_canvas_cursor()
            except Exception:
                pass

        def _canvas_click(s, e):
            self._on_canvas_click(s, e)

        def _canvas_move(s, e):
            self._on_canvas_move(s, e)

        def _canvas_wheel(s, e):
            self._on_canvas_wheel(s, e)

        def _canvas_down(s, e):
            self._on_canvas_mouse_down(s, e)

        def _canvas_up(s, e):
            self._on_canvas_mouse_up(s, e)

        def _canvas_lost_cap(s, e):
            self._end_canvas_pan(restore_cursor=True)

        if btn_man is not None:
            btn_man.Click += RoutedEventHandler(_manual)
        if btn_c is not None:
            btn_c.Click += RoutedEventHandler(_cancel)
        if btn_ok is not None:
            btn_ok.Click += RoutedEventHandler(_crear)
        try:
            win.Focusable = True
            win.PreviewKeyDown += KeyEventHandler(_on_key)
            win.PreviewKeyUp += KeyEventHandler(_on_key_up)
            win.LostKeyboardFocus += KeyboardFocusChangedEventHandler(
                _on_lost_focus
            )
        except Exception:
            try:
                win.KeyDown += KeyEventHandler(_on_key)
                win.KeyUp += KeyEventHandler(_on_key_up)
            except Exception:
                pass
        if cv is not None:
            try:
                cv.Focusable = True
                # Sin Background el Canvas no recibe hit-test en zonas vacías
                cv.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
                self._update_canvas_cursor()
            except Exception:
                pass
            try:
                cv.PreviewMouseLeftButtonDown += MouseButtonEventHandler(
                    _canvas_click
                )
            except Exception:
                try:
                    cv.MouseLeftButtonDown += MouseButtonEventHandler(_canvas_click)
                except Exception:
                    pass
            try:
                cv.PreviewMouseDown += MouseButtonEventHandler(_canvas_down)
            except Exception:
                try:
                    cv.MouseDown += MouseButtonEventHandler(_canvas_down)
                except Exception:
                    pass
            try:
                cv.PreviewMouseUp += MouseButtonEventHandler(_canvas_up)
            except Exception:
                try:
                    cv.MouseUp += MouseButtonEventHandler(_canvas_up)
                except Exception:
                    pass
            try:
                cv.LostMouseCapture += MouseEventHandler(_canvas_lost_cap)
            except Exception:
                pass
            try:
                cv.MouseMove += MouseEventHandler(_canvas_move)
            except Exception:
                pass
            try:
                cv.PreviewMouseWheel += MouseWheelEventHandler(_canvas_wheel)
            except Exception:
                try:
                    cv.MouseWheel += MouseWheelEventHandler(_canvas_wheel)
                except Exception:
                    pass
            try:
                cv.PreviewKeyDown += KeyEventHandler(_on_key)
                cv.PreviewKeyUp += KeyEventHandler(_on_key_up)
            except Exception:
                pass

        try:
            fid = _element_id_int(self._floor.Id)
            sub = win.FindName(u"TxtSubtitle")
            if sub is not None:
                sub.Text = (
                    u"Floor Id {0} · paños por 2 puntos en canvas · "
                    u"planta a escala real."
                ).format(fid)
            hint = win.FindName(u"TxtHint")
            if hint is not None:
                _ahorro_hint = (
                    u"Ahorro / Ø / esp. por paño. "
                    if _FEATURE_AHORRO_FIERRO
                    else u"Ø / esp. por paño. "
                )
                hint.Text = (
                    u"Tabs Superior / Inferior: dibuje paños en cada cara (2 clics). "
                    u"Sin polígonos en una cara → no se crea Area Rein. ahí. "
                    u"Clic en un paño lo activa. Principal = luz menor del paño. "
                    + _ahorro_hint
                    + u"Ctrl+clic fusión (misma cara) · Supr elimina · Esc cancela. "
                    u"Rueda = zoom · clic rueda = pan · Ctrl+0 = reset vista."
                )
        except Exception:
            pass

    def _stop_view_redraw_timer(self):
        t = getattr(self, u"_view_redraw_timer", None)
        self._view_redraw_timer = None
        self._view_redraw_pending = False
        if t is None:
            return
        try:
            t.Stop()
        except Exception:
            pass

    def _stop_size_redraw_timer(self):
        t = getattr(self, u"_size_redraw_timer", None)
        self._size_redraw_timer = None
        self._size_redraw_pending = False
        if t is None:
            return
        try:
            t.Stop()
        except Exception:
            pass

    def _resolve_manual_path(self):
        """Ruta a ``manual_usuario.html`` en la carpeta del pushbutton."""
        candidates = []
        try:
            import bimtools_paths

            pb = bimtools_paths.get_pushbutton_dir()
            if pb:
                candidates.append(os.path.join(pb, u"manual_usuario.html"))
        except Exception:
            pass
        # Fallback: *.tab/Armadura.panel/*AreaReinLosaSketch*.pushbutton/
        try:
            ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            for tab_name in os.listdir(ext_dir):
                if not tab_name.endswith(u".tab"):
                    continue
                panel = os.path.join(ext_dir, tab_name, u"Armadura.panel")
                if not os.path.isdir(panel):
                    continue
                for pb_name in os.listdir(panel):
                    if u"AreaReinLosaSketch" not in pb_name:
                        continue
                    candidates.append(
                        os.path.join(panel, pb_name, u"manual_usuario.html")
                    )
        except Exception:
            pass
        seen = set()
        for path in candidates:
            try:
                ap = os.path.normpath(os.path.abspath(path))
            except Exception:
                continue
            if ap in seen:
                continue
            seen.add(ap)
            if os.path.isfile(ap):
                return ap
        return None

    def _open_manual(self):
        """Abre el HTML del manual en el navegador / app asociada."""
        path = self._resolve_manual_path()
        if not path:
            _mostrar_aviso(
                self._uiapp,
                u"No se encontró manual_usuario.html.",
                content=u"Debe estar en la carpeta del pushbutton de la herramienta.",
            )
            return
        try:
            os.startfile(path)
        except Exception as ex:
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo abrir el manual.",
                content=_as_unicode(ex),
            )

    def _cache_ui_refs(self):
        """Cachea FindName de controles usados en hot paths."""
        win = self._win
        if win is None:
            return
        try:
            self._ui_cv_plan = win.FindName(u"CvPlan")
        except Exception:
            self._ui_cv_plan = None
        try:
            self._ui_txt_canvas_header = win.FindName(u"TxtCanvasHeader")
        except Exception:
            self._ui_txt_canvas_header = None

    def _get_cv_plan(self):
        cv = getattr(self, u"_ui_cv_plan", None)
        if cv is not None:
            return cv
        try:
            if self._win is not None:
                cv = self._win.FindName(u"CvPlan")
                self._ui_cv_plan = cv
                return cv
        except Exception:
            pass
        return None

    def _build_sketch_polylines_cache(self):
        """Polylines mm del Sketch (loops); fijo mientras no cambie el Floor."""
        plane = self._plane
        loops = []
        for loop in self._loops or []:
            pts = _loop_to_polyline_mm(loop, plane)
            if pts:
                loops.append(pts)
        self._sketch_loop_polylines_mm = loops
        return loops

    def _ensure_sketch_polylines_cache(self):
        cache = getattr(self, u"_sketch_loop_polylines_mm", None)
        if cache is not None:
            return cache
        return self._build_sketch_polylines_cache()

    def _mark_snap_geo_dirty(self):
        self._snap_geo_dirty = True

    def _ensure_snap_geometry(self):
        """Rebuild verts/segs/índice solo si paños / pick / overlays invalidaron."""
        if not getattr(self, u"_snap_geo_dirty", True):
            return
        try:
            self._rebuild_snap_geometry()
        except Exception:
            self._snap_verts = []
            self._snap_segs = []
            self._snap_cell_index = None
        self._snap_geo_dirty = False

    def _schedule_size_redraw(self):
        """Coalesce SizeChanged (~12 ms); full redraw solo si cw/ch cambian."""
        # En carga off-screen el prepare hace el único paint; SizeChanged sobra.
        if not getattr(self, u"_ui_revealed", False):
            return
        self._size_redraw_pending = True
        t = getattr(self, u"_size_redraw_timer", None)
        if t is not None:
            try:
                if not t.IsEnabled:
                    t.Start()
                return
            except Exception:
                self._size_redraw_timer = None
        try:
            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(12)

            def _tick(sender, args):
                try:
                    sender.Stop()
                except Exception:
                    pass
                if not getattr(self, u"_size_redraw_pending", False):
                    return
                self._flush_size_redraw()

            t.Tick += _tick
            self._size_redraw_timer = t
            t.Start()
        except Exception:
            self._flush_size_redraw()

    def _flush_size_redraw(self):
        self._size_redraw_pending = False
        cv = self._get_cv_plan()
        if cv is None:
            return
        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            return
        if cw < 40 or ch < 40:
            return
        last_cw = float(getattr(self, u"_last_canvas_cw", 0.0) or 0.0)
        last_ch = float(getattr(self, u"_last_canvas_ch", 0.0) or 0.0)
        if abs(cw - last_cw) < 0.5 and abs(ch - last_ch) < 0.5:
            return
        try:
            self._redraw_canvas()
        except Exception:
            pass

    def _overlay_host(self):
        """Canvas HUD (snap/escala) o CvPlan si aún no hay capas."""
        hud = getattr(self, u"_hud_layer", None)
        if hud is not None:
            return hud
        return self._get_cv_plan()

    def _update_hud_scale_bar(self, cw=None, ch=None, scale=None):
        """Barra de escala en pantalla (HUD); se actualiza en cada zoom/pan."""
        hud = getattr(self, u"_hud_layer", None)
        if hud is None:
            return
        xf = self._view_xform
        try:
            if cw is None:
                cw = float(xf[u"cw"]) if xf else 0.0
            if ch is None:
                ch = float(xf[u"ch"]) if xf else 0.0
            if scale is None:
                scale = float(xf[u"scale"]) if xf else 0.0
            cw = float(cw)
            ch = float(ch)
            scale = float(scale)
        except Exception:
            return
        if cw < 40 or ch < 40 or scale < 1e-12:
            return
        try:
            remove = []
            for child in list(hud.Children):
                try:
                    if getattr(child, u"Tag", None) == _HUD_SCALE_TAG:
                        remove.append(child)
                except Exception:
                    pass
            for child in remove:
                hud.Children.Remove(child)
        except Exception:
            pass
        try:
            bar_mm = 2000.0
            bar_px = bar_mm * scale
            sx, sy = 12.0, ch - 18.0
            sl = WpfLine()
            sl.X1, sl.Y1, sl.X2, sl.Y2 = sx, sy, sx + bar_px, sy
            sl.Stroke = _brush(u"#95B8CC")
            sl.StrokeThickness = 2
            sl.Tag = _HUD_SCALE_TAG
            sl.IsHitTestVisible = False
            hud.Children.Add(sl)
            for xx in (sx, sx + bar_px):
                t = WpfLine()
                t.X1, t.Y1, t.X2, t.Y2 = xx, sy - 4, xx, sy + 4
                t.Stroke = _brush(u"#95B8CC")
                t.StrokeThickness = 1.5
                t.Tag = _HUD_SCALE_TAG
                t.IsHitTestVisible = False
                hud.Children.Add(t)
            stb = TextBlock()
            stb.Text = u"2.00 m · {:.1f} px/m".format(scale * 1000.0)
            stb.Foreground = _brush(u"#64748b")
            stb.FontSize = 10
            stb.Tag = _HUD_SCALE_TAG
            stb.IsHitTestVisible = False
            WpfCanvas.SetLeft(stb, sx)
            WpfCanvas.SetTop(stb, sy - 18)
            hud.Children.Add(stb)
        except Exception:
            pass

    def _apply_scene_view_transform(self):
        """
        Aplica zoom/pan vía MatrixTransform sobre la capa de escena.
        No recrea children; actualiza _view_xform + barra de escala HUD.
        """
        scene = getattr(self, u"_scene_layer", None)
        base = getattr(self, u"_scene_base", None)
        if scene is None or base is None:
            return False
        cv = self._get_cv_plan()
        if cv is None:
            return False
        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            return False
        if abs(float(base.get(u"cw", 0.0)) - cw) > 0.5:
            return False
        if abs(float(base.get(u"ch", 0.0)) - ch) > 0.5:
            return False
        if not self._apply_view_to_xform(cw, ch):
            return False
        xf = self._view_xform
        try:
            zoom = float(self._view_zoom) if self._view_zoom else 1.0
            ox0 = float(base[u"ox0"])
            oy0 = float(base[u"oy0"])
            ox = float(xf[u"ox"])
            oy = float(xf[u"oy"])
            # px = zoom * px_fit + (ox - ox0*zoom)  (idem en Y)
            mtx = Matrix(
                zoom,
                0.0,
                0.0,
                zoom,
                ox - ox0 * zoom,
                oy - oy0 * zoom,
            )
            mt = getattr(self, u"_scene_matrix_transform", None)
            if mt is None:
                mt = MatrixTransform(mtx)
                self._scene_matrix_transform = mt
                scene.RenderTransform = mt
            else:
                mt.Matrix = mtx
                if scene.RenderTransform is not mt:
                    scene.RenderTransform = mt
        except Exception:
            return False
        try:
            self._update_hud_scale_bar(cw, ch, float(xf[u"scale"]))
        except Exception:
            pass
        return True

    def _try_apply_view_transform_only(self):
        """view_only: transform al instante; False → hace falta rebuild."""
        if not self._apply_scene_view_transform():
            return False
        if self._hover_snap is not None or self._pick_pt1 is not None:
            try:
                self._refresh_snap_overlay()
            except Exception:
                pass
        return True

    def _flush_view_redraw(self):
        """Aplica transform (o rebuild view_only) pendiente."""
        self._view_redraw_pending = False
        t = getattr(self, u"_view_redraw_timer", None)
        if t is not None:
            try:
                t.Stop()
            except Exception:
                pass
        if self._win is None:
            return
        try:
            if self._try_apply_view_transform_only():
                return
            self._redraw_canvas(view_only=True)
        except Exception:
            try:
                self._redraw_canvas()
            except Exception:
                pass

    def _schedule_view_redraw(self):
        """
        Zoom/pan: MatrixTransform inmediato (sin Clear/rebuild de escena).
        Fallback coalescido (~12 ms) solo si aún no hay capa de escena.
        """
        if self._apply_scene_view_transform():
            if self._hover_snap is not None or self._pick_pt1 is not None:
                try:
                    self._refresh_snap_overlay()
                except Exception:
                    pass
            return
        self._view_redraw_pending = True
        t = getattr(self, u"_view_redraw_timer", None)
        if t is not None:
            try:
                if not t.IsEnabled:
                    t.Start()
                return
            except Exception:
                self._view_redraw_timer = None
        try:
            t = DispatcherTimer()
            t.Interval = TimeSpan.FromMilliseconds(12)

            def _tick(sender, args):
                try:
                    sender.Stop()
                except Exception:
                    pass
                if not getattr(self, u"_view_redraw_pending", False):
                    return
                self._flush_view_redraw()

            t.Tick += _tick
            self._view_redraw_timer = t
            t.Start()
        except Exception:
            self._flush_view_redraw()

    def _apply_view_to_xform(self, cw, ch):
        """Sincroniza scale/ox/oy de _view_xform tras zoom/pan sin redibujar."""
        xf = self._view_xform
        if xf is None:
            return False
        try:
            cw = float(cw)
            ch = float(ch)
            if cw < 40 or ch < 40:
                return False
            min_x = float(xf[u"min_x"])
            max_x = float(xf[u"max_x"])
            min_y = float(xf[u"min_y"])
            max_y = float(xf[u"max_y"])
            fit_scale = float(xf.get(u"fit_scale") or 0.0)
            if fit_scale < 1e-12:
                return False
            zoom = float(self._view_zoom) if self._view_zoom else 1.0
            if zoom < 0.25:
                zoom = 0.25
            if zoom > 16.0:
                zoom = 16.0
            self._view_zoom = zoom
            scale = fit_scale * zoom
            bbox_cx = (min_x + max_x) / 2.0
            bbox_cy = (min_y + max_y) / 2.0
            cx_mm = bbox_cx + float(self._view_pan_x or 0.0)
            cy_mm = bbox_cy + float(self._view_pan_y or 0.0)
            ox = cw / 2.0 - (cx_mm - min_x) * scale
            oy = ch / 2.0 - (max_y - cy_mm) * scale
            xf[u"scale"] = scale
            xf[u"ox"] = ox
            xf[u"oy"] = oy
            xf[u"cw"] = cw
            xf[u"ch"] = ch
            return True
        except Exception:
            return False

    def _ensure_ctx_geo_cache(self):
        """Lazy: unión muros/vigas en mm (lista para el primer paint visible)."""
        cache = getattr(self, u"_ctx_geo_cache", None)
        if cache is not None:
            return cache
        wall_geo, beam_geo, wall_pts, beam_pts = _build_wall_beam_geo_mm(
            self._overlays
        )
        cache = {
            u"wall_geo": wall_geo,
            u"beam_geo": beam_geo,
            u"wall_pts": wall_pts,
            u"beam_pts": beam_pts,
            u"union_pending": False,
        }
        self._ctx_geo_cache = cache
        return cache

    def _schedule_plant_grids_load(self):
        """Ejes (Grid) en idle: solo snap, sin redibujar."""
        if getattr(self, u"_ctx_grids_scheduled", False):
            return
        if getattr(self, u"_ctx_grids_loaded", False):
            return
        win = self._win
        if win is None:
            return
        self._ctx_grids_scheduled = True

        def _run():
            self._load_plant_grids()

        try:
            from System import Action
            from System.Windows.Threading import DispatcherPriority

            win.Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle, Action(_run)
            )
        except Exception:
            self._ctx_grids_scheduled = False
            try:
                _run()
            except Exception:
                pass

    def _load_plant_grids(self):
        if getattr(self, u"_ctx_grids_loaded", False):
            return
        try:
            grids = _collect_grid_overlays_mm(
                self._doc, self._floor, self._plane
            )
            if grids:
                self._overlays = list(self._overlays or []) + list(grids)
                self._snap_geo_dirty = True
            self._ctx_grids_loaded = True
        except Exception:
            self._ctx_grids_loaded = True

    def _reset_canvas_view(self):
        """Restaura zoom/pan al encuadre automático (fit)."""
        self._end_canvas_pan(restore_cursor=True)
        self._stop_view_redraw_timer()
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        if not self._try_apply_view_transform_only():
            self._redraw_canvas()
        self._set_status(u"Vista restablecida (fit).")

    def _update_canvas_cursor(self):
        """Panning → SizeAll; Ctrl → flecha+plus; else Cross."""
        cv = self._get_cv_plan()
        if cv is None:
            return

        if bool(self._panning):
            try:
                cv.Cursor = Cursors.SizeAll
            except Exception:
                try:
                    cv.Cursor = Cursors.Hand
                except Exception:
                    pass
            return

        ctrl = _ctrl_modifier_down()
        self._cursor_ctrl = bool(ctrl)
        try:
            if ctrl:
                cv.Cursor = _get_arrow_plus_cursor()
            else:
                cv.Cursor = Cursors.Cross
        except Exception:
            try:
                cv.Cursor = Cursors.Cross
            except Exception:
                pass

    def _end_canvas_pan(self, restore_cursor=True):
        """Termina arrastre de pan (botón medio) y libera captura."""
        was = bool(self._panning)
        self._panning = False
        cv = self._get_cv_plan()
        if cv is not None:
            try:
                if cv.IsMouseCaptured:
                    cv.ReleaseMouseCapture()
            except Exception:
                pass
            if restore_cursor:
                self._update_canvas_cursor()
        # Último frame de pan sin esperar al timer
        if was and getattr(self, u"_view_redraw_pending", False):
            try:
                self._flush_view_redraw()
            except Exception:
                pass
        return was

    def _on_canvas_mouse_down(self, sender, e):
        """Inicio de pan con botón medio (rueda); no inicia paño ni merge."""
        try:
            btn = e.ChangedButton
        except Exception:
            return
        if btn != MouseButton.Middle:
            return
        cv = sender
        if cv is None or self._view_xform is None:
            try:
                e.Handled = True
            except Exception:
                pass
            return
        try:
            pos = e.GetPosition(cv)
            self._pan_last_x = float(pos.X)
            self._pan_last_y = float(pos.Y)
        except Exception:
            return
        self._panning = True
        try:
            cv.CaptureMouse()
        except Exception:
            pass
        # Durante pan no cambiar a flecha+plus aunque Ctrl esté pulsado
        try:
            cv.Cursor = Cursors.SizeAll
        except Exception:
            try:
                cv.Cursor = Cursors.Hand
            except Exception:
                pass
        try:
            e.Handled = True
        except Exception:
            pass

    def _on_canvas_mouse_up(self, sender, e):
        """Fin de pan con botón medio."""
        try:
            btn = e.ChangedButton
        except Exception:
            return
        if btn != MouseButton.Middle:
            return
        self._end_canvas_pan(restore_cursor=True)
        try:
            e.Handled = True
        except Exception:
            pass

    def _on_canvas_wheel(self, sender, e):
        """Zoom hacia el cursor (CAD): el punto mm bajo el ratón permanece fijo."""
        cv = sender
        xf = self._view_xform
        if cv is None or xf is None:
            return
        try:
            delta = int(e.Delta)
        except Exception:
            return
        if delta == 0:
            return

        zoom_min = 0.25
        zoom_max = 16.0
        # Paso suave por notch (120); pow acumula deltas fraccionarios (precision touchpad)
        step = 1.06

        zoom = float(self._view_zoom) if self._view_zoom else 1.0
        if zoom < zoom_min:
            zoom = zoom_min
        if zoom > zoom_max:
            zoom = zoom_max
        try:
            zoom_new = zoom * math.pow(step, float(delta) / 120.0)
        except Exception:
            zoom_new = zoom * (step if delta > 0 else (1.0 / step))
        if zoom_new < zoom_min:
            zoom_new = zoom_min
        if zoom_new > zoom_max:
            zoom_new = zoom_max
        if abs(zoom_new - zoom) < 1e-12:
            try:
                e.Handled = True
            except Exception:
                pass
            return

        try:
            pos = e.GetPosition(cv)
            mx = float(pos.X)
            my = float(pos.Y)
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            return
        if cw < 40 or ch < 40:
            return

        try:
            scale = float(xf[u"scale"])
            if scale < 1e-12:
                return
            min_x = float(xf[u"min_x"])
            max_y = float(xf[u"max_y"])
            max_x = float(xf.get(u"max_x", min_x))
            min_y = float(xf.get(u"min_y", max_y))
            ox = float(xf[u"ox"])
            oy = float(xf[u"oy"])
        except Exception:
            return

        # Punto mundo bajo el cursor (con transform actual)
        xmm = min_x + (mx - ox) / scale
        ymm = max_y - (my - oy) / scale

        actual_factor = zoom_new / zoom
        scale_new = scale * actual_factor

        # Centro de vista (mm) tal que (xmm,ymm) quede en (mx,my) tras el zoom
        cx_mm = xmm + (cw / 2.0 - mx) / scale_new
        cy_mm = ymm - (ch / 2.0 - my) / scale_new
        bbox_cx = (min_x + max_x) / 2.0
        bbox_cy = (min_y + max_y) / 2.0
        self._view_pan_x = cx_mm - bbox_cx
        self._view_pan_y = cy_mm - bbox_cy
        self._view_zoom = zoom_new
        # Xform al instante: siguientes notches / snap usan la vista actual
        self._apply_view_to_xform(cw, ch)

        try:
            e.Handled = True
        except Exception:
            pass
        self._schedule_view_redraw()

    def _cancel_current_pano_pick(self):
        """Cancela el paño en definición (punto A / preview)."""
        had_pick = self._pick_pt1 is not None
        had = had_pick or self._hover_snap is not None
        self._pick_pt1 = None
        self._hover_snap = None
        if had_pick:
            self._mark_snap_geo_dirty()
        if had:
            self._set_status(u"Creación de paño cancelada.")
        self._redraw_canvas()

    def _canvas_to_mm(self, pos):
        xf = self._view_xform
        if xf is None:
            return None
        try:
            scale = float(xf[u"scale"])
            if scale < 1e-12:
                return None
            xmm = float(xf[u"min_x"]) + (float(pos.X) - float(xf[u"ox"])) / scale
            ymm = float(xf[u"max_y"]) - (float(pos.Y) - float(xf[u"oy"])) / scale
            return (xmm, ymm)
        except Exception:
            return None

    def _mm_to_px(self, xmm, ymm):
        xf = self._view_xform
        if xf is None:
            return None
        try:
            scale = float(xf[u"scale"])
            px = float(xf[u"ox"]) + (float(xmm) - float(xf[u"min_x"])) * scale
            py = float(xf[u"oy"]) + (float(xf[u"max_y"]) - float(ymm)) * scale
            return (px, py)
        except Exception:
            return None

    def _rebuild_snap_geometry(self):
        verts = []
        segs = []
        for pts in self._ensure_sketch_polylines_cache() or []:
            _append_ring_snap(verts, segs, pts, include_midpoints=True)
        for ov in self._overlays or []:
            pts = ov.get(u"pts") or []
            if ov.get(u"kind") == _CTX_GRID or ov.get(u"closed") is False:
                _append_polyline_snap(verts, segs, pts, include_midpoints=True)
            else:
                _append_ring_snap(verts, segs, pts, include_midpoints=True)
        try:
            _append_wall_beam_intersection_snap(verts, self._overlays)
        except Exception:
            pass
        for pano in self._panos_for_face():
            _append_ring_snap(verts, segs, pano.get(u"pts") or [], include_midpoints=True)
        for ar in self._existing_ars or []:
            for ring in ar.get(u"loops") or []:
                _append_ring_snap(verts, segs, ring, include_midpoints=True)
            if not ar.get(u"loops"):
                _append_ring_snap(verts, segs, ar.get(u"pts") or [], include_midpoints=True)
        # Guías ortogonales desde punto A (útil para rectángulos)
        if self._pick_pt1 is not None:
            ax, ay = float(self._pick_pt1[0]), float(self._pick_pt1[1])
            span = 50000.0  # 50 m guía
            segs.append(((ax - span, ay), (ax + span, ay)))
            segs.append(((ax, ay - span), (ax, ay + span)))
            verts.append((ax, ay))
        self._snap_verts = verts
        self._snap_segs = segs
        try:
            self._snap_cell_index = _build_snap_cell_index(verts, segs)
        except Exception:
            self._snap_cell_index = None

    def _snap_thresh_mm(self):
        xf = self._view_xform
        if xf is None:
            return 0.0
        try:
            scale = float(xf[u"scale"])
            if scale < 1e-12:
                return 0.0
            return float(_SNAP_PX) / scale
        except Exception:
            return 0.0

    def _resolve_snap(self, pt_mm):
        if pt_mm is None:
            return None, None
        self._ensure_snap_geometry()
        thresh = self._snap_thresh_mm()
        if thresh <= 0:
            return (float(pt_mm[0]), float(pt_mm[1])), None
        return _snap_point_mm(
            pt_mm,
            self._snap_verts,
            self._snap_segs,
            thresh,
            getattr(self, u"_snap_cell_index", None),
        )

    def _clear_snap_overlay(self, cv):
        if cv is None:
            return
        try:
            remove = []
            for child in list(cv.Children):
                try:
                    if getattr(child, u"Tag", None) == _SNAP_TAG:
                        remove.append(child)
                except Exception:
                    pass
            for child in remove:
                cv.Children.Remove(child)
        except Exception:
            pass

    def _refresh_snap_overlay(self):
        cv = self._overlay_host()
        if cv is None or self._view_xform is None:
            return
        self._clear_snap_overlay(cv)
        hover = self._hover_snap
        if hover is None:
            return
        try:
            xmm, ymm, kind = hover[0], hover[1], hover[2]
        except Exception:
            return
        pxpy = self._mm_to_px(xmm, ymm)
        if pxpy is None:
            return
        px, py = pxpy
        # Preview rectángulo A → hover
        if self._pick_pt1 is not None:
            try:
                pts = rect_from_two_points_mm(self._pick_pt1, (xmm, ymm))
                if pts and len(pts) >= 4:
                    poly = WpfPolygon()
                    pc = PointCollection()
                    for qx, qy in pts:
                        qpx = self._mm_to_px(qx, qy)
                        if qpx is None:
                            continue
                        pc.Add(WpfPoint(qpx[0], qpx[1]))
                    if pc.Count >= 3:
                        poly.Points = pc
                        poly.Fill = _brush(u"#5BC0DE", 45)
                        poly.Stroke = _brush(u"#5BC0DE")
                        poly.StrokeThickness = 1.4
                        try:
                            dashes = DoubleCollection()
                            dashes.Add(5)
                            dashes.Add(3)
                            poly.StrokeDashArray = dashes
                        except Exception:
                            pass
                        poly.Tag = _SNAP_TAG
                        poly.IsHitTestVisible = False
                        cv.Children.Add(poly)
            except Exception:
                pass
        # Marcador snap
        try:
            if kind == u"vertex":
                r = 6.0
                el = WpfEllipse()
                el.Width = r * 2.0
                el.Height = r * 2.0
                el.Fill = _brush(u"#fbbf24", 200)
                el.Stroke = _brush(u"#E8F4F8")
                el.StrokeThickness = 1.5
                el.Tag = _SNAP_TAG
                el.IsHitTestVisible = False
                WpfCanvas.SetLeft(el, px - r)
                WpfCanvas.SetTop(el, py - r)
                cv.Children.Add(el)
            elif kind == u"edge":
                s = 9.0
                rect = WpfRectangle()
                rect.Width = s
                rect.Height = s
                rect.Fill = _brush(u"#4ade80", 180)
                rect.Stroke = _brush(u"#E8F4F8")
                rect.StrokeThickness = 1.2
                rect.Tag = _SNAP_TAG
                rect.IsHitTestVisible = False
                WpfCanvas.SetLeft(rect, px - s * 0.5)
                WpfCanvas.SetTop(rect, py - s * 0.5)
                cv.Children.Add(rect)
            else:
                r = 4.0
                el = WpfEllipse()
                el.Width = r * 2.0
                el.Height = r * 2.0
                el.Fill = _brush(u"#95B8CC", 160)
                el.Stroke = _brush(u"#E8F4F8", 180)
                el.StrokeThickness = 1.0
                el.Tag = _SNAP_TAG
                el.IsHitTestVisible = False
                WpfCanvas.SetLeft(el, px - r)
                WpfCanvas.SetTop(el, py - r)
                cv.Children.Add(el)
            # cruz
            for dx0, dy0, dx1, dy1 in (
                (-10, 0, 10, 0),
                (0, -10, 0, 10),
            ):
                ln = WpfLine()
                ln.X1, ln.Y1, ln.X2, ln.Y2 = px + dx0, py + dy0, px + dx1, py + dy1
                ln.Stroke = _brush(u"#fbbf24" if kind == u"vertex" else u"#4ade80")
                ln.StrokeThickness = 1.0
                ln.Tag = _SNAP_TAG
                ln.IsHitTestVisible = False
                cv.Children.Add(ln)
        except Exception:
            pass

    def _on_canvas_move(self, sender, e):
        if self._view_xform is None:
            return
        cv = self._get_cv_plan()
        if cv is None:
            return
        try:
            pos = e.GetPosition(cv)
        except Exception:
            return

        # Pan con botón medio: delta px → mm vía scale de _view_xform
        if self._panning:
            try:
                mx = float(pos.X)
                my = float(pos.Y)
                dx_px = mx - float(self._pan_last_x)
                dy_px = my - float(self._pan_last_y)
                self._pan_last_x = mx
                self._pan_last_y = my
                scale = float(self._view_xform.get(u"scale") or 0.0)
                if scale > 1e-12 and (abs(dx_px) > 1e-9 or abs(dy_px) > 1e-9):
                    # Arrastrar derecha/abajo mueve el contenido con el cursor
                    self._view_pan_x = float(self._view_pan_x or 0.0) - (
                        dx_px / scale
                    )
                    self._view_pan_y = float(self._view_pan_y or 0.0) + (
                        dy_px / scale
                    )
                    try:
                        cw = float(cv.ActualWidth)
                        ch = float(cv.ActualHeight)
                    except Exception:
                        cw = ch = 0.0
                    self._apply_view_to_xform(cw, ch)
                    self._schedule_view_redraw()
            except Exception:
                pass
            try:
                e.Handled = True
            except Exception:
                pass
            return

        # Ctrl puede cambiar sin KeyDown/Up (Alt-Tab, etc.)
        try:
            ctrl_now = _ctrl_modifier_down()
            if bool(ctrl_now) != bool(getattr(self, u"_cursor_ctrl", False)):
                self._update_canvas_cursor()
        except Exception:
            pass

        raw = self._canvas_to_mm(pos)
        if raw is None:
            return
        snapped, kind = self._resolve_snap(raw)
        if snapped is None:
            return
        # Overlay solo con hit de snap o mientras se define el 2º punto
        if kind is None and self._pick_pt1 is None:
            if self._hover_snap is not None:
                self._hover_snap = None
                self._refresh_snap_overlay()
            return
        new_hover = (snapped[0], snapped[1], kind)
        prev = self._hover_snap
        if (
            prev is not None
            and abs(prev[0] - new_hover[0]) < 0.05
            and abs(prev[1] - new_hover[1]) < 0.05
            and prev[2] == new_hover[2]
        ):
            return
        self._hover_snap = new_hover
        self._refresh_snap_overlay()

    def _on_canvas_click(self, sender, e):
        # No iniciar paño durante pan ni con botón que no sea izquierdo
        if self._panning:
            try:
                e.Handled = True
            except Exception:
                pass
            return
        try:
            if e.ChangedButton != MouseButton.Left:
                return
        except Exception:
            pass
        cv = self._get_cv_plan()
        if cv is None or self._view_xform is None:
            self._set_status(u"Espere a que el canvas termine de dibujar.")
            return
        try:
            pos = e.GetPosition(cv)
            e.Handled = True
        except Exception:
            return
        try:
            cv.Focus()
        except Exception:
            pass
        raw = self._canvas_to_mm(pos)
        if raw is None:
            return

        # Ctrl+clic: selección de fusión (no inicia rectángulo de 2 puntos)
        ctrl = False
        try:
            mods = Keyboard.Modifiers
            ctrl = (mods & ModifierKeys.Control) == ModifierKeys.Control
        except Exception:
            ctrl = False
        if ctrl:
            self._toggle_merge_at_point(raw)
            return

        # Clic sin Ctrl sobre un paño (sin pick en curso) → paño activo / cards
        if self._pick_pt1 is None:
            hit = self._hit_pano_at_mm(raw)
            if hit is not None:
                self._activate_pano(hit.get(u"id"), clear_merge=True)
                return

        pt, kind = self._resolve_snap(raw)
        if pt is None:
            return
        snap_lbl = u""
        if kind == u"vertex":
            snap_lbl = u" · snap vértice"
        elif kind == u"edge":
            snap_lbl = u" · snap arista"
        if self._pick_pt1 is None:
            self._pick_pt1 = pt
            self._hover_snap = None
            self._mark_snap_geo_dirty()
            self._set_status(
                u"Esquina A marcada ({0:.0f}, {1:.0f}) mm{2}. "
                u"Indique la esquina opuesta.".format(pt[0], pt[1], snap_lbl)
            )
            self._redraw_canvas()
            return
        # Segundo punto → crear paño
        pts = None
        try:
            pts = rect_from_two_points_mm(self._pick_pt1, pt)
        except Exception:
            pts = None
        self._pick_pt1 = None
        self._hover_snap = None
        self._mark_snap_geo_dirty()
        if not pts:
            self._set_status(u"No se creó paño: distancia insuficiente.")
            self._redraw_canvas()
            return
        area = 0.0
        try:
            if shoelace_area_m2 is not None:
                area = float(shoelace_area_m2(pts))
        except Exception:
            area = abs((pts[1][0] - pts[0][0]) * (pts[2][1] - pts[0][1])) / 1.0e6
        face = _normalize_face_id(getattr(self, u"_active_face", u"inferior"))
        seq_map = getattr(self, u"_pano_seq_by_face", None) or {}
        face_seq = int(seq_map.get(face) or 0) + 1
        seq_map[face] = face_seq
        self._pano_seq_by_face = seq_map
        self._pano_seq += 1
        pid = u"P{}".format(self._pano_seq)
        settings = self._tool_default_pano_settings(face)
        pill = _FACE_PILL.get(face) or u"SUP"
        pano = {
            u"id": pid,
            u"face": face,
            u"label": u"Paño {0} {1}".format(pill, face_seq),
            u"pts": pts,
            u"area_m2": area,
            u"layer_cfg": settings[u"layer_cfg"],
            u"ahorro_inferior": settings[u"ahorro_inferior"],
            u"ahorro_superior": settings[u"ahorro_superior"],
        }
        self._panos.append(pano)
        self._pano_merge.clear()
        # Al agregar polígono → activar toggle de malla de esa cara
        self._set_face_modeling_toggle(face, True, writeback=False)
        self._mark_snap_geo_dirty()
        self._update_create_button()
        n_face = len(self._panos_for_face(face))
        self._activate_pano(
            pid,
            clear_merge=True,
            status=u"{0} creado ({1:.1f} m²){2} · {3} paño(s) {4}.".format(
                pano[u"label"], area, snap_lbl, n_face, pill
            ),
        )

    def _hit_pano_at_mm(self, pt_mm):
        """Paño de la cara activa bajo el punto (último dibujado gana), o None."""
        if pt_mm is None or point_in_polygon_mm is None:
            return None
        hit = None
        for pano in self._panos_for_face():
            pts = pano.get(u"pts") or []
            if len(pts) < 3:
                continue
            try:
                if point_in_polygon_mm(pt_mm, pts):
                    hit = pano
            except Exception:
                continue
        return hit

    def _toggle_merge_at_point(self, pt_mm):
        """Ctrl+clic: alterna paño en el conjunto de fusión.

        Tras añadir un paño, si quedan ≥2 y forman un componente conectado
        (toque/solape), fusiona automáticamente solo geometría. Si están
        disjuntos, mantiene la selección.
        """
        # Cancelar pick de 2 puntos pendiente (Ctrl no dibuja)
        if self._pick_pt1 is not None:
            self._pick_pt1 = None
            self._hover_snap = None
            self._mark_snap_geo_dirty()
        pano = self._hit_pano_at_mm(pt_mm)
        if pano is None:
            self._set_status(u"Ctrl+clic: pulse dentro de un paño de usuario para fusionar.")
            return
        pid = pano.get(u"id")
        added = False
        if pid in self._pano_merge:
            self._pano_merge.discard(pid)
            self._set_status(
                u"Quitado de fusión: {0} · {1} en fusión.".format(
                    pano.get(u"label") or pid, len(self._pano_merge)
                )
            )
        else:
            self._pano_merge.add(pid)
            added = True
            self._set_status(
                u"Añadido a fusión: {0} · {1} en fusión.".format(
                    pano.get(u"label") or pid, len(self._pano_merge)
                )
            )
        # Cards: último clic (con o sin quedar en fusión) define el activo
        self._activate_pano(pid, clear_merge=False, redraw=False, status=None)
        n = len(self._pano_merge)

        # Auto-fusión solo tras toggle-add con ≥2 (no al quitar)
        if added and n >= 2:
            if self._merge_selection_is_connected():
                self._redraw_canvas()
                self._merge_selected_panos()
                return
            self._set_status(
                u"{0} paños en fusión, pero no se tocan/solapan. "
                u"Ctrl+clic para ajustar · Supr elimina · Esc limpia.".format(n)
            )
            self._redraw_canvas()
            return

        if n >= 2:
            self._set_status(
                u"{0} paños en fusión. Ctrl+clic otro que se toque fusiona · "
                u"Supr elimina · Esc limpia.".format(n)
            )
        elif n == 1:
            self._set_status(
                u"1 paño en fusión. Supr elimina · Ctrl+clic otro que se toque/solape "
                u"(fusión automática) · Esc limpia."
            )
        self._redraw_canvas()

    def _merge_selection_is_connected(self):
        """True si los paños en _pano_merge forman un único componente por toque/solape."""
        ids = list(self._pano_merge or [])
        if len(ids) < 2:
            return False
        id_set = set(ids)
        polys = [
            p.get(u"pts")
            for p in (self._panos or [])
            if p.get(u"id") in id_set
        ]
        polys = [pts for pts in polys if pts and len(pts) >= 3]
        if len(polys) < 2:
            return False
        if polygons_form_single_component_mm is None:
            # Sin checker: permitir intento (union_polygons_mm decidirá)
            return True
        try:
            return bool(polygons_form_single_component_mm(polys))
        except Exception:
            return False

    def _merge_selected_panos(self):
        """Une paños seleccionados (unión booleana) en un solo paño de canvas.

        Solo geometría: no crea AreaReinforcement. El usuario usa Crear después.
        """
        ids = list(self._pano_merge or [])
        if len(ids) < 2:
            self._set_status(u"Seleccione al menos 2 paños con Ctrl+clic para fusionar.")
            return
        if union_polygons_mm is None:
            self._set_status(u"Fusión no disponible (módulo de geometría).")
            return
        id_set = set(ids)
        selected = [p for p in (self._panos or []) if p.get(u"id") in id_set]
        if len(selected) < 2:
            self._pano_merge.clear()
            self._set_status(u"Selección de fusión inválida.")
            self._redraw_canvas()
            return
        faces = set(self._pano_face(p) for p in selected)
        if len(faces) != 1:
            self._set_status(
                u"Solo se pueden fusionar paños de la misma cara (Superior o Inferior)."
            )
            return
        merge_face = faces.pop()
        polys = [p.get(u"pts") for p in selected]
        merged_pts = None
        try:
            merged_pts = union_polygons_mm(polys)
        except Exception:
            merged_pts = None
        if not merged_pts or len(merged_pts) < 3:
            self._set_status(
                u"No se pueden fusionar: los paños no se tocan ni solapan "
                u"(o la unión falló)."
            )
            return
        area = 0.0
        try:
            if shoelace_area_m2 is not None:
                area = float(shoelace_area_m2(merged_pts))
        except Exception:
            area = 0.0
        # Settings: activo si estaba en la fusión; si no, primer paño fusionado
        src_settings = None
        act = getattr(self, u"_active_pano_id", None)
        for p in selected:
            if act is not None and p.get(u"id") == act:
                src_settings = self._pano_settings_copy(p)
                break
        if src_settings is None and selected:
            src_settings = self._pano_settings_copy(selected[0])
        if src_settings is None:
            src_settings = self._tool_default_pano_settings()

        # Sustituir: quitar fusionados, insertar uno nuevo editable
        kept = [p for p in (self._panos or []) if p.get(u"id") not in id_set]
        for pid in id_set:
            self._pano_selected.discard(pid)
            self._pano_merge.discard(pid)
        self._pano_seq += 1
        seq_map = getattr(self, u"_pano_seq_by_face", None) or {}
        face_seq = int(seq_map.get(merge_face) or 0) + 1
        seq_map[merge_face] = face_seq
        self._pano_seq_by_face = seq_map
        new_id = u"P{}".format(self._pano_seq)
        pill = _FACE_PILL.get(merge_face) or u"SUP"
        new_pano = {
            u"id": new_id,
            u"face": merge_face,
            u"label": u"Paño {0} {1}".format(pill, face_seq),
            u"pts": merged_pts,
            u"area_m2": area,
            u"layer_cfg": self._clamp_layer_cfg_to_face(
                src_settings[u"layer_cfg"], merge_face
            ),
            u"ahorro_inferior": src_settings[u"ahorro_inferior"]
            if merge_face == u"inferior"
            else False,
            u"ahorro_superior": src_settings[u"ahorro_superior"]
            if merge_face == u"superior"
            else False,
        }
        kept.append(new_pano)
        self._panos = kept
        self._pano_merge.clear()
        self._set_face_modeling_toggle(merge_face, True, writeback=False)
        self._mark_snap_geo_dirty()
        self._update_create_button()
        n_src = len(selected)
        self._activate_pano(
            new_id,
            clear_merge=True,
            status=u"Fusión OK: {0} paños → {1} ({2:.1f} m²). Pulse Crear Area Reinf.".format(
                n_src, new_pano[u"label"], area
            ),
        )

    def _remove_panos(self, ids):
        """Elimina paños por id y refresca lista, botón Crear, snap y canvas."""
        if not ids:
            return
        id_set = set(ids)
        removed = []
        kept = []
        for p in self._panos or []:
            pid = p.get(u"id")
            if pid in id_set:
                removed.append(p)
                self._pano_selected.discard(pid)
                self._pano_merge.discard(pid)
            else:
                kept.append(p)
        if not removed:
            return
        self._panos = kept
        act = getattr(self, u"_active_pano_id", None)
        if act is not None and act in id_set:
            face = _normalize_face_id(getattr(self, u"_active_face", u"inferior"))
            face_kept = [p for p in kept if self._pano_face(p) == face]
            pick = face_kept[-1] if face_kept else None
            self._active_pano_id = pick.get(u"id") if pick else None
            if self._active_pano_id is not None:
                self._pano_selected = set([self._active_pano_id])
            else:
                self._pano_selected = set()
        # Sin paños en una cara → apagar su toggle
        self._sync_face_toggles_from_panos()
        self._mark_snap_geo_dirty()
        self._update_create_button()
        if len(removed) == 1:
            self._set_status(
                u"Eliminado {0}.".format(removed[0].get(u"label") or removed[0].get(u"id"))
            )
        else:
            self._set_status(
                u"Eliminados {0} paños · quedan {1}.".format(
                    len(removed), len(self._panos)
                )
            )
        self._sync_cards_from_active()
        self._redraw_canvas()

    def _update_create_button(self):
        btn = self._win.FindName(u"BtnCrear") if self._win is not None else None
        n = len(self._panos or [])
        if btn is None:
            return
        try:
            if n <= 0:
                btn.Content = u"Crear Area Reinf."
                btn.IsEnabled = False
            elif n == 1:
                btn.Content = u"Crear 1 Area Reinf."
                btn.IsEnabled = True
            else:
                btn.Content = u"Crear {} Area Reinf.".format(n)
                btn.IsEnabled = True
        except Exception:
            pass

    def _selected_panos(self):
        """Todos los paños del canvas (ya no hay checklist de selección en UI)."""
        return list(self._panos or [])

    def _pano_by_id(self, pid):
        if pid is None:
            return None
        for p in self._panos or []:
            if p.get(u"id") == pid:
                return p
        return None

    def _tool_default_bar_id(self):
        default_bar = (
            self._bar_types[0][2].Id if self._bar_types else ElementId.InvalidElementId
        )
        for dmm, _lab, bt in self._bar_types or []:
            if dmm == 12:
                return bt.Id
        return default_bar

    def _copy_layer_cfg(self, layer_cfg):
        out = {}
        for key in _LAYER_KEYS:
            src = (layer_cfg or {}).get(key) or {}
            out[key] = {
                u"active": bool(src.get(u"active", True)),
                u"bar_type_id": src.get(u"bar_type_id"),
                u"spacing_mm": int(src.get(u"spacing_mm") or 150),
            }
        return out

    def _pano_face(self, pano):
        if pano is None:
            return _normalize_face_id(getattr(self, u"_active_face", u"inferior"))
        return _normalize_face_id(pano.get(u"face"))

    def _panos_for_face(self, face_id=None):
        face = _normalize_face_id(
            face_id if face_id is not None else getattr(self, u"_active_face", None)
        )
        return [p for p in (self._panos or []) if self._pano_face(p) == face]

    def _default_layer_cfg(self, face=None):
        """Defaults: solo la cara indicada ON (Ø12 si hay, Esp 150)."""
        bar_id = self._tool_default_bar_id()
        face = _normalize_face_id(
            face if face is not None else getattr(self, u"_active_face", u"inferior")
        )
        keys_on = set(_keys_for_face(face))
        cfg = {}
        for key in _LAYER_KEYS:
            cfg[key] = {
                u"active": key in keys_on,
                u"bar_type_id": bar_id,
                u"spacing_mm": 150,
            }
        return cfg

    def _clamp_layer_cfg_to_face(self, layer_cfg, face):
        """Deja activas solo las capas de ``face``; el resto OFF."""
        face = _normalize_face_id(face)
        keys_on = set(_keys_for_face(face))
        out = self._copy_layer_cfg(layer_cfg or self._default_layer_cfg(face))
        for key in _LAYER_KEYS:
            src = out.get(key) or {}
            src[u"active"] = bool(src.get(u"active")) and key in keys_on
            out[key] = src
        return out

    def _tool_default_pano_settings(self, face=None):
        face = _normalize_face_id(
            face if face is not None else getattr(self, u"_active_face", u"inferior")
        )
        return {
            u"face": face,
            u"layer_cfg": self._default_layer_cfg(face),
            u"ahorro_inferior": False,
            u"ahorro_superior": False,
        }

    def _pano_settings_copy(self, pano):
        if not pano:
            return self._tool_default_pano_settings()
        face = self._pano_face(pano)
        cfg = pano.get(u"layer_cfg")
        if not cfg:
            cfg = self._default_layer_cfg(face)
        return {
            u"face": face,
            u"layer_cfg": self._clamp_layer_cfg_to_face(cfg, face),
            u"ahorro_inferior": bool(pano.get(u"ahorro_inferior")),
            u"ahorro_superior": bool(pano.get(u"ahorro_superior")),
        }

    def _ensure_pano_settings(self, pano):
        if pano is None:
            return
        if u"face" not in pano or pano.get(u"face") not in (u"superior", u"inferior"):
            pano[u"face"] = _normalize_face_id(
                getattr(self, u"_active_face", u"inferior")
            )
        face = self._pano_face(pano)
        if not pano.get(u"layer_cfg"):
            pano[u"layer_cfg"] = self._default_layer_cfg(face)
        else:
            pano[u"layer_cfg"] = self._clamp_layer_cfg_to_face(
                pano.get(u"layer_cfg"), face
            )
        if u"ahorro_inferior" not in pano:
            pano[u"ahorro_inferior"] = False
        if u"ahorro_superior" not in pano:
            pano[u"ahorro_superior"] = False
        # Ahorro solo aplica a la cara del paño
        if face == u"superior":
            pano[u"ahorro_inferior"] = False
        else:
            pano[u"ahorro_superior"] = False
        if not _FEATURE_AHORRO_FIERRO:
            pano[u"ahorro_inferior"] = False
            pano[u"ahorro_superior"] = False
        # Compat: limpiar flags/UI antiguos de pata L
        for _k in (u"pata_edges", u"pata_l_auto"):
            if _k in pano:
                try:
                    del pano[_k]
                except Exception:
                    pano[_k] = None

    def _outline_pts_mm(self):
        """Anillo exterior de la losa en mm (plano Sketch)."""
        out = []
        try:
            loops = getattr(self, u"_sketch_loop_polylines_mm", None)
            if loops and loops[0]:
                for p in loops[0]:
                    out.append((float(p[0]), float(p[1])))
        except Exception:
            out = []
        if len(out) >= 3:
            return out
        out = []
        try:
            plane = self._plane
            for c in self._curves or []:
                seg = _sample_curve_mm(c, plane)
                if not seg:
                    continue
                if out and seg:
                    # evita duplicar vértice compartido
                    try:
                        if abs(out[-1][0] - seg[0][0]) < 1e-6 and abs(
                            out[-1][1] - seg[0][1]
                        ) < 1e-6:
                            seg = seg[1:]
                    except Exception:
                        pass
                out.extend(seg)
        except Exception:
            out = []
        return out if len(out) >= 3 else []

    def _sketch_hole_rings_mm(self):
        """Huecos del Sketch (loops interiores) en mm."""
        holes = []
        try:
            loops = getattr(self, u"_sketch_loop_polylines_mm", None)
            if not loops:
                try:
                    loops = self._ensure_sketch_polylines_cache()
                except Exception:
                    loops = None
            for hole in (loops or [])[1:]:
                pts = []
                for p in hole or []:
                    try:
                        pts.append((float(p[0]), float(p[1])))
                    except Exception:
                        continue
                if len(pts) >= 3:
                    holes.append(pts)
        except Exception:
            holes = []
        return holes

    def _pata_hole_rings_mm(self):
        """
        Anillos para matching de pata L además del outline exterior:
        shafts/pasadas (overlays) + huecos del Sketch.
        """
        rings = []
        try:
            raw = _collect_pasada_rings_mm(
                self._overlays, sketch_holes=self._sketch_hole_rings_mm()
            )
            for ring in raw or []:
                pts = []
                for p in ring or []:
                    try:
                        pts.append((float(p[0]), float(p[1])))
                    except Exception:
                        continue
                if len(pts) >= 3:
                    rings.append(pts)
        except Exception:
            rings = []
        return rings

    def _select_combo_by_tag(self, cmb, tag):
        if cmb is None or tag is None:
            return
        try:
            for i in range(cmb.Items.Count):
                it = cmb.Items[i]
                if it is None:
                    continue
                if it.Tag == tag:
                    cmb.SelectedItem = it
                    return
        except Exception:
            pass
        # ElementId: comparar por IntegerValue
        try:
            tag_iv = int(tag.IntegerValue)
        except Exception:
            return
        try:
            for i in range(cmb.Items.Count):
                it = cmb.Items[i]
                if it is None or it.Tag is None:
                    continue
                try:
                    if int(it.Tag.IntegerValue) == tag_iv:
                        cmb.SelectedItem = it
                        return
                except Exception:
                    continue
        except Exception:
            pass

    def _bar_diam_mm_from_id(self, bar_type_id):
        """Diámetro nominal (mm) del RebarBarType, o None."""
        if bar_type_id is None:
            return None
        try:
            tag_iv = int(bar_type_id.IntegerValue)
        except Exception:
            try:
                tag_iv = int(bar_type_id)
            except Exception:
                return None
        for dmm, _lab, bt in self._bar_types or []:
            try:
                if int(bt.Id.IntegerValue) == tag_iv:
                    return int(dmm) if dmm else None
            except Exception:
                continue
        return None

    def _face_cfg_summary_line(self, layer_cfg, keys, short):
        """Resumen compacto de cara: ``Inf Ø12@150`` o ``Inf —`` si OFF."""
        cfg = layer_cfg or {}
        face_on = any(bool((cfg.get(k) or {}).get(u"active")) for k in keys)
        if not face_on:
            return u"{0} —".format(short)
        major_key = keys[0]
        for k in keys:
            if k.endswith(u"_major"):
                major_key = k
                break
        src = cfg.get(major_key) or {}
        esp = int(src.get(u"spacing_mm") or 150)
        dmm = self._bar_diam_mm_from_id(src.get(u"bar_type_id"))
        if dmm:
            return u"{0} Ø{1}@{2}".format(short, dmm, esp)
        return u"{0} @{1}".format(short, esp)

    def _pano_ahorro_tags(self, pano, layer_cfg=None, face_id=None):
        """
        Tags AF de la cara del paño: ``[u'AF Inf']`` o ``[u'AF Sup']``.

        ``face_id`` opcional; por defecto la cara del paño.
        """
        cfg = layer_cfg
        if cfg is None and pano is not None:
            cfg = pano.get(u"layer_cfg")
        cfg = cfg or {}
        face = _normalize_face_id(
            face_id
            if face_id is not None
            else (self._pano_face(pano) if pano is not None else u"inferior")
        )
        tags = []
        if face == u"inferior":
            inf_on = any(
                bool((cfg.get(k) or {}).get(u"active"))
                for k in (u"interior_major", u"interior_minor")
            )
            if inf_on and pano is not None and bool(pano.get(u"ahorro_inferior")):
                tags.append(u"AF Inf")
        else:
            sup_on = any(
                bool((cfg.get(k) or {}).get(u"active"))
                for k in (u"exterior_major", u"exterior_minor")
            )
            if sup_on and pano is not None and bool(pano.get(u"ahorro_superior")):
                tags.append(u"AF Sup")
        return tags

    def _pano_cfg_hint_summary(self, pano):
        """Hint Capas: solo la malla de la cara del paño + Ahorro."""
        self._ensure_pano_settings(pano)
        cfg = pano.get(u"layer_cfg") or {}
        face = self._pano_face(pano)
        if face == u"inferior":
            mesh = self._face_cfg_summary_line(
                cfg, (u"interior_major", u"interior_minor"), u"Inf"
            )
        else:
            mesh = self._face_cfg_summary_line(
                cfg, (u"exterior_major", u"exterior_minor"), u"Sup"
            )
        if not _FEATURE_AHORRO_FIERRO:
            return mesh
        af = self._pano_ahorro_tags(pano, cfg, face_id=face)
        ahorro = u"sí" if af else u"no"
        return u"{0} · Ahorro: {1}".format(mesh, ahorro)

    def _pano_title_short(self, title):
        """``Paño 3`` → ``P3``; otro texto se acorta."""
        t = _as_unicode(title) if title is not None else u""
        t = t.strip() or u"Paño"
        low = t.lower()
        prefix = u"paño"
        if low.startswith(prefix):
            num = t[len(prefix) :].strip()
            if num:
                return u"P{0}".format(num)
            return u"P"
        if len(t) <= 6:
            return t
        return t[:6]

    def _face_cfg_slash_bits(self, layer_cfg, keys, short):
        """Cara compacta: ``Inf 12/150`` o ``Inf —``."""
        cfg = layer_cfg or {}
        face_on = any(bool((cfg.get(k) or {}).get(u"active")) for k in keys)
        if not face_on:
            return u"{0} —".format(short)
        major_key = keys[0]
        for k in keys:
            if k.endswith(u"_major"):
                major_key = k
                break
        src = cfg.get(major_key) or {}
        esp = int(src.get(u"spacing_mm") or 150)
        dmm = self._bar_diam_mm_from_id(src.get(u"bar_type_id"))
        if dmm:
            return u"{0} {1}/{2}".format(short, dmm, esp)
        return u"{0} —/{1}".format(short, esp)

    def _luz_menor_m_from_pts(self, pts):
        """Luz menor AABB en metros (1 decimal útil), o None."""
        if not pts or luz_menor_mm_from_polygon is None:
            return None
        try:
            lm_mm = luz_menor_mm_from_polygon(pts)
        except Exception:
            lm_mm = None
        if lm_mm is None or lm_mm <= 0.0:
            return None
        return float(lm_mm) / 1000.0

    def _canvas_labels_compact(self, pts=None):
        """
        True → etiqueta corta (zoom out / escala baja).

        Usa ``_view_xform['scale']`` (px/mm = fit·zoom). Si el paño es
        muy pequeño en fit, también compacta.
        """
        xf = getattr(self, u"_view_xform", None) or {}
        try:
            scale = float(xf.get(u"scale") or 0.0)
        except Exception:
            scale = 0.0
        # ~38 px/m → scale 0.038; por debajo la multilínea satura
        if scale > 0.0 and scale < 0.038:
            return True
        try:
            zoom = float(self._view_zoom) if self._view_zoom else 1.0
        except Exception:
            zoom = 1.0
        if zoom < 0.8:
            return True
        if pts:
            try:
                fit = float(xf.get(u"fit_scale") or scale or 0.0)
            except Exception:
                fit = scale
            lm_m = self._luz_menor_m_from_pts(pts)
            if lm_m is not None and fit > 0.0:
                if (lm_m * 1000.0) * fit < 72.0:
                    return True
        return False

    def _pano_cfg_canvas_parts(self, pano, pts=None, compact=None):
        """
        Partes tipográficas del bloque de info en planta.

        Solo la malla de la cara del paño (Inf o Sup), no ambas.

        Returns dict:
          title, title_short, face, mesh, mesh_c, lm_line, lm_short, af (list),
          compact (bool), compact_line
        """
        empty = {
            u"title": u"",
            u"title_short": u"",
            u"face": u"inferior",
            u"mesh": u"",
            u"mesh_c": u"",
            u"inf": u"",
            u"sup": u"",
            u"lm_line": u"",
            u"lm_short": u"",
            u"af": [],
            u"compact": False,
            u"compact_line": u"",
        }
        if pano is None:
            return empty
        self._ensure_pano_settings(pano)
        cfg = pano.get(u"layer_cfg") or {}
        face = self._pano_face(pano)
        title = pano.get(u"label") or pano.get(u"id") or u"Paño"
        title = _as_unicode(title)
        ring = pts if pts is not None else (pano.get(u"pts") or [])
        if compact is None:
            compact = self._canvas_labels_compact(ring)
        if face == u"inferior":
            keys = (u"interior_major", u"interior_minor")
            short = u"Inf"
        else:
            keys = (u"exterior_major", u"exterior_minor")
            short = u"Sup"
        mesh = self._face_cfg_summary_line(cfg, keys, short)
        mesh_c = self._face_cfg_slash_bits(cfg, keys, short)
        af = self._pano_ahorro_tags(pano, cfg, face_id=face)
        lm_m = self._luz_menor_m_from_pts(ring)
        lm_line = u""
        lm_short = u""
        if lm_m is not None:
            lm_line = u"Luz menor {0:.1f} m".format(lm_m)
            lm_short = u"Lm {0:.1f}m".format(lm_m)
        title_short = self._pano_title_short(title)
        bits = [title_short, mesh_c]
        if lm_short:
            bits.append(lm_short)
        if af:
            bits.append(u"/".join(af))
        return {
            u"title": title,
            u"title_short": title_short,
            u"face": face,
            u"mesh": mesh,
            u"mesh_c": mesh_c,
            # Compat: solo la cara del paño rellena inf o sup
            u"inf": mesh if face == u"inferior" else u"",
            u"sup": mesh if face == u"superior" else u"",
            u"lm_line": lm_line,
            u"lm_short": lm_short,
            u"af": af,
            u"compact": bool(compact),
            u"compact_line": u" · ".join(bits),
        }

    def _pano_cfg_canvas_label(self, pano, pts=None, compact=None):
        """Texto en planta (solo malla de la cara del paño + luz menor)."""
        parts = self._pano_cfg_canvas_parts(pano, pts=pts, compact=compact)
        if not parts.get(u"title"):
            return u""
        if parts.get(u"compact"):
            return parts.get(u"compact_line") or u""
        lines = [parts[u"title"]]
        mesh = parts.get(u"mesh") or u""
        if mesh:
            lines.append(mesh)
        if parts.get(u"lm_line"):
            lines.append(parts[u"lm_line"])
        af = parts.get(u"af") or []
        if af:
            lines.append(u" · ".join(af))
        return u"\n".join(lines)

    def _set_face_modeling_toggle(self, face_id, on, writeback=True):
        """Activa/desactiva el toggle de malla de una cara (y sincroniza visual)."""
        face_id = _normalize_face_id(face_id)
        face = (self._face_ui or {}).get(face_id) or {}
        chk = face.get(u"chk")
        if chk is None:
            return
        want = bool(on)
        prev_sync = getattr(self, u"_ui_syncing", False)
        if not writeback:
            self._ui_syncing = True
        try:
            try:
                if bool(chk.IsChecked) != want:
                    chk.IsChecked = want
            except Exception:
                try:
                    chk.IsChecked = want
                except Exception:
                    pass
            try:
                parts = face.get(u"toggle_parts") or {}
                if parts:
                    thumb_xform = parts.get(u"thumb_xform")
                    track_fill = parts.get(u"track_fill")
                    track_border = parts.get(u"track_border")
                    if thumb_xform is not None:
                        thumb_xform.X = 18.0 if want else 0.0
                    if want:
                        ar, ag, ab = parts.get(u"accent") or (0x5B, 0xC0, 0xDE)
                        if track_fill is not None:
                            track_fill.Color = Color.FromRgb(ar, ag, ab)
                        if track_border is not None:
                            track_border.Color = Color.FromRgb(ar, ag, ab)
                    else:
                        if track_fill is not None:
                            track_fill.Color = Color.FromRgb(18, 38, 54)
                        if track_border is not None:
                            track_border.Color = Color.FromRgb(33, 70, 92)
            except Exception:
                pass
            try:
                body = face.get(u"body")
                if body is not None:
                    body.IsEnabled = want
                    body.Opacity = 1.0 if want else 0.45
            except Exception:
                pass
        finally:
            if not writeback:
                self._ui_syncing = prev_sync

    def _sync_face_toggles_from_panos(self):
        """Toggle ON si la cara tiene ≥1 paño; OFF si no tiene."""
        for face_id in (u"inferior", u"superior"):
            has = len(self._panos_for_face(face_id)) > 0
            self._set_face_modeling_toggle(face_id, has, writeback=False)

    def _set_cards_enabled(self, enabled):
        """Habilita/deshabilita cards laterales (sin paño activo → disabled)."""
        en = bool(enabled)
        for g_id, face in (self._face_ui or {}).items():
            chk = face.get(u"chk")
            achk = face.get(u"ahorro_chk")
            ahost = face.get(u"ahorro_host")
            body = face.get(u"body")
            try:
                if chk is not None:
                    chk.IsEnabled = en
            except Exception:
                pass
            face_on = True
            if en and chk is not None:
                try:
                    face_on = bool(chk.IsChecked)
                except Exception:
                    face_on = True
            # Ø/Esp siguen el toggle de malla; Ahorro queda usable con paño activo
            dim_body = en and face_on
            try:
                if body is not None:
                    body.IsEnabled = dim_body
                    body.Opacity = 1.0 if dim_body else 0.45
            except Exception:
                pass
            try:
                target = ahost if ahost is not None else achk
                if target is not None:
                    target.IsEnabled = en
                    target.Opacity = 1.0 if en else 0.45
            except Exception:
                pass
        for ui in (self._layer_ui or {}).values():
            for key in (u"cmb_bar", u"cmb_esp"):
                w = ui.get(key)
                if w is None:
                    continue
                try:
                    w.IsEnabled = en
                except Exception:
                    pass

    def _update_layers_hint(self, pano):
        tb = None
        try:
            if self._win is not None:
                tb = self._win.FindName(u"TxtLayersHint")
        except Exception:
            tb = None
        if tb is None:
            return
        try:
            if pano is None:
                tb.Text = (
                    u"Cree o seleccione un paño en la planta para editar mallas."
                )
            else:
                label = pano.get(u"label") or pano.get(u"id") or u"Paño"
                tb.Text = u"{0} · activo\n{1}".format(
                    label, self._pano_cfg_hint_summary(pano)
                )
        except Exception:
            pass

    def _apply_settings_to_ui(self, layer_cfg, ahorro_inferior, ahorro_superior):
        """Empuja settings al UI de cards (llamar con ``_ui_syncing``)."""
        cfg = layer_cfg or self._default_layer_cfg()
        for g_id, keys in (
            (u"inferior", (u"interior_major", u"interior_minor")),
            (u"superior", (u"exterior_major", u"exterior_minor")),
        ):
            face = (self._face_ui or {}).get(g_id) or {}
            # Toggle malla: ON solo si hay ≥1 polígono en esa cara
            face_on = len(self._panos_for_face(g_id)) > 0
            self._set_face_modeling_toggle(g_id, face_on, writeback=False)
            achk = face.get(u"ahorro_chk")
            if achk is not None:
                try:
                    achk.IsChecked = bool(
                        ahorro_inferior if g_id == u"inferior" else ahorro_superior
                    )
                except Exception:
                    pass
            # Visual ahorro (por si IsChecked no dispara evento)
            try:
                ap = face.get(u"ahorro_parts") or {}
                if ap:
                    a_on = bool(
                        ahorro_inferior if g_id == u"inferior" else ahorro_superior
                    )
                    thumb_xform = ap.get(u"thumb_xform")
                    track_fill = ap.get(u"track_fill")
                    track_border = ap.get(u"track_border")
                    if thumb_xform is not None:
                        thumb_xform.X = 18.0 if a_on else 0.0
                    if a_on:
                        ar, ag, ab = ap.get(u"accent") or (0x5B, 0xC0, 0xDE)
                        if track_fill is not None:
                            track_fill.Color = Color.FromRgb(ar, ag, ab)
                        if track_border is not None:
                            track_border.Color = Color.FromRgb(ar, ag, ab)
                    else:
                        if track_fill is not None:
                            track_fill.Color = Color.FromRgb(18, 38, 54)
                        if track_border is not None:
                            track_border.Color = Color.FromRgb(33, 70, 92)
            except Exception:
                pass

        for key in _LAYER_KEYS:
            ui = (self._layer_ui or {}).get(key) or {}
            src = cfg.get(key) or {}
            self._select_combo_by_tag(ui.get(u"cmb_bar"), src.get(u"bar_type_id"))
            esp = int(src.get(u"spacing_mm") or 150)
            self._select_combo_by_tag(ui.get(u"cmb_esp"), esp)

    def _sync_cards_from_active(self):
        """Paño activo → cards (o defaults deshabilitados si no hay activo)."""
        pano = self._pano_by_id(getattr(self, u"_active_pano_id", None))
        self._ui_syncing = True
        try:
            if pano is None:
                self._apply_settings_to_ui(self._default_layer_cfg(), False, False)
                self._set_cards_enabled(False)
                self._update_layers_hint(None)
            else:
                self._ensure_pano_settings(pano)
                self._apply_settings_to_ui(
                    pano.get(u"layer_cfg"),
                    bool(pano.get(u"ahorro_inferior")),
                    bool(pano.get(u"ahorro_superior")),
                )
                self._set_cards_enabled(True)
                self._update_layers_hint(pano)
        finally:
            self._ui_syncing = False

    def _write_active_from_cards(self):
        """Cards → settings del paño activo (solo su cara)."""
        if getattr(self, u"_ui_syncing", False):
            return
        pano = self._pano_by_id(getattr(self, u"_active_pano_id", None))
        if pano is None:
            return
        face = self._pano_face(pano)
        pano[u"face"] = face
        pano[u"layer_cfg"] = self._clamp_layer_cfg_to_face(
            self._read_layer_cfg(), face
        )
        ahorro_inf, ahorro_sup = self._read_ahorro_flags()
        if face == u"superior":
            pano[u"ahorro_superior"] = bool(ahorro_sup)
            pano[u"ahorro_inferior"] = False
        else:
            pano[u"ahorro_inferior"] = bool(ahorro_inf)
            pano[u"ahorro_superior"] = False

    def _activate_pano(self, pid, clear_merge=True, redraw=True, status=None):
        """
        Define el paño activo para las cards laterales.

        - Clic sin Ctrl: ``clear_merge=True`` (selección simple).
        - Ctrl+clic: ``clear_merge=False``; el último clic sigue siendo activo.
        """
        pano = self._pano_by_id(pid)
        if pano is None:
            self._active_pano_id = None
            self._pano_selected = set()
            self._sync_cards_from_active()
            if redraw:
                self._redraw_canvas()
            return
        self._ensure_pano_settings(pano)
        face = self._pano_face(pano)
        if getattr(self, u"_active_face", None) != face:
            self._active_face = face
            try:
                self._apply_face_tab_visuals(face)
            except Exception:
                pass
            self._mark_snap_geo_dirty()
        self._active_pano_id = pid
        if clear_merge:
            try:
                self._pano_merge.clear()
            except Exception:
                self._pano_merge = set()
            self._pano_selected = set([pid])
        else:
            self._pano_selected.add(pid)
        self._sync_cards_from_active()
        if status is not None:
            self._set_status(status)
        elif clear_merge:
            self._set_status(
                u"Activo: {0} · edite malla {1}.".format(
                    pano.get(u"label") or pid,
                    _FACE_PILL.get(face) or face,
                )
            )
        if redraw:
            self._redraw_canvas()

    def _build_layer_panels(self):
        pnl = self._win.FindName(u"PnlLayers")
        if pnl is None:
            return
        try:
            pnl.Children.Clear()
        except Exception:
            while pnl.Children.Count > 0:
                pnl.Children.RemoveAt(pnl.Children.Count - 1)

        self._layer_ui = {}
        self._face_ui = {}
        self._borde_inf_ui = {}

        default_esp = 150
        default_bar = (
            self._bar_types[0][2].Id if self._bar_types else ElementId.InvalidElementId
        )
        for dmm, _lab, bt in self._bar_types:
            if dmm == 12:
                default_bar = bt.Id
                break

        from System.Windows import (
            CornerRadius,
            FontWeights,
            GridLength,
            GridUnitType,
            TextWrapping,
        )
        from System.Windows.Controls import (
            Border,
            CheckBox,
            ColumnDefinition,
            ComboBox,
            Grid,
            Orientation,
            StackPanel,
        )

        def _on_layer_change(s, e):
            if not getattr(self, u"_ui_syncing", False):
                self._write_active_from_cards()
                pano = self._pano_by_id(getattr(self, u"_active_pano_id", None))
                self._update_layers_hint(pano)
            self._redraw_canvas()

        def _apply_combo_style(cmb):
            for style_key in (u"ComboStretch", u"Combo"):
                try:
                    cmb.Style = self._win.FindResource(style_key)
                    return
                except Exception:
                    pass

        def _parse_hex_rgb(hex_color):
            h = (_as_unicode(hex_color) or u"#5BC0DE").lstrip(u"#")
            if len(h) < 6:
                return (0x5B, 0xC0, 0xDE)
            try:
                return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            except Exception:
                return (0x5B, 0xC0, 0xDE)

        def _sync_face_toggle_visual(parts, checked):
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
                    if track_fill is not None:
                        track_fill.Color = Color.FromRgb(18, 38, 54)
                    if track_border is not None:
                        track_border.Color = Color.FromRgb(33, 70, 92)
            except Exception:
                pass

        def _apply_face_toggle(chk, label_text, accent_hex, parts):
            """CheckBox → switch track+thumb + etiqueta (ON/OFF de cara)."""
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
            track.Child = thumb
            host.Children.Add(track)

            lbl = TextBlock()
            lbl.Text = _as_unicode(label_text or u"")
            lbl.FontSize = 11.0
            try:
                lbl.FontWeight = FontWeights.SemiBold
            except Exception:
                pass
            lbl.VerticalAlignment = VerticalAlignment.Center
            lbl.Foreground = _brush(u"#E8F4F8")
            lbl.TextWrapping = TextWrapping.NoWrap
            host.Children.Add(lbl)

            chk.Content = host
            ar, ag, ab = _parse_hex_rgb(accent_hex)
            parts.clear()
            parts[u"thumb_xform"] = thumb_xform
            parts[u"track_fill"] = track_fill
            parts[u"track_border"] = track_border
            parts[u"accent"] = (ar, ag, ab)
            try:
                on = bool(chk.IsChecked)
            except Exception:
                on = False
            _sync_face_toggle_visual(parts, on)

        def _make_layer_row(key, face_chk, is_last):
            """Fila Principal/Secundaria: etiqueta + Ø|Esp en una fila horizontal."""
            border = Border()
            border.Background = _brush(u"#0E1B32")
            border.BorderBrush = _brush(_LAYER_COLORS[key])
            border.BorderThickness = Thickness(1)
            border.CornerRadius = CornerRadius(4)
            border.Padding = Thickness(6, 6, 6, 6)
            border.Margin = Thickness(0, 0, 0, 0 if is_last else 6)

            sp = StackPanel()
            sp.Orientation = Orientation.Vertical

            title_tb = TextBlock()
            title_tb.Text = _LAYER_LABELS[key]
            title_tb.Foreground = _brush(u"#E8F4F8")
            title_tb.FontSize = 11
            try:
                title_tb.FontWeight = FontWeights.SemiBold
            except Exception:
                pass
            title_tb.Margin = Thickness(0, 0, 0, 4)
            tip = _LAYER_LABEL_TOOLTIPS.get(key)
            if tip:
                try:
                    title_tb.ToolTip = tip
                except Exception:
                    pass

            row = Grid()
            row.Margin = Thickness(0, 0, 0, 0)
            for w_star in (False, True, False, True):
                cd = ColumnDefinition()
                if w_star:
                    cd.Width = GridLength(1.0, GridUnitType.Star)
                else:
                    cd.Width = GridLength(1.0, GridUnitType.Auto)
                row.ColumnDefinitions.Add(cd)

            lb_o = TextBlock()
            lb_o.Text = u"Ø"
            lb_o.Foreground = _brush(u"#95B8CC")
            lb_o.FontSize = 11
            try:
                lb_o.FontWeight = FontWeights.SemiBold
            except Exception:
                pass
            lb_o.Width = 14
            lb_o.Margin = Thickness(0, 0, 4, 0)
            lb_o.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(lb_o, 0)

            cmb_o = ComboBox()
            _apply_combo_style(cmb_o)
            cmb_o.MinWidth = 0
            cmb_o.Margin = Thickness(0, 0, 6, 0)
            cmb_o.VerticalAlignment = VerticalAlignment.Center
            for _dmm, lab, bt in self._bar_types:
                it = ComboBoxItem()
                it.Content = lab
                it.Tag = bt.Id
                cmb_o.Items.Add(it)
                if bt.Id == default_bar:
                    cmb_o.SelectedItem = it
            if cmb_o.SelectedIndex < 0 and cmb_o.Items.Count > 0:
                cmb_o.SelectedIndex = 0
            Grid.SetColumn(cmb_o, 1)

            lb_e = TextBlock()
            lb_e.Text = u"Esp."
            lb_e.Foreground = _brush(u"#95B8CC")
            lb_e.FontSize = 11
            try:
                lb_e.FontWeight = FontWeights.SemiBold
            except Exception:
                pass
            lb_e.Width = 28
            lb_e.Margin = Thickness(0, 0, 4, 0)
            lb_e.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(lb_e, 2)

            cmb_e = ComboBox()
            _apply_combo_style(cmb_e)
            cmb_e.MinWidth = 0
            cmb_e.VerticalAlignment = VerticalAlignment.Center
            for esp in _SPACING_OPTS_MM:
                it = ComboBoxItem()
                it.Content = u"{}".format(esp)
                it.Tag = esp
                cmb_e.Items.Add(it)
                if esp == default_esp:
                    cmb_e.SelectedItem = it
            if cmb_e.SelectedIndex < 0 and cmb_e.Items.Count > 0:
                cmb_e.SelectedIndex = 0
            Grid.SetColumn(cmb_e, 3)

            row.Children.Add(lb_o)
            row.Children.Add(cmb_o)
            row.Children.Add(lb_e)
            row.Children.Add(cmb_e)

            sp.Children.Add(title_tb)
            sp.Children.Add(row)
            border.Child = sp

            self._layer_ui[key] = {
                u"face_chk": face_chk,
                u"cmb_bar": cmb_o,
                u"cmb_esp": cmb_e,
            }

            cmb_o.SelectionChanged += SelectionChangedEventHandler(_on_layer_change)
            cmb_e.SelectionChanged += SelectionChangedEventHandler(_on_layer_change)
            return border

        from System.Windows import Visibility

        tabs_row = Grid()
        tabs_row.Margin = Thickness(0, 0, 0, 8)
        for _i in range(len(_LAYER_GROUPS)):
            cd = ColumnDefinition()
            cd.Width = GridLength(1.0, GridUnitType.Star)
            tabs_row.ColumnDefinitions.Add(cd)

        content_host = StackPanel()
        content_host.Orientation = Orientation.Vertical

        for gi, group in enumerate(_LAYER_GROUPS):
            g_color = group[u"color"]
            g_id = group[u"id"]

            tab = Border()
            tab.Cursor = Cursors.Hand
            tab.CornerRadius = CornerRadius(4, 4, 0, 0)
            tab.Padding = Thickness(8, 8, 8, 8)
            tab.Margin = Thickness(0, 0, 4, 0) if gi == 0 else Thickness(4, 0, 0, 0)
            tab.Background = _brush(u"#0a1620")
            tab.BorderBrush = _brush(u"#21465C")
            tab.BorderThickness = Thickness(1)
            tab_inner = StackPanel()
            tab_inner.Orientation = Orientation.Horizontal
            tab_inner.HorizontalAlignment = HorizontalAlignment.Center

            pill = Border()
            pill.Background = _brush(u"#0E1B32")
            pill.BorderBrush = _brush(u"#21465C")
            pill.BorderThickness = Thickness(1)
            pill.CornerRadius = CornerRadius(3)
            pill.Padding = Thickness(6, 2, 6, 2)
            pill.Margin = Thickness(0, 0, 6, 0)
            pill_tb = TextBlock()
            pill_tb.Text = group[u"pill"]
            pill_tb.Foreground = _brush(u"#64748b")
            pill_tb.FontSize = 10
            try:
                pill_tb.FontWeight = FontWeights.Bold
            except Exception:
                pass
            pill.Child = pill_tb

            label_tb = TextBlock()
            label_tb.Text = group[u"title"]
            label_tb.Foreground = _brush(u"#95B8CC")
            label_tb.FontSize = 12
            label_tb.VerticalAlignment = VerticalAlignment.Center

            tab_inner.Children.Add(pill)
            tab_inner.Children.Add(label_tb)
            tab.Child = tab_inner

            panel = Border()
            panel.Background = _brush(u"#0a1620")
            panel.BorderBrush = _brush(g_color)
            panel.BorderThickness = Thickness(1)
            panel.CornerRadius = CornerRadius(0, 4, 4, 4)
            panel.Padding = Thickness(8)
            panel.Visibility = Visibility.Collapsed

            g_sp = StackPanel()
            g_sp.Orientation = Orientation.Vertical

            # Toggle malla: OFF hasta que se agregue un polígono en esa cara
            face_chk = CheckBox()
            face_chk.IsChecked = False
            face_chk.VerticalAlignment = VerticalAlignment.Center
            face_chk.Cursor = Cursors.Hand
            face_chk.Margin = Thickness(0, 0, 0, 6)
            try:
                face_chk.Style = self._win.FindResource(u"BimToolsToggleMini")
            except Exception:
                pass
            toggle_parts = {}
            _apply_face_toggle(face_chk, group[u"title"], g_color, toggle_parts)

            body = StackPanel()
            body.Orientation = Orientation.Vertical
            body.IsEnabled = False
            try:
                body.Opacity = 0.45
            except Exception:
                pass

            # Ahorro de fierro: toggle por cara (default OFF); desactivado de momento
            ahorro_chk = None
            ahorro_host = None
            ahorro_parts = {}
            if _FEATURE_AHORRO_FIERRO:
                ahorro_chk = CheckBox()
                ahorro_chk.IsChecked = False
                ahorro_chk.VerticalAlignment = VerticalAlignment.Center
                ahorro_chk.Cursor = Cursors.Hand
                ahorro_chk.Margin = Thickness(0, 0, 0, 0)
                try:
                    ahorro_chk.Style = self._win.FindResource(u"BimToolsToggleMini")
                except Exception:
                    pass
                _apply_face_toggle(
                    ahorro_chk,
                    u"Ahorro de fierro",
                    _AHORRO_TOGGLE_ACCENT,
                    ahorro_parts,
                )
                try:
                    ahorro_chk.ToolTip = _AHORRO_TIP
                except Exception:
                    pass

                ahorro_sub = TextBlock()
                ahorro_sub.Text = _AHORRO_TIP
                ahorro_sub.Foreground = _brush(u"#64748b")
                ahorro_sub.FontSize = 9
                ahorro_sub.TextWrapping = TextWrapping.Wrap
                ahorro_sub.Margin = Thickness(28, 2, 0, 0)

                ahorro_host = StackPanel()
                ahorro_host.Orientation = Orientation.Vertical
                ahorro_host.Margin = Thickness(0, 0, 0, 6)
                ahorro_host.Children.Add(ahorro_chk)
                ahorro_host.Children.Add(ahorro_sub)

                def _make_ahorro_handler(parts, chk_ref):
                    def _on_ahorro_toggle(s, e):
                        try:
                            on = bool(chk_ref.IsChecked)
                        except Exception:
                            on = False
                        _sync_face_toggle_visual(parts, on)
                        _on_layer_change(s, e)

                    return _on_ahorro_toggle

                ahorro_handler = _make_ahorro_handler(ahorro_parts, ahorro_chk)
                ahorro_chk.Checked += RoutedEventHandler(ahorro_handler)
                ahorro_chk.Unchecked += RoutedEventHandler(ahorro_handler)

            def _make_face_handler(parts, body_panel, chk_ref):
                def _on_face_toggle(s, e):
                    try:
                        on = bool(chk_ref.IsChecked)
                    except Exception:
                        on = False
                    _sync_face_toggle_visual(parts, on)
                    try:
                        body_panel.IsEnabled = on
                        body_panel.Opacity = 1.0 if on else 0.45
                    except Exception:
                        pass
                    _on_layer_change(s, e)

                return _on_face_toggle

            face_handler = _make_face_handler(toggle_parts, body, face_chk)
            face_chk.Checked += RoutedEventHandler(face_handler)
            face_chk.Unchecked += RoutedEventHandler(face_handler)

            # Orden: toggle malla → [Ahorro de fierro] → Principal/Secundaria
            g_sp.Children.Add(face_chk)
            if ahorro_host is not None:
                g_sp.Children.Add(ahorro_host)

            keys = group[u"keys"]
            for ki, key in enumerate(keys):
                body.Children.Add(_make_layer_row(key, face_chk, ki == len(keys) - 1))
            g_sp.Children.Add(body)

            # Inferior: card adicional Barras en borde de losa (UI; lógica luego)
            if g_id == u"inferior":
                try:
                    self._build_borde_losa_inferior_card(
                        g_sp,
                        apply_combo_style=_apply_combo_style,
                        apply_face_toggle=_apply_face_toggle,
                        sync_face_toggle_visual=_sync_face_toggle_visual,
                        default_bar=default_bar,
                        default_esp=default_esp,
                    )
                except Exception:
                    self._borde_inf_ui = {}

            panel.Child = g_sp

            def _make_tab_handler(fid):
                def _on(s, e):
                    self._set_active_face(fid)

                return _on

            tab.MouseLeftButtonDown += MouseButtonEventHandler(
                _make_tab_handler(g_id)
            )

            Grid.SetColumn(tab, gi)
            tabs_row.Children.Add(tab)
            content_host.Children.Add(panel)
            self._face_ui[g_id] = {
                u"tab": tab,
                u"pill": pill,
                u"pill_tb": pill_tb,
                u"label_tb": label_tb,
                u"panel": panel,
                u"chk": face_chk,
                u"ahorro_chk": ahorro_chk,
                u"ahorro_host": ahorro_host,
                u"body": body,
                u"keys": keys,
                u"toggle_parts": toggle_parts,
                u"ahorro_parts": ahorro_parts,
            }

        pnl.Children.Add(tabs_row)
        pnl.Children.Add(content_host)

        self._set_active_face(getattr(self, u"_active_face", u"inferior"))

    def _build_borde_losa_inferior_card(
        self,
        parent,
        apply_combo_style,
        apply_face_toggle,
        sync_face_toggle_visual,
        default_bar,
        default_esp,
    ):
        """
        Card Inferior «Barras en borde de losa».

        Solo UI por ahora (toggle + Ø/Esp); el modelado se definirá después.
        """
        from System.Windows import CornerRadius, TextWrapping
        from System.Windows.Controls import (
            Border,
            CheckBox,
            ComboBox,
            Orientation,
            StackPanel,
        )

        self._borde_inf_ui = {}
        if parent is None:
            return

        card = Border()
        card.Background = _brush(u"#0E1B32")
        card.BorderBrush = _brush(_BORDE_LOSA_ACCENT)
        card.BorderThickness = Thickness(1)
        card.CornerRadius = CornerRadius(4)
        card.Padding = Thickness(8)
        card.Margin = Thickness(0, 10, 0, 0)

        root = StackPanel()
        root.Orientation = Orientation.Vertical

        chk = CheckBox()
        chk.IsChecked = False
        chk.Cursor = Cursors.Hand
        chk.Margin = Thickness(0, 0, 0, 6)
        try:
            chk.Style = self._win.FindResource(u"BimToolsToggleMini")
        except Exception:
            pass
        toggle_parts = {}
        apply_face_toggle(
            chk, u"Barras en borde de losa", _BORDE_LOSA_ACCENT, toggle_parts
        )

        body = StackPanel()
        body.Orientation = Orientation.Vertical
        body.IsEnabled = False
        try:
            body.Opacity = 0.45
        except Exception:
            pass

        tip = TextBlock()
        tip.Text = _BORDE_LOSA_TIP
        tip.Foreground = _brush(u"#64748b")
        tip.FontSize = 9
        tip.TextWrapping = TextWrapping.Wrap
        tip.Margin = Thickness(0, 0, 0, 6)
        body.Children.Add(tip)

        row = StackPanel()
        row.Orientation = Orientation.Horizontal

        lb_o = TextBlock()
        lb_o.Text = u"Ø"
        lb_o.Foreground = _brush(u"#95B8CC")
        lb_o.FontSize = 11
        lb_o.Width = 14
        lb_o.Margin = Thickness(0, 0, 4, 0)
        lb_o.VerticalAlignment = VerticalAlignment.Center

        cmb_o = ComboBox()
        apply_combo_style(cmb_o)
        cmb_o.MinWidth = 72
        cmb_o.Margin = Thickness(0, 0, 8, 0)
        cmb_o.VerticalAlignment = VerticalAlignment.Center
        for _dmm, lab, bt in self._bar_types:
            it = ComboBoxItem()
            it.Content = lab
            it.Tag = bt.Id
            cmb_o.Items.Add(it)
            if bt.Id == default_bar:
                cmb_o.SelectedItem = it
        if cmb_o.SelectedIndex < 0 and cmb_o.Items.Count > 0:
            cmb_o.SelectedIndex = 0

        lb_e = TextBlock()
        lb_e.Text = u"Esp."
        lb_e.Foreground = _brush(u"#95B8CC")
        lb_e.FontSize = 11
        lb_e.Width = 28
        lb_e.Margin = Thickness(0, 0, 4, 0)
        lb_e.VerticalAlignment = VerticalAlignment.Center

        cmb_e = ComboBox()
        apply_combo_style(cmb_e)
        cmb_e.MinWidth = 64
        cmb_e.VerticalAlignment = VerticalAlignment.Center
        for esp in _SPACING_OPTS_MM:
            it = ComboBoxItem()
            it.Content = u"{}".format(esp)
            it.Tag = esp
            cmb_e.Items.Add(it)
            if esp == default_esp:
                cmb_e.SelectedItem = it
        if cmb_e.SelectedIndex < 0 and cmb_e.Items.Count > 0:
            cmb_e.SelectedIndex = 0

        for el in (lb_o, cmb_o, lb_e, cmb_e):
            row.Children.Add(el)
        body.Children.Add(row)

        def _on_borde_toggle(s, e):
            try:
                on = bool(chk.IsChecked)
            except Exception:
                on = False
            sync_face_toggle_visual(toggle_parts, on)
            try:
                body.IsEnabled = on
                body.Opacity = 1.0 if on else 0.45
            except Exception:
                pass

        chk.Checked += RoutedEventHandler(_on_borde_toggle)
        chk.Unchecked += RoutedEventHandler(_on_borde_toggle)

        root.Children.Add(chk)
        root.Children.Add(body)
        card.Child = root
        parent.Children.Add(card)

        self._borde_inf_ui = {
            u"card": card,
            u"chk": chk,
            u"body": body,
            u"toggle_parts": toggle_parts,
            u"cmb_bar": cmb_o,
            u"cmb_esp": cmb_e,
            u"tip": tip,
        }

    def _apply_face_tab_visuals(self, face_id):
        """Actualiza chrome de tabs Superior/Inferior + panel visible."""
        face_id = _normalize_face_id(face_id)
        from System.Windows import Visibility

        for g in _LAYER_GROUPS:
            g_id = g[u"id"]
            ui = self._face_ui.get(g_id) or {}
            on = g_id == face_id
            accent = g[u"color"]
            tab = ui.get(u"tab")
            pill = ui.get(u"pill")
            pill_tb = ui.get(u"pill_tb")
            label_tb = ui.get(u"label_tb")
            panel = ui.get(u"panel")
            if tab is not None:
                try:
                    tab.Background = _brush(u"#0E1B32" if on else u"#0a1620")
                    tab.BorderBrush = _brush(accent if on else u"#21465C")
                    tab.BorderThickness = Thickness(1, 1, 1, 2 if on else 1)
                except Exception:
                    pass
            if pill is not None:
                try:
                    pill.BorderBrush = _brush(accent if on else u"#21465C")
                except Exception:
                    pass
            if pill_tb is not None:
                try:
                    pill_tb.Foreground = _brush(accent if on else u"#64748b")
                except Exception:
                    pass
            if label_tb is not None:
                try:
                    label_tb.Foreground = _brush(u"#E8F4F8" if on else u"#95B8CC")
                    label_tb.FontWeight = (
                        FontWeights.SemiBold if on else FontWeights.Normal
                    )
                except Exception:
                    pass
            if panel is not None:
                try:
                    panel.Visibility = (
                        Visibility.Visible if on else Visibility.Collapsed
                    )
                except Exception:
                    pass

    def _set_active_face(self, face_id):
        """Cambia tab Superior/Inferior: visual, paños de esa cara y cards."""
        if face_id not in (u"superior", u"inferior"):
            return
        prev = getattr(self, u"_active_face", None)
        self._active_face = face_id
        self._apply_face_tab_visuals(face_id)

        if prev == face_id:
            return

        # Cambiar de cara: cancelar pick/fusión y mostrar paños de esa cara
        had_pick = self._pick_pt1 is not None
        self._pick_pt1 = None
        self._hover_snap = None
        if had_pick:
            self._mark_snap_geo_dirty()
        try:
            self._pano_merge.clear()
        except Exception:
            self._pano_merge = set()

        face_panos = self._panos_for_face(face_id)
        cur = self._pano_by_id(getattr(self, u"_active_pano_id", None))
        if cur is not None and self._pano_face(cur) == face_id:
            self._sync_cards_from_active()
        elif face_panos:
            self._activate_pano(
                face_panos[-1].get(u"id"),
                clear_merge=True,
                redraw=False,
                status=u"Cara {0}: {1} paño(s). Dibuje o edite.".format(
                    _FACE_PILL.get(face_id) or face_id,
                    len(face_panos),
                ),
            )
        else:
            self._active_pano_id = None
            self._pano_selected = set()
            self._sync_cards_from_active()
            self._set_status(
                u"Cara {0}: sin paños. 2 clics para definir polígonos "
                u"(sin polígonos → no se crea AR en esta cara).".format(
                    _FACE_PILL.get(face_id) or face_id
                )
            )
        self._mark_snap_geo_dirty()
        self._redraw_canvas()

    def _set_status(self, text):
        try:
            tb = self._win.FindName(u"TxtStatus")
            if tb is not None:
                tb.Text = _as_unicode(text)
        except Exception:
            pass

    def _major_direction(self):
        """Fallback global: arista más larga del Sketch (losa)."""
        return direccion_arista_mas_larga(self._curves)

    def _major_direction_for_pano(self, pts):
        """
        Major del paño: luz menor (AABB plano-mm → XYZ).
        Solo si el polígono es degenerado usa arista más larga de la losa.
        """
        dir_mm = None
        try:
            if span_direction_from_polygon_mm is not None:
                dir_mm = span_direction_from_polygon_mm(pts)
        except Exception:
            dir_mm = None
        if dir_mm is None:
            return self._major_direction()
        return _plane_mm_dir_to_xyz(dir_mm[0], dir_mm[1], self._plane)

    def _read_layer_cfg(self):
        """
        layer_cfg por clave AR. ``active`` sigue el toggle de malla de la cara.
        """
        cfg = {}
        for key in _LAYER_KEYS:
            ui = self._layer_ui.get(key) or {}
            active = False
            try:
                face_chk = ui.get(u"face_chk")
                if face_chk is not None:
                    active = bool(face_chk.IsChecked)
            except Exception:
                active = False
            bar_id = ElementId.InvalidElementId
            try:
                it = ui[u"cmb_bar"].SelectedItem
                if it is not None and it.Tag is not None:
                    bar_id = it.Tag
            except Exception:
                pass
            esp = 150
            try:
                it = ui[u"cmb_esp"].SelectedItem
                if it is not None and it.Tag is not None:
                    esp = int(it.Tag)
            except Exception:
                pass
            cfg[key] = {
                u"active": active,
                u"bar_type_id": bar_id,
                u"spacing_mm": esp,
            }
        return cfg

    def _read_ahorro_flags(self):
        """Ahorro por cara desde el toggle «Ahorro de fierro» (Inferior / Superior)."""
        if not _FEATURE_AHORRO_FIERRO:
            return False, False
        ahorro_inferior = False
        ahorro_superior = False
        for g_id in (u"inferior", u"superior"):
            face = (self._face_ui or {}).get(g_id) or {}
            ahorro_on = False
            try:
                achk = face.get(u"ahorro_chk")
                if achk is not None:
                    ahorro_on = bool(achk.IsChecked)
            except Exception:
                ahorro_on = False
            # Solo aplica si la cara tiene malla activa (toggle o paños)
            face_on = False
            try:
                chk = face.get(u"chk")
                if chk is not None:
                    face_on = bool(chk.IsChecked)
            except Exception:
                face_on = False
            if not face_on:
                ahorro_on = False
            if g_id == u"inferior":
                ahorro_inferior = ahorro_on
            else:
                ahorro_superior = ahorro_on
        return ahorro_inferior, ahorro_superior

    def _dispose_crear_event(self):
        try:
            if self._crear_event is not None:
                self._crear_event.Dispose()
        except Exception:
            pass
        self._crear_event = None

    def _snapshot_crear_request(self):
        """
        Captura paños con settings propios (sin widgets WPF).

        Cada paño lleva su ``layer_cfg`` + ``ahorro_inferior`` / ``ahorro_superior``.
        Returns:
            (dict, None) si ok; (None, unicode) si validación falla.
        """
        # Persistir cards del activo antes del snapshot
        try:
            self._write_active_from_cards()
        except Exception:
            pass

        selected = self._selected_panos()
        if not selected:
            return None, (
                u"Defina al menos un paño en Superior o Inferior "
                u"(2 clics en el canvas de la cara correspondiente)."
            )

        panos = []
        for pano in selected:
            self._ensure_pano_settings(pano)
            face = self._pano_face(pano)
            layer_copy = self._clamp_layer_cfg_to_face(pano.get(u"layer_cfg"), face)
            label = pano.get(u"label") or pano.get(u"id") or u"paño"
            pts_src = pano.get(u"pts") or []
            pts_copy = []
            for pt in pts_src:
                try:
                    pts_copy.append((float(pt[0]), float(pt[1])))
                except Exception:
                    continue
            major = self._major_direction_for_pano(pts_src)
            major_xyz = None
            if major is not None:
                try:
                    major_xyz = (
                        float(major.X),
                        float(major.Y),
                        float(major.Z),
                    )
                except Exception:
                    major_xyz = None
            panos.append(
                {
                    u"id": pano.get(u"id"),
                    u"face": face,
                    u"label": pano.get(u"label"),
                    u"pts": pts_copy,
                    u"major_xyz": major_xyz,
                    u"layer_cfg": layer_copy,
                    u"ahorro_inferior": (
                        bool(pano.get(u"ahorro_inferior"))
                        if (
                            _FEATURE_AHORRO_FIERRO
                            and face == u"inferior"
                        )
                        else False
                    ),
                    u"ahorro_superior": (
                        bool(pano.get(u"ahorro_superior"))
                        if (
                            _FEATURE_AHORRO_FIERRO
                            and face == u"superior"
                        )
                        else False
                    ),
                }
            )

        outline_pts = []
        try:
            for p in self._outline_pts_mm() or []:
                outline_pts.append((float(p[0]), float(p[1])))
        except Exception:
            outline_pts = []
        hole_rings = []
        try:
            hole_rings = list(self._pata_hole_rings_mm() or [])
        except Exception:
            hole_rings = []

        return (
            {
                u"panos": panos,
                u"doc": self._doc,
                u"floor": self._floor,
                u"uidoc": self._uidoc,
                u"uiapp": self._uiapp,
                u"plane": self._plane,
                u"outline_pts_mm": outline_pts,
                u"pata_hole_rings_mm": hole_rings,
            },
            None,
        )

    def _redraw_canvas(self, view_only=False):
        """
        Redibuja el canvas de planta.

        Escena en coords fit (zoom=1, pan=0) dentro de _scene_layer;
        zoom/pan solo actualiza MatrixTransform (+ HUD escala).
        view_only=True: intenta transform sin rebuild; si falla, rebuild completo.
        """
        # Un full redraw absorbe cualquier view-redraw pendiente
        if not view_only:
            self._view_redraw_pending = False
            t = getattr(self, u"_view_redraw_timer", None)
            if t is not None:
                try:
                    t.Stop()
                except Exception:
                    pass

        cv = self._get_cv_plan()
        if cv is None:
            return

        try:
            cw = float(cv.ActualWidth)
            ch = float(cv.ActualHeight)
        except Exception:
            cw = ch = 0.0
        # Sin Show() Actual* suele ser 0; usar Width/Height forzados en prepare offline.
        if cw < 40 or ch < 40:
            try:
                cw = float(cv.Width or 0.0)
                ch = float(cv.Height or 0.0)
            except Exception:
                cw = ch = 0.0
        if cw < 40 or ch < 40:
            return

        # Zoom/pan fluido: no Clear ni recrear children
        if view_only and self._try_apply_view_transform_only():
            return

        self._last_canvas_cw = cw
        self._last_canvas_ch = ch

        plane = self._plane
        all_pts = []
        loop_polylines = list(self._ensure_sketch_polylines_cache() or [])
        for pts in loop_polylines:
            all_pts.extend(pts)
        for ov in self._overlays or []:
            pts = ov.get(u"pts") or []
            all_pts.extend(pts)
        for pano in self._panos or []:
            all_pts.extend(pano.get(u"pts") or [])
        for ar in self._existing_ars or []:
            for ring in ar.get(u"loops") or []:
                all_pts.extend(ring)
            if not ar.get(u"loops"):
                all_pts.extend(ar.get(u"pts") or [])
        if self._pick_pt1 is not None:
            all_pts.append(self._pick_pt1)
        if not all_pts:
            try:
                cv.Children.Clear()
            except Exception:
                pass
            self._scene_layer = None
            self._hud_layer = None
            self._scene_base = None
            self._scene_matrix_transform = None
            self._view_xform = None
            return

        min_x, min_y, max_x, max_y = _bbox_mm(all_pts)
        bw = max(max_x - min_x, 1.0)
        bh = max(max_y - min_y, 1.0)
        pad = 36.0
        fit_scale = min((cw - 2 * pad) / bw, (ch - 2 * pad) / bh)
        if fit_scale < 1e-12:
            fit_scale = 1e-12
        zoom = float(self._view_zoom) if self._view_zoom else 1.0
        if zoom < 0.25:
            zoom = 0.25
        if zoom > 16.0:
            zoom = 16.0
        self._view_zoom = zoom
        scale = fit_scale * zoom
        # Centro de vista = centro bbox + pan (mm); zoom=1/pan=0 ≡ fit centrado
        bbox_cx = (min_x + max_x) / 2.0
        bbox_cy = (min_y + max_y) / 2.0
        cx_mm = bbox_cx + float(self._view_pan_x or 0.0)
        cy_mm = bbox_cy + float(self._view_pan_y or 0.0)
        # Base fit (escena); ox/oy actuales van a _view_xform + MatrixTransform
        ox0 = cw / 2.0 - (bbox_cx - min_x) * fit_scale
        oy0 = ch / 2.0 - (max_y - bbox_cy) * fit_scale
        ox = cw / 2.0 - (cx_mm - min_x) * scale
        oy = ch / 2.0 - (max_y - cy_mm) * scale
        self._view_xform = {
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"ox": ox,
            u"oy": oy,
            u"scale": scale,
            u"fit_scale": fit_scale,
            u"cw": cw,
            u"ch": ch,
        }
        self._scene_base = {
            u"ox0": ox0,
            u"oy0": oy0,
            u"fit_scale": fit_scale,
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"cw": cw,
            u"ch": ch,
        }

        try:
            cv.Children.Clear()
        except Exception:
            return
        scene = WpfCanvas()
        scene.IsHitTestVisible = False
        hud = WpfCanvas()
        hud.IsHitTestVisible = False
        try:
            cv.Children.Add(scene)
            cv.Children.Add(hud)
        except Exception:
            return
        self._scene_layer = scene
        self._hud_layer = hud
        self._scene_matrix_transform = None

        # Geometría de escena siempre en coords fit (zoom/pan = transform)
        def to_px(xmm, ymm):
            return (
                ox0 + (xmm - min_x) * fit_scale,
                oy0 + (max_y - ymm) * fit_scale,
            )

        def _add_polygon(pts, fill_hex, stroke_hex, stroke_w=1.2, dashed=False, fill_a=200):
            if not pts or len(pts) < 3:
                return
            poly = WpfPolygon()
            pc = PointCollection()
            for xmm, ymm in pts:
                px, py = to_px(xmm, ymm)
                pc.Add(WpfPoint(px, py))
            poly.Points = pc
            poly.Fill = _brush(fill_hex, fill_a)
            poly.Stroke = _brush(stroke_hex)
            poly.StrokeThickness = stroke_w
            if dashed:
                try:
                    dashes = DoubleCollection()
                    dashes.Add(4)
                    dashes.Add(3)
                    poly.StrokeDashArray = dashes
                except Exception:
                    pass
            scene.Children.Add(poly)

        def _pano_centroid_mm(pts):
            """Centroide por media de vértices (mm plano)."""
            n = float(len(pts))
            return (
                sum(float(q[0]) for q in pts) / n,
                sum(float(q[1]) for q in pts) / n,
            )

        def _major_symbol_geom(pts):
            """
            Geometría del símbolo Major (luz menor): dir unitaria, centroide,
            half-seg mm y perpendicular unitaria in-plane.
            """
            if not pts or len(pts) < 3:
                return None
            dir_mm = None
            try:
                if span_direction_from_polygon_mm is not None:
                    dir_mm = span_direction_from_polygon_mm(pts)
            except Exception:
                dir_mm = None
            if dir_mm is None:
                return None
            dx = float(dir_mm[0])
            dy = float(dir_mm[1])
            L = (dx * dx + dy * dy) ** 0.5
            if L < 1e-9:
                return None
            dx /= L
            dy /= L
            cx, cy = _pano_centroid_mm(pts)
            xs = [float(q[0]) for q in pts]
            ys = [float(q[1]) for q in pts]
            pw = max(xs) - min(xs)
            ph = max(ys) - min(ys)
            extent_long = max(pw, ph)
            if extent_long < 1.0:
                return None
            seg_len = extent_long * 0.50
            half = seg_len * 0.5
            nx, ny = -dy, dx
            return {
                u"dx": dx,
                u"dy": dy,
                u"nx": nx,
                u"ny": ny,
                u"cx": cx,
                u"cy": cy,
                u"pw": pw,
                u"ph": ph,
                u"min_x": min(xs),
                u"max_x": max(xs),
                u"min_y": min(ys),
                u"max_y": max(ys),
                u"half": half,
                u"seg_len": seg_len,
            }

        def _prepare_pano_info_block(pts, pano, is_active=False, merge_sel=False):
            """
            Construye la tarjeta de info (sin añadir al scene).

            Returns (border, tw_fit_px, th_fit_px) o None.
            Centrado posterior usa tw/th ya con shrink si el paño es chico.
            """
            if not pts or pano is None:
                return None
            from System import Double as _Dbl
            from System.Windows import (
                CornerRadius,
                HorizontalAlignment,
                Size as _Sz,
                TextAlignment as _TA,
            )
            from System.Windows.Controls import Border, StackPanel
            from System.Windows.Media import ScaleTransform

            parts = self._pano_cfg_canvas_parts(pano, pts=pts)
            compact = bool(parts.get(u"compact"))
            if merge_sel:
                accent = _ACTIVE_STROKE if is_active else _MERGE_STROKE
            elif is_active:
                accent = _ACTIVE_STROKE
            else:
                accent = u"#E8F4F8"
            body_fg = u"#B8D4E0" if not is_active else u"#D0EAF4"
            af_fg = u"#7DD3C0"
            panel = StackPanel()
            try:
                panel.HorizontalAlignment = HorizontalAlignment.Center
            except Exception:
                pass

            def _tb(text, size, weight, fg):
                t = TextBlock()
                t.Text = text
                t.FontSize = size
                try:
                    t.FontWeight = weight
                except Exception:
                    pass
                t.Foreground = _brush(fg)
                try:
                    t.TextAlignment = _TA.Center
                    t.HorizontalAlignment = HorizontalAlignment.Center
                except Exception:
                    pass
                return t

            if compact:
                line = parts.get(u"compact_line") or u""
                if not line:
                    return None
                panel.Children.Add(
                    _tb(line, 8.0, FontWeights.SemiBold, accent)
                )
            else:
                title = parts.get(u"title") or u"Paño"
                panel.Children.Add(
                    _tb(
                        title,
                        10.5 if is_active else 10.0,
                        FontWeights.Bold,
                        accent,
                    )
                )
                # Solo malla de la cara del paño (Inf o Sup), no ambas.
                for key in (u"mesh", u"lm_line"):
                    s = parts.get(key) or u""
                    if s:
                        panel.Children.Add(
                            _tb(s, 8.0, FontWeights.Normal, body_fg)
                        )
                af = parts.get(u"af") or []
                if af:
                    panel.Children.Add(
                        _tb(
                            u" · ".join(af),
                            7.5,
                            FontWeights.SemiBold,
                            af_fg,
                        )
                    )

            border = Border()
            border.Child = panel
            border.Background = _brush(u"#071018", 175)
            try:
                border.CornerRadius = CornerRadius(3.0)
            except Exception:
                pass
            border.Padding = Thickness(5, 3, 5, 3)
            if is_active:
                border.BorderBrush = _brush(_ACTIVE_STROKE, 230)
                border.BorderThickness = Thickness(1.3)
            elif merge_sel:
                border.BorderBrush = _brush(_MERGE_STROKE, 210)
                border.BorderThickness = Thickness(1.1)
            else:
                border.BorderBrush = _brush(u"#21465C", 150)
                border.BorderThickness = Thickness(0.8)
            try:
                border.IsHitTestVisible = False
            except Exception:
                pass
            try:
                border.Measure(
                    _Sz(_Dbl.PositiveInfinity, _Dbl.PositiveInfinity)
                )
                tw = float(border.DesiredSize.Width)
                th = float(border.DesiredSize.Height)
            except Exception:
                tw, th = (96.0, 40.0) if not compact else (120.0, 16.0)

            # Paño muy chico: encoger tarjeta para que quepa en el AABB
            try:
                xs = [float(q[0]) for q in pts]
                ys = [float(q[1]) for q in pts]
                pano_w_px = max(max(xs) - min(xs), 1.0) * fit_scale
                pano_h_px = max(max(ys) - min(ys), 1.0) * fit_scale
                max_tw = max(pano_w_px * 0.82, 20.0)
                max_th = max(pano_h_px * 0.82, 14.0)
                ui_scale = 1.0
                if tw > max_tw or th > max_th:
                    ui_scale = min(
                        max_tw / max(tw, 1.0),
                        max_th / max(th, 1.0),
                    )
                    ui_scale = max(0.5, min(1.0, ui_scale))
                if ui_scale < 0.999:
                    border.LayoutTransform = ScaleTransform(ui_scale, ui_scale)
                    try:
                        border.Measure(
                            _Sz(
                                _Dbl.PositiveInfinity,
                                _Dbl.PositiveInfinity,
                            )
                        )
                        tw = float(border.DesiredSize.Width)
                        th = float(border.DesiredSize.Height)
                    except Exception:
                        tw *= ui_scale
                        th *= ui_scale
            except Exception:
                pass
            return border, tw, th

        def _place_pano_info_block(pts, border, tw, th):
            """Centra la tarjeta en el centroide del paño (fit-px)."""
            if border is None or not pts:
                return
            cx, cy = _pano_centroid_mm(pts)
            px, py = to_px(cx, cy)
            WpfCanvas.SetLeft(border, px - float(tw) * 0.5)
            WpfCanvas.SetTop(border, py - float(th) * 0.5)
            scene.Children.Add(border)

        def _add_major_luz_menor_symbol(
            pts, is_active=False, label_tw=0.0, label_th=0.0
        ):
            """
            Símbolo Major estilo AR Revit: segmento // luz menor + ticks
            perpendiculares. Desplazado en la perpendicular in-plane para
            no cruzar la tarjeta de info centrada en el paño.
            """
            g = _major_symbol_geom(pts)
            if g is None:
                return
            dx = g[u"dx"]
            dy = g[u"dy"]
            nx = g[u"nx"]
            ny = g[u"ny"]
            cx = g[u"cx"]
            cy = g[u"cy"]
            half = g[u"half"]
            seg_len = g[u"seg_len"]
            pw = g[u"pw"]
            ph = g[u"ph"]
            cap = max(seg_len * 0.10, min(pw, ph) * 0.06)
            if cap > seg_len * 0.22:
                cap = seg_len * 0.22
            ch = cap * 0.5

            # Holgura = semi-extensión AABB etiqueta en dir. perp. pantalla + pad
            # to_px invierte Y: (nx, ny) mm → (nx, -ny) fit-px
            ltw = float(label_tw or 0.0)
            lth = float(label_th or 0.0)
            ux, uy = float(nx), -float(ny)
            ulen = (ux * ux + uy * uy) ** 0.5
            if ulen > 1e-12:
                ux /= ulen
                uy /= ulen
            pad_px = 8.0 if (ltw > 0.5 or lth > 0.5) else 0.0
            half_lab_px = 0.5 * (abs(ux) * ltw + abs(uy) * lth)
            need_mm = 0.0
            if fit_scale > 1e-12:
                need_mm = (half_lab_px + pad_px) / fit_scale

            # Lado con más espacio libre dentro del AABB del paño
            minx = g[u"min_x"]
            maxx = g[u"max_x"]
            miny = g[u"min_y"]
            maxy = g[u"max_y"]
            dots = []
            for x, y in (
                (minx, miny),
                (minx, maxy),
                (maxx, miny),
                (maxx, maxy),
            ):
                dots.append((x - cx) * nx + (y - cy) * ny)
            free_pos = max(dots)
            free_neg = -min(dots)
            if free_pos > free_neg + 1e-6:
                side = 1.0
                free = free_pos
            elif free_neg > free_pos + 1e-6:
                side = -1.0
                free = free_neg
            else:
                side = 1.0
                free = free_pos

            # Mantener ticks dentro del AABB; si no hay sitio, reducir offset
            max_off = max(0.0, free - ch - 2.0)
            off_mm = min(need_mm, max_off) if max_off > 0.0 else 0.0

            cx_i = cx + nx * off_mm * side
            cy_i = cy + ny * off_mm * side

            # Paño estrecho: acortar segmento desde el centro ya offset
            maj_dots = []
            for x, y in (
                (minx, miny),
                (minx, maxy),
                (maxx, miny),
                (maxx, maxy),
            ):
                maj_dots.append((x - cx_i) * dx + (y - cy_i) * dy)
            half_free_maj = min(-min(maj_dots), max(maj_dots))
            half_max = max(0.0, half_free_maj - 2.0)
            if half > half_max:
                half = half_max
                seg_len = half * 2.0
                cap = max(seg_len * 0.10, min(pw, ph) * 0.06)
                if cap > seg_len * 0.22:
                    cap = seg_len * 0.22
                ch = cap * 0.5

            x1 = cx_i - dx * half
            y1 = cy_i - dy * half
            x2 = cx_i + dx * half
            y2 = cy_i + dy * half
            stroke = _ACTIVE_STROKE if is_active else u"#5BC0DE"
            thick = 2.4 if is_active else 1.7
            alpha = 235 if is_active else 195

            def _seg(ax, ay, bx, by, w=None):
                pax, pay = to_px(ax, ay)
                pbx, pby = to_px(bx, by)
                ln = WpfLine()
                ln.X1 = pax
                ln.Y1 = pay
                ln.X2 = pbx
                ln.Y2 = pby
                ln.Stroke = _brush(stroke, alpha)
                ln.StrokeThickness = thick if w is None else w
                scene.Children.Add(ln)

            if half < 1.0:
                return
            _seg(x1, y1, x2, y2)
            _seg(x1 - nx * ch, y1 - ny * ch, x1 + nx * ch, y1 + ny * ch)
            _seg(x2 - nx * ch, y2 - ny * ch, x2 + nx * ch, y2 + ny * ch)

        def _mid_layer(scene_ref, to_px_ref, add_poly_ref, sw_fn=None):
            # Usa closures scene/to_px/_add_polygon del redraw.
            # Paños definidos por el usuario (2 puntos → rectángulo / fusionados).
            # Relleno/borde recortados por pasadas + huecos Sketch (Exclude).
            # Grosor = mismo escalado que muros/vigas/pasadas (sw_fn).
            def _line_w(base=1.2):
                if sw_fn is not None:
                    try:
                        return float(sw_fn(base))
                    except Exception:
                        pass
                try:
                    return float(base) * 0.5
                except Exception:
                    return 0.6

            pano_stroke = _line_w(1.2)
            try:
                sketch_holes = (
                    list(loop_polylines[1:])
                    if loop_polylines and len(loop_polylines) > 1
                    else []
                )
                pasada_rings = _collect_pasada_rings_mm(
                    self._overlays, sketch_holes=sketch_holes
                )
                pasada_mask = (
                    _union_polygons_geometry(pasada_rings, to_px)
                    if pasada_rings
                    else None
                )
            except Exception:
                pasada_mask = None

            def _add_pano_poly(
                pts, fill_hex, stroke_hex, stroke_w=None, dashed=False, fill_a=200
            ):
                geo = _pano_geometry_cut_by_pasadas(pts, pasada_mask, to_px)
                if geo is None:
                    return
                try:
                    b = geo.Bounds
                    if b is None or float(b.Width) < 0.25 or float(b.Height) < 0.25:
                        return
                except Exception:
                    pass
                _add_path_geometry(
                    scene,
                    geo,
                    fill_hex,
                    stroke_hex,
                    stroke_w=pano_stroke if stroke_w is None else stroke_w,
                    fill_a=fill_a,
                    dashed=dashed,
                )

            try:
                face_panos = self._panos_for_face()
                for i, pano in enumerate(face_panos):
                    pts = pano.get(u"pts") or []
                    if len(pts) < 3:
                        continue
                    pid = pano[u"id"]
                    sel = pid in (self._pano_selected or set())
                    merge_sel = pid in (self._pano_merge or set())
                    is_active = pid == getattr(self, u"_active_pano_id", None)
                    col = _PANO_COLORS[i % len(_PANO_COLORS)]
                    if merge_sel:
                        _add_pano_poly(
                            pts,
                            _MERGE_FILL,
                            _ACTIVE_STROKE if is_active else _MERGE_STROKE,
                            dashed=False,
                            fill_a=120,
                        )
                    elif is_active:
                        _add_pano_poly(
                            pts,
                            _ACTIVE_FILL,
                            _ACTIVE_STROKE,
                            dashed=False,
                            fill_a=110,
                        )
                    else:
                        _add_pano_poly(
                            pts,
                            col,
                            col,
                            dashed=not sel,
                            fill_a=90 if sel else 40,
                        )
                    try:
                        # Medir tarjeta → offset "I" perp. → centrar etiqueta en paño
                        info = None
                        try:
                            info = _prepare_pano_info_block(
                                pts,
                                pano,
                                is_active=is_active,
                                merge_sel=merge_sel,
                            )
                        except Exception:
                            info = None
                        ltw = float(info[1]) if info else 0.0
                        lth = float(info[2]) if info else 0.0
                        try:
                            _add_major_luz_menor_symbol(
                                pts,
                                is_active=is_active,
                                label_tw=ltw,
                                label_th=lth,
                            )
                        except Exception:
                            pass
                        try:
                            if info is not None:
                                _place_pano_info_block(
                                    pts, info[0], info[1], info[2]
                                )
                        except Exception:
                            pass
                    except Exception:
                        pass
            except Exception:
                pass

            # Punto A pendiente (primer clic)
            try:
                if self._pick_pt1 is not None:
                    px, py = to_px(self._pick_pt1[0], self._pick_pt1[1])
                    r = 5.0
                    el = WpfEllipse()
                    el.Width = r * 2.0
                    el.Height = r * 2.0
                    el.Fill = _brush(u"#5BC0DE")
                    el.Stroke = _brush(u"#E8F4F8")
                    el.StrokeThickness = 1.5
                    WpfCanvas.SetLeft(el, px - r)
                    WpfCanvas.SetTop(el, py - r)
                    scene.Children.Add(el)
                    tb = TextBlock()
                    tb.Text = u"A"
                    tb.Foreground = _brush(u"#5BC0DE")
                    tb.FontSize = 11
                    tb.FontWeight = FontWeights.SemiBold
                    WpfCanvas.SetLeft(tb, px + 7)
                    WpfCanvas.SetTop(tb, py - 10)
                    scene.Children.Add(tb)
            except Exception:
                pass


        hdr = getattr(self, u"_ui_txt_canvas_header", None)
        if hdr is None and self._win is not None:
            hdr = self._win.FindName(u"TxtCanvasHeader")
            self._ui_txt_canvas_header = hdr
        header_text = None
        try:
            nw, nb, npas = _count_ctx(self._overlays)
            npas += int(getattr(self, u"_sketch_holes", 0) or 0)
            face = _normalize_face_id(getattr(self, u"_active_face", u"inferior"))
            n_face = len(self._panos_for_face(face))
            n_sup = len(self._panos_for_face(u"superior"))
            n_inf = len(self._panos_for_face(u"inferior"))
            act_p = self._pano_by_id(getattr(self, u"_active_pano_id", None))
            act_lbl = u"—"
            if act_p is not None:
                act_lbl = act_p.get(u"label") or act_p.get(u"id") or u"—"
            n_ar = len(self._existing_ars or [])
            header_text = (
                u"PLANTA · {:.0f}×{:.0f} mm · "
                u"cara {} · {} paño(s) · SUP {} · INF {} · activo {} · "
                u"AR {} · muros {} · vigas {}"
            ).format(
                bw,
                bh,
                _FACE_PILL.get(face) or face,
                n_face,
                n_sup,
                n_inf,
                act_lbl,
                n_ar,
                nw,
                nb,
            )
        except Exception:
            header_text = None

        paint_planta_context_layers(
            scene=scene,
            hud=hud,
            to_px=to_px,
            add_polygon=_add_polygon,
            loop_polylines=loop_polylines,
            overlays=self._overlays,
            existing_ars=self._existing_ars,
            curves_outer=self._curves,
            plane=plane,
            ctx_geo_cache=self._ensure_ctx_geo_cache(),
            ox0=ox0,
            oy0=oy0,
            min_x=min_x,
            max_x=max_x,
            min_y=min_y,
            max_y=max_y,
            fit_scale=fit_scale,
            bw=bw,
            bh=bh,
            cw=cw,
            ch=ch,
            major_xyz=self._major_direction(),
            mid_layer_callback=_mid_layer,
            header_tb=hdr,
            header_text=header_text,
            context_line_scale=0.5,
        )

        # Snap en mm: solo si paños / pick / overlays invalidaron
        try:
            self._ensure_snap_geometry()
        except Exception:
            self._snap_verts = []
            self._snap_segs = []
            self._snap_cell_index = None
            self._snap_geo_dirty = False

        # Zoom/pan actual + barra escala HUD
        try:
            self._apply_scene_view_transform()
        except Exception:
            pass
        try:
            self._refresh_snap_overlay()
        except Exception:
            pass

    def _execute_crear(self, request=None):
        """Crea ARs desde snapshot (sin controles WPF abiertos)."""
        req = request if request is not None else self._crear_request
        if not req:
            return
        doc = req.get(u"doc")
        floor = req.get(u"floor")
        uidoc = req.get(u"uidoc")
        uiapp = req.get(u"uiapp") or self._uiapp
        plane = req.get(u"plane")
        panos = req.get(u"panos") or []
        outline_pts = list(req.get(u"outline_pts_mm") or [])
        hole_rings = list(req.get(u"pata_hole_rings_mm") or [])

        if doc is None or floor is None:
            _mostrar_aviso(uiapp, u"Documento o Floor no disponible.")
            return
        if not panos:
            _mostrar_aviso(uiapp, u"Defina al menos un paño.")
            return

        any_ahorro = False
        for _p in panos:
            _cfg = _p.get(u"layer_cfg") or {}
            _ai = bool(_p.get(u"ahorro_inferior"))
            _as = bool(_p.get(u"ahorro_superior"))
            if _ai:
                if any(
                    (_cfg.get(k) or {}).get(u"active")
                    for k in (u"interior_major", u"interior_minor")
                ):
                    any_ahorro = True
            if _as:
                if any(
                    (_cfg.get(k) or {}).get(u"active")
                    for k in (u"exterior_major", u"exterior_minor")
                ):
                    any_ahorro = True

        creados = []  # Rebar libres post RemoveAreaSystem (o AR si dissolve falla)
        # Create y post en fases separadas: un Regenerate para todo el batch.
        pending_post = []  # dict(ar_id, tx_label, pbar_label, ahorro)
        errores = []
        mra_avisos = []
        mra_ok_total = 0
        mra_tipo_aviso_hecho = False
        tag_avisos = []
        tag_ok_total = 0
        tag_familia_aviso_hecho = False
        pata_avisos = []

        # Lookups una vez por ExternalEvent / batch Create (no por AR).
        area_type_id = _default_area_type_id(doc)
        if area_type_id == ElementId.InvalidElementId:
            _mostrar_aviso(uiapp, u"No hay AreaReinforcementType en el proyecto.")
            return
        bars = _bar_types_sorted(doc)
        if not bars:
            _mostrar_aviso(uiapp, u"No hay RebarBarType en el proyecto.")
            return
        tag_map = None
        try:
            from enfierrado_shaft_hashtag import _collect_rebar_tag_symbol_map

            tag_map = _collect_rebar_tag_symbol_map(doc, _REBAR_TAG_FAMILY_NAME)
        except Exception:
            tag_map = None
        # BIP de capas resueltos una vez (cache módulo).
        _resolved_layer_bip_enums()

        n_jobs_est = 0
        for _p in panos:
            n_jobs_est += _estimate_pano_job_count(
                _p.get(u"layer_cfg") or {},
                bool(_p.get(u"ahorro_superior")),
                bool(_p.get(u"ahorro_inferior")),
            )
        n_jobs_est = max(1, n_jobs_est)
        # Create + post: ProgressBar con dos fases.
        total = max(1, n_jobs_est * 2)
        # ProgressBar pyRevit (standalone: la UI WPF ya se cerró en Crear).
        pbar = _AreaReinLosaCrearProgress(total)
        pbar.__enter__()
        job_i = 0
        try:
            tg_name = (
                u"Arainco: Area Rein. losa (ahorro fierro)"
                if any_ahorro
                else u"Arainco: Area Rein. losa (paños)"
            )
            tg = TransactionGroup(doc, tg_name)
            tg.Start()
            try:
                for pano in panos:
                    pts = pano.get(u"pts") or []
                    label = pano.get(u"label") or u"paño"
                    # Offset hacia adentro solo para curvas de Create; pts UI intactos.
                    if inset_polygon_mm is None:
                        errores.append(
                            u"{0}: no se pudo aplicar inset {1:g} mm "
                            u"(módulo no disponible).".format(label, _AR_INSET_MM)
                        )
                        continue
                    pts_ar = inset_polygon_mm(pts, _AR_INSET_MM)
                    if not pts_ar or len(pts_ar) < 3:
                        errores.append(
                            u"{0}: no se pudo aplicar inset {1:g} mm "
                            u"(paño demasiado estrecho o polígono inválido).".format(
                                label, _AR_INSET_MM
                            )
                        )
                        continue
                    major_t = pano.get(u"major_xyz")
                    if major_t is not None:
                        try:
                            major = XYZ(
                                float(major_t[0]),
                                float(major_t[1]),
                                float(major_t[2]),
                            )
                        except Exception:
                            major = direccion_arista_mas_larga(self._curves)
                    else:
                        major = direccion_arista_mas_larga(self._curves)

                    layer_cfg = pano.get(u"layer_cfg") or {}
                    ahorro_superior = bool(pano.get(u"ahorro_superior"))
                    ahorro_inferior = bool(pano.get(u"ahorro_inferior"))

                    # Jobs por cara con settings propios del paño
                    create_jobs = _build_pano_create_jobs(
                        pts_ar,
                        layer_cfg,
                        major,
                        plane,
                        label,
                        ahorro_superior,
                        ahorro_inferior,
                        errores,
                    )

                    # Fase 1: Create (+ capas). Rollback por job si Create falla.
                    for job in create_jobs:
                        job_i += 1
                        pbar.update(
                            job_i,
                            label=job.get(u"pbar_label") or label,
                        )
                        curves = _poly_mm_to_curves(job[u"pts"], plane)
                        if not curves:
                            errores.append(
                                u"{0}: polígono inválido para Create.".format(
                                    job.get(u"tx_label") or label
                                )
                            )
                            continue
                        job_ahorro = bool(job.get(u"ahorro"))
                        tx_name = (
                            u"Arainco: Area Rein. ahorro {0}".format(
                                job.get(u"tx_label") or label
                            )
                            if job_ahorro
                            else u"Arainco: Area Rein. {0}".format(
                                job.get(u"tx_label") or label
                            )
                        )
                        t = Transaction(doc, tx_name)
                        t.Start()
                        try:
                            ar, err = crear_area_reinforcement(
                                doc,
                                floor,
                                curves,
                                major,
                                job[u"layer_cfg"],
                                area_type_id=area_type_id,
                                bars=bars,
                            )
                            if ar is None:
                                t.RollBack()
                                errores.append(
                                    u"{0}: {1}".format(
                                        job.get(u"tx_label") or label,
                                        err or u"error",
                                    )
                                )
                                continue
                            # Create committed; post hará RemoveAreaSystem + MRA.
                            t.Commit()
                            _pano_pts = []
                            try:
                                for _pt in pts or []:
                                    _pano_pts.append(
                                        (float(_pt[0]), float(_pt[1]))
                                    )
                            except Exception:
                                _pano_pts = list(pts or [])
                            pending_post.append(
                                {
                                    u"ar_id": ar.Id,
                                    u"tx_label": job.get(u"tx_label") or label,
                                    u"pbar_label": job.get(u"pbar_label") or label,
                                    u"ahorro": job_ahorro,
                                    u"face": _normalize_face_id(
                                        pano.get(u"face")
                                    ),
                                    u"pano_pts": _pano_pts,
                                }
                            )
                        except Exception as ex:
                            try:
                                if t.HasStarted():
                                    t.RollBack()
                            except Exception:
                                pass
                            errores.append(
                                u"{0}: {1}".format(
                                    job.get(u"tx_label") or label,
                                    _as_unicode(ex),
                                )
                            )

                # Fase 2: un solo Regenerate para materializar hijos de todos los AR.
                if pending_post:
                    t_regen = Transaction(doc, u"Arainco: Area Rein. regenerar")
                    t_regen.Start()
                    try:
                        doc.Regenerate()
                        t_regen.Commit()
                    except Exception as ex_regen:
                        try:
                            if t_regen.HasStarted():
                                t_regen.RollBack()
                        except Exception:
                            pass
                        errores.append(
                            u"Regenerate batch: {0}".format(_as_unicode(ex_regen))
                        )

                # Fase 3: post por AR — RemoveAreaSystem + Show Middle/stamps/tags/MRA.
                # Create ya committed; si post falla, el AR puede quedar sin anotar.
                for item in pending_post:
                    job_i += 1
                    pbar_label = item.get(u"pbar_label") or item.get(u"tx_label")
                    pbar.update(
                        job_i,
                        label=u"Post {0}".format(pbar_label),
                    )
                    ar = None
                    try:
                        ar = doc.GetElement(item[u"ar_id"])
                    except Exception:
                        ar = None
                    if ar is None or not isinstance(ar, AreaReinforcement):
                        errores.append(
                            u"{0}: AR no disponible tras Create.".format(
                                item.get(u"tx_label") or u"paño"
                            )
                        )
                        continue
                    tx_post = (
                        u"Arainco: Area Rein. post ahorro {0}".format(
                            item.get(u"tx_label") or u"paño"
                        )
                        if bool(item.get(u"ahorro"))
                        else u"Arainco: Area Rein. post {0}".format(
                            item.get(u"tx_label") or u"paño"
                        )
                    )
                    free_rebars = []
                    # Patas L: aristas del paño ∩ outline / shafts / huecos
                    pata_ctx = None
                    try:
                        _ppt = item.get(u"pano_pts") or []
                        if (
                            _FEATURE_PATA_L
                            and _ppt
                            and (outline_pts or hole_rings)
                        ):
                            pata_ctx = {
                                u"floor": floor,
                                u"plane": plane,
                                u"face": item.get(u"face"),
                                u"outline_pts": outline_pts,
                                u"hole_rings": hole_rings,
                                u"pano_pts": _ppt,
                                u"enabled": True,
                            }
                    except Exception:
                        pata_ctx = None
                    t = Transaction(doc, tx_post)
                    t.Start()
                    # Patas L pueden quedar fuera del host: silenciar warning Revit
                    try:
                        if attach_rebar_outside_host_swallower is not None:
                            attach_rebar_outside_host_swallower(t)
                    except Exception:
                        pass
                    try:
                        (
                            mra_ok_total,
                            mra_tipo_aviso_hecho,
                            tag_ok_total,
                            tag_familia_aviso_hecho,
                        ) = _post_create_area_reinforcement(
                            doc,
                            ar,
                            uidoc,
                            mra_avisos,
                            mra_ok_total,
                            mra_tipo_aviso_hecho,
                            tag_avisos,
                            tag_ok_total,
                            tag_familia_aviso_hecho,
                            tag_map=tag_map,
                            skip_regenerate=True,
                            allow_retry_regenerate=True,
                            out_free_rebars=free_rebars,
                            pata_ctx=pata_ctx,
                            pata_avisos=pata_avisos,
                        )
                        t.Commit()
                        if free_rebars:
                            creados.extend(free_rebars)
                        else:
                            try:
                                ar_keep = doc.GetElement(item[u"ar_id"])
                                if ar_keep is not None and isinstance(
                                    ar_keep, AreaReinforcement
                                ):
                                    creados.append(ar_keep)
                            except Exception:
                                pass
                    except Exception as ex:
                        try:
                            if t.HasStarted():
                                t.RollBack()
                        except Exception:
                            pass
                        errores.append(
                            u"{0}: post {1}".format(
                                item.get(u"tx_label") or u"paño",
                                _as_unicode(ex),
                            )
                        )
                        # AR creado queda en modelo aunque falle el post.
                        try:
                            ar_keep = doc.GetElement(item[u"ar_id"])
                            if ar_keep is not None and isinstance(
                                ar_keep, AreaReinforcement
                            ):
                                creados.append(ar_keep)
                        except Exception:
                            pass

                if creados:
                    tg.Assimilate()
                else:
                    tg.RollBack()
            except Exception as ex:
                try:
                    tg.RollBack()
                except Exception:
                    pass
                _mostrar_aviso(
                    uiapp,
                    u"Error en TransactionGroup.",
                    content=_as_unicode(ex),
                )
                return
            finally:
                if finalizar_armadura_conjunto_guid_ejecucion is not None:
                    try:
                        finalizar_armadura_conjunto_guid_ejecucion()
                    except Exception:
                        pass
        finally:
            try:
                pbar.__exit__(None, None, None)
            except Exception:
                pass

        if creados:
            try:
                if uidoc is not None:
                    ids_sel = []
                    for el in creados:
                        try:
                            el2 = doc.GetElement(el.Id)
                        except Exception:
                            el2 = el
                        if el2 is not None:
                            ids_sel.append(el2.Id)
                    if ids_sel:
                        uidoc.Selection.SetElementIds(List[ElementId](ids_sel))
            except Exception:
                pass
            # Resumen MRA / etiquetas (soft-fail: no bloquea creación)
            try:
                parts = [
                    u"Rebar: {0}".format(len(creados)),
                    u"MRA: {0}".format(int(mra_ok_total)),
                    u"Etiquetas: {0}".format(int(tag_ok_total)),
                ]
                extra = []
                for av in (pata_avisos or [])[:3]:
                    extra.append(av)
                for av in (mra_avisos or [])[:4]:
                    extra.append(av)
                for av in (tag_avisos or [])[:2]:
                    extra.append(av)
                for av in (errores or [])[:3]:
                    extra.append(av)
                self._set_status(
                    u" · ".join(parts)
                    + (u" — " + u" | ".join(extra) if extra else u"")
                )
            except Exception:
                pass
            if mra_avisos and int(mra_ok_total) <= 0:
                try:
                    _mostrar_aviso(
                        uiapp,
                        u"Armadura creada, pero sin Multi-Rebar Annotation.",
                        content=u"\n".join((mra_avisos or [])[:8]),
                    )
                except Exception:
                    pass
        else:
            _mostrar_aviso(
                uiapp,
                u"No se pudo crear AreaReinforcement.",
                content=u"\n".join(errores[:10]) if errores else u"",
            )

    def _force_canvas_size_for_paint(self, ww, wh):
        """Fija tamaño del canvas cuando el layout * aún no midió (sin Show)."""
        cv = self._get_cv_plan()
        if cv is None:
            return False
        try:
            from System.Windows import Rect, Size
        except Exception:
            return False
        try:
            ww = max(900.0, float(ww))
            wh = max(640.0, float(wh))
        except Exception:
            ww, wh = 1280.0, 800.0
        # Columna derecha ~360 + márgenes; filas chrome ~200
        cw = max(400.0, ww - 360.0 - 56.0)
        ch = max(300.0, wh - 200.0)
        try:
            cv.Width = cw
            cv.Height = ch
            cv.Measure(Size(cw, ch))
            cv.Arrange(Rect(0.0, 0.0, cw, ch))
            cv.UpdateLayout()
        except Exception:
            try:
                cv.Width = cw
                cv.Height = ch
            except Exception:
                return False
        try:
            aw = float(getattr(cv, u"ActualWidth", 0.0) or 0.0)
            ah = float(getattr(cv, u"ActualHeight", 0.0) or 0.0)
        except Exception:
            aw = ah = 0.0
        if aw < 40.0 or ah < 40.0:
            # Actual* a veces sigue en 0 offline; Width/Height bastan para paint
            try:
                if float(cv.Width or 0.0) >= 40.0 and float(cv.Height or 0.0) >= 40.0:
                    return True
            except Exception:
                pass
            return False
        return True

    def _layout_window_offline(self):
        """Measure/Arrange sin Show(): permite pintar el canvas antes de levantar HWND."""
        win = self._win
        if win is None:
            return
        w = h = 1280.0
        try:
            from System.Windows import Rect, Size, SystemParameters, WindowStartupLocation

            wa = SystemParameters.WorkArea
            w = max(900.0, float(wa.Width))
            h = max(640.0, float(wa.Height))
            win.WindowStartupLocation = WindowStartupLocation.Manual
            win.WindowState = WindowState.Normal
            win.Width = w
            win.Height = h
            try:
                win.ShowInTaskbar = False
            except Exception:
                pass
            win.Measure(Size(w, h))
            win.Arrange(Rect(0.0, 0.0, w, h))
            win.UpdateLayout()
        except Exception:
            try:
                from System.Windows import Rect, Size

                w = h = 1280.0
                win.Width = w
                win.Height = h
                win.Measure(Size(w, h))
                win.Arrange(Rect(0.0, 0.0, w, h))
                win.UpdateLayout()
            except Exception:
                pass
        # Grid * casi nunca entrega ActualWidth al canvas sin HWND: forzar siempre.
        self._force_canvas_size_for_paint(w, h)

    def _canvas_paint_size(self):
        """Tamaño usable para paint (Actual* o Width/Height forzado)."""
        cv = self._get_cv_plan()
        if cv is None:
            return 0.0, 0.0
        try:
            cw = float(getattr(cv, u"ActualWidth", 0.0) or 0.0)
            ch = float(getattr(cv, u"ActualHeight", 0.0) or 0.0)
        except Exception:
            cw = ch = 0.0
        if cw >= 40.0 and ch >= 40.0:
            return cw, ch
        try:
            cw = float(cv.Width or 0.0)
            ch = float(cv.Height or 0.0)
        except Exception:
            cw = ch = 0.0
        return cw, ch

    def _prepare_ui_content(self):
        """Monta geo + canvas + snap sin mostrar la ventana."""
        if getattr(self, u"_ui_prepare_done", False):
            return
        try:
            _set_wait_cursor(True)
        except Exception:
            pass
        self._size_redraw_pending = False
        try:
            t = getattr(self, u"_size_redraw_timer", None)
            if t is not None:
                t.Stop()
        except Exception:
            pass
        try:
            self._cache_ui_refs()
        except Exception:
            pass
        try:
            self._ensure_ctx_geo_cache()
        except Exception:
            pass
        # Reintentar layout si el canvas aún no tiene tamaño pintable.
        cw, ch = self._canvas_paint_size()
        if cw < 40.0 or ch < 40.0:
            try:
                ww = float(self._win.Width) if self._win is not None else 1280.0
                wh = float(self._win.Height) if self._win is not None else 800.0
            except Exception:
                ww, wh = 1280.0, 800.0
            self._force_canvas_size_for_paint(ww, wh)
        try:
            self._redraw_canvas()
        except Exception:
            pass
        # Si Actual* era 0, _redraw_canvas salió sin pintar: usar Width/Height.
        try:
            cv = self._get_cv_plan()
            nch = 0
            if cv is not None:
                try:
                    nch = int(cv.Children.Count)
                except Exception:
                    nch = 0
            if nch <= 0:
                self._redraw_canvas_with_forced_size()
        except Exception:
            pass
        try:
            self._ensure_snap_geometry()
        except Exception:
            pass
        self._ui_prepare_done = True

    def _redraw_canvas_with_forced_size(self):
        """Paint usando Width/Height del canvas si Actual* sigue en 0 (offline)."""
        cv = self._get_cv_plan()
        if cv is None:
            return
        try:
            cw = float(cv.ActualWidth or 0.0)
            ch = float(cv.ActualHeight or 0.0)
        except Exception:
            cw = ch = 0.0
        if cw >= 40.0 and ch >= 40.0:
            self._redraw_canvas()
            return
        try:
            fw = float(cv.Width or 0.0)
            fh = float(cv.Height or 0.0)
        except Exception:
            return
        if fw < 40.0 or fh < 40.0:
            return
        # Temporalmente reportar tamaño forzado vía Width (Actual puede seguir 0).
        # Parche mínimo: asignar MinWidth/MinHeight y volver a medir.
        try:
            from System.Windows import Rect, Size

            cv.MinWidth = fw
            cv.MinHeight = fh
            cv.Measure(Size(fw, fh))
            cv.Arrange(Rect(0.0, 0.0, fw, fh))
            cv.UpdateLayout()
        except Exception:
            pass
        try:
            self._redraw_canvas()
        except Exception:
            pass

    def show(self, on_before_reveal=None):
        """Prepara todo sin Show(); levanta la ventana solo cuando el canvas está listo."""
        win = self._win
        if win is None:
            return
        self._on_before_reveal = on_before_reveal
        try:
            _set_wait_cursor(True)
        except Exception:
            pass

        self._ui_revealed = False
        self._ui_reveal_retry = False
        # 1) Layout sin HWND  2) paint completo  3) Owner + un solo Show ya listo
        self._layout_window_offline()
        self._prepare_ui_content()
        try:
            from System.Windows.Interop import WindowInteropHelper

            hwnd = revit_main_hwnd(self._uiapp)
            if hwnd is not None:
                WindowInteropHelper(win).Owner = hwnd
        except Exception:
            pass
        try:
            hwnd = revit_main_hwnd(self._uiapp)
            position_wpf_window_top_left_at_active_view(
                win, self._uidoc, hwnd
            )
        except Exception:
            pass
        try:
            win.ShowActivated = True
        except Exception:
            pass
        try:
            win.Opacity = 1.0
        except Exception:
            pass
        try:
            win.WindowState = WindowState.Maximized
        except Exception:
            pass
        _register_singleton(win)
        try:
            win.Show()
        except Exception:
            try:
                cb = getattr(self, u"_on_before_reveal", None)
                if cb is not None:
                    self._on_before_reveal = None
                    cb()
                else:
                    _set_wait_cursor(False)
            except Exception:
                pass
            raise
        self._ui_revealed = True
        # Ajustar al tamaño real maximizado sin dejar el canvas vacío un frame.
        try:
            win.UpdateLayout()
        except Exception:
            pass
        try:
            from System.Windows import FrameworkElement

            cv = self._get_cv_plan()
            if cv is not None:
                cv.ClearValue(FrameworkElement.WidthProperty)
                cv.ClearValue(FrameworkElement.HeightProperty)
                try:
                    cv.ClearValue(FrameworkElement.MinWidthProperty)
                    cv.ClearValue(FrameworkElement.MinHeightProperty)
                except Exception:
                    pass
            win.UpdateLayout()
        except Exception:
            pass
        try:
            self._cache_ui_refs()
            self._redraw_canvas()
        except Exception:
            pass
        try:
            win.Activate()
            win.Focus()
            cv = self._get_cv_plan()
            if cv is not None:
                cv.Focus()
        except Exception:
            pass
        try:
            cb = getattr(self, u"_on_before_reveal", None)
            if cb is not None:
                self._on_before_reveal = None
                cb()
            else:
                _set_wait_cursor(False)
        except Exception:
            try:
                _set_wait_cursor(False)
            except Exception:
                pass
        self._on_before_reveal = None


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------


def run(revit):
    """Punto de entrada pyRevit / scripts."""
    uiapp = revit
    try:
        uidoc = uiapp.ActiveUIDocument
    except Exception:
        uidoc = None
    if uidoc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    doc = uidoc.Document
    if doc is None:
        _mostrar_aviso(uiapp, u"No hay documento activo.")
        return
    if doc.IsFamilyDocument:
        _mostrar_aviso(uiapp, u"Abra un proyecto (no un family document).")
        return

    try:
        active_view = uidoc.ActiveView
    except Exception:
        active_view = None
    if not _vista_es_planta(active_view):
        _mostrar_aviso(
            uiapp,
            u"Esta herramienta solo funciona en vistas de planta.",
            content=u"Abra una planta (ViewPlan) y vuelva a ejecutar.",
        )
        return

    if _focus_existing(uiapp):
        return

    floor = _pick_floor(uidoc, doc, uiapp)
    if floor is None:
        return

    loops = obtener_loops_sketch(floor, doc)
    if not loops:
        _mostrar_aviso(
            uiapp,
            u"La losa no tiene Sketch válido.",
            content=u"Se requiere Floor.SketchId con Profile (loop exterior).",
        )
        return

    outer = loops[0]
    plane = _plane_from_curves(outer)
    if plane is None:
        _mostrar_aviso(uiapp, u"No se pudo obtener el plano del Sketch.")
        return

    try:
        _set_wait_cursor(True)
        try:
            ctrl = AreaReinLosaSketchController(
                uiapp, uidoc, doc, floor, outer, loops, plane
            )
            # Cursor Wait hasta reveal (ventana ya completa y visible).
            ctrl.show(on_before_reveal=lambda: _set_wait_cursor(False))
        finally:
            _set_wait_cursor(False)
    except Exception as ex:
        _set_wait_cursor(False)
        _unregister_singleton()
        _mostrar_aviso(
            uiapp,
            u"Error al abrir la UI.",
            content=_as_unicode(ex),
        )
        return


def run_pyrevit(revit):
    run(revit)
