# -*- coding: utf-8 -*-
"""
Arainco: Armado fundación aislada (Sketch).

Selección de fundación aislada + canvas planta / sección + tabs Inferior /
Superior / Lateral (ø y sep. por luz en mallas). Colocar crea Rebar con el
motor CreateFromCurves + Maximum Spacing (ø/sep independientes por luz).

Revit 2024–2026 · IronPython (pyRevit).
"""

from __future__ import print_function

import math
import os
import sys

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System import AppDomain, EventHandler
from System.Windows import (
    CornerRadius,
    FontWeights,
    HorizontalAlignment,
    Point as WpfPoint,
    RoutedEventHandler,
    Thickness,
    VerticalAlignment,
    WindowState,
)
from System.Windows.Controls import (
    Border,
    Canvas as WpfCanvas,
    Orientation,
    SelectionChangedEventHandler,
    StackPanel,
    TextBlock,
)
from System.Windows.Input import (
    MouseButtonEventHandler,
    MouseEventHandler,
    MouseWheelEventHandler,
    TextCompositionEventHandler,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import Color, SolidColorBrush, TranslateTransform
from System.Windows.Shapes import Ellipse as WpfEllipse
from System.Windows.Shapes import Line as WpfLine
from System.Windows.Shapes import Polygon as WpfPolygon
from System.Windows.Shapes import Rectangle as WpfRectangle

from Autodesk.Revit.DB import BuiltInCategory, ViewPlan, WallFoundation
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

from bimtools_ui_tokens import ACCENT_PRIMARY, WINDOW_CHROME_TITLE
from bimtools_wpf_dark_theme import BIMTOOLS_DARK_STYLES_XML
from bimtools_instruction_dialog import show_message_dialog
from revit_wpf_window_position import (
    position_wpf_window_top_left_at_active_view,
    revit_main_hwnd,
)
from armado_fundacion_aislada_crear import (
    DESCUENTO_PATA_U_MM,
    OFFSET_PRIMERA_LATERAL_MM,
    REC_HORIZONTAL_EJE_MM,
    REC_LATERAL_CARA_MM,
    REC_PLANTA_MALLA_MM,
    crear_armadura_fundacion_aislada,
)
from geometria_fundacion_cara_inferior import (
    RECUBRIMIENTO_EXTREMOS_MM,
    _iter_curvas_en_curveloop,
    elegir_loop_mayor_perimetro,
    extraer_curvas_perimetrales_cara_inferior,
    largo_gancho_u_tabla_mm,
    rango_z_caras_laterales_o_bbox,
)

try:
    from barras_bordes_losa_gancho_empotramiento import (
        _build_bar_type_entries,
        _rebar_nominal_diameter_mm,
        element_id_to_int,
    )
except Exception:
    _build_bar_type_entries = None
    _rebar_nominal_diameter_mm = None

    def element_id_to_int(eid):
        if eid is None:
            return None
        try:
            return int(eid.IntegerValue)
        except Exception:
            pass
        try:
            return int(eid.Value)
        except Exception:
            return None


_DIALOG_TITLE = u"Arainco: Fundación aislada"
_SINGLETON_KEY = u"Arainco.ArmadoFundacionAisladaSketch.ActiveWindow"
_SINGLETON_CTRL_KEY = u"Arainco.ArmadoFundacionAisladaSketch.ActiveController"
_FT_TO_MM = 304.8
_FOUNDATION_CAT_ID = int(BuiltInCategory.OST_StructuralFoundation)
_BRUSH_CACHE = {}

_PLAN_PAD_FRAC = 0.08
_SECTION_PAD_PX = 28.0

_SEP_MM_MIN = 100
_SEP_MM_MAX = 300
_SEP_MM_STEP = 10
_SEP_MM_DEFAULT = 150
_SEP_LAT_DEFAULT = 200
# Preview: mismos recubrimientos que armado_fundacion_aislada_crear.
_COVER_PREVIEW_MM = float(REC_HORIZONTAL_EJE_MM)
_REC_PLANTA_PREVIEW_MM = float(REC_PLANTA_MALLA_MM)
_REC_EXTREMOS_PREVIEW_MM = float(RECUBRIMIENTO_EXTREMOS_MM)

_CARD_KEYS = (u"inferior", u"superior", u"lateral")
_DIR_KEYS = (u"luz_mayor", u"luz_menor")
_CARD_UI = {
    u"inferior": {
        u"chk": u"ChkInferior",
        u"panel": u"PanelInferior",
        u"title": u"Inferior",
        u"color": u"#4ade80",
        u"dirs": {
            u"luz_mayor": {
                u"cmb": u"CmbInfMayorDiam",
                u"sep": u"TxtInfMayorSepMm",
                u"label": u"Luz mayor",
            },
            u"luz_menor": {
                u"cmb": u"CmbInfMenorDiam",
                u"sep": u"TxtInfMenorSepMm",
                u"label": u"Luz menor",
            },
        },
    },
    u"superior": {
        u"chk": u"ChkSuperior",
        u"panel": u"PanelSuperior",
        u"title": u"Superior",
        u"color": u"#5BC0DE",
        u"dirs": {
            u"luz_mayor": {
                u"cmb": u"CmbSupMayorDiam",
                u"sep": u"TxtSupMayorSepMm",
                u"label": u"Luz mayor",
            },
            u"luz_menor": {
                u"cmb": u"CmbSupMenorDiam",
                u"sep": u"TxtSupMenorSepMm",
                u"label": u"Luz menor",
            },
        },
    },
    u"lateral": {
        u"chk": u"ChkLateral",
        u"panel": u"PanelLateral",
        u"title": u"Lateral",
        u"color": u"#fbbf24",
        u"cmb": u"CmbLatDiam",
        u"sep": u"TxtLatSepMm",
    },
}


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


def _normalize_sep_textbox(tb, default_val=_SEP_MM_DEFAULT):
    if tb is None:
        return
    try:
        s = _as_unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            tb.Text = _as_unicode(int(default_val))
            return
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        tb.Text = _as_unicode(int(default_val))
        return
    n = max(_SEP_MM_MIN, min(_SEP_MM_MAX, n))
    nmax = int((_SEP_MM_MAX - _SEP_MM_MIN) // _SEP_MM_STEP)
    steps = int(round((n - _SEP_MM_MIN) / float(_SEP_MM_STEP)))
    steps = max(0, min(nmax, steps))
    n = _SEP_MM_MIN + steps * _SEP_MM_STEP
    tb.Text = _as_unicode(int(n))


def _leer_sep_mm(tb, default_val=_SEP_MM_DEFAULT):
    if tb is None:
        return float(default_val)
    try:
        s = _as_unicode(tb.Text).replace(u"mm", u"").strip()
        if not s:
            return float(default_val)
        n = int(round(float(s.replace(u",", u"."))))
    except Exception:
        return float(default_val)
    n = max(_SEP_MM_MIN, min(_SEP_MM_MAX, n))
    nmax = int((_SEP_MM_MAX - _SEP_MM_MIN) // _SEP_MM_STEP)
    steps = int(round((n - _SEP_MM_MIN) / float(_SEP_MM_STEP)))
    steps = max(0, min(nmax, steps))
    return float(_SEP_MM_MIN + steps * _SEP_MM_STEP)


def _sep_text_is_digits_only(text):
    if text is None:
        return False
    s = _as_unicode(text)
    if not s:
        return True
    for ch in s:
        if ch < u"0" or ch > u"9":
            return False
    return True


def _parse_hex_rgb(hex_color):
    h = (_as_unicode(hex_color) or u"#5BC0DE").lstrip(u"#")
    if len(h) < 6:
        return (0x5B, 0xC0, 0xDE)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return (0x5B, 0xC0, 0xDE)


def _apply_toggle_switch(chk, label_text, parts, accent_hex=None):
    """CheckBox BimToolsToggleMini -> track+thumb + etiqueta."""
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


def _diam_sep_row_xaml(cmb_name, sep_name, tip_diam, tip_sep, default_sep):
    """Fila ø + @ + caja numérica de separación (sin steppers)."""
    return (
        u'<Grid HorizontalAlignment="Stretch">'
        u'<Grid.ColumnDefinitions>'
        u'<ColumnDefinition Width="*"/>'
        u'<ColumnDefinition Width="Auto"/>'
        u'<ColumnDefinition Width="96"/>'
        u"</Grid.ColumnDefinitions>"
        u'<ComboBox Grid.Column="0" x:Name="{cmb}" Style="{{StaticResource Combo}}" '
        u'IsEditable="False" IsReadOnly="True" MinWidth="80" ToolTip="{tip_d}">'
        u"<ComboBox.ItemContainerStyle>"
        u'<Style TargetType="ComboBoxItem" BasedOn="{{StaticResource ComboItem}}"/>'
        u"</ComboBox.ItemContainerStyle>"
        u"</ComboBox>"
        u'<TextBlock Grid.Column="1" Text="@" FontSize="12" FontWeight="Bold" '
        u'Foreground="#95B8CC" VerticalAlignment="Center" '
        u'HorizontalAlignment="Center" Margin="4,0,4,0"/>'
        u'<TextBox Grid.Column="2" x:Name="{sep}" '
        u'Style="{{StaticResource BimToolsTextBoxDark}}" '
        u'Text="{def_sep}" Height="24" Padding="6,0,6,0" '
        u'VerticalContentAlignment="Center" HorizontalContentAlignment="Center" '
        u'MaxLength="4" ToolTip="{tip_s}"/>'
        u"</Grid>"
    ).format(
        cmb=cmb_name,
        sep=sep_name,
        tip_d=tip_diam,
        tip_s=tip_sep,
        def_sep=_as_unicode(int(default_sep)),
    )


def _mesh_panel_xaml(panel_name, title, mayor_pfx, menor_pfx, chk_name):
    """Contenido de un tab Inferior/Superior (toggle + luz mayor/menor)."""
    row_mayor = _diam_sep_row_xaml(
        u"Cmb{0}Diam".format(mayor_pfx),
        u"Txt{0}SepMm".format(mayor_pfx),
        u"Diámetro — {0} · luz mayor".format(title),
        u"Separación luz mayor (mm). Solo números; rango 100–300.",
        _SEP_MM_DEFAULT,
    )
    row_menor = _diam_sep_row_xaml(
        u"Cmb{0}Diam".format(menor_pfx),
        u"Txt{0}SepMm".format(menor_pfx),
        u"Diámetro — {0} · luz menor".format(title),
        u"Separación luz menor (mm). Solo números; rango 100–300.",
        _SEP_MM_DEFAULT,
    )
    return (
        u'<StackPanel x:Name="{panel}" Visibility="Collapsed">'
        u'<CheckBox x:Name="{chk}" Style="{{StaticResource BimToolsToggleMini}}" '
        u'IsChecked="True" Margin="0,0,0,10" VerticalAlignment="Center"/>'
        u'<TextBlock Text="Luz mayor" Foreground="#95B8CC" FontSize="10" '
        u'FontWeight="SemiBold" Margin="0,0,0,4"/>'
        u"{row_mayor}"
        u'<TextBlock Text="Luz menor" Foreground="#95B8CC" FontSize="10" '
        u'FontWeight="SemiBold" Margin="0,8,0,4"/>'
        u"{row_menor}"
        u"</StackPanel>"
    ).format(
        panel=panel_name, chk=chk_name, row_mayor=row_mayor, row_menor=row_menor
    )


_CARD_XAML = (
    u'<Border Background="#0a1620" BorderBrush="#21465C" BorderThickness="1" '
    u'CornerRadius="4" Padding="8" Margin="0,0,0,0">'
    u"<StackPanel>"
    u'<Grid x:Name="PnlTabs" Margin="0,0,0,10"/>'
    u'<StackPanel x:Name="PnlTabBody">'
    + _mesh_panel_xaml(
        u"PanelInferior", u"Inferior", u"InfMayor", u"InfMenor", u"ChkInferior"
    )
    + _mesh_panel_xaml(
        u"PanelSuperior", u"Superior", u"SupMayor", u"SupMenor", u"ChkSuperior"
    )
    + (
        u'<StackPanel x:Name="PanelLateral" Visibility="Collapsed">'
        u'<CheckBox x:Name="ChkLateral" Style="{StaticResource BimToolsToggleMini}" '
        u'IsChecked="True" Margin="0,0,0,10" VerticalAlignment="Center"/>'
        + _diam_sep_row_xaml(
            u"CmbLatDiam",
            u"TxtLatSepMm",
            u"Diámetro — armadura lateral",
            u"Distanciamiento vertical (mm). Solo números; rango 100–300.",
            _SEP_LAT_DEFAULT,
        )
        + u"</StackPanel>"
    )
    + u"</StackPanel></StackPanel></Border>"
)



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


def _register_singleton(win, ctrl=None):
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, win)
    except Exception:
        pass
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_CTRL_KEY, ctrl)
    except Exception:
        pass


def _unregister_singleton():
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_KEY, None)
    except Exception:
        pass
    try:
        AppDomain.CurrentDomain.SetData(_SINGLETON_CTRL_KEY, None)
    except Exception:
        pass


def _vista_es_planta(view):
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


# ---------------------------------------------------------------------------
# Selección / geometría
# ---------------------------------------------------------------------------


class FundacionAisladaSelectionFilter(ISelectionFilter):
    """Structural Foundation; excluye WallFoundation."""

    def AllowElement(self, elem):
        try:
            if elem is None:
                return False
            if isinstance(elem, WallFoundation):
                return False
            cat = elem.Category
            if cat is None:
                return False
            return element_id_to_int(cat.Id) == _FOUNDATION_CAT_ID
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


def _sample_curve_xy_mm(curve, n_seg=12):
    """Puntos XY (mm) a lo largo de una curva Revit."""
    pts = []
    if curve is None:
        return pts
    try:
        n = max(2, int(n_seg))
        for i in range(n + 1):
            t = float(i) / float(n)
            try:
                p = curve.Evaluate(t, True)
            except Exception:
                if i == 0:
                    p = curve.GetEndPoint(0)
                elif i == n:
                    p = curve.GetEndPoint(1)
                else:
                    continue
            pts.append((float(p.X) * _FT_TO_MM, float(p.Y) * _FT_TO_MM))
    except Exception:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            pts = [
                (float(p0.X) * _FT_TO_MM, float(p0.Y) * _FT_TO_MM),
                (float(p1.X) * _FT_TO_MM, float(p1.Y) * _FT_TO_MM),
            ]
        except Exception:
            pass
    return pts


def _loop_to_polygon_mm(curve_loop):
    """Polígono cerrado XY mm desde un CurveLoop (sin duplicar cierre)."""
    pts = []
    for c in _iter_curvas_en_curveloop(curve_loop):
        seg = _sample_curve_xy_mm(c, n_seg=16)
        if not seg:
            continue
        if pts:
            # evitar duplicar extremo compartido
            last = pts[-1]
            first = seg[0]
            if abs(last[0] - first[0]) < 0.5 and abs(last[1] - first[1]) < 0.5:
                seg = seg[1:]
        pts.extend(seg)
    if len(pts) >= 3:
        first = pts[0]
        last = pts[-1]
        if abs(first[0] - last[0]) < 0.5 and abs(first[1] - last[1]) < 0.5:
            pts = pts[:-1]
    return pts if len(pts) >= 3 else None


def _bbox_fallback_polygon_mm(elem):
    try:
        bb = elem.get_BoundingBox(None)
        if bb is None:
            return None
        x0 = float(bb.Min.X) * _FT_TO_MM
        y0 = float(bb.Min.Y) * _FT_TO_MM
        x1 = float(bb.Max.X) * _FT_TO_MM
        y1 = float(bb.Max.Y) * _FT_TO_MM
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    except Exception:
        return None


def extract_foundation_plan_polygons_mm(elem):
    """
    Lista de polígonos XY mm (exterior primero).

    Returns:
        list[list[(x,y)]] o lista vacía.
    """
    out = []
    r = None
    try:
        r = extraer_curvas_perimetrales_cara_inferior(elem)
    except Exception:
        r = None
    if r is not None:
        loops, _z = r
        ordered = []
        best = elegir_loop_mayor_perimetro(loops) if loops else None
        if best is not None:
            ordered.append(best)
            for cl in loops or []:
                if cl is best:
                    continue
                ordered.append(cl)
        else:
            ordered = list(loops or [])
        for cl in ordered:
            poly = _loop_to_polygon_mm(cl)
            if poly:
                out.append(poly)
    if not out:
        fb = _bbox_fallback_polygon_mm(elem)
        if fb:
            out.append(fb)
    return out


def _poly_edge_lengths_dirs(poly):
    """Lista (length, ux, uy, x0, y0, x1, y1) por arista del polígono."""
    out = []
    if not poly or len(poly) < 2:
        return out
    n = len(poly)
    for i in range(n):
        x0, y0 = float(poly[i][0]), float(poly[i][1])
        x1, y1 = float(poly[(i + 1) % n][0]), float(poly[(i + 1) % n][1])
        dx, dy = x1 - x0, y1 - y0
        L = math.hypot(dx, dy)
        if L < 1.0:
            continue
        out.append((L, dx / L, dy / L, x0, y0, x1, y1))
    return out


def oriented_plan_frame_mm(plan_polygons):
    """
    Marco local de la zapata en planta (mm), alineado a aristas reales.

    Evita usar el AABB mundo (que en zapatas rotadas 45° da la diagonal).

    Returns dict:
      ox, oy          origen (proyección)
      u_hat, v_hat    unitarios: u // lado corto (luz menor), v // lado largo
      len_u, len_v    luces (mm) a lo largo de u y v
      u_min, u_max, v_min, v_max  rangos de proyección
    o None.
    """
    if not plan_polygons:
        return None
    poly = plan_polygons[0]
    if not poly or len(poly) < 3:
        return None
    edges = _poly_edge_lengths_dirs(poly)
    if not edges:
        return None

    # Dirección del lado más corto ≈ luz menor; la perpendicular ≈ luz mayor.
    edges_sorted = sorted(edges, key=lambda e: e[0])
    L_short, ux, uy = edges_sorted[0][0], edges_sorted[0][1], edges_sorted[0][2]
    # Buscar arista casi perpendicular (producto punto ~0) como lado largo.
    vx = vy = None
    L_long = L_short
    best_abs_dot = 1.0
    for L, ex, ey, _a, _b, _c, _d in edges:
        dot = abs(ex * ux + ey * uy)
        if dot < best_abs_dot:
            best_abs_dot = dot
            vx, vy = ex, ey
            L_long = L
    if vx is None or best_abs_dot > 0.35:
        # Fallback: girar 90° la dirección corta.
        vx, vy = -uy, ux
        L_long = L_short

    # Orientar v para que (u,v) sea dextrógiro en planta.
    if ux * vy - uy * vx < 0.0:
        vx, vy = -vx, -vy

    ox = float(poly[0][0])
    oy = float(poly[0][1])
    us = []
    vs = []
    for p in poly:
        px, py = float(p[0]), float(p[1])
        us.append((px - ox) * ux + (py - oy) * uy)
        vs.append((px - ox) * vx + (py - oy) * vy)
    u_min, u_max = min(us), max(us)
    v_min, v_max = min(vs), max(vs)
    len_u = max(1.0, u_max - u_min)
    len_v = max(1.0, v_max - v_min)

    # Asegurar u = lado corto, v = lado largo (como motor menor/mayor).
    if len_u > len_v + 1.0:
        ux, uy, vx, vy = vx, vy, ux, uy
        u_min, u_max, v_min, v_max = v_min, v_max, u_min, u_max
        len_u, len_v = len_v, len_u

    return {
        u"ox": ox,
        u"oy": oy,
        u"u_hat": (ux, uy),
        u"v_hat": (vx, vy),
        u"u_min": u_min,
        u"u_max": u_max,
        u"v_min": v_min,
        u"v_max": v_max,
        u"len_u": len_u,
        u"len_v": len_v,
        u"luz_menor_mm": len_u,
        u"luz_mayor_mm": len_v,
    }


def _frame_to_world(frame, u, v):
    """Punto mundo (x,y) mm desde coords locales (u,v) del marco orientado."""
    ox = float(frame[u"ox"])
    oy = float(frame[u"oy"])
    ux, uy = frame[u"u_hat"]
    vx, vy = frame[u"v_hat"]
    return (
        ox + float(u) * ux + float(v) * vx,
        oy + float(u) * uy + float(v) * vy,
    )


def extract_section_dims_mm(elem, plan_polygons):
    """
    (width_mm, height_mm) para el esquema de sección.

    Ancho = **luz menor** real (lado corto del perímetro orientado), no el AABB.
    Alto = rango Z laterales/bbox.
    """
    width_mm = 0.0
    frame = oriented_plan_frame_mm(plan_polygons)
    if frame is not None:
        width_mm = float(frame.get(u"luz_menor_mm") or 0.0)
    if width_mm < 1.0:
        # Fallback AABB (solo si no hay marco).
        if plan_polygons:
            xs = []
            ys = []
            for poly in plan_polygons:
                for x, y in poly:
                    xs.append(float(x))
                    ys.append(float(y))
            if xs and ys:
                width_mm = min(max(xs) - min(xs), max(ys) - min(ys))
    if width_mm < 1.0:
        try:
            bb = elem.get_BoundingBox(None)
            if bb is not None:
                dx = abs(float(bb.Max.X) - float(bb.Min.X)) * _FT_TO_MM
                dy = abs(float(bb.Max.Y) - float(bb.Min.Y)) * _FT_TO_MM
                width_mm = min(dx, dy)
        except Exception:
            width_mm = 1000.0
    if width_mm < 1.0:
        width_mm = 1000.0

    height_mm = 500.0
    try:
        z_lo, z_hi = rango_z_caras_laterales_o_bbox(elem)
        if z_lo is not None and z_hi is not None:
            height_mm = max(50.0, abs(float(z_hi) - float(z_lo)) * _FT_TO_MM)
    except Exception:
        pass
    return float(width_mm), float(height_mm)


def _element_label(elem):
    try:
        name = elem.Name
    except Exception:
        name = u""
    try:
        eid = element_id_to_int(elem.Id)
    except Exception:
        eid = u"?"
    name = _as_unicode(name).strip()
    if name:
        return u"{0}  [Id {1}]".format(name, eid)
    return u"Id {0}".format(eid)


def _pick_foundation(uidoc, uiapp):
    flt = FundacionAisladaSelectionFilter()
    try:
        ref_pick = uidoc.Selection.PickObject(
            ObjectType.Element,
            flt,
            u"Seleccione una fundación aislada.",
        )
    except Exception:
        return None
    if ref_pick is None:
        return None
    try:
        return uidoc.Document.GetElement(ref_pick.ElementId)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# UI — tamaño por defecto (caso más desfavorable: tab Inferior / Superior)
# chrome ≈32 + pad 36 + título/subtítulo ≈52 + rail INF/SUP ≈590
#   (sección 220 + fundación + tabs + toggle + luz mayor/menor) + hint ≈28
#   + footer ≈54  →  Height 820 / MinHeight 760
# ancho: pad 36 + planta ≥620 + rail 380 → Width 1040 / MinWidth 980
# ---------------------------------------------------------------------------

_XAML = u"""
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="__CHROME__"
  Height="820" Width="1040"
  MinHeight="760" MinWidth="980"
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
        <TextBlock x:Name="TxtTitle" Text="Arainco: Fundación aislada"
                   Foreground="#E8F4F8" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="TxtSubtitle" Margin="0,6,0,0" Foreground="#95B8CC"
                   FontSize="11" TextWrapping="Wrap"
                   Text="Tabs Inferior / Superior / Lateral · planta refleja el tab activo."/>
      </StackPanel>

      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="380"/>
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
              <TextBlock x:Name="TxtCanvasHeader"
                         Foreground="#64748b" FontSize="10" FontWeight="SemiBold"
                         VerticalAlignment="Center"
                         Text="PLANTA · CONTORNO (mm)"/>
            </Border>
            <Border Grid.Row="1" Background="#050E18" BorderBrush="Transparent"
                    BorderThickness="0" Padding="8,4,8,8">
              <Border Background="#050E18" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4">
                <Canvas x:Name="CvPlan" ClipToBounds="True" Background="#050E18"/>
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
                      BorderThickness="1" CornerRadius="4" Padding="8" Margin="0,0,0,10">
                <StackPanel>
                  <TextBlock Text="SECCIÓN · LUZ MENOR" Foreground="#64748b"
                             FontSize="10" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <Border Background="#050E18" BorderBrush="#21465C"
                          BorderThickness="1" CornerRadius="4" Height="220">
                    <Canvas x:Name="CvSection" ClipToBounds="True" Background="#050E18"/>
                  </Border>
                  <TextBlock x:Name="TxtSectionDims" Foreground="#64748b" FontSize="10"
                             Margin="0,6,0,0" TextWrapping="Wrap" Text=""/>
                </StackPanel>
              </Border>

              <Border Background="#0a1620" BorderBrush="#21465C"
                      BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,10">
                <StackPanel>
                  <TextBlock Text="Fundación" Foreground="#E8F4F8"
                             FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                  <TextBlock x:Name="TxtHost" Foreground="#95B8CC" FontSize="11"
                             TextWrapping="Wrap" Text="—"/>
                </StackPanel>
              </Border>

__CARDS__

            </StackPanel>
          </ScrollViewer>
        </Border>
      </Grid>

      <TextBlock Grid.Row="2" x:Name="TxtHint" Foreground="#64748b" FontSize="10"
                 TextWrapping="Wrap" Margin="0,8,0,0"
                 Text="Tabs Inferior / Superior / Lateral · la planta muestra el tab activo · rueda = zoom."/>

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
          <Button x:Name="BtnColocar" Content="Colocar armaduras"
                  Style="{StaticResource BtnPrimary}" MinWidth="180"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
""".replace(u"__CHROME__", WINDOW_CHROME_TITLE).replace(
    u"__STYLES__", BIMTOOLS_DARK_STYLES_XML
).replace(u"__CARDS__", _CARD_XAML)


class _ColocarArmaduraSketchHandler(IExternalEventHandler):
    def __init__(self):
        # Strong refs while ExternalEvent is queued / running (weakref alone
        # is unreliable with IronPython + WPF modeless).
        self._pending_ctrl = None
        self._pending_settings = None
        self._pending_foundation = None
        self._pending_doc = None
        self._pending_uiapp = None
        self._pending_uidoc = None
        self._pending_source_view = None

    def GetName(self):
        return u"AraincoColocarArmaduraFundacionAisladaSketch"

    def Execute(self, uiapp):
        ctrl = self._pending_ctrl
        settings = self._pending_settings
        foundation = self._pending_foundation
        doc = self._pending_doc
        app = uiapp or self._pending_uiapp
        uidoc = self._pending_uidoc
        source_view = self._pending_source_view
        try:
            if settings is None or foundation is None or doc is None:
                if app is not None:
                    _mostrar_aviso(
                        app,
                        u"No se pudo colocar la armadura.",
                        content=u"Faltan datos de colocación (reinicie la herramienta).",
                    )
                return
            if source_view is None and uidoc is not None:
                try:
                    source_view = uidoc.ActiveView
                except Exception:
                    source_view = None
            try:
                result = crear_armadura_fundacion_aislada(
                    doc,
                    foundation,
                    settings,
                    source_view=source_view,
                    uidoc=uidoc,
                )
            except Exception as ex:
                result = {
                    u"ok": False,
                    u"message": u"Error al crear armadura:\n{0}".format(
                        _as_unicode(ex)
                    ),
                }
            msg = result.get(u"message") or u""
            title = (
                u"Armadura colocada."
                if result.get(u"ok")
                else u"No se pudo colocar la armadura."
            )
            if app is not None:
                _mostrar_aviso(app, title, content=msg)
            if ctrl is not None:
                try:
                    if result.get(u"ok"):
                        ctrl._set_status(u"Armadura creada.")
                    else:
                        ctrl._set_status(u"Error al crear armadura.")
                except Exception:
                    pass
        finally:
            self._pending_ctrl = None
            self._pending_settings = None
            self._pending_foundation = None
            self._pending_doc = None
            self._pending_uiapp = None
            self._pending_uidoc = None
            self._pending_source_view = None
            if ctrl is not None:
                try:
                    ctrl._pending_settings = None
                except Exception:
                    pass
                try:
                    if getattr(ctrl, u"_win", None) is None and getattr(
                        ctrl, u"_colocar_event", None
                    ) is not None:
                        ctrl._colocar_event.Dispose()
                        ctrl._colocar_event = None
                except Exception:
                    pass


class ArmadoFundacionAisladaSketchController(object):
    def __init__(
        self,
        uiapp,
        uidoc,
        doc,
        foundation,
        plan_polygons,
        section_w,
        section_h,
        source_view=None,
    ):
        self._uiapp = uiapp
        self._uidoc = uidoc
        self._doc = doc
        self._foundation = foundation
        self._plan_polygons = list(plan_polygons or [])
        self._plan_frame = oriented_plan_frame_mm(self._plan_polygons)
        self._section_w = float(section_w or 0)
        self._section_h = float(section_h or 0)
        # Vista activa al ejecutar la herramienta (Section Filter / tipo Detail).
        self._source_view = source_view
        # Preferir luz menor del marco orientado (no AABB).
        if self._plan_frame is not None:
            try:
                lm = float(self._plan_frame.get(u"luz_menor_mm") or 0)
                if lm > 1.0:
                    self._section_w = lm
            except Exception:
                pass

        self._win = None
        self._view_zoom = 1.0
        self._view_pan_x = 0.0
        self._view_pan_y = 0.0
        self._scene_base = None
        self._panning = False
        self._pan_last = None
        self._ui_revealed = False

        self._ui_cv_plan = None
        self._ui_cv_section = None
        self._ui_txt_status = None
        self._ui_txt_host = None
        self._ui_txt_section_dims = None
        self._ui_txt_header = None
        self._bar_entries = []  # [(RebarBarType, label), ...]
        self._ui_syncing = False
        self._toggle_parts = {}  # card_key -> toggle visual parts
        self._active_tab = u"inferior"
        self._tab_ui = {}  # card_key -> {tab, pill, pill_tb, label_tb, panel}

        self._pending_settings = None
        self._colocar_handler = _ColocarArmaduraSketchHandler()
        self._colocar_event = ExternalEvent.Create(self._colocar_handler)

    def _set_status(self, text):
        try:
            if self._ui_txt_status is not None:
                self._ui_txt_status.Text = _as_unicode(text)
        except Exception:
            pass

    def _cache_ui_refs(self):
        win = self._win
        if win is None:
            return
        try:
            self._ui_cv_plan = win.FindName(u"CvPlan")
        except Exception:
            self._ui_cv_plan = None
        try:
            self._ui_cv_section = win.FindName(u"CvSection")
        except Exception:
            self._ui_cv_section = None
        try:
            self._ui_txt_status = win.FindName(u"TxtStatus")
        except Exception:
            self._ui_txt_status = None
        try:
            self._ui_txt_host = win.FindName(u"TxtHost")
        except Exception:
            self._ui_txt_host = None
        try:
            self._ui_txt_section_dims = win.FindName(u"TxtSectionDims")
        except Exception:
            self._ui_txt_section_dims = None
        try:
            self._ui_txt_header = win.FindName(u"TxtCanvasHeader")
        except Exception:
            self._ui_txt_header = None

    def _resolve_manual_path(self):
        candidates = []
        try:
            import bimtools_paths

            pb = bimtools_paths.get_pushbutton_dir()
            if pb:
                candidates.append(os.path.join(pb, u"manual_usuario.html"))
        except Exception:
            pass
        try:
            ext_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            for tab_name in os.listdir(ext_dir):
                if not tab_name.endswith(u".tab"):
                    continue
                panel = os.path.join(ext_dir, tab_name, u"Armadura.panel")
                if not os.path.isdir(panel):
                    continue
                for pb_name in os.listdir(panel):
                    if u"ArmadoFundacionAisladaSketch" not in pb_name:
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

    def _open_manual(self, sender=None, args=None):
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

    def _on_colocar(self, sender=None, args=None):
        settings = self.read_armadura_settings()
        if not any(settings[k][u"enabled"] for k in _CARD_KEYS):
            _mostrar_aviso(
                self._uiapp,
                u"Active al menos un grupo de armadura.",
                content=u"Active Inferior, Superior y/o Lateral con el toggle.",
            )
            return
        missing = []
        for key in _CARD_KEYS:
            s = settings[key]
            if not s.get(u"enabled"):
                continue
            if key == u"lateral":
                if s.get(u"bar_type") is None:
                    missing.append(u"Lateral")
            else:
                for dk in _DIR_KEYS:
                    d = s.get(dk) or {}
                    if d.get(u"bar_type") is None:
                        missing.append(
                            u"{0} · {1}".format(
                                (_CARD_UI.get(key) or {}).get(u"title") or key,
                                u"luz mayor" if dk == u"luz_mayor" else u"luz menor",
                            )
                        )
        if missing:
            _mostrar_aviso(
                self._uiapp,
                u"Falta diámetro de barra.",
                content=u"Seleccione un tipo en:\n" + u"\n".join(missing),
            )
            return
        if self._colocar_event is None or self._colocar_handler is None:
            _mostrar_aviso(
                self._uiapp,
                u"ExternalEvent no disponible.",
                content=u"Reinicie la herramienta e intente de nuevo.",
            )
            return
        self._pending_settings = settings
        source_view = self._source_view
        if source_view is None:
            try:
                source_view = self._uidoc.ActiveView if self._uidoc is not None else None
            except Exception:
                source_view = None
        try:
            self._colocar_handler._pending_ctrl = self
            self._colocar_handler._pending_settings = settings
            self._colocar_handler._pending_foundation = self._foundation
            self._colocar_handler._pending_doc = self._doc
            self._colocar_handler._pending_uiapp = self._uiapp
            self._colocar_handler._pending_uidoc = self._uidoc
            self._colocar_handler._pending_source_view = source_view
        except Exception:
            pass
        self._set_status(u"Colocando armadura…")
        try:
            self._colocar_event.Raise()
        except Exception as ex:
            self._pending_settings = None
            try:
                self._colocar_handler._pending_ctrl = None
                self._colocar_handler._pending_settings = None
                self._colocar_handler._pending_foundation = None
                self._colocar_handler._pending_doc = None
                self._colocar_handler._pending_uiapp = None
                self._colocar_handler._pending_uidoc = None
                self._colocar_handler._pending_source_view = None
            except Exception:
                pass
            _mostrar_aviso(
                self._uiapp,
                u"No se pudo iniciar la colocación.",
                content=_as_unicode(ex),
            )

    def _default_sep_for_card(self, card_key):
        if card_key == u"lateral":
            return _SEP_LAT_DEFAULT
        return _SEP_MM_DEFAULT

    def _iter_dir_metas(self, card_key):
        """Yield (dir_key_or_None, dir_meta dict with cmb/sep)."""
        meta = _CARD_UI.get(card_key) or {}
        dirs = meta.get(u"dirs")
        if dirs:
            for dk in _DIR_KEYS:
                dmeta = dirs.get(dk)
                if dmeta:
                    yield dk, dmeta
        else:
            yield None, {
                u"cmb": meta.get(u"cmb"),
                u"sep": meta.get(u"sep"),
                u"up": meta.get(u"up"),
                u"down": meta.get(u"down"),
            }

    def _find(self, name):
        if self._win is None or not name:
            return None
        try:
            return self._win.FindName(name)
        except Exception:
            return None

    def _on_sep_preview_text(self, sender, args):
        """Solo dígitos en cajas de separación."""
        try:
            txt = args.Text
        except Exception:
            return
        if not _sep_text_is_digits_only(txt):
            try:
                args.Handled = True
            except Exception:
                pass

    def _on_card_enabled_changed(self, sender=None, args=None):
        for key, meta in _CARD_UI.items():
            chk = self._find(meta[u"chk"])
            en = True
            try:
                if chk is not None:
                    en = bool(chk.IsChecked)
            except Exception:
                en = True
            parts = self._toggle_parts.get(key)
            if parts is not None:
                _sync_toggle_switch_visual(parts, en)
            for _dk, dmeta in self._iter_dir_metas(key):
                for nm in (dmeta.get(u"cmb"), dmeta.get(u"sep")):
                    ctrl = self._find(nm)
                    if ctrl is None:
                        continue
                    try:
                        ctrl.IsEnabled = en
                    except Exception:
                        pass
        if not self._ui_syncing:
            self._redraw_all()

    def _build_tabs(self):
        """Fila de tabs Inferior / Superior / Lateral (mismo patrón Area Rein.)."""
        from System.Windows import GridLength, GridUnitType
        from System.Windows.Controls import ColumnDefinition, Grid
        from System.Windows.Input import Cursors

        host = self._find(u"PnlTabs")
        if host is None:
            return
        try:
            host.Children.Clear()
            host.ColumnDefinitions.Clear()
        except Exception:
            pass

        self._tab_ui = {}
        for i, key in enumerate(_CARD_KEYS):
            meta = _CARD_UI[key]
            try:
                cd = ColumnDefinition()
                cd.Width = GridLength(1.0, GridUnitType.Star)
                host.ColumnDefinitions.Add(cd)
            except Exception:
                pass

            tab = Border()
            tab.Cursor = Cursors.Hand
            tab.CornerRadius = CornerRadius(4, 4, 0, 0)
            tab.Padding = Thickness(6, 8, 6, 8)
            if i == 0:
                tab.Margin = Thickness(0, 0, 3, 0)
            elif i == len(_CARD_KEYS) - 1:
                tab.Margin = Thickness(3, 0, 0, 0)
            else:
                tab.Margin = Thickness(3, 0, 3, 0)
            tab.Background = _brush(u"#0a1620")
            tab.BorderBrush = _brush(u"#21465C")
            tab.BorderThickness = Thickness(1)

            inner = StackPanel()
            inner.Orientation = Orientation.Horizontal
            try:
                inner.HorizontalAlignment = HorizontalAlignment.Center
            except Exception:
                pass

            pill = Border()
            pill.Background = _brush(u"#0E1B32")
            pill.BorderBrush = _brush(u"#21465C")
            pill.BorderThickness = Thickness(1)
            pill.CornerRadius = CornerRadius(3)
            pill.Padding = Thickness(5, 1, 5, 1)
            pill.Margin = Thickness(0, 0, 5, 0)
            pill_tb = TextBlock()
            pill_tb.Text = meta[u"title"][:3].upper()
            pill_tb.Foreground = _brush(u"#64748b")
            pill_tb.FontSize = 9
            try:
                pill_tb.FontWeight = FontWeights.Bold
            except Exception:
                pass
            pill.Child = pill_tb

            label_tb = TextBlock()
            label_tb.Text = meta[u"title"]
            label_tb.Foreground = _brush(u"#95B8CC")
            label_tb.FontSize = 11
            try:
                label_tb.FontWeight = FontWeights.SemiBold
                label_tb.VerticalAlignment = VerticalAlignment.Center
            except Exception:
                pass

            inner.Children.Add(pill)
            inner.Children.Add(label_tb)
            tab.Child = inner

            def _make(fid):
                def _on(s, e):
                    self._set_active_tab(fid)

                return _on

            try:
                tab.MouseLeftButtonDown += MouseButtonEventHandler(_make(key))
            except Exception:
                pass

            try:
                Grid.SetColumn(tab, i)
                host.Children.Add(tab)
            except Exception:
                pass

            self._tab_ui[key] = {
                u"tab": tab,
                u"pill": pill,
                u"pill_tb": pill_tb,
                u"label_tb": label_tb,
                u"panel": self._find(meta[u"panel"]),
            }

    def _set_active_tab(self, key):
        if key not in _CARD_KEYS:
            key = u"inferior"
        self._active_tab = key
        from System.Windows import Visibility

        for k, ui in (self._tab_ui or {}).items():
            active = k == key
            meta = _CARD_UI[k]
            color = meta.get(u"color") or ACCENT_PRIMARY
            tab = ui.get(u"tab")
            pill = ui.get(u"pill")
            pill_tb = ui.get(u"pill_tb")
            label_tb = ui.get(u"label_tb")
            panel = ui.get(u"panel")
            if tab is not None:
                try:
                    if active:
                        tab.Background = _brush(u"#0E1B32")
                        tab.BorderBrush = _brush(color)
                        tab.BorderThickness = Thickness(1, 1, 1, 0)
                    else:
                        tab.Background = _brush(u"#0a1620")
                        tab.BorderBrush = _brush(u"#21465C")
                        tab.BorderThickness = Thickness(1)
                except Exception:
                    pass
            if pill is not None:
                try:
                    if active:
                        pill.Background = _brush(color, 40)
                        pill.BorderBrush = _brush(color)
                    else:
                        pill.Background = _brush(u"#0E1B32")
                        pill.BorderBrush = _brush(u"#21465C")
                except Exception:
                    pass
            if pill_tb is not None:
                try:
                    pill_tb.Foreground = _brush(color if active else u"#64748b")
                except Exception:
                    pass
            if label_tb is not None:
                try:
                    label_tb.Foreground = _brush(
                        u"#E8F4F8" if active else u"#95B8CC"
                    )
                except Exception:
                    pass
            if panel is not None:
                try:
                    panel.Visibility = (
                        Visibility.Visible if active else Visibility.Collapsed
                    )
                except Exception:
                    pass
        if not self._ui_syncing:
            self._redraw_all()

    def _on_settings_changed(self, sender=None, args=None):
        if self._ui_syncing:
            return
        for key in _CARD_KEYS:
            default_sep = self._default_sep_for_card(key)
            for _dk, dmeta in self._iter_dir_metas(key):
                tb = self._find(dmeta.get(u"sep"))
                if tb is not None:
                    _normalize_sep_textbox(tb, default_sep)
        self._redraw_all()

    def _fill_one_combo(self, cmb, err):
        if cmb is None:
            return
        try:
            cmb.Items.Clear()
            cmb.IsEditable = False
        except Exception:
            pass
        if err:
            return
        for _bt, lbl in self._bar_entries:
            try:
                cmb.Items.Add(lbl)
            except Exception:
                pass
        sel_idx = 0
        for i, (b, lbl) in enumerate(self._bar_entries):
            dmm = None
            try:
                if b is not None and _rebar_nominal_diameter_mm is not None:
                    dmm = _rebar_nominal_diameter_mm(b)
            except Exception:
                dmm = None
            if dmm is not None and abs(float(dmm) - 8.0) < 0.6:
                sel_idx = i
                break
            if b is None and u"8" in _as_unicode(lbl):
                sel_idx = i
                break
        try:
            cmb.SelectedIndex = min(sel_idx, max(0, cmb.Items.Count - 1))
        except Exception:
            try:
                cmb.SelectedIndex = 0
            except Exception:
                pass

    def _cargar_combos_diametro(self):
        self._bar_entries = []
        err = None
        if _build_bar_type_entries is not None:
            try:
                entries, err = _build_bar_type_entries(self._doc)
                self._bar_entries = list(entries) if entries else []
            except Exception as ex:
                err = _as_unicode(ex)
                self._bar_entries = []
        else:
            err = u"No se pudo cargar tipos de barra."

        if err:
            self._set_status(err)

        for key in _CARD_KEYS:
            for _dk, dmeta in self._iter_dir_metas(key):
                self._fill_one_combo(self._find(dmeta.get(u"cmb")), err)

    def _init_cards(self):
        self._ui_syncing = True
        try:
            self._build_tabs()
            self._toggle_parts = {}
            for key in _CARD_KEYS:
                meta = _CARD_UI[key]
                chk = self._find(meta[u"chk"])
                parts = {}
                _apply_toggle_switch(
                    chk,
                    u"Incluir {0}".format(meta[u"title"].lower()),
                    parts,
                    accent_hex=meta.get(u"color") or ACCENT_PRIMARY,
                )
                self._toggle_parts[key] = parts

                default_sep = self._default_sep_for_card(key)
                for _dk, dmeta in self._iter_dir_metas(key):
                    tb = self._find(dmeta.get(u"sep"))
                    if tb is not None:
                        try:
                            tb.Text = _as_unicode(int(default_sep))
                        except Exception:
                            pass
                        _normalize_sep_textbox(tb, default_sep)
            self._cargar_combos_diametro()
            self._on_card_enabled_changed()
            self._set_active_tab(getattr(self, u"_active_tab", u"inferior"))
        finally:
            self._ui_syncing = False

    def _wire_cards(self):
        for key in _CARD_KEYS:
            meta = _CARD_UI[key]
            chk = self._find(meta[u"chk"])
            if chk is not None:
                try:
                    chk.Checked += RoutedEventHandler(self._on_card_enabled_changed)
                    chk.Unchecked += RoutedEventHandler(self._on_card_enabled_changed)
                except Exception:
                    pass
            for _dk, dmeta in self._iter_dir_metas(key):
                cmb = self._find(dmeta.get(u"cmb"))
                if cmb is not None:
                    try:
                        cmb.SelectionChanged += SelectionChangedEventHandler(
                            self._on_settings_changed
                        )
                    except Exception:
                        pass
                tb = self._find(dmeta.get(u"sep"))
                if tb is not None:
                    try:
                        tb.PreviewTextInput += TextCompositionEventHandler(
                            self._on_sep_preview_text
                        )
                    except Exception:
                        pass
                    try:
                        tb.LostFocus += RoutedEventHandler(self._on_settings_changed)
                    except Exception:
                        pass
                    try:
                        from System.Windows import DataObject, DataFormats

                        def _on_paste(sender, e):
                            try:
                                data = e.DataObject.GetData(DataFormats.Text)
                            except Exception:
                                data = None
                            if not _sep_text_is_digits_only(data):
                                try:
                                    e.CancelCommand()
                                except Exception:
                                    pass

                        DataObject.AddPastingHandler(tb, _on_paste)
                    except Exception:
                        pass

    def _read_dir_from_ui(self, dmeta, default_sep):
        tb = self._find(dmeta.get(u"sep"))
        sep = _leer_sep_mm(tb, default_sep)
        cmb = self._find(dmeta.get(u"cmb"))
        idx = -1
        label = u""
        try:
            if cmb is not None:
                idx = int(cmb.SelectedIndex)
                if idx >= 0 and idx < cmb.Items.Count:
                    label = _as_unicode(cmb.Items[idx])
        except Exception:
            idx = -1
        bar_type = None
        dmm = None
        if idx >= 0 and idx < len(self._bar_entries):
            bar_type, label = self._bar_entries[idx]
            try:
                if bar_type is not None and _rebar_nominal_diameter_mm is not None:
                    dmm = _rebar_nominal_diameter_mm(bar_type)
            except Exception:
                dmm = None
        return {
            u"spacing_mm": float(sep),
            u"diameter_mm": dmm,
            u"bar_type": bar_type,
            u"label": label,
        }

    def read_armadura_settings(self):
        """
        Snapshot de UI para creación de barras.

        Inferior/Superior:
            {enabled, luz_mayor: {...}, luz_menor: {...}}
        Lateral:
            {enabled, spacing_mm, diameter_mm, bar_type, label}
        """
        out = {}
        for key in _CARD_KEYS:
            meta = _CARD_UI[key]
            chk = self._find(meta[u"chk"])
            enabled = True
            try:
                if chk is not None:
                    enabled = bool(chk.IsChecked)
            except Exception:
                enabled = True
            default_sep = self._default_sep_for_card(key)
            if key == u"lateral":
                d = self._read_dir_from_ui(
                    {
                        u"cmb": meta.get(u"cmb"),
                        u"sep": meta.get(u"sep"),
                    },
                    default_sep,
                )
                d[u"enabled"] = enabled
                out[key] = d
            else:
                entry = {u"enabled": enabled}
                for dk, dmeta in self._iter_dir_metas(key):
                    entry[dk] = self._read_dir_from_ui(dmeta, default_sep)
                out[key] = entry
        return out

    def _update_status_from_settings(self):
        settings = self.read_armadura_settings()
        parts = []
        for key in _CARD_KEYS:
            s = settings[key]
            if not s[u"enabled"]:
                continue
            meta = _CARD_UI[key]
            if key == u"lateral":
                dmm = s.get(u"diameter_mm")
                dlab = (
                    u"ø{0:.0f}".format(float(dmm))
                    if dmm is not None
                    else (s.get(u"label") or u"—")
                )
                parts.append(
                    u"Lat {0}@{1:.0f}".format(dlab, float(s[u"spacing_mm"]))
                )
            else:
                short = meta[u"title"][:3]
                for dk in _DIR_KEYS:
                    d = s.get(dk) or {}
                    dmm = d.get(u"diameter_mm")
                    dlab = (
                        u"ø{0:.0f}".format(float(dmm))
                        if dmm is not None
                        else (d.get(u"label") or u"—")
                    )
                    tag = u"May" if dk == u"luz_mayor" else u"Men"
                    parts.append(
                        u"{0}{1} {2}@{3:.0f}".format(
                            short, tag, dlab, float(d.get(u"spacing_mm") or 0)
                        )
                    )
        if not parts:
            self._set_status(u"Ningún grupo activo.")
        else:
            self._set_status(u" · ".join(parts))

    def _positions_along(self, length_mm, spacing_mm, cover_mm):
        """
        Ejes tipo Maximum Spacing con includeFirst/Last:
        primera en cover, última en length-cover, paso <= spacing.
        """
        L = float(length_mm)
        e = max(1.0, float(spacing_mm))
        c = max(0.0, float(cover_mm))
        if L <= 2.0 * c + 1.0:
            return [L * 0.5]
        start = c
        end = L - c
        xs = []
        x = start
        guard = 0
        while x <= end + 0.5 and guard < 500:
            xs.append(x)
            x += e
            guard += 1
        if not xs:
            xs = [L * 0.5]
        # Asegurar barra en el extremo (includeLastBar) si el paso no cae exacto.
        if xs and abs(xs[-1] - end) > 0.5:
            if end - xs[-1] > 0.5 * e:
                xs.append(end)
            else:
                xs[-1] = end
        return xs

    def _extremos_eje_mm(self, d_mm):
        """Recorte de extremo al eje: 50 + ø/2 (motor)."""
        try:
            d = float(d_mm or 0.0)
        except Exception:
            d = 0.0
        return float(_REC_EXTREMOS_PREVIEW_MM) + 0.5 * d

    def _diameter_inf_mm_for_lateral(self, settings):
        """ø malla inferior (máx. de ambas luces) si el grupo está activo; si no, 0."""
        inf = (settings or {}).get(u"inferior") or {}
        if not inf.get(u"enabled"):
            return 0.0
        d_inf = 0.0
        for dk in _DIR_KEYS:
            raw = (inf.get(dk) or {}).get(u"diameter_mm")
            if raw is None:
                continue
            try:
                d_inf = max(d_inf, float(raw))
            except Exception:
                pass
        return d_inf

    def _lateral_plan_inset_mm(self, settings, d_lat_mm=None):
        """
        Offset en planta al eje lateral (motor):
        50 + ø malla inferior + ø lateral / 2.
        """
        if d_lat_mm is None:
            lat = (settings or {}).get(u"lateral") or {}
            d_lat_mm = float(lat.get(u"diameter_mm") or 8.0)
        d_inf = self._diameter_inf_mm_for_lateral(settings)
        return float(REC_LATERAL_CARA_MM) + float(d_inf) + 0.5 * float(d_lat_mm)

    def _lateral_z_axis_positions_mm(self, h_mm, sep_mm, d_mm=None):
        """
        Laterales en altura (motor): primera a 100 mm desde cara inferior;
        array = h − 200 mm → última a h − 100 mm; Maximum Spacing.
        """
        clear = float(OFFSET_PRIMERA_LATERAL_MM)
        # d_mm reservado por compatibilidad; el motor fija 100 / h-200.
        return self._positions_along(h_mm, sep_mm, clear)

    def _pata_u_preview_mm(self, d_mm, h_mm, solo_inferior):
        """Longitud de pata U en mm para esquema de sección."""
        try:
            d = float(d_mm or 0.0)
        except Exception:
            d = 0.0
        try:
            h = float(h_mm or 0.0)
        except Exception:
            h = 0.0
        descuento = float(DESCUENTO_PATA_U_MM) + 0.5 * d
        leg_max = max(0.0, h - descuento)
        if not solo_inferior:
            return leg_max
        hook_mm = None
        try:
            hook_mm = largo_gancho_u_tabla_mm(d)
        except Exception:
            hook_mm = None
        if hook_mm is None:
            return leg_max
        try:
            eje_mm = float(hook_mm) - 0.5 * d
        except Exception:
            eje_mm = float(hook_mm)
        eje_mm = max(0.0, eje_mm)
        if leg_max > 0.0:
            return min(eje_mm, leg_max)
        return eje_mm

    def _add_bar_dot(self, cv, cx, cy, r, hex_color, alpha=255):
        el = WpfEllipse()
        el.Width = 2.0 * r
        el.Height = 2.0 * r
        el.Fill = _brush(hex_color, alpha)
        stroke_a = 180 if int(alpha) >= 200 else max(60, int(alpha) - 20)
        el.Stroke = _brush(u"#071018", stroke_a)
        el.StrokeThickness = 0.6
        try:
            WpfCanvas.SetLeft(el, cx - r)
            WpfCanvas.SetTop(el, cy - r)
            cv.Children.Add(el)
        except Exception:
            pass

    def _add_section_bar_line(self, cv, x0, y0, x1, y1, hex_color, alpha=255, thickness=1.6):
        """Barra vista de canto en sección (dirección en el plano del corte)."""
        try:
            ln = WpfLine()
            ln.X1 = float(x0)
            ln.Y1 = float(y0)
            ln.X2 = float(x1)
            ln.Y2 = float(y1)
            ln.Stroke = _brush(hex_color, alpha)
            ln.StrokeThickness = float(thickness)
            cv.Children.Add(ln)
        except Exception:
            pass

    def _section_layer_alpha(self, key, active, settings):
        """Opacidad sección: activa plena; otras tenues; desactivada aún más tenue."""
        layer = (settings or {}).get(key) or {}
        enabled = bool(layer.get(u"enabled"))
        if key == active:
            return 255 if enabled else 90
        if enabled:
            return 75
        return 40

    def _on_cancel(self, sender=None, args=None):
        try:
            if self._win is not None:
                self._win.Close()
        except Exception:
            pass

    def _on_closed(self, sender=None, args=None):
        _unregister_singleton()
        # Si hay colocación en cola, el ExternalEvent se libera tras Execute.
        if getattr(self, u"_pending_settings", None) is not None:
            return
        try:
            ev = getattr(self, u"_colocar_event", None)
            if ev is not None:
                ev.Dispose()
        except Exception:
            pass
        self._colocar_event = None

    def _canvas_size(self, cv):
        try:
            w = float(cv.ActualWidth or 0)
            h = float(cv.ActualHeight or 0)
        except Exception:
            w, h = 0.0, 0.0
        if w < 40:
            try:
                w = float(cv.RenderSize.Width or 0)
            except Exception:
                pass
        if h < 40:
            try:
                h = float(cv.RenderSize.Height or 0)
            except Exception:
                pass
        return max(40.0, w), max(40.0, h)

    def _plan_bbox_mm(self):
        xs = []
        ys = []
        for poly in self._plan_polygons:
            for x, y in poly:
                xs.append(float(x))
                ys.append(float(y))
        if not xs:
            return 0.0, 1000.0, 0.0, 1000.0
        return min(xs), max(xs), min(ys), max(ys)

    def _redraw_plan(self):
        cv = self._ui_cv_plan
        if cv is None:
            return
        try:
            cv.Children.Clear()
        except Exception:
            return

        cw, ch = self._canvas_size(cv)
        min_x, max_x, min_y, max_y = self._plan_bbox_mm()
        span_x = max(1.0, max_x - min_x)
        span_y = max(1.0, max_y - min_y)
        pad = _PLAN_PAD_FRAC
        fit = min(
            (cw * (1.0 - 2.0 * pad)) / span_x,
            (ch * (1.0 - 2.0 * pad)) / span_y,
        )
        fit = max(1e-6, fit)
        scale = fit * max(0.05, float(self._view_zoom))
        cx_mm = 0.5 * (min_x + max_x) + float(self._view_pan_x)
        cy_mm = 0.5 * (min_y + max_y) + float(self._view_pan_y)
        ox = cw / 2.0 - (cx_mm - min_x) * scale
        oy = ch / 2.0 - (max_y - cy_mm) * scale

        self._scene_base = {
            u"min_x": min_x,
            u"max_x": max_x,
            u"min_y": min_y,
            u"max_y": max_y,
            u"ox": ox,
            u"oy": oy,
            u"scale": scale,
            u"fit": fit,
            u"cw": cw,
            u"ch": ch,
        }

        def to_px(xmm, ymm):
            return (
                ox + (float(xmm) - min_x) * scale,
                oy + (max_y - float(ymm)) * scale,
            )

        from System.Windows.Media import PointCollection

        for i, poly in enumerate(self._plan_polygons):
            if not poly or len(poly) < 3:
                continue
            wp = WpfPolygon()
            pc = PointCollection()
            for xmm, ymm in poly:
                px, py = to_px(xmm, ymm)
                pc.Add(WpfPoint(px, py))
            wp.Points = pc
            if i == 0:
                wp.Fill = _brush(u"#1a3a4d", 160)
                wp.Stroke = _brush(u"#5BC0DE")
                wp.StrokeThickness = 1.6
            else:
                wp.Fill = _brush(u"#071018", 120)
                wp.Stroke = _brush(u"#64748b")
                wp.StrokeThickness = 1.0
            try:
                cv.Children.Add(wp)
            except Exception:
                pass

        active = getattr(self, u"_active_tab", u"inferior") or u"inferior"
        settings = self.read_armadura_settings()
        layer = settings.get(active) or {}
        meta = _CARD_UI.get(active) or {}
        color = meta.get(u"color") or ACCENT_PRIMARY
        enabled = bool(layer.get(u"enabled"))

        if enabled:
            if active in (u"inferior", u"superior"):
                self._draw_plan_mesh_bars(cv, to_px, layer, color)
            elif active == u"lateral":
                self._draw_plan_lateral_bars(cv, to_px, layer, color, settings)

        try:
            if self._ui_txt_header is not None:
                title = meta.get(u"title") or active
                state = u"" if enabled else u" · (desactivado)"
                fr = self._plan_frame
                if fr is not None:
                    dim_a = float(fr.get(u"luz_menor_mm") or span_x)
                    dim_b = float(fr.get(u"luz_mayor_mm") or span_y)
                else:
                    dim_a, dim_b = span_x, span_y
                self._ui_txt_header.Text = (
                    u"PLANTA · {0}{1}  ·  {2:.0f} × {3:.0f} mm"
                ).format(title.upper(), state, dim_a, dim_b)
        except Exception:
            pass

    def _draw_plan_line(self, cv, to_px, x0, y0, x1, y1, hex_color, thickness=1.2):
        try:
            px0, py0 = to_px(x0, y0)
            px1, py1 = to_px(x1, y1)
            ln = WpfLine()
            ln.X1 = px0
            ln.Y1 = py0
            ln.X2 = px1
            ln.Y2 = py1
            ln.Stroke = _brush(hex_color)
            ln.StrokeThickness = thickness
            cv.Children.Add(ln)
        except Exception:
            pass

    def _draw_plan_mesh_bars(self, cv, to_px, layer, color):
        """
        Preview fiel al motor, en el marco de aristas reales (soporta rotación):
        - u // luz menor, v // luz mayor
        - Offset planta / array: 100 mm; extremos: 50 + ø/2
        """
        frame = self._plan_frame
        if frame is None:
            return
        rec_planta = float(_REC_PLANTA_PREVIEW_MM)
        len_u = float(frame[u"len_u"])
        len_v = float(frame[u"len_v"])
        u0 = float(frame[u"u_min"])
        v0 = float(frame[u"v_min"])

        d_may = layer.get(u"luz_mayor") or {}
        d_men = layer.get(u"luz_menor") or {}
        sep_may = float(d_may.get(u"spacing_mm") or _SEP_MM_DEFAULT)
        sep_men = float(d_men.get(u"spacing_mm") or _SEP_MM_DEFAULT)
        dmm_may = float(d_may.get(u"diameter_mm") or 8.0)
        dmm_men = float(d_men.get(u"diameter_mm") or 8.0)
        ext_may = self._extremos_eje_mm(dmm_may)
        ext_men = self._extremos_eje_mm(dmm_men)
        color_men = u"#86efac" if color == u"#4ade80" else u"#7dd3fc"

        # Luz menor: barras // u (lado corto), array a lo largo de v
        u_a = u0 + ext_men
        u_b = u0 + len_u - ext_men
        if u_b > u_a:
            for dv in self._positions_along(len_v, sep_men, rec_planta):
                vv = v0 + dv
                p0 = _frame_to_world(frame, u_a, vv)
                p1 = _frame_to_world(frame, u_b, vv)
                self._draw_plan_line(
                    cv, to_px, p0[0], p0[1], p1[0], p1[1], color_men, 1.2
                )

        # Luz mayor: barras // v (lado largo), array a lo largo de u
        v_a = v0 + ext_may
        v_b = v0 + len_v - ext_may
        if v_b > v_a:
            for du in self._positions_along(len_u, sep_may, rec_planta):
                uu = u0 + du
                p0 = _frame_to_world(frame, uu, v_a)
                p1 = _frame_to_world(frame, uu, v_b)
                self._draw_plan_line(
                    cv, to_px, p0[0], p0[1], p1[0], p1[1], color, 1.5
                )

    def _draw_plan_lateral_bars(self, cv, to_px, layer, color, settings=None):
        """
        Laterales en planta: rectángulo orientado al eje
        (inset 50 + ø_inf + ø_lat/2 en el marco de aristas).
        """
        frame = self._plan_frame
        if frame is None:
            return
        d_lat = float((layer or {}).get(u"diameter_mm") or 8.0)
        inset = self._lateral_plan_inset_mm(settings or {}, d_lat)
        u0 = float(frame[u"u_min"]) + inset
        u1 = float(frame[u"u_max"]) - inset
        v0 = float(frame[u"v_min"]) + inset
        v1 = float(frame[u"v_max"]) - inset
        if u1 <= u0 or v1 <= v0:
            return
        corners = (
            _frame_to_world(frame, u0, v0),
            _frame_to_world(frame, u1, v0),
            _frame_to_world(frame, u1, v1),
            _frame_to_world(frame, u0, v1),
        )
        for i in range(4):
            a = corners[i]
            b = corners[(i + 1) % 4]
            self._draw_plan_line(cv, to_px, a[0], a[1], b[0], b[1], color, 1.6)

    def _redraw_section(self):
        cv = self._ui_cv_section
        if cv is None:
            return
        try:
            cv.Children.Clear()
        except Exception:
            return

        cw, ch = self._canvas_size(cv)
        w_mm = max(1.0, float(self._section_w))
        h_mm = max(1.0, float(self._section_h))
        pad = _SECTION_PAD_PX
        usable_w = max(20.0, cw - 2.0 * pad)
        usable_h = max(20.0, ch - 2.0 * pad - 8.0)
        scale = min(usable_w / w_mm, usable_h / h_mm)
        rw = w_mm * scale
        rh = h_mm * scale
        left = (cw - rw) * 0.5
        top = (ch - rh) * 0.5

        rect = WpfRectangle()
        rect.Width = rw
        rect.Height = rh
        rect.Fill = _brush(u"#1a3a4d", 180)
        rect.Stroke = _brush(u"#5BC0DE")
        rect.StrokeThickness = 1.5
        try:
            WpfCanvas.SetLeft(rect, left)
            WpfCanvas.SetTop(rect, top)
            cv.Children.Add(rect)
        except Exception:
            pass

        settings = self.read_armadura_settings()
        rec_h = float(REC_HORIZONTAL_EJE_MM)
        rec_planta = float(_REC_PLANTA_PREVIEW_MM)
        active = getattr(self, u"_active_tab", u"inferior") or u"inferior"
        inf_on = bool((settings.get(u"inferior") or {}).get(u"enabled"))
        sup_on = bool((settings.get(u"superior") or {}).get(u"enabled"))

        def mm_to_px_x(xmm):
            return left + float(xmm) * scale

        def mm_to_px_y(ymm_from_bottom):
            return top + rh - float(ymm_from_bottom) * scale

        def _draw_mesh_section(layer_key, alpha, from_bottom):
            """
            Corte por luz menor (motor):
            - Luz menor → línea (eje) + patas U en extremos si aplica.
            - Luz mayor → círculos, array planta 100 mm, sep. luz mayor.
            Elevación: 50+ø/2; 2ª capa +ø de luz menor.
            """
            layer = settings.get(layer_key) or {}
            if not layer.get(u"enabled"):
                return
            d_men = layer.get(u"luz_menor") or {}
            d_may = layer.get(u"luz_mayor") or {}
            sep_may = float(d_may.get(u"spacing_mm") or _SEP_MM_DEFAULT)
            dmm_men = float(d_men.get(u"diameter_mm") or 8.0)
            dmm_may = float(d_may.get(u"diameter_mm") or 8.0)
            ext_men = self._extremos_eje_mm(dmm_men)
            color_men = _CARD_UI[layer_key][u"color"]
            color_may = u"#86efac" if layer_key == u"inferior" else u"#7dd3fc"

            if from_bottom:
                y_men = rec_h + 0.5 * dmm_men
                y_may = rec_h + 0.5 * dmm_may + dmm_men
                leg_up = True
                solo_inf = not sup_on
                draw_u = True  # inferior siempre intenta U
            else:
                y_men = h_mm - rec_h - 0.5 * dmm_men
                y_may = h_mm - rec_h - 0.5 * dmm_may - dmm_men
                leg_up = False
                solo_inf = False
                draw_u = inf_on  # superior en U solo si hay inferior

            # Luz menor: tramo horizontal (extremos 50+ø/2)
            x0 = ext_men
            x1 = w_mm - ext_men
            if x1 <= x0:
                x0, x1 = 0.0, w_mm
            py = mm_to_px_y(y_men)
            thick = max(1.4, min(3.0, dmm_men * scale))
            self._add_section_bar_line(
                cv,
                mm_to_px_x(x0),
                py,
                mm_to_px_x(x1),
                py,
                color_men,
                alpha,
                thick,
            )
            if draw_u:
                leg = self._pata_u_preview_mm(dmm_men, h_mm, solo_inf)
                if leg > 1.0:
                    for xx in (x0, x1):
                        if leg_up:
                            y1_leg = min(h_mm - rec_h, y_men + leg)
                        else:
                            y1_leg = max(rec_h, y_men - leg)
                        self._add_section_bar_line(
                            cv,
                            mm_to_px_x(xx),
                            mm_to_px_y(y_men),
                            mm_to_px_x(xx),
                            mm_to_px_y(y1_leg),
                            color_men,
                            alpha,
                            thick,
                        )

            # Luz mayor: puntos (array planta 100 mm)
            r = max(2.0, min(5.0, 0.5 * dmm_may * scale))
            for x_mm in self._positions_along(w_mm, sep_may, rec_planta):
                self._add_bar_dot(
                    cv, mm_to_px_x(x_mm), mm_to_px_y(y_may), r, color_may, alpha,
                )

        def _draw_inf(alpha):
            _draw_mesh_section(u"inferior", alpha, True)

        def _draw_sup(alpha):
            _draw_mesh_section(u"superior", alpha, False)

        def _draw_lat(alpha):
            """
            Laterales (motor):
            - X: 50 + ø_inf + ø_lat/2.
            - Y: primera a 100 mm; última a h−100; Maximum Spacing.
            """
            layer = settings.get(u"lateral") or {}
            if not layer.get(u"enabled"):
                return
            sep = float(layer.get(u"spacing_mm") or _SEP_LAT_DEFAULT)
            dmm = float(layer.get(u"diameter_mm") or 8.0)
            r = max(2.0, min(5.0, 0.5 * dmm * scale))
            inset = self._lateral_plan_inset_mm(settings, dmm)
            x_left = inset
            x_right = w_mm - inset
            for y_mm in self._lateral_z_axis_positions_mm(h_mm, sep, dmm):
                self._add_bar_dot(
                    cv, mm_to_px_x(x_left), mm_to_px_y(y_mm), r,
                    _CARD_UI[u"lateral"][u"color"], alpha,
                )
                self._add_bar_dot(
                    cv, mm_to_px_x(x_right), mm_to_px_y(y_mm), r,
                    _CARD_UI[u"lateral"][u"color"], alpha,
                )

        # Primero capas no activas (tenues); luego la activa encima.
        drawers = {
            u"inferior": _draw_inf,
            u"superior": _draw_sup,
            u"lateral": _draw_lat,
        }
        for key in _CARD_KEYS:
            if key == active:
                continue
            drawers[key](self._section_layer_alpha(key, active, settings))
        drawers[active](self._section_layer_alpha(active, active, settings))

        try:
            ln = WpfLine()
            ln.X1 = left
            ln.Y1 = top + rh + 10
            ln.X2 = left + rw
            ln.Y2 = top + rh + 10
            ln.Stroke = _brush(u"#64748b")
            ln.StrokeThickness = 1.0
            cv.Children.Add(ln)
        except Exception:
            pass

        try:
            if self._ui_txt_section_dims is not None:
                meta = _CARD_UI.get(active) or {}
                self._ui_txt_section_dims.Text = (
                    u"Luz menor {1:.0f} mm · Alto {2:.0f} mm · editando {0}"
                ).format(meta.get(u"title") or active, w_mm, h_mm)
        except Exception:
            pass

    def _redraw_all(self):
        self._redraw_plan()
        self._redraw_section()
        self._update_status_from_settings()

    def _on_plan_wheel(self, sender, args):
        try:
            delta = int(args.Delta)
        except Exception:
            return
        factor = 1.12 if delta > 0 else (1.0 / 1.12)
        nz = float(self._view_zoom) * factor
        self._view_zoom = max(0.2, min(8.0, nz))
        self._redraw_plan()
        try:
            args.Handled = True
        except Exception:
            pass

    def _begin_pan(self, args):
        self._panning = True
        try:
            self._pan_last = args.GetPosition(self._ui_cv_plan)
        except Exception:
            self._pan_last = None
        try:
            self._ui_cv_plan.CaptureMouse()
        except Exception:
            pass
        try:
            args.Handled = True
        except Exception:
            pass

    def _on_plan_down(self, sender, args):
        try:
            self._begin_pan(args)
        except Exception:
            pass

    def _on_plan_move(self, sender, args):
        if not self._panning or self._scene_base is None:
            return
        try:
            pos = args.GetPosition(self._ui_cv_plan)
            last = self._pan_last
            if last is None:
                self._pan_last = pos
                return
            dx_px = float(pos.X) - float(last.X)
            dy_px = float(pos.Y) - float(last.Y)
            scale = float(self._scene_base.get(u"scale") or 1.0)
            if scale < 1e-9:
                return
            self._view_pan_x -= dx_px / scale
            self._view_pan_y += dy_px / scale
            self._pan_last = pos
            self._redraw_plan()
            try:
                args.Handled = True
            except Exception:
                pass
        except Exception:
            pass

    def _on_plan_up(self, sender, args):
        if not self._panning:
            return
        self._panning = False
        self._pan_last = None
        try:
            if self._ui_cv_plan is not None:
                self._ui_cv_plan.ReleaseMouseCapture()
        except Exception:
            pass

    def _on_size_changed(self, sender=None, args=None):
        if not self._ui_revealed:
            return
        self._redraw_all()

    def _wire_events(self):
        win = self._win
        if win is None:
            return
        try:
            btn_man = win.FindName(u"BtnManual")
            if btn_man is not None:
                btn_man.Click += RoutedEventHandler(self._open_manual)
        except Exception:
            pass
        try:
            btn_c = win.FindName(u"BtnCancelar")
            if btn_c is not None:
                btn_c.Click += RoutedEventHandler(self._on_cancel)
        except Exception:
            pass
        try:
            btn_p = win.FindName(u"BtnColocar")
            if btn_p is not None:
                btn_p.Click += RoutedEventHandler(self._on_colocar)
        except Exception:
            pass
        try:
            win.Closed += EventHandler(self._on_closed)
        except Exception:
            pass
        try:
            from System.Windows import SizeChangedEventHandler

            win.SizeChanged += SizeChangedEventHandler(self._on_size_changed)
        except Exception:
            try:
                win.SizeChanged += EventHandler(self._on_size_changed)
            except Exception:
                pass

        cv = self._ui_cv_plan
        if cv is not None:
            try:
                cv.MouseWheel += MouseWheelEventHandler(self._on_plan_wheel)
            except Exception:
                pass
            try:
                cv.MouseLeftButtonDown += MouseButtonEventHandler(self._on_plan_down)
                cv.MouseLeftButtonUp += MouseButtonEventHandler(self._on_plan_up)
                cv.MouseMove += MouseEventHandler(self._on_plan_move)
            except Exception:
                pass
            try:
                from System.Windows.Input import MouseButton

                def _md(s, e):
                    try:
                        if e.ChangedButton == MouseButton.Middle:
                            self._begin_pan(e)
                    except Exception:
                        pass

                def _mu(s, e):
                    try:
                        if e.ChangedButton == MouseButton.Middle:
                            self._on_plan_up(s, e)
                    except Exception:
                        pass

                cv.MouseDown += MouseButtonEventHandler(_md)
                cv.MouseUp += MouseButtonEventHandler(_mu)
            except Exception:
                pass

        self._wire_cards()

    def show(self):
        self._win = XamlReader.Parse(_XAML)
        win = self._win
        self._cache_ui_refs()
        self._wire_events()
        self._init_cards()

        try:
            if self._ui_txt_host is not None:
                self._ui_txt_host.Text = _element_label(self._foundation)
        except Exception:
            pass

        hwnd = None
        try:
            hwnd = revit_main_hwnd(self._uiapp)
        except Exception:
            pass
        try:
            from System.Windows.Interop import WindowInteropHelper

            if hwnd:
                helper = WindowInteropHelper(win)
                helper.Owner = hwnd
        except Exception:
            pass
        try:
            position_wpf_window_top_left_at_active_view(win, self._uidoc, hwnd)
        except Exception:
            pass

        _register_singleton(win, self)
        try:
            win.Show()
        except Exception:
            _unregister_singleton()
            raise

        self._ui_revealed = True
        try:
            win.UpdateLayout()
        except Exception:
            pass
        self._redraw_all()
        try:
            win.Activate()
            win.Focus()
            if self._ui_cv_plan is not None:
                self._ui_cv_plan.Focus()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------


def run(revit):
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
            content=u"Abra una planta (ViewPlan) o una planta de cimentación "
            u"y vuelva a ejecutar.",
        )
        return

    if _focus_existing(uiapp):
        return

    foundation = _pick_foundation(uidoc, uiapp)
    if foundation is None:
        return

    try:
        from geometria_fundacion_cara_inferior import clear_face_cache

        clear_face_cache()
    except Exception:
        pass

    plan_polys = extract_foundation_plan_polygons_mm(foundation)
    if not plan_polys:
        _mostrar_aviso(
            uiapp,
            u"No se pudo obtener el contorno de la fundación.",
            content=u"Se requiere geometría de cara inferior o BoundingBox usable.",
        )
        return

    sec_w, sec_h = extract_section_dims_mm(foundation, plan_polys)

    try:
        ctrl = ArmadoFundacionAisladaSketchController(
            uiapp,
            uidoc,
            doc,
            foundation,
            plan_polys,
            sec_w,
            sec_h,
            source_view=active_view,
        )
        ctrl.show()
    except Exception as ex:
        _unregister_singleton()
        _mostrar_aviso(
            uiapp,
            u"Error al abrir la UI.",
            content=_as_unicode(ex),
        )


def run_pyrevit(revit):
    run(revit)
